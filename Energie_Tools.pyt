import arcpy, os
import xlwt, xlrd, csv
import numpy as np
from os.path import basename, dirname, exists, join
from arcpy import env,mapping,Delete_management,CreateFileGDB_management, \
     CreateFeatureDataset_management,\
     DefineProjection_management,FeatureClassToFeatureClass_conversion,\
     CheckInExtension, CheckOutExtension,Describe, GetInstallInfo,SpatialReference
from arcpy.conversion import FeatureClassToFeatureClass
from arcpy.management import CreateFeatureDataset, CreateFileGDB, DefineProjection

from comtypes.client import CreateObject, GetModule

ARCOBJECTS_DIR = join(GetInstallInfo()['InstallDir'], 'com')
GetModule(join(ARCOBJECTS_DIR, 'esriDataSourcesFile.olb'))
GetModule(join(ARCOBJECTS_DIR, 'esriGeoDatabase.olb'))
GetModule(join(ARCOBJECTS_DIR, 'esriSystem.olb'))

from comtypes.gen.esriDataSourcesFile import ShapefileWorkspaceFactory
from comtypes.gen.esriDataSourcesGDB import FileGDBWorkspaceFactory
from comtypes.gen.esriGeoDatabase import esriDatasetType, \
    esriNetworkAttributeDataType, esriNetworkAttributeUnits, \
    esriNetworkAttributeUsageType, esriNetworkDatasetType, \
    esriNetworkEdgeConnectivityPolicy, esriNetworkEdgeDirection, \
    esriNetworkElementType, DENetworkDataset, EdgeFeatureSource, \
    EvaluatedNetworkAttribute, IDataElement, IDataset, IDatasetContainer3, \
    IDEGeoDataset, IDENetworkDataset2, IEdgeFeatureSource, IEnumDataset, \
    IEvaluatedNetworkAttribute, IFeatureDatasetExtension, \
    IFeatureDatasetExtensionContainer, IFeatureWorkspace, IGeoDataset, \
    INetworkAttribute2, INetworkBuild, INetworkDataset, \
    INetworkConstantEvaluator, INetworkEvaluator, INetworkFieldEvaluator2, \
    INetworkSource, IWorkspace, IWorkspaceFactory, IWorkspaceExtension3, \
    IWorkspaceExtensionManager, NetworkConstantEvaluator, \
    NetworkFieldEvaluator
from comtypes.gen.esriSystem import IArray, IUID, UID
ARRAY_GUID = "{8F2B6061-AB00-11D2-87F4-0000F8751720}"


def create_gdb_for_network(path, GDB_PATH, FDS_PATH, GDB_NAME, FDS_NAME, FC_NAME, shp_name):
     
    spatial_ref = Describe(shp_name).spatialReference
    #print("{0} : {1}".format(shp_name, spatial_ref.name))    
    if exists(GDB_PATH):
        Delete_management(GDB_PATH)
        #print ("Database deleted")
    CreateFileGDB_management(path, GDB_NAME)
    CreateFeatureDataset_management(GDB_PATH , FDS_NAME)
    DefineProjection_management(FDS_PATH, spatial_ref)
    FeatureClassToFeatureClass_conversion(shp_name, FDS_PATH, FC_NAME)
    #print ('Database created')

def create_gdb_network_dataset(path, GDB_PATH, FDS_PATH, GDB_NAME, FDS_NAME, FC_NAME, shp_name, ND_NAME, cost_field):
    
    # create an empty data element for a buildable network dataset
    net = new_obj(DENetworkDataset, IDENetworkDataset2)
    net.Buildable = True
    net.NetworkType = esriNetworkDatasetType(1)

    # open the feature class and ctype to the IGeoDataset interface
    gdb_ws_factory = new_obj(FileGDBWorkspaceFactory, IWorkspaceFactory)
    gdb_workspace = ctype(gdb_ws_factory.OpenFromFile(GDB_PATH, 0),IFeatureWorkspace)
    gdb_feat_ds = ctype(gdb_workspace.OpenFeatureDataset(FDS_NAME),IGeoDataset)

    # copy the feature dataset's extent and spatial reference to the network dataset data element
    net_geo_elem = ctype(net, IDEGeoDataset)
    net_geo_elem.Extent = gdb_feat_ds.Extent
    net_geo_elem.SpatialReference = gdb_feat_ds.SpatialReference

    # specify the name of the network dataset
    net_element = ctype(net, IDataElement)
    net_element.Name = ND_NAME

    edge_net = new_obj(EdgeFeatureSource, INetworkSource)
    edge_net.Name = FC_NAME
    edge_net.ElementType = esriNetworkElementType(2)
    
    # set the edge feature source's connectivity settings, connect network through any coincident vertex
    edge_feat = ctype(edge_net, IEdgeFeatureSource)
    edge_feat.ClassConnectivityGroup = 1
    edge_feat.ClassConnectivityPolicy = esriNetworkEdgeConnectivityPolicy(0)
    edge_feat.UsesSubtypes = False

    source_array = new_obj(ARRAY_GUID, IArray)
    source_array.Add(edge_net)
    net.Sources = source_array

    add_network_attributes(net, edge_net, cost_field)

    # get the feature dataset extension and create the network dataset based on the data element.
    data_xc = ctype(gdb_feat_ds, IFeatureDatasetExtensionContainer)
    data_ext = ctype(data_xc.FindExtension(esriDatasetType(19)),IFeatureDatasetExtension)
    data_cont = ctype(data_ext, IDatasetContainer3)
    net_dataset = ctype(data_cont.CreateDataset(net), INetworkDataset)

    #  build network
    net_build = ctype(net_dataset, INetworkBuild)
    net_build.BuildNetwork(net_geo_elem.Extent)
    
def new_obj(class_, interface):
    pointer = CreateObject(class_, interface=interface)
    return pointer

def ctype(obj, interface):
    pointer = obj.QueryInterface(interface)
    return pointer

def add_network_attributes(net, edge_net, cost_field):
    
    attribute_array = new_obj(ARRAY_GUID, IArray)
    language = 'VBScript'
   
    # 1) cost attribute
    len_eval_attr = new_obj(EvaluatedNetworkAttribute,IEvaluatedNetworkAttribute)
    len_attr = ctype(len_eval_attr, INetworkAttribute2)
    len_attr.DataType = esriNetworkAttributeDataType(2)  # double
    len_attr.Name = 'Length'
    len_attr.Units = esriNetworkAttributeUnits(9)  # meters
    len_attr.UsageType = esriNetworkAttributeUsageType(0)  # cost
    len_attr.UseByDefault = True

    len_expr = cost_field
    set_evaluator_logic(len_eval_attr, edge_net, len_expr, '', language)
    set_evaluator_constants(len_eval_attr, 0)
    attribute_array.Add(len_eval_attr)
   
    net.Attributes = attribute_array


def set_evaluator_logic(eval_attr, edge_net, expression, pre_logic, lang):
    """This function implements the same logic for an attribute in
    both the along and against directions
    """

    # for esriNetworkEdgeDirection 1=along, 2=against
    for direction in (1, 2):
        net_eval = new_obj(NetworkFieldEvaluator, INetworkFieldEvaluator2)
        net_eval.SetLanguage(lang)
        net_eval.SetExpression(expression, pre_logic)
        eval_attr.Evaluator.setter(
            eval_attr, edge_net, esriNetworkEdgeDirection(direction),
            ctype(net_eval, INetworkEvaluator))


def set_evaluator_constants(eval_attr, constant):
    #This function sets all evaluator constants to the same value

    # for ConstantValue False means traversable (that is, not restricted),
    # for esriNetworkElementType 1=junction, 2=edge, 3=turn
    for element_type in (1, 2, 3):
        const_eval = new_obj(NetworkConstantEvaluator, INetworkConstantEvaluator)
        const_eval.ConstantValue = constant
        eval_attr.DefaultEvaluator.setter(
            eval_attr, esriNetworkElementType(element_type),
            ctype(const_eval, INetworkEvaluator))

def GetRoutes(network, analysis_layer, impedance_attribute, num_facilities, facilities, incidents, outRoutesFC):
    
    NAResultObject = arcpy.na.MakeClosestFacilityLayer(network,analysis_layer,
                                                   impedance_attribute,"TRAVEL_TO",
                                                   "", num_facilities, "", "ALLOW_UTURNS")

    #Get the layer object from the result object. The closest facility layer can
    #now be referenced using the layer object.
    outNALayer = NAResultObject.getOutput(0)

    #Get the names of all the sublayers within the closest facility layer.
    subLayerNames = arcpy.na.GetNAClassNames(outNALayer)

    #Stores the layer names that will be used
    facilitiesLayerName = subLayerNames["Facilities"]
    incidentsLayerName = subLayerNames["Incidents"]
    routesLayerName = subLayerNames["CFRoutes"]

    arcpy.na.AddLocations(outNALayer, facilitiesLayerName, facilities, "", "")
    arcpy.na.AddLocations(outNALayer, incidentsLayerName, incidents, "", "")

    arcpy.na.Solve(outNALayer)
    
    if exists(outRoutesFC):    
            arcpy.Delete_management(outRoutesFC)
    RoutesSubLayer = arcpy.mapping.ListLayers(outNALayer, routesLayerName)[0]
    arcpy.management.CopyFeatures(RoutesSubLayer, outRoutesFC)

def SpatialJoin(input_feature, join_feature, output_feature, merge_rule, keep, connectivity_rule, attributes):
    fieldmappings = arcpy.FieldMappings()
    fieldmappings.addTable(input_feature)
    fieldmappings.addTable(join_feature)
    fnames = arcpy.ListFields(join_feature)
    join_type = keep
    for f in fnames:
        if f.name in attributes: 
            fidx = fieldmappings.findFieldMapIndex(f.name)
            #print f.name
            fmap = fieldmappings.getFieldMap(fidx)
            fmap.mergeRule = merge_rule
            fieldmappings.replaceFieldMap(fidx,fmap)       

    arcpy.SpatialJoin_analysis(input_feature, join_feature, output_feature, "JOIN_ONE_TO_ONE", join_type, fieldmappings, connectivity_rule)

def Delete(feature):
    if exists(feature):
            arcpy.Delete_management(feature)

def LoadCurve(efhconsumption, mfhconsumption, ghconsumption, gkoconsumption, gmkconsumption, losses, path, env): #could be changed to one list
        
    #read .csv-files with the hourly percentages into lists. Could be optimized by iterating a list
    daytables = []    
    
    daytables.append([]) #4-dimensionale daytables[category][weekday][temperature][hour]
    weekday = 1
    while weekday <= 7:
        daytables[0].append([])
        with open(os.path.join(env, 'lcdata/efh' + str(weekday) + '.csv'))as csvfile:  #open file for each weekday
            csvdata = csv.reader(csvfile, delimiter=';')
            for row in csvdata:     #one row for each temperature zone
                row = list(map(float, row)) #convert data into floats
                daytables[0][weekday - 1].append(row) #
        weekday = weekday + 1
        csvfile.close()  
    
    daytables.append([])
    weekday = 1
    while weekday <= 7:
        daytables[1].append([])
        with open(os.path.join(env, 'lcdata/mfh' + str(weekday) + '.csv')) as csvfile:
            csvdata = csv.reader(csvfile, delimiter=';')
            for row in csvdata:
                row = list(map(float, row))
                daytables[1][weekday - 1].append(row)
        weekday = weekday + 1
        csvfile.close()
    
    daytables.append([])
    weekday = 1
    while weekday <= 7:
        daytables[2].append([])
        with open(os.path.join(env, 'lcdata/gh' + str(weekday) + '.csv')) as csvfile:
            csvdata = csv.reader(csvfile, delimiter=';')
            for row in csvdata:
                row = list(map(float, row))
                daytables[2][weekday - 1].append(row)
        weekday = weekday + 1
        csvfile.close()
    
    daytables.append([])
    weekday = 1
    while weekday <= 7:
        daytables[3].append([])
        with open(os.path.join(env, 'lcdata/gko' + str(weekday) + '.csv')) as csvfile:
            csvdata = csv.reader(csvfile, delimiter=';')
            for row in csvdata:
                row = list(map(float, row))
                daytables[3][weekday - 1].append(row)
        weekday = weekday + 1
        csvfile.close()
    
    daytables.append([])
    weekday = 1
    while weekday <= 7:
        daytables[4].append([])
        with open(os.path.join(env, 'lcdata/gmk' + str(weekday) + '.csv')) as csvfile:
            csvdata = csv.reader(csvfile, delimiter=';')
            for row in csvdata:
                row = list(map(float, row))
                daytables[4][weekday - 1].append(row)
        weekday = weekday + 1
        csvfile.close()   

    
    #calculation of the daily consumption

    
    #oder of the categories: EFH (0), MFH (1), GKO (2), GH (3), GMK (4)
    #order of the weekdays: Mo (0), Tu (1), We (2), Th (3), Fr (4), Sa (5), Su (6)
    #order of the sigmoid-parameters: A (0), B (1), C (2), D (3)
    
    #get sigmoid-parameters into lists
    with open(os.path.join(env, 'lcdata/factors.csv')) as csvfile:
        factors = []
        csvdata = csv.reader(csvfile, delimiter=';')
        for row in csvdata:
            row = list(map(float, row))
            factors.append(row) #each row contains data for one of the categories
        csvfile.close()
    
    #get weekday-faktors into lists
    with open(os.path.join(env, 'lcdata/wdfactors.csv')) as csvfile:
        wdfactors = []
        csvdata = csv.reader(csvfile, delimiter=';')
        for row in csvdata:
            row = list(map(float, row))
            wdfactors.append(row) #each row contains data for one of the weekdays
        csvfile.close()
            
    #get daily mean temperature into lists
    with open(os.path.join(env, 'lcdata/temp.csv')) as csvfile:
        temperatures = [] 
        csvdata = csv.reader(csvfile, delimiter=';')
        for row in csvdata:
            temperatures.append(float(row[2])) 
        csvfile.close()
            
            
    #save consumption into list
    categoryconsumption = (efhconsumption, mfhconsumption, gkoconsumption, ghconsumption, gmkconsumption)


    #save weekdays in list with 356 entries (0 to 6, mo to su)
    i = 1
    k = 0
    weekday = []
    while i <= 365:
        weekday.append(k)
        k += 1
        if k == 7:
            k = 0
        i += 1
    
            
    category = 0
    hdays = []          #result of the sigmoid-function
    hdaysum = []        #sum of the results
    loadcurves = []     #list contains load for each day in the year
    for consumption in categoryconsumption: #do it for all categories
        hdays.append([])
        hdaysum.append([])
        warmwaterheatfactor = 1 #number of houses with warm water supply / number of houses without warm water supply
        factorA = factors[category][0]
        factorB = factors[category][1]
        factorC = factors[category][2]
        factorDv = factors[category][3] * warmwaterheatfactor
        tetaA0 = 40
        i = 0   #counting variable for day set to 0
        
        for temperature in temperatures: #For each day in the year use the temperature for calculating the sigmoid-function
            factorweekday = wdfactors[category][weekday[i]] #get weekday factor from list
            tetaA = temperature
            hdays[category].append(factorweekday * (factorA / (1 + (factorB / (tetaA - tetaA0)) ** factorC) + factorDv)) #this ist the sigmoid-function
            i += 1 #next day
        
        #build sum of all results
        hdaysum[category] = 0
        for hday in hdays[category]:
            hdaysum[category] = hdaysum[category] + hday
            
        category = category + 1 #next category
            
    #divide the yearly consumption to the days by using the sigmoid-function
    category = 0        
    for consumption in categoryconsumption: #do it for all categories 
        loadcurves.append([])
        normalizedconsumption = consumption / hdaysum[category]
        for hday in hdays[category]:
            loadcurves[category].append(hday * normalizedconsumption)
        category = category + 1   

    #identify temperature zone    
    k = 0
    finalloadcurves = []
    for daytable in daytables:
        finalloadcurve = []
        for dayload in loadcurves[k]:
            hourload = []
            hour = 0
            
            if temperatures[k] < -15:
                temperaturezone = 0
            elif temperatures[k] < -10 and temperatures[k] >= -15:
                temperaturezone = 1
            elif temperatures[k] < -5 and temperatures[k] >= -10:
                temperaturezone = 2
            elif temperatures[k] < 0 and temperatures[k] >= -5:
                temperaturezone = 3
            elif temperatures[k] < 5 and temperatures[k] >= 0:
                temperaturezone = 4
            elif temperatures[k] < 10 and temperatures[k] >= 5:
                temperaturezone = 5
            elif temperatures[k] < 15 and temperatures[k] >= 10:
                temperaturezone = 6
            elif temperatures[k] < 20 and temperatures[k] >= 15:
                temperaturezone = 7
            elif temperatures[k] < 25 and temperatures[k] >= 20:
                temperaturezone = 8
            else:
                temperaturezone = 9

            
            while hour < 24:
                hourload.append(dayload * daytable[weekday[k]][temperaturezone][hour])
                hour += 1  
            finalloadcurve = finalloadcurve + hourload
        k += 1
        
        finalloadcurves.append(finalloadcurve)

    lossescurve = []
    for element in finalloadcurves[0]:
        lossescurve.append(losses/len(finalloadcurves[0]))
    
    finalloadcurvesum = [sum(x) for x in zip(finalloadcurves[0], finalloadcurves[1], finalloadcurves[2], finalloadcurves[3], finalloadcurves[4], lossescurve)] #not very good code, should be iterated

    #print('Summe Lastgang:', sum(finalloadcurvesum), 'kWh')
    #print('Fehler:', round(((efhconsumption + mfhconsumption + ghconsumption + gkoconsumption + gmkconsumption + losses - sum(finalloadcurvesum)) / (efhconsumption + mfhconsumption + ghconsumption + gkoconsumption + gmkconsumption + losses)) * 100, 6), "%")
    #print('Rechenzeit:', round((t2-t1)*1000), 'Millisekunden (VBA 45000 ms)')    
    
    
    finalloadcurvesumtransposed = []   
    i=1
    for element in finalloadcurvesum:
        finalloadcurvesumtransposed.append((i, element))
        i += 1
        
    #print(finalloadcurvesumtransposed)        
        
    with open(os.path.join(env, 'loadcurve.csv'), 'w') as csvfile:
        writer = csv.writer(csvfile, delimiter=',', lineterminator='\n')
        writer.writerow(('hour', 'demand'))
        for element in finalloadcurvesumtransposed:
            writer.writerow(element)
        csvfile.close()




                
           
class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Grid_Estimation"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
           
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Grid_Estimation"
        self.description = ""
        self.canRunInBackground = False
        
    def getParameterInfo(self):
        param0 = arcpy.Parameter(
            displayName="Input Building Polygons",
            name="input0",
            datatype="Feature Layer",
            parameterType="Required",
            direction="Input")
	param0.filter.list = ["Polygon"]
        
        param1 = arcpy.Parameter(
            displayName = "Input Road Polylines",
            name = "input1",
            datatype = "Feature Layer",
            parameterType = "Required",
            direction= "Input")
        param1.filter.list = ["Polyline"]

	param2 = arcpy.Parameter(
            displayName = "Input Source Feature",
            name = "input2",
            datatype = "Feature Layer",
            parameterType = "Required",
            direction= "Input")
        param2.filter.list = ["Polygon"]

        param3 = arcpy.Parameter(
            displayName = "Vorlauftemperatur (°C)",
            name = "input3",
            datatype = "GPLong",
            parameterType = "Required",
            direction= "Input")

        param4 = arcpy.Parameter(
            displayName = "Rücklauftemperatur (°C)",
            name = "input4",
            datatype = "GPLong",
            parameterType = "Required",
            direction= "Input")

	param5 = arcpy.Parameter(
            displayName = "Output Line Feature",
            name = "out_features",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction= "Output")
        param5.filter.list = ["Polyline"]
        
        params = [param0 , param1, param2, param3, param4, param5]

        return params


    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        
        buildings = parameters[0].valueAsText
        streets = parameters[1].valueAsText
        source = parameters[2].valueAsText
        htemp = parameters[3].valueAsText
        ltemp = parameters[4].valueAsText     
	Net =  parameters[5].valueAsText
	#workbook = parameters[6].valueAsText
	arcpy.env.workspace =  arcpy.Describe(buildings).path
	workbook = os.path.join(arcpy.env.workspace, "dashboard.xls")
        
        Delete("building.shp")
	Delete("centroid.shp")
	Delete("centroidToRd.shp")
	Delete("ConnectingLines.shp")
	Delete("target.shp")
	Delete("b_centroid.shp")
	Delete("sc_centroid.shp")
	Delete("junctions1.shp")
	Delete("junctions2.shp")
	Delete("junctions.shp")
	Delete("segments.shp")
	
	
	arcpy.MakeFeatureLayer_management(buildings, "building")
        arcpy.SelectLayerByAttribute_management("building",  "NEW_SELECTION", '"WVBR" <> 0' )
        building = os.path.join(arcpy.env.workspace, 'building.shp')
        arcpy.CopyFeatures_management("building", building)

        segments = os.path.join(arcpy.env.workspace, "segments.shp")

         
	### ------------ Draw connecting lines ----------------------------###
	
	
        arcpy.AddMessage("Drawing Connecting lines...")
	"""
	# Draw lines connecting building centroids and the road network
        
	fc_centroid = os.path.join(arcpy.env.workspace, "centroid.shp")  # building centroids
        fc_connectingLines = os.path.join(arcpy.env.workspace, "connectingLines.shp")
        target = os.path.join(arcpy.env.workspace, "target.shp") # save copy of centroid without source for use as NA facilties
	arcpy.FeatureToPoint_management(building, "b_centroid.shp", "CENTROID")
        arcpy.CopyFeatures_management("b_centroid.shp", target)

        arcpy.FeatureToPoint_management(source, "sc_centroid.shp", "CENTROID")
	arcpy.Merge_management(["sc_centroid.shp", "b_centroid.shp"], fc_centroid) # add also the source point to connect line to it 

	arcpy.AddXY_management(fc_centroid) # Adding POINT_X and POINT_Y
	arcpy.Near_analysis(fc_centroid, streets, "", "LOCATION", "NO_ANGLE" ) # Adding NEAR_X and NEAR_Y
	arcpy.XYToLine_management(fc_centroid, fc_connectingLines, "POINT_X", "POINT_Y", "NEAR_X", "NEAR_Y")
        """ 
        

        # remove lines that intersect with more than one/two buildings 
       
        arcpy.FeatureToPoint_management(source, "sc_centroid.shp", "CENTROID")
        
        target = os.path.join(arcpy.env.workspace, "b_centroid.shp")
        arcpy.FeatureToPoint_management(building, target, "CENTROID")
        
        fc_centroid = os.path.join(arcpy.env.workspace, "centroid.shp")
        arcpy.Merge_management(["sc_centroid.shp", "b_centroid.shp"], fc_centroid)
        
        def RemoveLine(fc, cnt):
            arcpy.AddField_management(fc, "count", "INTEGER" )

            with arcpy.da.UpdateCursor(fc, ['count', 'SHAPE@']) as cursor:
                 for row in cursor:
                      
                      arcpy.Delete_management("layer")
                      arcpy.MakeFeatureLayer_management(building, "layer")
                      arcpy.SelectLayerByLocation_management("layer", 'INTERSECT', row[1])
                      result = arcpy.GetCount_management("layer")
                      row[0] = int(result.getOutput(0))
                      
                      cursor.updateRow(row)
                      
            arcpy.Delete_management("tempLayer")        
            if cnt == 1:    
                arcpy.MakeFeatureLayer_management(fc, "tempLayer")
                arcpy.SelectLayerByAttribute_management("tempLayer",  "NEW_SELECTION", '"count" > 1')
                arcpy.DeleteFeatures_management("tempLayer")
              
            else:    
                arcpy.MakeFeatureLayer_management(fc, "tempLayer")
                arcpy.SelectLayerByAttribute_management("tempLayer",  "NEW_SELECTION", '"count" > 2')
                arcpy.DeleteFeatures_management("tempLayer")

                
        centroidToRd = os.path.join(arcpy.env.workspace, "centroidToRd.shp")
        arcpy.AddXY_management(fc_centroid) # Adding POINT_X and POINT_Y
        arcpy.Near_analysis(fc_centroid, streets, "", "LOCATION", "NO_ANGLE" ) # Adding NEAR_X and NEAR_Y
        arcpy.XYToLine_management(fc_centroid, centroidToRd, "POINT_X", "POINT_Y", "NEAR_X", "NEAR_Y")

        RemoveLine(centroidToRd, 1)

        fc_connectingLines = os.path.join(arcpy.env.workspace, "ConnectingLines.shp")
        arcpy.CopyFeatures_management(centroidToRd, fc_connectingLines)


        cnt = 1
        itr = 0
        arr = [centroidToRd]

        #iterate through centroid points until each one is connected to closest centroid or directly to road

        while (cnt > 0  and itr < 5):

            itr +=1
            
            unique_name = arcpy.CreateUniqueName("feeder.shp")
            arcpy.Delete_management("connected")  
            arcpy.MakeFeatureLayer_management(fc_centroid, "connected")
            arcpy.SelectLayerByLocation_management("connected", 'INTERSECT', fc_connectingLines, invert_spatial_relationship = "NOT_INVERT")

            arcpy.Delete_management("disconnected") 
            arcpy.MakeFeatureLayer_management(fc_centroid, "disconnected")
            arcpy.SelectLayerByLocation_management("disconnected", 'INTERSECT', fc_connectingLines, invert_spatial_relationship = "INVERT")
            cnt = arcpy.GetCount_management("disconnected")
            #print cnt

            
            arcpy.AddXY_management("disconnected") # Adding POINT_X and POINT_Y
            arcpy.Near_analysis("disconnected", "connected", "", "LOCATION", "NO_ANGLE" ) # Adding NEAR_X and NEAR_Y
            arcpy.XYToLine_management("disconnected", unique_name, "POINT_X", "POINT_Y", "NEAR_X", "NEAR_Y")

            
            RemoveLine(unique_name, 2)
            
            fcnt = arcpy.GetCount_management(unique_name)
            if int(fcnt.getOutput(0)) > 0: 
                arr.append(unique_name)
                arcpy.Delete_management("ConnectingLines.shp")
                fc_connectingLines = os.path.join(arcpy.env.workspace, "ConnectingLines.shp")
                arcpy.Merge_management(arr, fc_connectingLines)
            else:
                arcpy.Delete_management(unique_name)
         
	# snap the lines to street network to avoid topology errors
	arcpy.Snap_edit(fc_connectingLines, [[streets, "VERTEX", "0.5 Meters"], [streets, "EDGE", "0.5 Meters"]])

        # create road segments
        
        segments = os.path.join(arcpy.env.workspace, "segments.shp")	
        arcpy.Intersect_analysis([streets, fc_connectingLines], "junctions1.shp", "ALL", "", "POINT")
        arcpy.Intersect_analysis(streets, "junctions2.shp", "ALL", "", "POINT")
        arcpy.Merge_management(["junctions1.shp", "junctions2.shp"], "junctions.shp")
        arcpy.SplitLineAtPoint_management(streets, "junctions.shp" , segments, "1 Meters")
        arcpy.AddField_management(segments, "Shape_Leng", "DOUBLE")
        arcpy.CalculateField_management(segments, "Shape_Leng", "!shape.length@meters!","PYTHON_9.3")

        # merge connectingLines and segments to create network dataset
        Delete("merged_lines1.shp")
	arcpy.Merge_management([segments, fc_connectingLines], "merged_lines1.shp")
	arcpy.CalculateField_management("merged_lines1.shp", "Shape_Leng", "!shape.length@meters!","PYTHON_9.3")
        arcpy.CheckOutExtension('Network')


	arcpy.AddMessage("Running Network Analysis I...")
        ### ------------------------------------- Network Analysis I (Length)----------------------------------###

        Delete("Routes1.shp")
        Delete("Routes1_Attr.shp")
        Delete("dissolved1st.shp")
        Delete("grid1.shp")
        
        GDB_NAME = "network_GDB.gdb"
        FDS_NAME = "network_FDS"
        GDB_PATH = os.path.join(arcpy.env.workspace, GDB_NAME)
        FDS_PATH = os.path.join(GDB_PATH, FDS_NAME)
        ND_NAME = 'energie_net'
        FC_NAME = 'net_edges'
        shp_name = os.path.join(arcpy.env.workspace, "merged_lines1.shp")
        
        create_gdb_for_network(arcpy.env.workspace, GDB_PATH, FDS_PATH, GDB_NAME, FDS_NAME, FC_NAME, shp_name)
        create_gdb_network_dataset(arcpy.env.workspace, GDB_PATH, FDS_PATH, GDB_NAME, FDS_NAME, FC_NAME, shp_name, ND_NAME, '[Shape]')
  
        network = os.path.join(FDS_PATH, ND_NAME)
        count = arcpy.GetCount_management(building)
        
        Routes = os.path.join(arcpy.env.workspace, "Routes1.shp")
        GetRoutes(network, "ClosestFacility1", "Length" , count, target,  "sc_centroid.shp",Routes)
        
        # spatial join routes with target points to get attributes
        attributes = ["EBF", "WVBRpEBF", "WVBR", "leistungkW", "Bcount"]
        
        Routes_Attr = os.path.join(arcpy.env.workspace, "Routes1_Attr.shp") 
        SpatialJoin(Routes, target, Routes_Attr,  "first", "KEEP_ALL","CLOSEST", attributes)

        arcpy.Dissolve_management(Routes_Attr, "dissolved1st.shp")
        arcpy.AddField_management("dissolved1st.shp", "Length", "DOUBLE" )
        arcpy.CalculateField_management("dissolved1st.shp", "Length", "!shape.length@meters!","PYTHON_9.3")
        
        with arcpy.da.SearchCursor("dissolved1st.shp", ['Length']) as cursor:
            for row in cursor:
                length_old = row[0]
        
        # spatial join segments and routes_attr to obtain cumulative impedance
      
        grid = os.path.join(arcpy.env.workspace, "grid1.shp") 
        SpatialJoin(segments, Routes_Attr , grid,  "sum", "KEEP_ALL", "COMPLETELY_WITHIN", attributes)

        arcpy.AddMessage("Running Network Analysis II...")
            
        ### ------------------------------------- Network Analysis II (kW) ----------------------------------###
            
        length_new = 0
        iterator = 0
        
        while (length_old > length_new and iterator < 6 ):
            iterator = iterator + 1
            if (iterator > 1):
                    length_old = length_new
                
            Delete("merged_lines2.shp")
                
            GDB_NAME = "network_GDB.gdb"
            FDS_NAME = "network_FDS"
            GDB_PATH = os.path.join(arcpy.env.workspace, GDB_NAME)
            FDS_PATH = os.path.join(GDB_PATH, FDS_NAME)
            ND_NAME = 'energie_net'
            FC_NAME = 'net_edges'
            shp_name = os.path.join(arcpy.env.workspace, "merged_lines2.shp")
                
            #merge and add a field which can used as cost
            expression = "f(!leistungkW!, !Shape_Leng!)"
            codeblock = """def f(x,y):
                            if x == 0:
                                z = y 
                            else:
                                z = y / x
                            return z"""

                              
            arcpy.Merge_management([grid, fc_connectingLines], shp_name)
            
            with arcpy.da.UpdateCursor(shp_name, ['Shape_Leng', 'leistungkW']) as cursor:
                for row in cursor:
                    if row[0] > 50 and row[1] < 50:
                        cursor.deleteRow()
                       
            arcpy.CalculateField_management(shp_name, "Shape_Leng", "!shape.length@meters!","PYTHON_9.3")
            arcpy.AddField_management(shp_name, "Length_kW", "DOUBLE" )
            arcpy.CalculateField_management(shp_name, "Length_kW", expression, "PYTHON_9.3", codeblock)
            create_gdb_for_network(arcpy.env.workspace, GDB_PATH, FDS_PATH, GDB_NAME, FDS_NAME, FC_NAME, shp_name)
            create_gdb_network_dataset(arcpy.env.workspace, GDB_PATH, FDS_PATH, GDB_NAME, FDS_NAME, FC_NAME, shp_name, ND_NAME, '[Length_kW]')
                
            network = os.path.join(FDS_PATH, ND_NAME)
            count = arcpy.GetCount_management(building)
            Delete("Routes2.shp")
            Routes = os.path.join(arcpy.env.workspace, "Routes2.shp")        
            GetRoutes(network,"ClosestFacility2", "Length" , count,  target, "sc_centroid.shp", Routes)
                
            # spatial join routes with target points to get attributes
            Delete("Routes2_Attr.shp")
            Routes_Attr = os.path.join(arcpy.env.workspace, "Routes2_Attr.shp") 
            SpatialJoin(Routes, target, Routes_Attr,  "first", "KEEP_ALL", "CLOSEST", attributes)
            arcpy.AddField_management(Routes_Attr, "Bcount", "SHORT")
            arcpy.CalculateField_management(Routes_Attr, "Bcount", 1,"PYTHON_9.3")

            Delete("dissolved.shp")
            arcpy.Dissolve_management(Routes_Attr, "dissolved.shp")
            arcpy.AddField_management("dissolved.shp", "Length", "DOUBLE" )
            arcpy.CalculateField_management("dissolved.shp", "Length", "!shape.length@meters!","PYTHON_9.3")
            with arcpy.da.SearchCursor("dissolved.shp", ['Length']) as cursor:
                    for row in cursor:
                        length_new = row[0]

            # spatial join segments and routes_attr to obtain cumulative impedance
            Delete("grid2.shp")
            grid = os.path.join(arcpy.env.workspace, "grid2.shp") 
            SpatialJoin(segments, Routes_Attr , grid,  "sum", "KEEP_ALL", "COMPLETELY_WITHIN", attributes)

       
        

        # merge connecting lines to the net result
        Delete("connectingLines_attr.shp")
        fc_connectingLines_attr = os.path.join(arcpy.env.workspace, "connectingLines_attr.shp")
        arcpy.MakeFeatureLayer_management(target, 'target_lyr') 
        arcpy.MakeFeatureLayer_management(Routes_Attr, 'routes_lyr')
        arcpy.SelectLayerByLocation_management('target_lyr', "INTERSECT", 'routes_lyr', 5.0)  # search radius of 5 meter
        net_count = int(arcpy.GetCount_management('target_lyr').getOutput(0))
        attributes = ["EBF", "WVBRpEBF", "WVBR", "leistungkW"]
        SpatialJoin(fc_connectingLines, 'target_lyr' , fc_connectingLines_attr,  "first", "KEEP_COMMON", "INTERSECT", attributes)
        arcpy.AddField_management(fc_connectingLines_attr, "Bcount", "SHORT")
        arcpy.CalculateField_management(fc_connectingLines_attr, "Bcount", 1,"PYTHON_9.3")           
        arcpy.Merge_management([grid, fc_connectingLines_attr], Net)

       
        # do not display segments with kW value of zero
        with arcpy.da.UpdateCursor(Net, ['leistungkW']) as cursor:
                for row in cursor:
                    if row[0] == 0:
                        cursor.deleteRow()

        
        #arcpy.CopyFeatures_management("grid2.shp", Net)
        arcpy.CheckInExtension('Network')
        
        arcpy.AddMessage("Calculating Parameters...")
        ### ------------------------------------- Calculations ----------------------------------### 

        
        # calculating factor, GLF
        arcpy.AddField_management(Net, "GLF", "DOUBLE" )
        expression = "f(!Bcount!)"
        codeblock = """def f(n):
                        a = 0.4497
                        b = 0.5512
                        c = 53.8483
                        d = 1.7627

                        y = a + (b / (1+ pow(n/c,d)))
                        return y"""

        arcpy.CalculateField_management(Net, "GLF", expression, "PYTHON_9.3", codeblock)
        
        # calculating volume
        arcpy.AddField_management(Net, "kW_GLF", "DOUBLE" )
        arcpy.CalculateField_management(Net, "kW_GLF", "!leistungkW!*!GLF!", "PYTHON_9.3")

        arcpy.AddField_management(Net, "Volume", "DOUBLE" )

        #piecewise linear interpolation
        t =[0, 10, 20, 30, 40, 50, 60, 70, 80 , 90, 100] #temperature
        d = [0.99984, 0.9997, 0.99821, 0.99565, 0.99222, 0.98803, 0.9832, 0.97778, 0.97182, 0.96535, 0.9584] # density sample
        c = [4.2176, 4.1921, 4.1818, 4.1784, 4.1785, 4.1806, 4.1843, 4.1895, 4.1963, 4.205, 4.2159] #capacity sample
        density = np.interp(int(htemp), t, d)
        cp = np.interp(int(htemp), t, c)
        
        with arcpy.da.UpdateCursor(Net, ['kW_GLF', 'Volume']) as cursor:
                for row in cursor:
                    row[1] = row[0] / ( density * cp * ( int(htemp) - int(ltemp) ) ) 
                    cursor.updateRow(row)

       
        # Diameter,Velocity and losses
        
        db = r"D:\Jigeeshu\losses.xlsx" # remember to copy the file in correct folder


        xl_workbook = xlrd.open_workbook(db)
        sheet = xl_workbook.sheet_by_index(0)
        l = [[sheet.cell_value(r,c) for r in range(sheet.nrows)] for c in range(sheet.ncols)]

        DN = l[3] # innendurchmesser
        
 
        K = int(htemp) - 10  # outside temperature is considered 10 C assuming below ground installation
        
        arcpy.AddField_management(Net, "DN", "DOUBLE" )
        arcpy.AddField_management(Net, "Velocity", "DOUBLE" )
        arcpy.AddField_management(Net, "Verlust", "DOUBLE" )
        arcpy.AddField_management(Net, "Preis", "DOUBLE" )

        arcpy.CalculateField_management(Net, "Shape_Leng", "!shape.length@meters!","PYTHON_9.3")
         
        with arcpy.da.UpdateCursor(Net, ['Volume', 'DN', 'Velocity', 'Shape_leng', 'Verlust', 'Preis']) as cursor:
                    for row in cursor: 
                        v = 0
                        vtemp = []
                        n = -1
                        try: 
                            while (v < 0.5 or v > 1):
                                n = n + 1
                                r = DN[n]/2
                                vtemp.append( row[0] * 1000 / (3.142 *  pow(r,2)))
                                v = vtemp[-1]
                                Diameter = DN[n]
                            
                        except:
                            
                                diff = []
                                n = -1
                                for i in range(len(vtemp)):
                                    if (vtemp[i]-1) > 0:
                                        n = n + 1
                                        diff.append(vtemp[i]-1)
                                if (n > 0):
                                    v = vtemp[len(diff)-1]
                                    Diameter = DN[n]
                                else:
                                    v = vtemp[0]
                                    Diameter = DN[0]
                                
                        
                        #losses - not propagating and Cost
                        for i in range(len(l[3])):
                            if Diameter == l[3][i]:
                                #print (l[3][i], l[4][i])
                                loss = (l[4][i]*K*row[3])/1000
                                cost = l[9][i]

                        row[1] = Diameter
                        row[2] = v
                        row[4] = loss
                        row[5] = cost
                       
                        cursor.updateRow(row)

        arcpy.AddMessage("Preparing Dashboard...")

        ### ------------------------------------- Dashboard ----------------------------------### 
        
        book = xlwt.Workbook()
        sh = book.add_sheet("Sheet1")

        sh.write(0,0,'Typ')
        sh.write(0,1,'WVBR')
        sh.write(0,2,'max_kwGLF')
        sh.write(0,3,'Building Count')
        sh.write(0,4,'DN')
        sh.write(0,5,'DN_length')
        sh.write(0,6, 'Total_Cost')

        typ = ['EFH', 'MFH', 'GH', 'GKO', 'GMK']

        l = []
        with arcpy.da.SearchCursor(building, ['Lastprofil']) as cursor:
            for row in cursor:
                l.append(str(row[0]))



        ul = list(set(l))

        WVBR_values = []

        for i in range(len(typ)):
            sh.write(i+1,0,typ[i])
            total = 0
            losses  = 0
            with arcpy.da.SearchCursor(building, ['WVBR', 'Lastprofil']) as cursor:
                for row in cursor:
                    if typ[i] in ul:
                        if row[1] == typ[i]:
                            total = total + row[0]
                    else:
                        total = 0
            WVBR_values.append(total)
            sh.write(i+1,1,total)

        verlust = 0 
        with arcpy.da.SearchCursor(Net, ['Verlust']) as cursor:
                for row in cursor:
                        verlust = verlust + row[0]
        sh.write(len(typ)+1, 0 , 'Verlust')
        sh.write(len(typ)+1, 1 , verlust*8760)
        WVBR_values.append(verlust*8760)

        l1 = []
        with arcpy.da.SearchCursor(Net, ['kW_GLF']) as cursor:
            for row in cursor:
                l1.append(row[0])

        sh.write(1,2, max(l1))
        arcpy.MakeFeatureLayer_management(building, "building_lyr")
        arcpy.SelectLayerByLocation_management("building_lyr", 'INTERSECT', Net)
        result = arcpy.GetCount_management("building_lyr")

        cnt = int(result.getOutput(0)) 
        sh.write(1,3, cnt) # count of building included in net

        l2 = []
        with arcpy.da.SearchCursor(Net, ['DN']) as cursor:
            for row in cursor:
                l2.append(row[0])

        ul1 = list(set(l2))

        arcpy.CalculateField_management(Net, "Shape_Leng", "!shape.length@meters!","PYTHON_9.3")
        for i in range(len(ul1)):
            sh.write(i+1,4,ul1[i])
            total = 0 
            with arcpy.da.SearchCursor(Net, ['Shape_Leng', 'DN']) as cursor:
                for row in cursor:
                    if row[1] == ul1[i]:
                        total = total + row[0]
            sh.write(i+1,5,total)


        book.save(workbook)

        arcpy.AddMessage("Generating Loadcurve...")

        ### ------------------------------------- Loadcurve ----------------------------------###

        loadcurve = os.path.join(arcpy.env.workspace, "loadcurve.xls")
        LoadCurve(WVBR_values[0], WVBR_values[1],WVBR_values[2],WVBR_values[3],WVBR_values[4],WVBR_values[5], loadcurve, arcpy.env.workspace)
