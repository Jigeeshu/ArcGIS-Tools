import arcpy
import xlrd
import os
import csv

def Delete(feature):
    if os.path.exists(feature):
            arcpy.Delete_management(feature)

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

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "HotSpot"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "HotSpot"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        param0 = arcpy.Parameter(
            displayName="Input Building Polygons",
            name="input1",
            datatype="Feature Layer",
            parameterType="Required",
            direction="Input")
	param0.filter.list = ["Polygon"]
	
        param1 = arcpy.Parameter(
            displayName = "Building Age Layer",
            name = "input6",
            datatype = "Feature Layer",
            parameterType = "Optional",
            direction= "Input")
        param1.filter.list = ["Polygon"]
         
        param2 = arcpy.Parameter(
            displayName = "Output Feature",
            name = "out_features",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction= "Output")
        param2.filter.list = ["Polygon"]
        
        params = [param0,  param1, param2]
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

        building = parameters[0].valueAsText
        data = r"D:\Jigeeshu\Tools\Tools\building_info_18_10_2016.xlsx"
        age_layer = parameters[1].valueAsText
        hotspot = parameters[1].valueAsText
        arcpy.env.workspace =  arcpy.Describe(building).path
        
        dicts = []
        arcpy.DeleteField_management(building, 
         ["WVBR", "EBF", "leistungkW","Faktor", "WVBRpEBF", "Bez", "Vlh", "Typ", "Lastprofil", "Strom", "Shape_Area"])  

        xl_workbook = xlrd.open_workbook(data)
        sheet = xl_workbook.sheet_by_index(0)
        l = [[sheet.cell_value(r,c) for r in range(sheet.nrows)] for c in range(sheet.ncols)]

        #l = [[x[i] for x in tl] for i in range(len(tl[0]))]
        """     
        #calculating faktor depending on 'Geschosse'
        expression = "f(!Geschosse!)"
        codeblock = def f(x):
                    if x == 0:
                        y = 0.8
                    elif x == 1:
                        y = 1.2
                    elif x == 2:
                        y = 2
                    elif x == 3:
                        y =  2.4
                    elif x >= 4:
                        y = x*0.8
                    return y
        
        arcpy.AddField_management(building, "Faktor", "FLOAT" )
        arcpy.CalculateField_management(building, "Faktor", expression,"PYTHON_9.3", codeblock)

        """
        arcpy.AddField_management(building, "WVBRpEBF", "FLOAT", '#', '#', '#', '#', 'NULLABLE')
        arcpy.AddField_management(building, "Bez", "TEXT" )
        arcpy.AddField_management(building, "Vlh", "SHORT" )
        arcpy.AddField_management(building, "Typ", "TEXT" )
        arcpy.AddField_management(building, "Lastprofil", "TEXT" )
        arcpy.AddField_management(building, "Strom", "SHORT" )

        fields = ['WVBRpEBF' ,"Bez", "Vlh", "Typ", "Lastprofil", "Strom"] #excel column names
        dicts = []
        n = []
        for f in fields:
            for i in range(len(l)):
                if(l[i][0] == f):
                   n.append (i)
            
            keys= l[0]
            values = l[n[-1]]
            dicts = dict(zip(keys, values))

            with arcpy.da.UpdateCursor(building, ['Funktion',f]) as cursor:
                for row in cursor:
                    if float(row[0]) in dicts.keys():
                        row[1] = dicts.get(float(row[0])) 
                    else:
                        row[1] = None     
                    cursor.updateRow(row)
                    
        

        ### ------------------adding WVBRpEBF value based on building age ------------------------------###
        """           
        #spatial join building age layer and building to get BAK and kWh
        joinFC = age_layer
        Delete("Gebaeude_alter.shp")
        output_join =  os.path.join(arcpy.env.workspace, "Gebaeude_alter.shp")

        def SpatialJoin(input_feature, join_feature, output_feature, merge_rule, connectivity_rule, attributes):
            fieldmappings = arcpy.FieldMappings()
            fieldmappings.addTable(input_feature)
            fieldmappings.addTable(join_feature)
            fnames = arcpy.ListFields(join_feature)
            for f in fnames:
                if f.name in attributes: 
                    fidx = fieldmappings.findFieldMapIndex(f.name)
                    fmap = fieldmappings.getFieldMap(fidx)
                    fmap.mergeRule = merge_rule
                    fieldmappings.replaceFieldMap(fidx,fmap)      

            arcpy.SpatialJoin_analysis(input_feature, join_feature, output_feature, "JOIN_ONE_TO_ONE", "KEEP_ALL", fieldmappings, connectivity_rule)


        attributes = ["BAK", "kWh_qm"]

        SpatialJoin(building,  joinFC, output_join,  "first", "INTERSECT", attributes)
        """
        
         # if BAK is not null and funktion is 1010 then change WVBRpEBF value
        expression = "f(!kWh_qm!, !Funktion!, !WVBRpEBF!)"
        codeblock = """def f(x,y,z):
                    if x != 0 and y == "1010":
                         return x
                    else:
                        return z"""

        arcpy.CalculateField_management(output_join, "WVBRpEBF", expression, "PYTHON_9.3", codeblock)

        arcpy.CopyFeatures_management(output_join, building)

        ### ------------------------------------------------------------------------------------------###

                
        #calculate EBF, WVBR and kW
        arcpy.AddField_management(building, "EBF", "DOUBLE")
        arcpy.AddField_management(building, "Shp_Area", "DOUBLE")
        arcpy.CalculateField_management(building, "Shp_Area", "!shape.area@squaremeters!","PYTHON_9.3")
        arcpy.CalculateField_management(building, "EBF", "!Shp_Area!*!Faktor!","PYTHON_9.3")
        
        arcpy.AddField_management(building, "WVBR", "DOUBLE")
        arcpy.CalculateField_management(building, "WVBR", "!WVBRpEBF!*!EBF!","PYTHON_9.3")

        expression = "f(!WVBR!, !Vlh!)"
        codeblock = """def f(x,y):
                        if y != 0 :
                            z = x/y
                        else:
                            z = 0
                        return z"""
            
        
        arcpy.AddField_management(building, "leistungkW", "DOUBLE")
        arcpy.CalculateField_management(building, "leistungkW", expression,"PYTHON_9.3", codeblock)
        
        
        # create fishnet - 100 X 100
        
        Delete("fishnet.shp")
        fishnet = os.path.join(arcpy.env.workspace, "fishnet.shp")
        desc = arcpy.Describe(building)
        origin = str(desc.extent.XMin) +" "+str(desc.extent.YMin)
        yAxis = str(desc.extent.XMin) + " " + str(desc.extent.YMin + 10)
        opp_corner = str(desc.extent.XMax) +" "+str(desc.extent.YMax)
        arcpy.CreateFishnet_management(fishnet,origin,yAxis,"100","100","0","0",opp_corner,"NO_LABELS","#","POLYGON")
        
        # clip fishnet
        
        Delete("fishnetClip.shp")
        Delete("fishnetClip_prj.shp")
        fishnetClip = os.path.join(arcpy.env.workspace, "fishnetClip.shp")
        fishnet_lyr = "fisnhet_lyr"
        arcpy.MakeFeatureLayer_management(fishnet, fishnet_lyr)
        arcpy.SelectLayerByLocation_management(fishnet_lyr, 'INTERSECT', building)
        arcpy.CopyFeatures_management(fishnet_lyr,fishnetClip)
        spatial_ref = arcpy.Describe(building).spatialReference
        sp = str(spatial_ref.name)
        crs = arcpy.SpatialReference(sp.replace("_", " "))
        arcpy.DefineProjection_management(fishnetClip, crs)
    

        
        
        # intersect building with fishnet
        
        Delete("Gbd_intersect.shp")
        intersected = os.path.join(arcpy.env.workspace, "Gbd_intersect.shp")
        arcpy.Intersect_analysis([building,fishnetClip], intersected, "ALL", "", "INPUT")
        arcpy.AddField_management(intersected, "EBF_sect", "DOUBLE")
        arcpy.CalculateField_management(intersected, "Shp_Area", "!shape.area@squaremeters!","PYTHON_9.3")
        arcpy.CalculateField_management(intersected, "EBF_sect", "!Shp_Area!*!Faktor!","PYTHON_9.3")
        arcpy.AddField_management(intersected, "WVBR_sect", "DOUBLE")
        arcpy.CalculateField_management(intersected, "WVBR_sect", "!WVBRpEBF!*!EBF_sect!","PYTHON_9.3")
        arcpy.AddField_management(intersected, "kW_sect", "DOUBLE")


        expression = "f(!WVBR_sect!, !Vlh!)"
        codeblock = """def f(x,y):
                        if y != 0 :
                            z = x/y
                        else:
                            z = 0
                        return z"""
        
        arcpy.CalculateField_management(intersected, "kW_sect", expression,"PYTHON_9.3", codeblock)

 
        # summarize for each fishnet cell
        Delete("summary")
        table = os.path.join(arcpy.env.workspace, "summary")
        arcpy.Statistics_analysis(intersected, table,  [["WVBR_sect", "SUM"],["EBF_sect", "SUM"]], "FID_fishne")
        arcpy.JoinField_management(fishnetClip, "FID", table, "FID_FISHNE", ["SUM_WVBR_SECT"])

        
        arcpy.AddField_management(fishnetClip, "WVBRpArea", "DOUBLE")
        arcpy.CalculateField_management(fishnetClip, "WVBRpArea", "!SUM_WVBR_S! / 1000","PYTHON_9.3")

        arcpy.CopyFeatures_management(fishnetClip, hotspot)
        
        
