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
            displayName = "Fishnet Layer",
            name = "input2",
            datatype = "Feature Layer",
            parameterType = "Required",
            direction= "Input")
        param1.filter.list = ["Polygon"]

        param2 = arcpy.Parameter(
            displayName = "Road Network",
            name = "input3",
            datatype = "Feature Layer",
            parameterType = "Required",
            direction= "Input")
        param2.filter.list = ["Line"]

        param3 = arcpy.Parameter(
            displayName = "Percentage of Potential User (>0%)",
            name = "input4",
            datatype = "GPLong",
            parameterType = "Required",
            direction= "Input")
        

        param4 = arcpy.Parameter(
            displayName = "Output Feature",
            name = "out_features",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction= "Output")
        param4.filter.list = ["Polygon"]
        
        params = [param0, param1,  param2, param3,  param4]
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
        fishnetClip = parameters[1].valueAsText
        data = r"D:\Jigeeshu\Tools\Tools\building_info_18_10_2016.xlsx"
        road = parameters[2].valueAsText
        percentage = parameters[3].valueAsText
        hotspot = parameters[4].valueAsText
        arcpy.env.workspace =  arcpy.Describe(building).path
        
        dicts = []
        arcpy.DeleteField_management(building, 
         ["WVBR", "EBF", "leistungkW", "WVBRpEBF", "Bez", "Vlh", "Typ", "Lastprofil", "Strom", "Shp_Area"])  
        # Faktor has been removed from the list

        
        xl_workbook = xlrd.open_workbook(data)
        sheet = xl_workbook.sheet_by_index(0)
        l = [[sheet.cell_value(r,c) for r in range(sheet.nrows)] for c in range(sheet.ncols)]

        #l = [[x[i] for x in tl] for i in range(len(tl[0]))]

        arcpy.AddMessage("Calculating Faktor...")
           
        #calculating faktor depending on 'Geschosse'
        expression = "f(!Geschosse!)"
        codeblock = """def f(x):
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
                    return y"""
        
        arcpy.AddField_management(building, "Faktor", "FLOAT" )
        arcpy.CalculateField_management(building, "Faktor", expression,"PYTHON_9.3", codeblock)

        
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

            with arcpy.da.UpdateCursor(building, ['Funktion',f]) as cursor:  # change field name if required 
                for row in cursor:
                    if float(row[0]) in dicts.keys():
                        row[1] = dicts.get(float(row[0])) 
                    else:
                        row[1] = None     
                    cursor.updateRow(row)
                     
        ### ------------------adding WVBRpEBF value based on building age ------------------------------###
        arcpy.AddMessage("Assign WVBRpEBF based on Building Age...")
        
        #spatial join building age layer and building to get BAK and kWh
        """
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
        
        
        
         # if BAK is not null and funktion is 1010 then change WVBRpEBF value
        expression = "f(!kWh_qm!, !Funktion!, !WVBRpEBF!)"
        codeblock = def f(x,y,z):
                    if x != 0 and y == "1010":
                         return x
                    else:
                        return z

        arcpy.CalculateField_management(building, "WVBRpEBF", expression, "PYTHON_9.3", codeblock) # output_join changed
        """
        #arcpy.CopyFeatures_management(output_join, building)

        ### ------------------------------------------------------------------------------------------###
        arcpy.AddMessage("Computing Hotspots ...")
        #calculate EBF, WVBR and kW
        arcpy.AddField_management(building, "EBF", "DOUBLE")
        arcpy.AddField_management(building, "Shp_Area", "DOUBLE")
        arcpy.CalculateField_management(building, "Shp_Area", "!shape.area@squaremeters!","PYTHON_9.3")
        arcpy.CalculateField_management(building, "EBF", "!Shp_Area!*!Faktor!","PYTHON_9.3")
        
        arcpy.AddField_management(building, "WVBR_100", "DOUBLE")
        arcpy.CalculateField_management(building, "WVBR_100", "!WVBRpEBF!*!EBF!","PYTHON_9.3")

        arcpy.AddField_management(building, "WVBR", "DOUBLE")
        with arcpy.da.UpdateCursor(building, ['WVBR','WVBR_100']) as cursor:
                for row in cursor:
                    row[0] = row[1] * int(percentage)/100
                    cursor.updateRow(row)

        expression = "f(!WVBR!, !Vlh!)"
        codeblock = """def f(x,y):
                        if y != 0 :
                            z = x/y
                        else:
                            z = 0
                        return z"""
            
        
        arcpy.AddField_management(building, "leistungkW", "DOUBLE")
        arcpy.CalculateField_management(building, "leistungkW", expression,"PYTHON_9.3", codeblock)

        #fishnetClip = r"C:\Users\Jigeeshu\Desktop\GIS_Data\Gebäudeumringe\sie02_fCopy.shp"
        
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

        #spatial join landuse grid with intersected/clipped buildings
        
        Delete(os.path.join(arcpy.env.workspace, "hotspot_temp.shp"))
        hotspot_temp = os.path.join(arcpy.env.workspace, "hotspot_temp.shp")
    
        attributes = ["WVBR_sect",  "EBF_sect", "kW_sect"]
        SpatialJoin(fishnetClip, intersected , hotspot_temp,  "sum", "KEEP_ALL", "CONTAINS", attributes)
        
        arcpy.AddField_management(hotspot_temp, "WVBRpArea", "DOUBLE")
        arcpy.AddField_management(hotspot_temp, "Shp_Area", "DOUBLE")
        arcpy.CalculateField_management(hotspot_temp, "Shp_Area", "!shape.area@hectares!","PYTHON_9.3")
        arcpy.CalculateField_management(hotspot_temp, "WVBRpArea", "!WVBR_sect! / !Shp_Area!","PYTHON_9.3")
        arcpy.CalculateField_management(hotspot_temp, "WVBRpArea", "!WVBRpArea! / 1000","PYTHON_9.3")  # MegaWatt
      

        
        ### ------------------------------------ Generalization ------------------------------------- ###
       
        arcpy.AddMessage("Generalizing...")
        arcpy.MakeFeatureLayer_management(hotspot_temp, "lyr")
        selected_hotspot = os.path.join(arcpy.env.workspace, "selected_hotspot.shp")
        arcpy.SelectLayerByAttribute_management("lyr", 'NEW_SELECTION', '"WVBRpArea" > 600 OR "OBJART_TXT" = \'AX_FlaecheGemischterNutzung\' OR "OBJART_TXT" = \'AX_FlaecheBesondererFunktionalerPraegung\'')
      #  arcpy.SelectLayerByAttribute_management("lyr",  "SUBSET_SELECTION", '"OBJART_TXT" = \'AX_FlaecheGemischterNutzung\' OR "OBJART_TXT" = \'AX_FlaecheBesondererFunktionalerPraegung\' ')
        arcpy.CopyFeatures_management("lyr", selected_hotspot)

        
        Delete(os.path.join(arcpy.env.workspace, "selected_hotspot_dissolved.shp"))
        selected_hotspot_dissolved = os.path.join(arcpy.env.workspace, "selected_hotspot_dissolved.shp")
        arcpy.Dissolve_management(selected_hotspot, selected_hotspot_dissolved, "", "","SINGLE_PART")


        Delete(os.path.join(arcpy.env.workspace, "selected_hotspot_dissolved_buffer15m.shp"))
        selected_hotspot_dissolved_buffer15m = os.path.join(arcpy.env.workspace, "selected_hotspot_dissolved_buffer15m.shp")
        arcpy.Buffer_analysis(selected_hotspot_dissolved, selected_hotspot_dissolved_buffer15m, "15 meters")

        
        Delete(os.path.join(arcpy.env.workspace, "selected_hotspot_buffer15m_dissolved.shp"))
        selected_hotspot_buffer15m_dissolved = os.path.join(arcpy.env.workspace, "selected_hotspot_buffer15m_dissolved.shp")
        arcpy.Dissolve_management(selected_hotspot_dissolved_buffer15m, selected_hotspot_buffer15m_dissolved, "", "","SINGLE_PART")

        
        
        arcpy.AddField_management(selected_hotspot_buffer15m_dissolved, "Area_ha", "DOUBLE")
        arcpy.CalculateField_management(selected_hotspot_buffer15m_dissolved, "Area_ha", "!shape.area@hectares!","PYTHON_9.3")

        
        arcpy.MakeFeatureLayer_management(selected_hotspot_buffer15m_dissolved, "lyr_dissolved")
        arcpy.SelectLayerByAttribute_management("lyr_dissolved", 'NEW_SELECTION', '"Area_ha" > 10')
        Delete(os.path.join(arcpy.env.workspace, "hotspot_consolidated.shp"))
        hotspot_consolidated = os.path.join(arcpy.env.workspace, "hotspot_consolidated.shp")
        arcpy.CopyFeatures_management("lyr_dissolved", hotspot_consolidated)

        
        arcpy.DeleteField_management(building,  ["count"])
        arcpy.AddField_management(building, "count", "INTEGER")
        arcpy.CalculateField_management(building, "count", "1","PYTHON_9.3")

        attributes =['WVBR','count']
        hotspot_consolidated_gebaeude_SPJ = os.path.join(arcpy.env.workspace, "hotspot_consolidated_gebaeude_SPJ.shp")
        SpatialJoin(hotspot_consolidated,  building, hotspot_consolidated_gebaeude_SPJ,  "sum", "KEEP_ALL","INTERSECT", attributes)
        arcpy.AddField_management(hotspot_consolidated_gebaeude_SPJ, "Area_ha", "DOUBLE")
        arcpy.CalculateField_management( hotspot_consolidated_gebaeude_SPJ, "Area_ha", "!shape.area@hectares!","PYTHON_9.3")
        arcpy.AddField_management(hotspot_consolidated_gebaeude_SPJ, "WVBR_ha", "DOUBLE")
        arcpy.CalculateField_management( hotspot_consolidated_gebaeude_SPJ, "WVBR_ha", "!WVBR! / (!Area_ha!*1000)","PYTHON_9.3")

        
        Delete(os.path.join(arcpy.env.workspace, "road_basisDLM_clipped.shp"))
        road_clipped = os.path.join(arcpy.env.workspace, "road_basisDLM_clipped.shp")
        arcpy.Clip_analysis(road, hotspot_consolidated_gebaeude_SPJ , road_clipped)
        arcpy.AddField_management(road_clipped, "length", "DOUBLE")
        arcpy.CalculateField_management( road_clipped, "length", "!shape.length@meters!","PYTHON_9.3")

        attributes =['length']
        
        SpatialJoin(hotspot_consolidated_gebaeude_SPJ,  road_clipped, hotspot,  "sum", "KEEP_ALL","INTERSECT", attributes)
        
