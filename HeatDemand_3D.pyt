import sys, string, os, arcgisscripting
import arcpy
import xlwt, xlrd
from xlrd import open_workbook
from os.path import basename, dirname, exists, join

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Tool"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        param0 = arcpy.Parameter(
            displayName="Input CityGML File",
            name="CityGML",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")


        param1 = arcpy.Parameter(
            displayName='Input Excel File (Air Temperature °C)',
            name='AirTemperature',
            datatype='DEFile',
            parameterType='Required',
            direction='Input')

        param1.filter.list = ["xls", "xlsx"]

        param2 = arcpy.Parameter(
            displayName = "Inside Constant Temperature °C",
            name = "InsideTemp",
            datatype = "GPDouble",
            parameterType = "Optional",
            direction= "Input")
        param2.value = 20.0

        param3 = arcpy.Parameter(
            displayName = "Select a folder to save output geodatabase",
            name = "output",
            datatype = "DEFolder",
            parameterType = "Required",
            direction= "Input")

        params = [param0 , param1, param2, param3]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        
        
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        data = parameters[1].valueAsText
        citygml= parameters[0].valueAsText
        insideT = parameters[0].valueAsText
        path = parameters[3].valueAsText
        output_gb = os.path.join(path, "output_gb.gdb")
        
        # reading temperature data from excel
        xl_workbook = xlrd.open_workbook(data)
        sheet = xl_workbook.sheet_by_index(0)

        airT = [[sheet.cell_value(r,c) for r in range(sheet.nrows)] for c in range(sheet.ncols)]
        diffT = []
        insideT = 20
        for i in range(1, len(airT[0])):
            diffT.append(float(insideT) - airT[0][i])


        #converting citygml data to Multipatch feature
        if exists(output_gb):
            arcpy.Delete_management(output_gb)
        gp = arcgisscripting.create()
        gp.CheckOutExtension("DataInteroperability")
        gp.QuickImport(citygml, output_gb)

        # creating 
        arcpy.CheckOutExtension('3D')

        wall = os.path.join(output_gb, "WallSurface_surface")
        roof = os.path.join(output_gb,"RoofSurface_surface")
        ground = os.path.join(output_gb, "GroundSurface_surface")
        building=os.path.join(output_gb,"Building_surface")
        building_part= os.path.join(output_gb, "BuildingPart_surface")
        shell = os.path.join(output_gb,"AllBuilding_Surface")
        arcpy.Merge_management([building, building_part], shell)

        
        arcpy.AddZInformation_3d(wall, 'SURFACE_AREA',  'NO_FILTER')
        arcpy.AddZInformation_3d(roof, 'SURFACE_AREA',  'NO_FILTER')
        arcpy.AddZInformation_3d(ground, 'SURFACE_AREA',  'NO_FILTER')
        arcpy.AddZInformation_3d(shell, 'Volume',  'NO_FILTER')
        
        arcpy.AddField_management(shell, "T_loss", "DOUBLE")
        arcpy.AddField_management(shell, "V_loss", "DOUBLE")
        arcpy.AddField_management(shell, "total_loss", "DOUBLE")

        
        dict_1 = {'Uwall' : 1.7, 'Uwindow': 2.7, 'Gwindow':0.76, 'Uroof': 1.5, 'Uground':1.2, 'Ratio':0.30}
        dict_2 = {'Uwall' : 1.7, 'Uwindow': 2.7, 'Gwindow':0.76, 'Uroof': 1.5, 'Uground':1.2, 'Ratio':0.25}
        dict_3 = {'Uwall' : 1.4, 'Uwindow': 2.7, 'Gwindow':0.76, 'Uroof': 1.3, 'Uground':1, 'Ratio':0.23}
        dict_4 = {'Uwall' : 1.2, 'Uwindow': 2.7, 'Gwindow':0.76, 'Uroof': 1.1, 'Uground':0.84, 'Ratio':0.28}
        dict_5 = {'Uwall' : 0.8, 'Uwindow': 2.7, 'Gwindow':0.76, 'Uroof': 0.45, 'Uground':0.6, 'Ratio':0.33}
        dict_6 = {'Uwall' : 0.4, 'Uwindow': 1.7, 'Gwindow':0.72, 'Uroof': 0.3, 'Uground':0.4, 'Ratio':0.35}

        dict_U = {'range1':dict_1, 'range2': dict_2, 'range3': dict_3, 'range4': dict_4, 'range5': dict_5 ,'range6': dict_6}


        with arcpy.da.UpdateCursor(shell, ['gml_id', 'T_loss', 'V_loss', 'total_loss', 'Volume']) as cursor1:
    
            # calculate transmission loss
            for everybuilding in cursor1:
                
                #surface area of building walls
                SArea_wall = 0
                with arcpy.da.SearchCursor(wall, ['gml_parent_id', 'SArea']) as cursor2:
                    for everywall in cursor2:
                        if (everywall[0]==everybuilding[0]):
                            SArea_wall = SArea_wall + everywall[1]

                loss_wall = dict_U['range5']['Uwall']*SArea_wall*0.67 # depends on building age -  condition to be added

                loss_window = dict_U['range5']['Uwindow']*SArea_wall*0.33
                # surface area of building roofs            
                SArea_roof = 0
                with arcpy.da.SearchCursor(roof, ['gml_parent_id', 'SArea']) as cursor2:
                    for everyroof in cursor2:
                        if (everyroof[0]==everybuilding[0]):
                            SArea_roof = SArea_roof + everyroof[1]
                loss_roof = dict_U['range5']['Uroof']*SArea_roof # depends on building age -  condition to be added

                #surface area of building ground
                SArea_ground = 0
                with arcpy.da.SearchCursor(ground , ['gml_parent_id', 'SArea']) as cursor2:
                    for everyground  in cursor2:
                        if (everyground[0]== everybuilding[0]):
                            SArea_ground  = SArea_ground  + everyground[1]

                loss_ground = dict_U['range5']['Uground']*SArea_ground   # depends on building age -  condition to be added

                total_loss = loss_wall + loss_roof + loss_ground + loss_window
                
                # calculate the hourly loss (degree hour)
                hour_Tloss = []
                for i in range(len(diffT)): 
                    if (i> 2880 and i<6553):
                        hour_Tloss.append(0)
                    else:
                        hour_Tloss.append(total_loss*diffT[i])

                everybuilding[1]= sum(hour_Tloss)/1000
                
                # calculate ventilation loss
                hour_Vloss = []
                for i in range(len(diffT)):
                    
                    if (i > 2880 and i<6553):
                        hour_Vloss.append(0)
                    else:
                        hour_Vloss.append(0.00033641 * 0.75 * 0.76* everybuilding[4] * diffT[i])
                    
                everybuilding[2] = sum(hour_Vloss)
                    
                everybuilding[3] = everybuilding[1] +  everybuilding[2]
                cursor1.updateRow(everybuilding) 
            
