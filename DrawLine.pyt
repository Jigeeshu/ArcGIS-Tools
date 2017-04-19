import arcpy, os
#import Shapely

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "ConnectLine"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "ConnectLine"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        param0 = arcpy.Parameter(
            displayName="Input Building Feature",
            name="input1",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Input")
	param0.filter.list = ["Polygon"]

        param1 = arcpy.Parameter(
            displayName = "Input Street Feature",
            name = "input2",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction= "Input")
        param1.filter.list = ["Polyline"]
		
	param2 = arcpy.Parameter(
            displayName = "Output Line Feature",
            name = "out_features",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction= "Output")
        param2.filter.list = ["Polyline"]
        
        params = [param0 , param1, param2]

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
        streets = parameters[1].valueAsText
	segments =  parameters[2].valueAsText
	scratchWorkspace = "C:/Users/Jigeeshu/Documents/ArcGIS/Default.gdb"
	
	
	tempFC = os.path.join(scratchWorkspace, "temp")

	arcpy.FeatureToPoint_management(building, tempFC, "CENTROID")
	arcpy.AddXY_management(tempFC)

	arcpy.Near_analysis(tempFC, streets, "", "LOCATION", "NO_ANGLE" )
	arcpy.XYToLine_management(tempFC, segments, "POINT_X", "POINT_Y", "NEAR_X", "NEAR_Y")
		
        return
