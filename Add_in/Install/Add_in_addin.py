import arcpy
import pythonaddins

class ButtonClass1(object):
    """Implementation for Add_in_addin.btn1 (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        mxd = arcpy.mapping.MapDocument('current')
        df = arcpy.mapping.ListDataFrames(mxd)[0]
