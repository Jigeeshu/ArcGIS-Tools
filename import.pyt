import arcpy
import xlrd
import xlwt
import datetime


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "ImportCables"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "ImportCables"
        self.description = "Converts excel sheet to another with replace values for fields with associated lookup tables"
        self.canRunInBackground = True

    def getParameterInfo(self):
        param1 = arcpy.Parameter(
            displayName='Input Excel Table',
            name='excel_in',
            datatype='DEFile',
            parameterType='Required',
            direction='Input')

        param1.filter.list = ["xls"]

        param2 = arcpy.Parameter(
            displayName = "Input Feature",
            name = "in_features",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction= "Input")
        param2.filter.list = ["Polyline"]
        
        params = [param1, param2]

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

         # parameters into python variable
        fname = parameters[0].valueAsText
        fc = parameters[1].valueAsText
                
        fields = [f.name for f in arcpy.ListFields(fc)] 
        
        xl_workbook = xlrd.open_workbook(fname)
        
        #sheet_names = xl_workbook.sheet_names()       
        #sheet = xl_workbook.sheet_by_name(sheet_names[0])
        sheet = xl_workbook.sheet_by_index(0)
        
        
        #getting values in python list 
        l = [[sheet.cell_value(r,c) for r in range(sheet.nrows)] for c in range(sheet.ncols)]
        li = list(enumerate(l))

        a = "a"
        b = "b"
        c = "c"
        d = "d"
        e = "e"
        f = "f"
        g = "g"
        h = "h"



        for i in range(len(l)):
            if l[i][0] == u'STATUS':
                a = i
            elif  l[i][0] == u'U_FARBE':
                b = i
            elif  l[i][0] == u'TYP':
                c = i
            elif  l[i][0] == u'U_BETRIEBSSPANNUNG':
                d = i
            elif  l[i][0] == u'U_FUNKTION':
                e = i
            elif  l[i][0] == u'U_NENNSPANNUNG_2':
                f = i
            elif  l[i][0] == u'U_QUERSCHNITT_2':
                g = i
            elif  li[i][1][0] == u'U_KABELKENNZEICHEN':
                h = li[i][0]
        
        
        wl = fc.split("\\")
        arcpy.env.workspace = "/".join(wl[0:len(wl)-2]) 
        path_log =  "/".join(wl[0:len(wl)-3]) 
        tbs = ["ELES_STATUS", "U_ELES_KABELBANDFARBE"  , "ELES_TYP_KABEL" , "U_ELES_BETRIEBSSPANNUNG" , "U_ELES_FUNKTION_KABEL"  ,"U_ELES_NENNSPANNUNG_2", "U_ELES_QUERSCHNITT_2_KABEL" ]
        
        
        dicts = []
        for tb in tbs:
            with arcpy.da.SearchCursor(tb,["CODE", "DESCRIPTION_G"]) as srows:
                dicts.append(dict([srow[0], srow[1]] for srow in srows))
        
           
        lookup = {
            a:dicts[0], b:dicts[1], c:dicts[2], d:dicts[3], e:dicts[4], f:dicts[5], g:dicts[6]
          }
        
        
                   
        for x in range(len(li)):
            if li[x][0] in lookup.keys():
                for i in range(len(li[x][1])): 
                    j = 0
                    while j < len(li[x][1]):
                        for k,v in lookup[x].items():
                            if l[x][j] == v:
                                l[x][j] = k
                            else:
                                l[x][j] = l[x][j]
                        j = j+1
                                      
        
        for i in range(len(l)):
            for j in range(len(l[0])):
                if l[i][j] == '':
                    l[i][j] = None

        log = open(path_log +"/import_log.txt","w")
        log.write("\n")
        d = datetime.datetime.now()
        log.write("Time: " + str(d) + "\n")
              
        for field in fields:
            for i in range(len(l)):
                if field == l[i][0]:
                    flist = [field, u'U_KABELKENNZEICHEN']
                    with arcpy.da.UpdateCursor(fc, flist) as cursor:
                        for j in range(1,len(l[i])):
                            try:
                                for row in cursor:   
                                        if row[1] == l[h][j]:
                                            row[0] = l[i][j]
                                            print row
                                            cursor.updateRow(row)
                                            break
                            except:
                                   log.write("\n")               
                                   log.write("Kabelkennzeichen [" + l[h][j]  + "]: fehlerhafter Wert '" +l[i][j] + "' fur Attribut " + l[i][0]+ ", (Zeile "+ str(j+1)+")"+"\n")
                                   
        
        return
