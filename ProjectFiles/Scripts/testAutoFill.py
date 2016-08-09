import arcpy

dict = {
    ## Key : Value Pairs    Key = entry   Value = filepath
    'ADC Grid'          : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorInfrequent.mdb/Reference/ADCVirginiaPeninsulaGridArea',
    'Street'            : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/Transportation/Road',
    'Hydrants'          : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_Hydrant_EAM',
    'Valves'            : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_SystemValve_EAM',
    'Meters'            : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_wServiceLocation_EAM',
    'Water Laterals'    : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_wLateral_EDM'
        }

def TextAutoFill(layer,field):
##    item_parts = {}
    featureClass = dict[layer]
    items=[]
    with arcpy.da.SearchCursor(featureClass, field) as cursor:
        for row in cursor:
            items.append(row) 
##            #item = ' '.join(map(lambda s: s.strip(), filter(lambda s: s is not None and s != '' and s != ' ', row)))
##            item = ''.join(map(lambda s: s.strip(), filter(lambda s: s is not None and s != '' and s != ' ', row)))
##            parts = {}
##            parts[field] = row
##            item_parts[item] = parts
##    items = item_parts.keys()
##    items.sort()
    items.sort()
    for x in items:
       print x[0]

#TextAutoFill('Meters','Address')
#TextAutoFill('Meters','ServiceLocationID')
#TextAutoFill('Valves','ACCT_NO')







##            
##            self.item_parts = {}
##            featureClass = dict[layer]
##            with arcpy.da.SearchCursor(featureClass, field) as cursor:
##                for row in cursor:
##                    item = ' '.join(map(lambda s: s.strip(), filter(lambda s: s is not None and s != '' and s != ' ', row)))
##                    parts = {}
##                    parts[field] = row
##                    self.item_parts[item] = parts
##            self.items = self.item_parts.keys()
##            self.items.sort()















def NumberAutoFill(layer,field):
    featureClass = dict[layer]
    items=[]
    with arcpy.da.SearchCursor(featureClass, field) as cursor:
        for row in cursor:
            items.append(row)#(filter(lambda s: s is not None and s != '' and s != ' ', str(row)))
    items.sort()
    for x in items:
       print x[0]#x


##    
##    item_parts = {}
##    featureClass = dict[layer]
##    with arcpy.da.SearchCursor(featureClass, field) as cursor:
##        for row in cursor:
##            item = row#''.join(map(lambda s: s.strip(), filter(lambda s: s is not None and s != '' and s != ' ', str(row))))
##            parts = {}
##            parts[field] = item
##            item_parts[item] = parts
##    items = item_parts.keys()
##    items.sort()
##    for x in items:
##       print x#[1:-1]#.replace('(','').replace(',)','')

NumberAutoFill('ADC Grid','MapNumber')
##NumberAutoFill('Valves','OBJECTID')
##NumberAutoFill('Hydrants','HYDRANT_NO')

'''
#.replace(',','').replace('(','').replace(')','')
############################################################################################################################
############################################################################################################################
############################################################################################################################


This is the class that I have written. It works but is there a better way to handle the field being an integer?

'''

class ComboBoxClass_3_Keyword(object):
    """Implementation for NNWW_Toolbars_addin.combobox_keyword (ComboBox)"""
    def __init__(self):
        self.editable = True
        self.enabled = True
        self.dropdownWidth = '123456789012345678901234567890'
        self.width = '123456789012345678901234567890'
    def onSelChange(self, selection):
        global CHOICE_keyword
        CHOICE_keyword = selection
    def onEditChange(self, text):
        global CHOICE_keyword
        CHOICE_keyword = text#.upper()
    def onFocus(self, focused):
        def SimpleAutoFill(layer,field):
            self.item_parts = {}
            featureClass = dict[layer]
            with arcpy.da.SearchCursor(featureClass, field) as cursor:
                for row in cursor:
                    item = ' '.join(map(lambda s: s.strip(), filter(lambda s: s is not None and s != '' and s != ' ', row)))
                    parts = {}
                    parts[field] = row
                    self.item_parts[item] = parts
            self.items = self.item_parts.keys()
            self.items.sort()
        def NumberAutoFill(layer,field):
            self.item_parts = {}
            featureClass = dict[layer]
            with arcpy.da.SearchCursor(featureClass, field) as cursor:
                for row in cursor:
                    item = ''.join(map(lambda s: s.strip().replace(',','').replace('(','').replace(')',''), filter(lambda s: s is not None and s != '' and s != ' ', str(row))))
                    parts = {}
                    parts[field] = row
                    self.item_parts[item] = parts
            self.items = self.item_parts.keys()
            self.items.sort()
        
        if focused:
            try:
                if CHOICE_layer == "Street" and CHOICE_field == "NAME":
                    self.street_parts = {}
                    #fieldNames = ('Prefix', 'Name', 'Type', 'Suffix')
                    fieldNames = ('Name', 'Type')
                    featureClass = dict['Street']
                    with arcpy.da.SearchCursor(featureClass, fieldNames) as cursor:
                        for row in cursor:
                            street = ' '.join(map(lambda s: s.strip(), filter(lambda s: s is not None and s != '' and s != ' ', row)))
                            parts = {}
                            #(parts['prefix'], parts['name'], parts['type'], parts['suffix']) = row
                            (parts['name'], parts['type']) = row
                            self.street_parts[street] = parts
                    self.items = self.street_parts.keys()
                    self.items.sort()
                else:
                    SimpleAutoFill(CHOICE_layer,CHOICE_field)
            except:
                NumberAutoFill(CHOICE_layer,CHOICE_field)
        else:
            self.items = []
