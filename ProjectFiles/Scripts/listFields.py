import arcpy

##    Engineering Drawing Reference = r'C:\GIS\OperationalSystems\Locations\Waterworks BaseMap\Data\Shapefiles\Drawings\cadindex.shp'
##    Button
##        Add Group Layer to bottom of ToC - Drawings & Historical Orthophotographs
##        Add all TIFs that are within the rectangle the user draws.
##            TIF location = r'S:\Docs\1910\'

dict = {
    ## Key : Value Pairs    Key = entry   Value = filepath
    ## Waterworks Feature Classes
    'ADC Grid'  : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorInfrequent.mdb/Reference/ADCVirginiaPeninsulaGridArea',
    'Street'    : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/Transportation/Road',
    'Hydrants'  : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_Hydrant_EAM',
    'Valves'    : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_SystemValve_EAM',
    'Meters'    : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_wServiceLocation_EAM',
    'Water Laterals'    : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_wLateral_EDM'
        }

def listFields(LAYER):
    featureClass = dict[LAYER]
    fieldList = arcpy.ListFields(featureClass)
    print "Layer:", LAYER
    print "Path:", dict[LAYER]
    for field in fieldList:
        if field.type == 'String':
            print field.name
    print "\n--------------------------------------------\n"

##listFields('ADC Grid')
##listFields('Street')
##listFields('Hydrants')
##listFields('Valves')
##listFields('Meters')
listFields('Water Laterals')


#self.items = ["ServiceLocationID", "TapID", "Subtype", "LifecycleStatus", "QCStatus", "MainDiameter", "TapDiameter", "DrawingNumber", "ProjectNumber", "SDCStatus", "LegacyAccountNumber", "District"]











