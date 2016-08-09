# Experimental_addin.py
# Barbara Gates
# 5/4/16

import arcpy
import pythonaddins


# Use Road feature class in Local Basemap data on user's C: drive.
RoadFeatureClass = r'C:\GIS\OperationalSystems\Locations\Waterworks BaseMap\Data\Databases\sdeVectorFrequent.mdb\Transportation\Road'

class ExperimentalComboBox(object):
    
    def __init__(self):
        # Constructor.  Initializes properties:
        #   items - List of choices displayed in combo box.
        #           If list is fixed, populate it here.
        #           If list is dynamic, populate it in one of the other functions instead.
        #   editable - True (False) = (Do not) allow user to type a new value in list of choices.
        #   enabled - True (False) = Tool is available (unavailable).
        #   value - User-entered (or user-selected) value.
        #   valueIsValid - True (False) = User-entered value is (not) in the required format.
        
        # Load street names from the Road feature class (Prefix, Name, Type, and Suffix fields) into memory:
        #
        #   street_parts = {
        #       'JEFFERSON AVE':          {'prefix': None, 'name': 'JEFFERSON',      'type': 'AVE', 'suffix': None},
        #       'N WILLIAM STYRON SQ S':  {'prefix': 'N',  'name': 'WILLIAM STYRON', 'type': 'SQ',  'suffix': 'S' },
        #       ...
        #   }
        self.street_parts = {}
        fieldNames = ('Prefix', 'Name', 'Type', 'Suffix')
        with arcpy.da.SearchCursor(RoadFeatureClass, fieldNames) as cursor:
            for row in cursor:
                street = ' '.join(map(lambda s: s.strip(), filter(lambda s: s is not None and s != '' and s != ' ', row)))
                parts = {}
                (parts['prefix'], parts['name'], parts['type'], parts['suffix']) = row
                self.street_parts[street] = parts
                
        self.items = self.street_parts.keys()
        self.items.sort()
        self.editable = True
        self.enabled = True
        self.dropdownWidth = 22 * 'W'
        self.width = 22 * 'W'
        self.userInput = ''
        
    def onSelChange(self, selection):
        # Occurs when user selects a new value in the combo box.
        # selection is the user-selected value.
        
        self.userInput = selection.upper()
        parts = self.street_parts[self.userInput]
        title = 'Your Selection'
        message = 'You selected %s.\n\n\tPrefix: \t%s\n\tName: \t%s\n\tType: \t%s\n\tSuffix: \t%s' % \
                  (self.userInput, parts['prefix'], parts['name'], parts['type'], parts['suffix'])
        type = 0  # "OK only".
        pythonaddins.MessageBox(message, title, type)
    
    def onEditChange(self, text):
        # Occurs when a new character is typed into the combo box.
        # text is the user-entered text.
        # Only applicable when editable property is True.
        
        self.userInput = text.upper()
    
    def onEnter(self):
        # Occurs when user presses ENTER in the edit box of the combo box.
        # Enables tool to wait until user finishes typing a value before processing its business logic.
        # Only applicable when editable property is True.
        
        parts = self.street_parts.get(self.userInput)
        if parts:
            title = 'Your Selection'
            message = 'You selected %s.\n\n\tPrefix: \t%s\n\tName: \t%s\n\tType: \t%s\n\tSuffix: \t%s' % \
                      (self.userInput, parts['prefix'], parts['name'], parts['type'], parts['suffix'])
            type = 0  # "OK only".
            pythonaddins.MessageBox(message, title, type)
        else:
            title = 'Error'
            message = '"%s" is not a valid street name.\nPlease select a valid name from the dropdown list.' % self.userInput
            type = 0  # "OK only".
            pythonaddins.MessageBox(message, title, type)            
