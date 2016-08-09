#-------------------------------------------------------------------------------
# Name:         NNWW_Toolbars_addin.py
# Purpose:      Source code for the Python Add-Ins designed for the Newport News
#               Waterworks ArcMap Toolbars set
#
# Author:       Brian Kingery
#
# Created:      3/31/2016
# Copyright:    (c) bkingery 2016
#
# Directions:   User must run Update Local GIS Features to maximize add-in.
#
#               If used on virtual environment, files must be copied.
#               1. Run FileCopy_Desktop_H.py on Desktop
#               2. Run FileCopy_H_Virtual.py on Virtual Environment
#-------------------------------------------------------------------------------

import arcpy, pythonaddins, string, os, time, datetime
from arcpy import env
from os import *

try:
    import xlrd, xlwt
    from xlwt import *
    from xlrd import *
    from xlwt.Utils import *
except:
    pass

def SendStatusMail_Search():
    try:
        import smtplib
        mailServer      = 'birch.nnww.local'
        mailRecipients  = ['Brian Kingery <bkingery@nnva.gov>',
                           'Marietta Washington <mvwashington@nnva.gov>']
        mailSender = '%s <%s@nnva.gov>' % (os.environ['USERNAME'], os.environ['USERNAME'])
        subject = 'NNWW Toolbars Python Add-in'
        body  = '\nSearch\r\n\n'
        body += 'This is an automatically generated message.  Please do not reply.\r\n'
        message = ''
        message += 'From: %s\r\n' % mailSender
        message += 'To: %s\r\n' % ', '.join(mailRecipients)
        message += 'Subject: %s\r\n\r\n' % subject
        message += '%s\r\n' % body
        server = smtplib.SMTP(mailServer)
        server.sendmail(mailSender, mailRecipients, message)
        server.quit()
    except:
        pass

def SendStatusMail_MakeLayer():
    try:
        import smtplib
        mailServer      = 'birch.nnww.local'
        mailRecipients  = ['Brian Kingery <bkingery@nnva.gov>',
                           'Marietta Washington <mvwashington@nnva.gov>']
        mailSender = '%s <%s@nnva.gov>' % (os.environ['USERNAME'], os.environ['USERNAME'])
        subject = 'NNWW Toolbars Python Add-in'
        body  = '\nMake Layer\r\n\n'
        body += 'This is an automatically generated message.  Please do not reply.\r\n'
        message = ''
        message += 'From: %s\r\n' % mailSender
        message += 'To: %s\r\n' % ', '.join(mailRecipients)
        message += 'Subject: %s\r\n\r\n' % subject
        message += '%s\r\n' % body
        server = smtplib.SMTP(mailServer)
        server.sendmail(mailSender, mailRecipients, message)
        server.quit()
    except:
        pass

def SendStatusMail_Intersection():
    try:
        import smtplib
        mailServer      = 'birch.nnww.local'
        mailRecipients  = ['Brian Kingery <bkingery@nnva.gov>',
                           'Marietta Washington <mvwashington@nnva.gov>']
        mailSender = '%s <%s@nnva.gov>' % (os.environ['USERNAME'], os.environ['USERNAME'])
        subject = 'NNWW Toolbars Python Add-in'
        body  = '\nStreet Intersection\r\n\n'
        body += 'This is an automatically generated message.  Please do not reply.\r\n'
        message = ''
        message += 'From: %s\r\n' % mailSender
        message += 'To: %s\r\n' % ', '.join(mailRecipients)
        message += 'Subject: %s\r\n\r\n' % subject
        message += '%s\r\n' % body
        server = smtplib.SMTP(mailServer)
        server.sendmail(mailSender, mailRecipients, message)
        server.quit()
    except:
        pass

def SendStatusMail_Drawings():
    try:
        import smtplib
        mailServer      = 'birch.nnww.local'
        mailRecipients  = ['Brian Kingery <bkingery@nnva.gov>',
                           'Marietta Washington <mvwashington@nnva.gov>']
        mailSender = '%s <%s@nnva.gov>' % (os.environ['USERNAME'], os.environ['USERNAME'])
        subject = 'NNWW Toolbars Python Add-in'
        body  = '\nDrawings\r\n\n'
        body += 'This is an automatically generated message.  Please do not reply.\r\n'
        message = ''
        message += 'From: %s\r\n' % mailSender
        message += 'To: %s\r\n' % ', '.join(mailRecipients)
        message += 'Subject: %s\r\n\r\n' % subject
        message += '%s\r\n' % body
        server = smtplib.SMTP(mailServer)
        server.sendmail(mailSender, mailRecipients, message)
        server.quit()
    except:
        pass

def SendStatusMail_NewProject():
    try:
        import smtplib
        mailServer      = 'birch.nnww.local'
        mailRecipients  = ['Brian Kingery <bkingery@nnva.gov>',
                           'Marietta Washington <mvwashington@nnva.gov>']
        mailSender = '%s <%s@nnva.gov>' % (os.environ['USERNAME'], os.environ['USERNAME'])
        subject = 'NNWW Toolbars Python Add-in'
        body  = '\nNew Project Data Extractor\r\n\n'
        body += 'This is an automatically generated message.  Please do not reply.\r\n'
        message = ''
        message += 'From: %s\r\n' % mailSender
        message += 'To: %s\r\n' % ', '.join(mailRecipients)
        message += 'Subject: %s\r\n\r\n' % subject
        message += '%s\r\n' % body
        server = smtplib.SMTP(mailServer)
        server.sendmail(mailSender, mailRecipients, message)
        server.quit()
    except:
        pass

def SendStatusMail_Help():
    try:
        import smtplib
        mailServer      = 'birch.nnww.local'
        mailRecipients  = ['Brian Kingery <bkingery@nnva.gov>',
                           'Marietta Washington <mvwashington@nnva.gov>']
        mailSender = '%s <%s@nnva.gov>' % (os.environ['USERNAME'], os.environ['USERNAME'])
        subject = 'NNWW Toolbars Python Add-in'
        body  = '\nUser Guide\r\n\n'
        body += 'This is an automatically generated message.  Please do not reply.\r\n'
        message = ''
        message += 'From: %s\r\n' % mailSender
        message += 'To: %s\r\n' % ', '.join(mailRecipients)
        message += 'Subject: %s\r\n\r\n' % subject
        message += '%s\r\n' % body
        server = smtplib.SMTP(mailServer)
        server.sendmail(mailSender, mailRecipients, message)
        server.quit()
    except:
        pass

##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##################################################                                      ##################################################
##################################################      NNWW (Search GIS Features)      ##################################################
##################################################                                      ##################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################

dict = {
    ## Key : Value Pairs    Key = entry   Value = filepath
    ## Waterworks Feature Classes - Local Files
    'ADC Grid'          : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorInfrequent.mdb/Reference/ADCVirginiaPeninsulaGridArea',
    'Street'            : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/Transportation/Road',
    'Hydrants'          : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_Hydrant_EAM',
    'Valves'            : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_SystemValve_EAM',
    'Meters'            : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_wServiceLocation_EAM',
    'Water Laterals'    : 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/WaterUtility/v_wLateral_EDM'
        }

## IF dictNetwork is used, THEN all of the SQL query statements must be reformatted to match an ArcSDE connection rather than the personal geodatabase used for Local files
##dictNetwork = {
##    ## Key : Value Pairs    Key = entry   Value = filepath
##    ## Waterworks Feature Classes - Network Files
##    'ADC Grid'          : '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Conway_sdeVector_sdeViewer.sde/sdeVector.SDEDATAOWNER.Reference/sdeVector.SDEDATAOWNER.ADCVirginiaPeninsulaGridArea',
##    'Street'            : '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Conway_sdeVector_sdeViewer.sde/sdeVector.SDEDATAOWNER.Transportation/sdeVector.SDEDATAOWNER.Road',
##    'Hydrants'          : '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_HYDRANT_EAM',
##    'Valves'            : '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_SYSTEMVALVE_EAM',
##    'Meters'            : '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_WSERVICELOCATION_EAM',
##    'Water Laterals'    : '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_WLATERAL_EDM'
##        }

layerList = (
    ## Layers from Add-In in current map
    'ADC Grid (Python Add-In Layer)',
    'Street (Python Add-In Layer)',
    'Hydrants (Python Add-In Layer)',
    'Valves (Python Add-In Layer)',
    'Meters (Python Add-In Layer)',
    'Water Laterals (Python Add-In Layer)',

    'ADC Grid (Python Add-In Layer)_temp',
    'Street (Python Add-In Layer)_temp',
    'Hydrants (Python Add-In Layer)_temp',
    'Valves (Python Add-In Layer)_temp',
    'Meters (Python Add-In Layer)_temp',
    'Water Laterals (Python Add-In Layer)_temp',
    )

def clearSelectedFeatures():
    # Clear already selected features
    mxd = arcpy.mapping.MapDocument("Current")
    activeDF = mxd.activeView
    df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]
    for lyr in arcpy.mapping.ListLayers(mxd,'',df):
        if lyr.name in layerList:
            arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")

##########################################################################################

class ComboBoxClass_1_Layer(object):
    """Implementation for NNWW_Toolbars_addin.combobox_layer (ComboBox)"""
    def __init__(self):
        self.items = ["ADC Grid", "Street", "Hydrants", "Valves", "Meters", "Water Laterals"]
        self.editable = True
        self.enabled = True
        self.dropdownWidth = '12345678901234567890'
        self.width = '12345678901234567890'
    def onSelChange(self, selection):
        global CHOICE_layer
        CHOICE_layer = selection
        #print "Layer selected:", CHOICE_layer
        combobox_field.value = ''
        combobox_keyword.value = ''
        combobox_field.refresh()
        combobox_keyword.refresh()
    def onEditChange(self, text):
        CHOICE_layer = text
    def onFocus(self, focused):
        pass
    def onEnter(self):
        pass
    def refresh(self):
        pass

class ComboBoxClass_2_Field(object):
    """Implementation for NNWW_Toolbars_addin.combobox_field (ComboBox)"""
    def __init__(self):
        self.editable = True
        self.enabled = True
        self.dropdownWidth = '123456789012345678901234567890'
        self.width = '123456789012345678901234567890'
    def onSelChange(self, selection):
        global CHOICE_field
        CHOICE_field = selection
        #print "Field selected:", CHOICE_field
        combobox_keyword.value = ''
        combobox_keyword.refresh()
    def onEditChange(self, text):
        CHOICE_field = text
    def onFocus(self, focused):
        self.items = []
        try:
            if focused:
                if CHOICE_layer == 'ADC Grid':
                    self.items = ["OBJECTID", "MapNumber", "GridCell"]
                if CHOICE_layer == 'Street':
                    self.items = ["OBJECTID", "NAME\n               (Ex - JEFFERSON)", "NAME + TYPE\n (Ex - JEFFERSON AVE)"]
                if CHOICE_layer == 'Hydrants':
                    self.items = ["OBJECTID", "HYDRANT_NO"]
                if CHOICE_layer == 'Valves':
                    self.items = ["OBJECTID", "ValveID", "NUMBER_", "HYDRANT_NO", "ACCT_NO", "FILENAME"]
                if CHOICE_layer == 'Meters':
                    self.items = ["OBJECTID", "ServiceLocationID", "ConnectionObjectID", "Address", "City", "ADCMapNumber", "ADCGridCell", "WDSMapNumber", "WDSGridCell", "CustomerServiceUnit", "MeterReadingUnit", "ReadingSequenceNumber", "Portion", "ACCT_NO", "SerialNumber", "BusinessPartnerID", "ContractAccount", "LastName", "Telephone"]
                if CHOICE_layer == 'Water Laterals':
                    self.items = ["ServiceLocationID", "TapID", "Subtype", "LifecycleStatus", "QCStatus", "MainDiameter", "TapDiameter", "DrawingNumber", "ProjectNumber", "SDCStatus", "LegacyAccountNumber", "District"]
                else:
                    pass
        except:
            print 'Select a Layer before selecting a Field.'
    def onEnter(self):
        pass
    def refresh(self):
        pass

class ComboBoxClass_3_Keyword(object):
    """Implementation for NNWW_Toolbars_addin.combobox_keyword (ComboBox)"""
    def __init__(self):
        self.editable = True
        self.enabled = True
        self.dropdownWidth = ''
        self.width = '123456789012345678901234567890'
    def onSelChange(self, selection):
        pass
    def onEditChange(self, text):
        global CHOICE_keyword
        CHOICE_keyword = self.value
    def onFocus(self, focused):
        pass
    def onEnter(self):
        pass
    def refresh(self):
        pass

##########################################################################################

class ButtonClass_4_Search(object):
    """Implementation for NNWW_Toolbars_addin.button_search (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        try:
            if CHOICE_layer == 'Street' and CHOICE_field == "NAME\n               (Ex - JEFFERSON)":
                print "Layer selected:", CHOICE_layer
                print "Field selected:", "NAME"
                print "Keyword:       ", CHOICE_keyword
                print "Searching..."
                featureClass    = dict[CHOICE_layer]
                newLayer        = CHOICE_layer + ' (Python Add-In Layer)'

                mxd = arcpy.mapping.MapDocument("Current")
                activeDF = mxd.activeView
                df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]

                clearSelectedFeatures()

                entry       = string.split(CHOICE_keyword.upper())
                roadName    = entry[:]
                roadName    = ' '.join(roadName)
                try:
                    # Make layer from feature class
                    arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                    # Apply search query
                    where_clause = "[NAME] = '" + roadName + "'"

                    arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                    # Get count of how many records met query
                    layerCount      = arcpy.GetCount_management(newLayer)
                    count           = int(layerCount.getOutput(0))
                    if count != 0:
                        # Zoom to layer
                        df.zoomToSelectedFeatures()
                        print 'Matches found: ' + str(count)
                        del mxd, df
                    else:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        del mxd, df
                except:
                    print 'No matches found'
                    # Add a pop up if no results
                    pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                    del mxd, df

            ##################################################################################

            elif CHOICE_layer == 'Street' and CHOICE_field == "NAME + TYPE\n (Ex - JEFFERSON AVE)":
                print "Layer selected:", CHOICE_layer
                print "Field selected:", "NAME + TYPE"
                print "Keyword:       ", CHOICE_keyword
                print "Searching..."
                featureClass    = dict[CHOICE_layer]
                newLayer        = CHOICE_layer + ' (Python Add-In Layer)'

                mxd = arcpy.mapping.MapDocument("Current")
                activeDF = mxd.activeView
                df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]

                clearSelectedFeatures()

                entry       = string.split(CHOICE_keyword.upper())
                roadName    = entry[:-1]
                roadName    = ' '.join(roadName)
                roadType    = entry[-1]
                try:
                    # Make layer from feature class
                    arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                    # Apply search query
                    where_clause = "[NAME] = '" + roadName + "' AND [TYPE] = '" + roadType + "'"

                    arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                    # Get count of how many records met query
                    layerCount      = arcpy.GetCount_management(newLayer)
                    count           = int(layerCount.getOutput(0))
                    if count != 0:
                        # Zoom to layer
                        df.zoomToSelectedFeatures()
                        print 'Matches found: ' + str(count)
                        del mxd, df
                    else:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        del mxd, df
                except:
                    print 'No matches found'
                    # Add a pop up if no results
                    pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                    del mxd, df

            ##################################################################################

            else:
                print "Layer selected:", CHOICE_layer
                print "Field selected:", CHOICE_field
                print "Keyword:       ", CHOICE_keyword
                print "Searching..."
                featureClass    = dict[CHOICE_layer]
                newLayer        = CHOICE_layer + ' (Python Add-In Layer)'

                mxd = arcpy.mapping.MapDocument("Current")
                activeDF = mxd.activeView
                df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]

                clearSelectedFeatures()

                try:
                    # Make layer from feature class
                    arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                    delimField = arcpy.AddFieldDelimiters(featureClass, CHOICE_field)

                    # Apply search query
                    where_clause = """{0} = {1}""".format(delimField, CHOICE_keyword)
                    #where_clause    = "[" + CHOICE_field + "] = " + CHOICE_keyword
                    arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                    # Get count of how many records met query
                    layerCount      = arcpy.GetCount_management(newLayer)
                    count           = int(layerCount.getOutput(0))
                    if count != 0:
                        # Zoom to layer
                        df.zoomToSelectedFeatures()
                        print 'Matches found: ' + str(count)
                        del mxd, df
                    else:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        del mxd, df
                except:
                    try:
                        # Make layer from feature class
                        arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                        delimField = arcpy.AddFieldDelimiters(featureClass, CHOICE_field)

                        where_clause = """{0} = '{1}'""".format(delimField, CHOICE_keyword)
                        arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                        layerCount      = arcpy.GetCount_management(newLayer)
                        count           = int(layerCount.getOutput(0))
                        if count != 0:
                            df.zoomToSelectedFeatures()
                            print 'Matches found: ' + str(count)
                            del mxd, df
                        else:
                            print 'No matches found'
                            # Add a pop up if no results
                            pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                            del mxd, df
                    except:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        del mxd, df
        except:
            print 'Please enter all required fields'
        try:
            SendStatusMail_Search()
        except:
            pass


class ButtonClass_5_MakeLayer(object):
    """Implementation for NNWW_Toolbars_addin.button_makelayer (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        try:
            if CHOICE_layer == 'Street' and CHOICE_field == "NAME\n               (Ex - JEFFERSON)":
                print "Layer selected:", CHOICE_layer
                print "Field selected:", "NAME"
                print "Keyword:       ", CHOICE_keyword
                print "Searching..."
                featureClass    = dict[CHOICE_layer]
                newLayer        = CHOICE_layer + ' (Python Add-In Layer)_temp'

                mxd = arcpy.mapping.MapDocument("Current")
                activeDF = mxd.activeView
                df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]

                entry       = string.split(CHOICE_keyword.upper())
                roadName    = entry[:]
                roadName    = ' '.join(roadName)
                try:
                    # Make layer from feature class
                    arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                    # Apply search query
                    where_clause = "[NAME] = '" + roadName + "'"

                    arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                    # Get count of how many records met query
                    layerCount      = arcpy.GetCount_management(newLayer)
                    count           = int(layerCount.getOutput(0))
                    if count != 0:
                        # Zoom to layer
                        df.zoomToSelectedFeatures()
                        print 'Matches found: ' + str(count)

                        featureclass = dict[CHOICE_layer]
                        selection    = CHOICE_layer + ' Selection (' + roadName + ')'
                        where_clause = "[NAME] = '" + roadName + "'"
                        arcpy.MakeFeatureLayer_management(featureclass, selection, where_clause)
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == newLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
                    else:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == newLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
                except:
                    print 'No matches found'
                    # Add a pop up if no results
                    pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                    for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                        if lyr.name == newLayer:
                             arcpy.mapping.RemoveLayer(df, lyr)
                    del mxd, df

            ##################################################################################

            elif CHOICE_layer == 'Street' and CHOICE_field == "NAME + TYPE\n (Ex - JEFFERSON AVE)":
                print "Layer selected:", CHOICE_layer
                print "Field selected:", "NAME + TYPE"
                print "Keyword:       ", CHOICE_keyword
                print "Searching..."
                featureClass    = dict[CHOICE_layer]
                newLayer        = CHOICE_layer + ' (Python Add-In Layer)_temp'

                mxd = arcpy.mapping.MapDocument("Current")
                activeDF = mxd.activeView
                df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]

                entry       = string.split(CHOICE_keyword.upper())
                roadName    = entry[:-1]
                roadName    = ' '.join(roadName)
                roadType    = entry[-1]
                try:
                    # Make layer from feature class
                    arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                    # Apply search query
                    where_clause = "[NAME] = '" + roadName + "' AND [TYPE] = '" + roadType + "'"

                    arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                    # Get count of how many records met query
                    layerCount      = arcpy.GetCount_management(newLayer)
                    count           = int(layerCount.getOutput(0))
                    if count != 0:
                        # Zoom to layer
                        df.zoomToSelectedFeatures()
                        print 'Matches found: ' + str(count)

                        featureclass = dict[CHOICE_layer]
                        selection    = CHOICE_layer + ' Selection (' + roadName + ' ' + roadType + ')'
                        where_clause = "[NAME] = '" + roadName + "' AND [TYPE] = '" + roadType + "'"
                        arcpy.MakeFeatureLayer_management(featureclass, selection, where_clause)
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == newLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
                    else:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == newLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
                except:
                    print 'No matches found'
                    # Add a pop up if no results
                    pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                    for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                        if lyr.name == newLayer:
                             arcpy.mapping.RemoveLayer(df, lyr)
                    del mxd, df

            ##################################################################################

            else:
                print "Layer selected:", CHOICE_layer
                print "Field selected:", CHOICE_field
                print "Keyword:       ", CHOICE_keyword
                print "Searching..."
                featureClass    = dict[CHOICE_layer]
                newLayer        = CHOICE_layer + ' (Python Add-In Layer)_temp'

                mxd = arcpy.mapping.MapDocument("Current")
                activeDF = mxd.activeView
                df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]

                try:
                    # Make layer from feature class
                    arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                    delimField = arcpy.AddFieldDelimiters(featureClass, CHOICE_field)

                    # Apply search query
                    where_clause = """{0} = {1}""".format(delimField, CHOICE_keyword)
                    #where_clause    = "[" + CHOICE_field + "] = " + CHOICE_keyword
                    arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                    # Get count of how many records met query
                    layerCount      = arcpy.GetCount_management(newLayer)
                    count           = int(layerCount.getOutput(0))
                    if count != 0:
                        # Zoom to layer
                        df.zoomToSelectedFeatures()
                        print 'Matches found: ' + str(count)

                        featureclass = dict[CHOICE_layer]
                        selection    = CHOICE_layer + ' Selection (' + CHOICE_field + ' = ' + CHOICE_keyword + ')'
                        where_clause = """{0} = {1}""".format(delimField, CHOICE_keyword)
                        arcpy.MakeFeatureLayer_management(featureclass, selection, where_clause)
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == newLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
                    else:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == newLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
                except:
                    try:
                        # Make layer from feature class
                        arcpy.MakeFeatureLayer_management(featureClass, newLayer)
                        delimField = arcpy.AddFieldDelimiters(featureClass, CHOICE_field)

                        where_clause = """{0} = '{1}'""".format(delimField, CHOICE_keyword)
                        arcpy.SelectLayerByAttribute_management(newLayer, "NEW_SELECTION", where_clause)
                        layerCount      = arcpy.GetCount_management(newLayer)
                        count           = int(layerCount.getOutput(0))
                        if count != 0:
                            df.zoomToSelectedFeatures()
                            print 'Matches found: ' + str(count)

                            featureclass = dict[CHOICE_layer]
                            selection    = CHOICE_layer + ' Selection (' + CHOICE_field + ' = ' + CHOICE_keyword + ')'
                            where_clause = """{0} = '{1}'""".format(delimField, CHOICE_keyword)
                            arcpy.MakeFeatureLayer_management(featureclass, selection, where_clause)
                            for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                                if lyr.name == newLayer:
                                     arcpy.mapping.RemoveLayer(df, lyr)
                            del mxd, df
                        else:
                            print 'No matches found'
                            # Add a pop up if no results
                            pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                            for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                                if lyr.name == newLayer:
                                     arcpy.mapping.RemoveLayer(df, lyr)
                            del mxd, df
                    except:
                        print 'No matches found'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found", "NNWW Toolbars")
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == newLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
        except:
            print 'Please enter all required fields'
        try:
            SendStatusMail_MakeLayer()
        except:
            pass

class ButtonClass_6_RemoveLayers(object):
    """Implementation for NNWW_Toolbars_addin.button_removelayers (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        mxd = arcpy.mapping.MapDocument("Current")
        activeDF = mxd.activeView
        df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]
        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
            if lyr.name in layerList:
                 arcpy.mapping.RemoveLayer(df, lyr)
        for lyr in arcpy.mapping.ListLayers(mxd,'* Selection (*',df):
            arcpy.mapping.RemoveLayer(df, lyr)
        combobox_layer.value = ''
        combobox_field.value = ''
        combobox_keyword.value = ''
        combobox_layer.refresh()
        combobox_field.refresh()
        combobox_keyword.refresh()
        print 'All Python Add-In layers removed.'

##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##################################################                                      ##################################################
##################################################    NNWW (Find Street Intersection)   ##################################################
##################################################                                      ##################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################

class ComboBoxClass_1_StreetOne(object):
    """Implementation for NNWW_Toolbars_addin.combobox_streetone (ComboBox)"""
    def __init__(self):
        self.editable = True
        self.enabled = True
        self.dropdownWidth = ''
        self.width = '123456789012345678901234567890'
    def onSelChange(self, selection):
        pass
    def onEditChange(self, text):
        global street1
        street1 = self.value
    def onFocus(self, focused):
        pass
    def onEnter(self):
        pass
    def refresh(self):
        pass

class ComboBoxClass_2_StreetTwo(object):
    """Implementation for NNWW_Toolbars_addin.combobox_streettwo (ComboBox)"""
    def __init__(self):
        self.editable = True
        self.enabled = True
        self.dropdownWidth = ''
        self.width = '123456789012345678901234567890'
    def onSelChange(self, selection):
        pass
    def onEditChange(self, text):
        global street2
        street2 = self.value
    def onFocus(self, focused):
        pass
    def onEnter(self):
        pass
    def refresh(self):
        pass

class ButtonClass_3_FindPoint(object):
    """Implementation for NNWW_Toolbars_addin.button_findpoint (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):

        Street = 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Databases/sdeVectorFrequent.mdb/Transportation/Road'
        streetLayer = 'Add-In Street Layer'

        mxd = arcpy.mapping.MapDocument("Current")
        activeDF = mxd.activeView
        df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]
        try:
            entry1       = string.split(street1.upper())
            roadName1    = entry1[:-1]
            roadName1    = ' '.join(roadName1)
            roadType1    = entry1[-1]

            entry2       = string.split(street2.upper())
            roadName2    = entry2[:-1]
            roadName2    = ' '.join(roadName2)
            roadType2    = entry2[-1]

            global street1Entry
            global street2Entry
            street1Entry = roadName1 + ' ' + roadType1
            street2Entry = roadName2 + ' ' + roadType2

            for lyr in arcpy.mapping.ListLayers(mxd,'Add-In Str*',df):
                arcpy.mapping.RemoveLayer(df, lyr)
            for lyr in arcpy.mapping.ListLayers(mxd,'Intersection',df):
                arcpy.mapping.RemoveLayer(df, lyr)

            try:
                clearSelectedFeatures()
                # Make layer from Street feature class
                arcpy.MakeFeatureLayer_management(Street, streetLayer)
                # Apply search query
                where_clause = "[NAME] = '" + roadName1 + "' AND [TYPE] = '" + roadType1 + "'"
                arcpy.SelectLayerByAttribute_management(streetLayer, "NEW_SELECTION", where_clause)
                # Get count of how many records met query
                layerCount      = arcpy.GetCount_management(streetLayer)
                count           = int(layerCount.getOutput(0))
                if count != 0:
                    print 'Add-In Street 1: ' + street1Entry
                    street1Selection = 'Add-In Street 1 - ' + street1Entry
                    arcpy.MakeFeatureLayer_management(Street, street1Selection, where_clause)
                    clearSelectedFeatures()
                    try:
                        # Apply search query
                        where_clause = "[NAME] = '" + roadName2 + "' AND [TYPE] = '" + roadType2 + "'"
                        arcpy.SelectLayerByAttribute_management(streetLayer, "NEW_SELECTION", where_clause)
                        # Get count of how many records met query
                        layerCount2      = arcpy.GetCount_management(streetLayer)
                        count2           = int(layerCount2.getOutput(0))
                        if count2 != 0:
                            print 'Add-In Street 2: ' + street2Entry
                            street2Selection = 'Add-In Street 2 - ' + street2Entry
                            arcpy.MakeFeatureLayer_management(Street, street2Selection, where_clause)
                            for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                                if lyr.name == streetLayer:
                                     arcpy.mapping.RemoveLayer(df, lyr)
                            #print 'Finding Intersection: ' + street1Entry + ' & ' + street2Entry
                            try:
                                arcpy.env.workspace = "in_memory"

                                global StreetIntersectionPoint
                                StreetIntersectionPoint = 'Intersection'

                                streetList = []
                                for lyr in arcpy.mapping.ListLayers(mxd,'Add-In Str*',df):
                                    streetList.append(lyr.name)

                                arcpy.Intersect_analysis(streetList, StreetIntersectionPoint, "ALL", "", "POINT")
                                del streetList

                                for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                                    if lyr.name == StreetIntersectionPoint:
                                        arcpy.SelectLayerByAttribute_management(lyr,"NEW_SELECTION")
                                        # Get count of how many records met query
                                        layerCount3      = arcpy.GetCount_management(lyr)
                                        count3           = int(layerCount3.getOutput(0))
                                        if count3 != 0:
                                            # Turn on Street Name Labels
                                            for lyr in arcpy.mapping.ListLayers(mxd,'Add-In Str*',df):
                                                lyr.showLabels = True
                                                arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")

                                            # Zoom to layer
                                            df.zoomToSelectedFeatures()

                                            # Update Symbology
                                            sourceLayerStreet1          = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Street1Symbology.lyr')
                                            sourceLayerStreet2          = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Street2Symbology.lyr')
                                            sourceLayerIntersection     = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Intersection.lyr')
                                            for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                                                if lyr.name == street1Selection:
                                                    updateLayer = lyr
                                                    arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayerStreet1, True)
                                                if lyr.name == street2Selection:
                                                    updateLayer = lyr
                                                    arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayerStreet2, True)
                                                if lyr.name == StreetIntersectionPoint:
                                                    updateLayer = lyr
                                                    arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayerIntersection, True)
                                            arcpy.RefreshTOC()
                                            arcpy.RefreshActiveView()
                                            clearSelectedFeatures()
                                            print 'Intersection located'
                                        else:
                                            arcpy.mapping.RemoveLayer(df, lyr)
                                            print 'The two roads do not intersect'
                                    # These 2 IF statements remove the 2 roads. Comment out if it is decided to keep layers
##                                    if lyr.name == street1Selection:
##                                         arcpy.mapping.RemoveLayer(df, lyr)
##                                    if lyr.name == street2Selection:
##                                         arcpy.mapping.RemoveLayer(df, lyr)
                                del mxd, df
                            except:
                                print 'Intersection finder failed'
                                # Add a pop up if no results
                                pythonaddins.MessageBox("Finding intersection failed", "NNWW Toolbars")
                        else:
                            for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                                if lyr.name == streetLayer:
                                     arcpy.mapping.RemoveLayer(df, lyr)
                                if lyr.name == street1Selection:
                                     arcpy.mapping.RemoveLayer(df, lyr)
                            del mxd, df
                            print 'No matches found for Street 2'
                            # Add a pop up if no results
                            pythonaddins.MessageBox("No matches found for Street 2", "NNWW Toolbars")
                    except:
                        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                            if lyr.name == streetLayer:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                            if lyr.name == street1Selection:
                                 arcpy.mapping.RemoveLayer(df, lyr)
                        del mxd, df
                        print 'No matches found for Street 2'
                        # Add a pop up if no results
                        pythonaddins.MessageBox("No matches found for Street 2", "NNWW Toolbars")
                else:
                    for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                        if lyr.name == streetLayer:
                             arcpy.mapping.RemoveLayer(df, lyr)
                    del mxd, df
                    print 'No matches found for Street 1'
                    # Add a pop up if no results
                    pythonaddins.MessageBox("No matches found for Street 1", "NNWW Toolbars")
            except:
                for lyr in arcpy.mapping.ListLayers(mxd,'',df):
                    if lyr.name == streetLayer:
                         arcpy.mapping.RemoveLayer(df, lyr)
                del mxd, df
                print 'No matches found for Street 1'
                # Add a pop up if no results
                pythonaddins.MessageBox("No matches found for Street 1", "NNWW Toolbars")
        except:
            print 'Please enter two valid street names'
##            # Add a pop up if no results
##            pythonaddins.MessageBox("Please enter two valid street names", "NNWW Toolbars")
        try:
            SendStatusMail_Intersection()
        except:
            pass

class ButtonClass_4_RemovePoint(object):
    """Implementation for NNWW_Toolbars_addin.button_removepoint (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        combobox_streetone.value = ''
        combobox_streettwo.value = ''
        combobox_streetone.refresh()
        combobox_streettwo.refresh()

        mxd = arcpy.mapping.MapDocument("Current")
        activeDF = mxd.activeView
        df = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]
        for lyr in arcpy.mapping.ListLayers(mxd,'Add-In Str*',df):
            arcpy.mapping.RemoveLayer(df, lyr)
        for lyr in arcpy.mapping.ListLayers(mxd, 'Intersection', df):
            arcpy.mapping.RemoveLayer(df, lyr)
        #arcpy.Delete_management(arcpy.env.workspace)
        print 'Street intersection layers removed'

##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##################################################                                      ##################################################
##################################################            NNWW (Drawings)           ##################################################
##################################################                                      ##################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################

class ButtonClass_1_AddCADIndex(object):
    """Implementation for NNWW_Toolbars_addin.button_addcadindex (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        mxd         = arcpy.mapping.MapDocument("Current")
        activeDF    = mxd.activeView
        df          = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]
        for lyr in arcpy.mapping.ListLayers(mxd, 'CAD Index', df):
            arcpy.mapping.RemoveLayer(df, lyr)
        cadIndex = 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Shapefiles/Drawings/cadindex.shp'
        cadIndexLayer = 'CAD Index'
        arcpy.MakeFeatureLayer_management(cadIndex, cadIndexLayer)
        # Update Symbology
        CADIndexLayer = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/CADIndex.lyr')
        for lyr in arcpy.mapping.ListLayers(mxd,'',df):
            if lyr.name == cadIndexLayer:
                updateLayer = lyr
                arcpy.mapping.UpdateLayer(df, updateLayer, CADIndexLayer, True)
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
        del mxd, df
        print 'CAD Index layer added.'
        try:
            SendStatusMail_Drawings()
        except:
            pass

class ToolClass_2_AddDrawings(object):
    """Implementation for NNWW_Toolbars_addin.tool_adddrawings (Tool)"""
    def __init__(self):
        self.enabled    = True
        self.shape      = "Rectangle"
        self.cursor     = 3
    def onRectangle(self, rectangle_geometry):
        arcpy.env.workspace = "in_memory"

        extent = rectangle_geometry
        projectArea = 'SearchBox'
        ProjectAreaPolygon = arcpy.CreateFishnet_management(projectArea,
                                '%f %f' %(extent.XMin, extent.YMin),
                                '%f %f' %(extent.XMin, extent.YMax),
                                0, 0, 1, 1,
                                '%f %f' %(extent.XMax, extent.YMax),'NO_LABELS',
                                '%f %f %f %f' %(extent.XMin, extent.YMin, extent.XMax, extent.YMax), 'POLYGON')

        # Make reference layer for CAD drawings
        EngineeringDrawingReference = 'C:/GIS/OperationalSystems/Locations/Waterworks BaseMap/Data/Shapefiles/Drawings/cadindex.shp'
        EngineeringDrawingReferenceLayer = 'Engineering Drawing Reference'
        arcpy.MakeFeatureLayer_management(EngineeringDrawingReference, EngineeringDrawingReferenceLayer)

        # Select all drawings that intersect the search box
        arcpy.SelectLayerByLocation_management(EngineeringDrawingReferenceLayer, 'intersect', projectArea)

        mxd         = arcpy.mapping.MapDocument("Current")
        activeDF    = mxd.activeView
        df          = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]
        matchcount = int(arcpy.GetCount_management(EngineeringDrawingReferenceLayer).getOutput(0))
        for lyr in arcpy.mapping.ListLayers(mxd, '', df):
            if lyr.name == 'SearchBox':
                arcpy.mapping.RemoveLayer(df, lyr)
            if lyr.name == 'Engineering Drawing Reference':
                arcpy.mapping.RemoveLayer(df, lyr)
            if lyr.name == 'Drawings & Historical Orthophotographs':
                arcpy.mapping.RemoveLayer(df, lyr)
        if matchcount == 0:
            print 'No drawings located within search box'
            # Add a pop up if no results
            pythonaddins.MessageBox("No drawings located within search box", "NNWW Toolbars")
            arcpy.RefreshTOC()
            arcpy.RefreshActiveView()
        else:
            selected_drawings = []
            with arcpy.da.SearchCursor(EngineeringDrawingReferenceLayer,'FILENAME') as cursor:
                for row in cursor:
                    selected_drawings.append(row)
            del cursor

            print 'Files found within search box: ' + str(matchcount) + '\n'
            print 'Please wait while files are processed...\n'

            ## Start
            ExecutionStartTime = datetime.datetime.now()
            print "Started: %s\n" % ExecutionStartTime.strftime('%A, %B %d, %Y %I:%M:%S %p')

            dwgCount = 0
            dwgCountFail = 0
            dxfCount = 0
            dxfCountFail = 0
            tifCount = 0
            tifCountFail = 0
            totalCount = 0

            dxfFiles = ['Annotation', 'MultiPatch', 'Point', 'Polygon', 'Polyline']

            for drawing in selected_drawings:
                drawing = str(drawing)
                drawing = drawing.replace("(u'","")
                drawing = drawing.replace("',)","")
                a = drawing.split('\\')
                title = a[-1].split('.')
                fileName = title[0].upper()
                fileType = title[1].upper()

                if fileType == 'DWG':
                    try:
                        print 'DWG Added: ' + a[-3] + '/' + a[-1].upper()
                        TEMPLayer           = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/tempGroup.lyr')
                        arcpy.mapping.AddLayer(df, TEMPLayer, "BOTTOM")
                        group_layer         = arcpy.mapping.ListLayers(mxd, 'tempGroup', df)[0]
                        group_layer.name    = a[-3] + '/' + fileName + '.' + fileType + ' -' # is the identifier for the layer
                        group_layer.visible = False
                        try:
                            for DWG in dwgFiles:
                                #DWGfile = 'S:/Docs/' + a[-3] + '/' + title[0] + '.' + title[1] + '/' + DWG
                                DWGfile = '//onyx/amw/Docs/' + a[-3] + '/' + title[0] + '.' + title[1] + '/' + DWG
                                arcpy.MakeFeatureLayer_management(DWGfile, DWG + ' (' + fileName + '.' + fileType + ')')
                                #print '                               /' + DXF + ' (' + fileName + '.' + fileType + ')'

                                for lyr in arcpy.mapping.ListLayers(mxd, DWG + ' (' + fileName + '.' + fileType + ')', df):
                                    moveLayer = lyr
                                    arcpy.mapping.AddLayerToGroup(df, group_layer, moveLayer, "BOTTOM")
                                    arcpy.mapping.RemoveLayer(df, lyr)
                            dwgCount += 1
                            totalCount += 1
##                            arcpy.RefreshTOC()
##                            arcpy.RefreshActiveView()
                        except:
                            print 'DWG File adding error' + ' <---------ERROR---------'
                            for lyr in arcpy.mapping.ListLayers(mxd, group_layer, df):
                                arcpy.mapping.RemoveLayer(df, lyr)
                            dwgCountFail += 1
                            totalCount += 1
                    except:
                        print a[-3] + '/' + a[-1].upper() + ' <---------ERROR---------'
                        dwgCountFail += 1
                        totalCount += 1

                elif fileType == 'DXF':
                    try:
                        print 'DXF Added: ' + a[-3] + '/' + a[-1].upper()
                        TEMPLayer           = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/tempGroup.lyr')
                        arcpy.mapping.AddLayer(df, TEMPLayer, "BOTTOM")
                        group_layer         = arcpy.mapping.ListLayers(mxd, 'tempGroup', df)[0]
                        group_layer.name    = a[-3] + '/' + fileName + '.' + fileType + ' -' # is the identifier for the layer
                        group_layer.visible = False
                        try:
                            for DXF in dxfFiles:
                                #DXFfile = 'S:/Docs/' + a[-3] + '/' + title[0] + '.' + title[1] + '/' + DXF
                                DXFfile = '//onyx/amw/Docs/' + a[-3] + '/' + title[0] + '.' + title[1] + '/' + DXF
                                arcpy.MakeFeatureLayer_management(DXFfile, DXF + ' (' + fileName + '.' + fileType + ')')
                                #print '                               /' + DXF + ' (' + fileName + '.' + fileType + ')'

                                for lyr in arcpy.mapping.ListLayers(mxd, DXF + ' (' + fileName + '.' + fileType + ')', df):
                                    moveLayer = lyr
                                    arcpy.mapping.AddLayerToGroup(df, group_layer, moveLayer, "BOTTOM")
                                    arcpy.mapping.RemoveLayer(df, lyr)
                            dxfCount += 1
                            totalCount += 1
##                            arcpy.RefreshTOC()
##                            arcpy.RefreshActiveView()
                        except:
                            print 'DXF File adding error' + ' <---------ERROR---------'
                            for lyr in arcpy.mapping.ListLayers(mxd, group_layer, df):
                                arcpy.mapping.RemoveLayer(df, lyr)
                            dxfCountFail += 1
                            totalCount += 1
                    except:
                        print a[-3] + '/' + a[-1].upper() + ' <---------ERROR---------'
                        dxfCountFail += 1
                        totalCount += 1

                elif fileType == 'TIF':
                    try:
                        print 'TIF Added: ' + a[-3] + '/' + a[-1].upper()
                        arcpy.MakeRasterLayer_management(drawing, drawing[:-4])
                        for lyr in arcpy.mapping.ListLayers(mxd, drawing[:-4], df):
                            lyr.name = a[-3] + '/' + a[-1].upper() + ' -' # is the identifier for the layer
                            lyr.visible = False
##                            arcpy.RefreshTOC()
##                            arcpy.RefreshActiveView()
                        tifCount += 1
                        totalCount += 1
                    except:
                        print a[-3] + '/' + a[-1].upper() + ' <---------ERROR---------'
                        tifCountFail += 1
                        totalCount += 1
            del selected_drawings

##            arcpy.RefreshTOC()
##            arcpy.RefreshActiveView()

            if dwgCount > 0 or dxfCount > 0 or tifCount > 0:
                temporaryGL = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/temp.lyr')
                arcpy.mapping.AddLayer(df, temporaryGL, "BOTTOM")
                MainGroupLayer = arcpy.mapping.Layer('//amazon/waterworks/Divisions/InfoTech/Shared/GIS/Python_Addins/addin_assistant/Projects/MyTools/NNWW_Toolbars/ProjectFiles/Drawings & Historical Orthophotographs.lyr')
                arcpy.mapping.AddLayer(df, MainGroupLayer, "TOP")

                temporaryGroupLayer   = arcpy.mapping.ListLayers(mxd, 'temp', df)[0]
                targetGroupLayer      = arcpy.mapping.ListLayers(mxd, 'Drawings & Historical Orthophotographs', df)[0]

                for lyr in arcpy.mapping.ListLayers(mxd, '* -',df):
                    moveLayer = lyr
                    arcpy.mapping.AddLayerToGroup(df, temporaryGroupLayer, moveLayer, "BOTTOM")
                    arcpy.mapping.RemoveLayer(df, lyr)

                ABC_Order = []
                for lyr in temporaryGroupLayer:
                    ABC_Order.append(lyr.name)

                for x in sorted(ABC_Order):
                    arcpy.mapping.AddLayerToGroup(df, targetGroupLayer, arcpy.mapping.ListLayers(mxd,x)[0],"BOTTOM")
                arcpy.mapping.RemoveLayer(df, temporaryGroupLayer)

                del ABC_Order

            for x in targetGroupLayer:
                y = x.longName.split()
                z = y[0]
                if z in dxfFiles:
                    arcpy.mapping.RemoveLayer(df,x)

            arcpy.RefreshTOC()
            arcpy.RefreshActiveView()

            ## Finish
            ExecutionEndTime = datetime.datetime.now()
            ElapsedTime = ExecutionEndTime - ExecutionStartTime
            print "\nEnded: %s\n" % ExecutionEndTime.strftime('%A, %B %d, %Y %I:%M:%S %p')
            print "Elapsed Time: %s" % str(ElapsedTime).split('.')[0]

            print ''
            print 'DWG        =',str(dwgCount)
            print 'DXF        =',str(dxfCount)
            print 'TIF        =',str(tifCount)
            print ''
            print 'DWG Errors =',str(dwgCountFail)
            print 'DXF Errors =',str(dxfCountFail)
            print 'TIF Errors =',str(tifCountFail)
            print ''
            print 'Total      =',str(totalCount)

            # Add a pop up message
            # No Errors
            if dwgCountFail == 0 and dxfCountFail == 0 and tifCountFail == 0:
                pythonaddins.MessageBox('\nTotal Files Processed\n' + str(totalCount) + '\n\nDWG: \t' + str(dwgCount) + '\nDXG: \t' + str(dxfCount) + '\nTIF: \t' + str(tifCount) + '\n\nElapsed Time\n%s\n' % str(ElapsedTime).split('.')[0], "NNWW Toolbars")
            # Errors
            if tifCountFail > 0 or dwgCountFail > 0 or dxfCountFail > 0:
                pythonaddins.MessageBox('\nTotal Files Processed\n' + str(totalCount) + '\n\nDWG: \t' + str(dwgCount) + '\nDXF: \t' + str(dxfCount) + '\nTIF: \t' + str(tifCount) + '\n\nErrors\n' + 'DWG: \t' + str(dwgCountFail) + '\nDXF: \t' + str(dxfCountFail) + '\nTIF: \t' + str(tifCountFail) + '\n\nElapsed Time\n%s\n' % str(ElapsedTime).split('.')[0], "NNWW Toolbars")
        try:
            SendStatusMail_Drawings()
        except:
            pass

class ButtonClass_3_RemoveDrawings(object):
    """Implementation for NNWW_Toolbars_addin.button_removedrawings (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        mxd         = arcpy.mapping.MapDocument("Current")
        activeDF    = mxd.activeView
        df          = arcpy.mapping.ListDataFrames(mxd, activeDF)[0]
        for lyr in arcpy.mapping.ListLayers(mxd, 'Drawings & Historical Orthophotographs', df):
            arcpy.mapping.RemoveLayer(df, lyr)
        for lyr in arcpy.mapping.ListLayers(mxd, 'CAD Index', df):
            arcpy.mapping.RemoveLayer(df, lyr)
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
        del mxd, df
        #arcpy.Delete_management(arcpy.env.workspace)
        print 'All drawings removed'

##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##################################################                                      ##################################################
##################################################          NNWW (User Guide)           ##################################################
##################################################                                      ##################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################

class ButtonClass_1_HelpDoc(object):
    """Implementation for NNWW_Toolbars_addin.button_helpdoc (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        print 'User Guide located at:\n\nR:\Divisions\InfoTech\Shared\GIS\NNWW_Toolbars_UserGuide.pdf'

        # Add a pop up
        pythonaddins.MessageBox('\n\t\tNewport News Waterworks\t\t\n\t\t      Toolbars User Guide\n\n\t\t            Located at:\n\n\t             R:/Divisions/InfoTech/Shared/GIS/\n\t               NNWW_Toolbars_UserGuide.pdf\n', "NNWW Toolbars")

        try:
            SendStatusMail_Help()
        except:
            pass

        #helpDoc = r'R:\Divisions\InfoTech\Shared\GIS\Python_Addins\addin_assistant\Projects\MyTools\NNWW_Toolbars\ProjectFiles\NNWW_Toolbar_UserGuide.pdf'
        #os.startfile(helpDoc)
        ###   OR   ###
        #webbrowser.open(helpDoc)

##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##################################################                                      ##################################################
##################################################  NNWW (New Project Data Extractor)   ##################################################
##################################################                                      ##################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################

class ToolClass_1_DataExtractor(object):
    """Implementation for NNWW_Toolbars_addin.tool_dataextractor (Tool)"""
    def __init__(self):
        self.enabled = True
        self.shape = "Rectangle"
        self.cursor = 3

    def onRectangle(self, rectangle_geometry):
        try:
            env.workspace = "//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/ExtractionFolder"
            env.overwriteoutput = True

            ## Get name of dataset by using timestamp as title
            timesnapshot = time.ctime()
            time_str = str(timesnapshot)
            time_list = string.split(time_str)
            month       = time_list[1]
            day         = time_list[2]
            year        = time_list[4]
            timestamp   = time_list[3]
            timestamp   = timestamp.replace(":", "")
            #timestamp   = int(timestamp)
            if timestamp[0:2] == '12':
                timestamp = timestamp[0:4] + 'pm'
            elif int(timestamp[0:2]) >= 13 and int(timestamp[0:2]) <= 24:
                x = str(int(timestamp[0:2]) - 12)
                timestamp = x + timestamp[2:4] + 'pm'
            elif timestamp[0:2] == '10' or timestamp[0:2] == '11':
                timestamp = timestamp[0:4] + 'am'
            else:
                timestamp = timestamp[1:4] + 'am'
            #title = day+month+year +"_"+ timestamp
            title = year+month+day +"_"+ timestamp
            print "Your project folder title is " + title
            print "The folder is located at location --> \\amazon\waterworks\Divisions\InfoTech\Shared\GIS\NewProjectApplication\ExtractionFolder"
            print "Please wait while data compiles..."

            event = str(title)
            projectFolder = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/ExtractionFolder/' + event

            ## Create New Project Folder
            if not os.path.exists(projectFolder):
                os.makedirs(projectFolder)

            ## Create New File GDB
            gdb = event + ".gdb"
            arcpy.CreateFileGDB_management(projectFolder, gdb)

            del env.workspace
            env.workspace = projectFolder +"/"+ gdb

            ## Make new dataset
            ## http://resources.arcgis.com/en/help/arcgis-rest-api/index.html#/Projected_coordinate_systems/02r3000000vt000000/
            ## NAD_1983_StatePlane_Virginia_South_FIPS_4502_Feet
            dataset = "Dataset"
            sr = arcpy.SpatialReference(2284)
            arcpy.CreateFeatureDataset_management(env.workspace, dataset, sr)

            del env.workspace
            env.workspace = projectFolder +"/"+ gdb +"/"+ dataset

            ## Make copy of polygon feature class

            extent = rectangle_geometry
            # Create a fishnet with 1 row and 1 columns.
            projectarea = 'ProjectAreaPolygon'
            ProjectAreaPolygon = arcpy.CreateFishnet_management(projectarea,
                                    '%f %f' %(extent.XMin, extent.YMin),
                                    '%f %f' %(extent.XMin, extent.YMax),
                                    0, 0, 1, 1,
                                    '%f %f' %(extent.XMax, extent.YMax),'NO_LABELS',
                                    '%f %f %f %f' %(extent.XMin, extent.YMin, extent.XMax, extent.YMax), 'POLYGON')
            #arcpy.RefreshActiveView()
            #return ProjectAreaPolygon


            #########################################################################


            datasetPolygon = projectarea
            datasetPolygonLayer = datasetPolygon + "_lyr"
            ## Have to make a layer to use SelectLayerByLocation_management
            arcpy.MakeFeatureLayer_management(datasetPolygon, datasetPolygonLayer)

            ##Used to find out what fields are in the layer
            ##
            ##fieldList = arcpy.ListFields(""""layer"""")
            ##for field in fieldList:
            ##    print field.name


            ##
            ## Hydrant
            ##

            hydrant = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_HYDRANT_EAM'
            hydrantLayer = "V_HYDRANT_EAM_lyr"
            arcpy.MakeFeatureLayer_management(hydrant, hydrantLayer)
            arcpy.SelectLayerByLocation_management(hydrantLayer, 'intersect', datasetPolygonLayer)
            # If features matched criteria write them to a new feature class
            matchcount = int(arcpy.GetCount_management(hydrantLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                ## Make copy of intersected features
                infc = hydrantLayer
                outfc = "Hydrant"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["STATUS","SYMBOLCD","GPS_X_COORD","GPS_Y_COORD","FILENAME","COLL_CODE","OPER_MOVED","POLYGONID",
                          "SCALE","ANGLE","PIPESIZC","PIPESIZN","OWNER","OPERDIST","DEPBURYC","DEPBURYN","VLOCC","VLOCN",
                          "LOCDIR","BARRELSZ","VALVESZ","OPENCODE","TURNSRQD","HYDRCONN","MAP_","GROUP_NO","SEQ_NO",
                          "N_DRAWING","DATEINSP","DATE_COMPLETED","TIME_24HR","STAT_PRESS","RESID_PRESS","AT_20_PSI","REMARKS",
                          "DISTRICT","HYDRANT_SEQ_NO"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Mainbreak
            ##
            Mainbreak = '//nile/applications/Databases/Engineering/Data/Distribution_GeoDB.mdb/Mainbreaks'
            MainbreakLayer = "MainBrks_lyr"
            arcpy.MakeFeatureLayer_management(Mainbreak, MainbreakLayer)
            arcpy.SelectLayerByLocation_management(MainbreakLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(MainbreakLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = MainbreakLayer
                outfc = "Mainbreak"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["OBJECT_","TIMEOUT","SIZE_","TYPE","TOT_LABR","TOT_PART","TOT_SUBL","TOT_TOOL","DATEOUT","LOCLEAK",
                          "TYPELK","SOILTY","SUBTYACT","COMMT","PLYWRAP","PAVINGWO","PAVINGCST","ID_NO","TOTREPCS","COST_MILE",
                          "Key5","abanProjNu"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Valve
            ##
            valve = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_SYSTEMVALVE_EAM'
            valveLayer = "V_SYSTEMVALVE_EAM_lyr"
            arcpy.MakeFeatureLayer_management(valve, valveLayer)
            arcpy.SelectLayerByLocation_management(valveLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(valveLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = valveLayer
                outfc = "Valve"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["ValveID","ACCT_NO","STATUS","SYMBOLCD","GPS_X_COORD","GPS_Y_COORD","FILENAME","COLL_CODE","OPER_MOVED","POLYGONID",
                          "SCALE","ANGLE","SEQ_NO","SYSTEM","END_CONNECT","MAIN_SIZE","PIPE_SEQ_NO","PAR_PIPE_NO","STATUS_CODE","MSUTIL_PAGE",
                          "MSUTIL_GRID","WWMAPS_PAGE","WWMAPS_GRID","OPENS_CODE","TURNS","SET_IN","BELOW_ST_GR","INSPECT_FREQ","GEARING_CODE",
                          "BYPASS_CODE","BYPASS_SIZE","BYPASS_TURNS","BOX_TYPE","DEPTH_TO_NUT","INSP_GROUP","INSP_SEQ_NO","TYPE_CODE","SUFFIX",
                          "LOC_FEET","LOC_DIR","LOC_FROM","ST_NAME_1","ST_NAME_2","ST_NAME_3","STREET_TYPE","FROM_STREET","XLOC_FEET","XLOC_DIR",
                          "XLOC_FROM","XST_NAME_1","XST_NAME_2","XST_NAME_3","XSUF","TO_STREET","DISTRICT","SEGMENT_NO","REPLACE_DATE","LAST_CHG_DATE",
                          "LAST_CHG_USER","LAST_CHG_PROG","PURGE_FLAG","REMARKS","REMARKS2"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Abandoned Pipe
            ##
            abandonedPipe = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_WPRESSURIZEDMAINABANDONED'
            abandonedPipeLayer = "V_WPRESSURIZEDMAINABANDONED_lyr"
            arcpy.MakeFeatureLayer_management(abandonedPipe, abandonedPipeLayer)
            arcpy.SelectLayerByLocation_management(abandonedPipeLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(abandonedPipeLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = abandonedPipeLayer
                outfc = "AbandonedPipe"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["ProjectID", "ProjectType","PrimaryReference","SecondaryReference","Subsystem","LifecycleStatus","QCStatus","WorkingPressure",
                          "CFactor","LiningType","WireClass","CathodicProtection","ExteriorCoating","JointType","GroundSurfaceType","OriginStation",
                          "DestinationStation","Owner","Comment","Length3D","LegacyObjectID","AbandonedProjectID"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Main
            ##
            main = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_WPRESSURIZEDMAIN'
            mainLayer = "V_WPRESSURIZEDMAIN_lyr"
            arcpy.MakeFeatureLayer_management(main, mainLayer)
            arcpy.SelectLayerByLocation_management(mainLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(mainLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = mainLayer
                outfc = "Main"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["FacilityID","SecondaryReference","LifecycleStatus","QCStatus","Manufacturer","WorkingPressure","CFactor",
                          "WireClass","CathodicProtection","ExteriorCoating","JointType","GroundSurfaceType","OriginStation","DestinationStation",
                          "AbandonedDate","AbandonedReference","ReplacementDate","ReplacementDiameter","RetirementNumber","Owner","LineType",
                          "Comment","Length3D"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Waterworks Easements
            ##
            easement = '//nile/applications/Databases/Engineering/Data/Distribution_GeoDB.mdb/Easements'
            easementLayer = "WaterworksEasement_lyr"
            arcpy.MakeFeatureLayer_management(easement, easementLayer)
            arcpy.SelectLayerByLocation_management(easementLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(easementLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = easementLayer
                outfc = "Easement"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                #No fields to delete
            ##
            ## Fiber Cable
            ##
            fiber = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/sdeVector.SDEDATAOWNER.OtherUtility/sdeVector.SDEDATAOWNER.FiberOpticLine'
            fiberLayer = "FiberOpticLine_lyr"
            arcpy.MakeFeatureLayer_management(fiber, fiberLayer)
            arcpy.SelectLayerByLocation_management(fiberLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(fiberLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = fiberLayer
                outfc = "FiberCable"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                #No fields to delete
            ##
            ## Fireflow
            ##
            fireflow = '//nile/applications/Databases/Engineering/Data/Distribution_GeoDB.mdb/Fireflows'
            fireflowLayer = "FIREFLOW_lyr"
            arcpy.MakeFeatureLayer_management(fireflow, fireflowLayer)
            arcpy.SelectLayerByLocation_management(fireflowLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(fireflowLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = fireflowLayer
                outfc = "Fireflow"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                #No fields to delete
            ##
            ## DSI Prioritized Area
            ##
            dsi = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/sdeVector.SDEDATAOWNER.DSIPriority/sdeVector.sdeDataOwner.DSIPrioritizedArea'
            dsiLayer = "DSIPrioritizedArea_lyr"
            arcpy.MakeFeatureLayer_management(dsi, dsiLayer)
            arcpy.SelectLayerByLocation_management(dsiLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(dsiLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = dsiLayer
                outfc = "DSIPrioritizedArea"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["AreaID","PavementConstructionCategory","Comment","PlannedPavingDate","PlannedCompleteDate","CompleteDate",
                          "FireDeptForm","MsUtilityDate","MsUtilityTicket","DesignMeetingDate","PaybackYears","LegacyObjectID","ProjectSource",
                          "aProjectWeight","LifeExpectancy","ImprovementPointsAge","ImprovementPointsBreak","ImprovementPointsCost",
                          "ImprovementPointsOther","DataCalculationDate"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Current Project
            ##
            currentProject = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_CONSTRUCTIONPROJECT_EAM'
            currentProjectLayer = "V_CONSTRUCTIONPROJECT_EAM_lyr"
            arcpy.MakeFeatureLayer_management(currentProject, currentProjectLayer)
            arcpy.SelectLayerByLocation_management(currentProjectLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(currentProjectLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = currentProjectLayer
                outfc = "CurrentProject"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["MODELDT","GISEDITOR","WKORDER","STATUS","ADDRESS","DVLPR","ENGINEER","J_PRELIM","J_AGMTDT","L_ESTMAT","L_ESTLBR","L_ESTFIX",
                          "L_ESTEQP","L_ESTSUB","L_ESTLAB","L_ESTESM","L_ESTSPR","L_ESTSPV","L_ESTLGM","L_ESTTOT","L_ACTMAT","L_ACTLBR","L_ACTFIX","L_ACTEQP",
                          "L_ACTSUB","L_ACTLAB","L_ACTESM","L_ACTSPR","L_ACTSPV","L_ACTLGM","L_ACTTOT","M_ESTFLO","M_ACTFLO","M_LTR_DT","M_AGREXE","M_FM_TAP",
                          "M_TO_TAP","M_OBJ1","M_OBJ2","M_OBJ3","M_TAP1","M_TAPSZ1","M_TAP2","M_TAPSZ2","M_TAP3","M_TAPSZ3","M_TAP4","M_TAPSZ4","M_TAP5","M_TAPSZ5",
                          "M_TAP6","M_TAPSZ6","M_PAGES","M_MAP","M_DRW_WW","M_DRW_CN","M_PI_WW","M_PI_CN","M_SVC_WW","M_SVC_CN","M_OTH_WW","M_OTH_CN","M_FLOSEQ","M_CGWAIV",
                          "M_MODIF","M_COMMENT","N_INSPCTR","N_INSPDT","N_FLCTR","N_ST_WW","N_ST_CN","N_MSUTIL","N_PRMT1","N_PRMT2","N_PRMT3","N_PRMT4","O_FLSHDT","O_REVDT",
                          "O_BLO_CK","O_BL_QTY","O_REVAPP","O_REV_BY","O_VLV_CK","O_WTR_ON","O_CLRTY","O_COMMT","P_COMMT","O_CMP_WW","O_CMP_CN","T_LAB_ST","T_LABDTW","T_LABDTC",
                          "T_LABCST","T_AB_STA","T_AB_DT","T_AB_DESC","T_TIEINW","T_TIEINC","T_PLDT_W","T_PLDT_C","T_CHSVC","T_SVC_DT","T_PWKCMP","UP_AFD_S","UP_AFD_D","UP_MB_S",
                          "UP_MB_D","UP_MBAMT","UP_ASB_S","UP_ASB_D","UP_CST_S","UP_CST_D","UP_CST_A","UP_ES_S","UP_ES_D","UP_SLV_S","UP_SLV_D","UP_INS_S","UP_INS_D","US_AFD_S",
                          "US_AFD_D","US_MB_S","US_MB_D","US_MBAMT","US_ASB_S","US_ASB_D","US_CST_S","US_CST_D","US_CST_A","US_ES_S","US_ES_D","US_SLV_S","US_SLV_D","US_INS_S",
                          "US_INS_D","U_INSDLY","U_INPLCE","U_INACCT","U_INSAPP","U_FINSTL","U_PREINS","U_60DAY","U_FADJ_W","U_FADJ_C","U_FININS","U_DOAPPV","U_CEAPPV","V_HYD_S",
                          "V_HYD_D","V_MAP_S","V_MAP_D","V_AB_S","V_AB_D","V_DRAW_S","V_DRAW_D","V_WWST_S","V_WWST_D","V_GIS_S","V_INVDT","V_APRCVD","M_PMRCVD","V_PRMT_D","FYEAR"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Meter
            ##
            meter = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_WSERVICELOCATION_EAM'
            meterLayer = "V_WSERVICELOCATION_EAM_lyr"
            arcpy.MakeFeatureLayer_management(meter, meterLayer)
            arcpy.SelectLayerByLocation_management(meterLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(meterLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                infc = meterLayer
                outfc = "Meter"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    # ConnectionObjectID added to list 20151120 to drop in the Excel document
                    # ServiceLocationID dropped 20151120 from list so it is in the Excel document
                    dropFields = ["ConnectionObjectID","SubType","LifecycleStatus","QCStatus","AddressAddendum","City","ZipCode",
                           "CrossStreet","Lot","Subdivision","ADCMapNumber","ADCGridCell","WDSMapNumber","WDSGridCell","CustomerServiceUnit",
                           "MeterReadingUnit","ReadingSequenceNumber","LocationCode","LocationNote","PremiseType","BackflowCount","Portion",
                           "WaterQualityConcern","WaterQualityWorkOrderDate","ACCT_NO","GPSCollectionStatus","GPSQCStatus","GPSXCoordinate",
                           "GPSYCoordinate","SerialNumber","MeterStatus","AcquisitionDate","AcquisitionCost","NextInspectionDate","NextReplacementYear",
                           "MeasurementPrecision","MeasurementUnit","Note","NoteAdditionalText","BusinessPartnerID","ContractAccount","Title",
                           "Name2","Name3","Name4","Telephone","Extension","MobileTelephone"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass
            ##
            ## Water Lateral
            ##
            lateral = '//amazon/waterworks/Divisions/InfoTech/Shared/GIS/NewProjectApplication/Hidden/Conway_sdeVector_sdeViewer.sde/SDEVECTOR.SDEDATAOWNER.V_WLATERAL_EDM'
            lateralLayer = "V_WLATERAL_EDM_lyr"
            arcpy.MakeFeatureLayer_management(lateral, lateralLayer)
            arcpy.SelectLayerByLocation_management(lateralLayer, 'intersect', datasetPolygonLayer)
            matchcount = int(arcpy.GetCount_management(lateralLayer).getOutput(0))
            if matchcount == 0:
                pass
            else:
                # Joins the meters to the laterals based on the JOINFIELD
                JOINFIELD   = "ServiceLocationID"
                arcpy.AddJoin_management(lateralLayer, JOINFIELD, meterLayer, JOINFIELD,"KEEP_ALL")


                infc = lateralLayer
                outfc = "Lateral"
                arcpy.FeatureClassToFeatureClass_conversion(infc, env.workspace, outfc)
                try:
                    dropFields = ["FacilityID","LifecycleStatus","QCStatus","SDCDate","SDCStatus","MeterFeeAmount","CorporationType","CorporationSize",
                                  "OriginalTapNumber","OriginalServiceDate","InstallationPaidDate","InstallationFeeAmount","Installer","Contractor",
                                  "Owner","Comments","LegacyAccountNumber",
                                        # These are the new fields that have to be deleted from the AddJoin operation
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ServiceLocationID",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ConnectionObjectID",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_SubType",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_LifecycleStatus",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_QCStatus",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_AddressAddendum",
                                        #SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Address
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_PremiseAdditionalD",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_City",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ZipCode",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_CrossStreet",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Lot",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Subdivision",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_JURIS",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ADCMapNumber",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ADCGridCell",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_WDSMapNumber",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_WDSGridCell",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_CustomerServiceUni",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_MeterReadingUnit",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ReadingSequenceNum",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Location",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_LocationCode",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_LocationNote",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_PremiseType",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_PremiseTypeDescrip",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_InstallationType",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_MeterPresent",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_CustomerPresent",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_BackflowPresent",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_BackflowCount",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Portion",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_RateCategory",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_WaterQualityConcer",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_WaterQualityWorkOr",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ACCT_NO",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_GPSCollectionStatu",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_GPSQCStatus",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_GPSXCoordinate",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_GPSYCoordinate",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_SerialNumber",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Manufacturer",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_MeterStatus",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_DeviceCategory",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_InstallationDate",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Model",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_AcquisitionDate",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_AcquisitionCost",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_NextInspectionDate",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_NextReplacementYea",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_MeasurementPrecisi",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_MeasurementUnit",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Note",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_NoteAdditionalText",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_BusinessPartnerID",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ContractAccount",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Title",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_FirstName",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_LastName",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Name",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Name2",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Name3",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Name4",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Telephone",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_Extension",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_MobileTelephone",
                                        "SDEVECTOR_SDEDATAOWNER_V_WSERVICELOCATION_EAM_ObjectID"]
                    arcpy.DeleteField_management(outfc, dropFields)
                except:
                    pass

            mxd = arcpy.mapping.MapDocument('Current')
            for df in arcpy.mapping.ListDataFrames(mxd):
                for lyr in arcpy.mapping.ListLayers(mxd, "", df):
                    if lyr.name == hydrantLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == MainbreakLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == valveLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == abandonedPipeLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == mainLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == easementLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == fiberLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == fireflowLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == dsiLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == currentProjectLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == meterLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == lateralLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
                    if lyr.name == datasetPolygonLayer:
                        arcpy.mapping.RemoveLayer(df, lyr)
            del mxd
            arcpy.RefreshActiveView()
            arcpy.RefreshTOC()
            #########################################################################

            os.makedirs(projectFolder +"/"+ "TEMP_Excel")


            ## Create Excel files from Tables of Feature Classes within Project Dataset
            fclist = arcpy.ListFeatureClasses()
            for fc in fclist:
                in_table = fc
                out_xls = projectFolder +"/"+ "TEMP_Excel" + "/" + fc + ".xls"
                arcpy.TableToExcel_conversion(in_table, out_xls)

            env.workspace = projectFolder +"/"+ "TEMP_Excel"


            ## Combining all Excel Files
            merged_book = xlwt.Workbook()
            for root, directories, filenames in os.walk(env.workspace):
                for filename in filenames:
                    datafile = os.path.join(root,filename)
                    book = xlrd.open_workbook(datafile)
                    for s in book.sheets():
                        x = merged_book.add_sheet(s.name)
                        for row in range(s.nrows):
                            for col in range(s.ncols):
                                cell_value = s.cell(row,col).value
                                x.write(row, col, cell_value)

            merged_book.save(projectFolder + "/" + event + "_All.xls")
            arcpy.Delete_management(env.workspace)
            print "Processing Complete!"
            # Add a pop up
            pythonaddins.MessageBox('\n\t\t       Processing Complete!\n\n\t\tYour project folder is located at:\n\nR:\Divisions\InfoTech\Shared\GIS\NewProjectApplication\ExtractionFolder\n\n\t\t            Folder Title:\n\n\t\t         ' + str(title) + '\n', "NNWW Toolbars")
            try:
                SendStatusMail_NewProject()
            except:
                pass
        except:
            print 'Extraction Error'

##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##################################################                                      ##################################################
##################################################        NNWW Toolbars Extension       ##################################################
##################################################                                      ##################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################

class ExtensionClass_NNWW(object):
    """Implementation for NNWW_Toolbars_addin.extensionnnww (Extension)"""
    def __init__(self):
        # For performance considerations, please remove all unused methods in this class.
        self.enabled = True
    def activeViewChanged(self):
        try:
            if arcpy.mapping.MapDocument('Current').activeView == 'PAGE_LAYOUT':
                button_addcadindex.enabled      = False
                tool_adddrawings.enabled        = False
                button_removedrawings.enabled   = False
                combobox_streetone.enabled      = False
                combobox_streettwo.enabled      = False
                button_findpoint.enabled        = False
                button_removepoint.enabled      = False
                #button_helpdoc.enabled          = False
                tool_dataextractor.enabled      = False
            else:
                button_addcadindex.enabled      = True
                tool_adddrawings.enabled        = True
                button_removedrawings.enabled   = True
                combobox_streetone.enabled      = True
                combobox_streettwo.enabled      = True
                button_findpoint.enabled        = True
                button_removepoint.enabled      = True
                #button_helpdoc.enabled          = True
                tool_dataextractor.enabled      = True
        except NameError:
            pass

##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
##########################################################################################################################################
















