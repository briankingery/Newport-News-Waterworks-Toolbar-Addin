#-------------------------------------------------------------------------------
# Name:         FileCopy_H_Virtual.py
# Purpose:      Copy directory from H drive to the Virtual C Drive
#
# Author:       Brian Kingery
#
# Created:      3/31/2016
# Copyright:    (c) bkingery 2016
#
# Directions:   Run FileCopy_H_Virtual.py on Virtual Environment
#-------------------------------------------------------------------------------

import os, shutil, sys, datetime

x = raw_input('You must be in the Virtual Environment to run this code.\n\nAre you there?\nYes or No: ')
if x.lower() == 'y' or x.lower() == 'yes':
    y = raw_input('Are you absolutely positive?\nYes or No: ')
    if y.lower() == 'y' or y.lower() == 'yes':
        ## Start
        ExecutionStartTime = datetime.datetime.now()
        print "Started: %s" % ExecutionStartTime.strftime('%A, %B %d, %Y %I:%M:%S %p')
        print "Processing\n"


        ## Delete C:\GIS if exists
        C_GIS = 'C:/GIS'
        if os.path.exists(C_GIS):
            shutil.rmtree(C_GIS)
        ## Move H:\GIS to C:\GIS
        src = '//amazon/home2/bkingery/GIS'
        dst = 'C:/'
        if os.path.exists(src):
            shutil.move(src, dst)
        else:
            print '\nGIS folder on H drive does not exist\n'


        ExecutionEndTime = datetime.datetime.now()
        ElapsedTime = ExecutionEndTime - ExecutionStartTime
        print "\nEnded: %s\n" % ExecutionEndTime.strftime('%A, %B %d, %Y %I:%M:%S %p')
        print "Elapsed Time: %s" % str(ElapsedTime).split('.')[0]
else:
    pass
    #sys.exit()
end = raw_input("Press 'Enter' to exit program")


