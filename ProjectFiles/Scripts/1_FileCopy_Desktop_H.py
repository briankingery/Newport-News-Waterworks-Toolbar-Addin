#-------------------------------------------------------------------------------
# Name:         FileCopy_Desktop_H.py
# Purpose:      Copy GIS directory from Desktop C drive to H Drive
#
# Author:       Brian Kingery
#
# Created:      3/31/2016
# Copyright:    (c) bkingery 2016
#
# Directions:   Run FileCopy_Desktop_H.py on Desktop
#-------------------------------------------------------------------------------

import os, shutil, sys, datetime

x = raw_input('You must be on the Desktop to run this code.\n\nAre you there?\nYes or No: ')
if x.lower() == 'y' or x.lower() == 'yes':
    y = raw_input('Are you absolutely positive?\nYes or No: ')
    if y.lower() == 'y' or y.lower() == 'yes':
        ## Start
        ExecutionStartTime = datetime.datetime.now()
        print "Started: %s" % ExecutionStartTime.strftime('%A, %B %d, %Y %I:%M:%S %p')
        print "Processing\n"


        ## Delete H:\GIS if exists
        H_GIS = '//amazon/home2/bkingery/GIS'
        if os.path.exists(H_GIS):
            shutil.rmtree(H_GIS)
        ## Copy C:\GIS to H:\
        src = 'C:/GIS'
        dst = '//amazon/home2/bkingery/GIS'
        shutil.copytree(src,dst)


        ExecutionEndTime = datetime.datetime.now()
        ElapsedTime = ExecutionEndTime - ExecutionStartTime
        print "\nEnded: %s\n" % ExecutionEndTime.strftime('%A, %B %d, %Y %I:%M:%S %p')
        print "Elapsed Time: %s" % str(ElapsedTime).split('.')[0]
else:
    pass
    #sys.exit()
end = raw_input("Press 'Enter' to exit program")


