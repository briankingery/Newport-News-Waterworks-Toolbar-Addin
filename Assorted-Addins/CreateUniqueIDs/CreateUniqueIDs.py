#-------------------------------------------------------------------------------
# Purpose:     Creates a unique ID for each record of an input feature class
#              or table.
#
# Author:      Ian Broad
# Website:     www.ianbroad.com
#
# Created:     07/09/2014
#-------------------------------------------------------------------------------

import arcpy
import uuid
import random

arcpy.env.overwriteOutput = True

def group_sort(group, sort, sort_type, option):
    global input_feature
    global sort_field
    global group_field
    global current_interval

    if group == "YES" and sort == "YES":
        groups = set(sorted([row[0] for row in arcpy.da.SearchCursor(input_feature, group_field)]))
        for group in groups:
            current_interval = value
            oids = [row[1] for row in sorted(arcpy.da.SearchCursor(input_feature, (sort_field, "OID@"), """{0} = '{1}'""".format(group_field, group)), reverse=sort_type)]
            add_id(oids)

    elif group == "YES" and sort == "NO":
        groups = set(sorted([row[0] for row in arcpy.da.SearchCursor(input_feature, group_field)]))
        for group in groups:
            current_interval = value
            oids = [row[0] for row in arcpy.da.SearchCursor(input_feature, ("OID@"), """{0} = '{1}'""".format(group_field, group))]
            add_id(oids)

    else:
        oids = [row[1] for row in sorted(arcpy.da.SearchCursor(input_feature, (sort_field, "OID@")), reverse=sort_type)]
        add_id(oids)

def add_id(oids):
    global input_feature
    global fields
    global prefix
    global suffix
    global value
    global current_interval

    for oid in oids:
        with arcpy.da.UpdateCursor(input_feature, (fields), """{0} = {1}""".format(str(fields[1]), oid)) as update_cursor:
            for row in update_cursor:
                row[0] = str(prefix) + str(current_interval) + str(suffix)
                update_cursor.updateRow(row)
                current_interval += value
                arcpy.SetProgressorPosition()

input_feature = arcpy.GetParameterAsText(0)
choice = arcpy.GetParameterAsText(1)
use_group = arcpy.GetParameterAsText(2)
group_field = arcpy.GetParameterAsText(3)
use_sort = arcpy.GetParameterAsText(4)
sort_field = arcpy.GetParameterAsText(5)
value = int(arcpy.GetParameterAsText(6))
option = arcpy.GetParameterAsText(7)
prefix = arcpy.GetParameterAsText(8)
suffix = arcpy.GetParameterAsText(9)

result = arcpy.GetCount_management(input_feature)
features = int(result.getOutput(0))

arcpy.SetProgressor("step", "Adding Unique IDs...", 0, features, 1)

fields = ["UID"]
oid_field = arcpy.Describe(input_feature).OIDFieldName
fields.append(str(oid_field))

if use_sort == "YES":
    fields.append(sort_field)
if use_group == "YES":
    fields.append(group_field)

if "UID" not in [f.name for f in arcpy.ListFields(input_feature)]:
    arcpy.AddField_management(input_feature, "UID", "TEXT")

if choice == "UNIVERSALLY UNIQUE IDENTIFER":
    with arcpy.da.UpdateCursor(input_feature, (fields)) as update_cursor:
        for row in update_cursor:
            row[0] = str(uuid.uuid4())
            update_cursor.updateRow(row)
            arcpy.SetProgressorPosition()

if choice == "RANDOM WITHIN RANGE":
    with arcpy.da.UpdateCursor(input_feature, (fields)) as update_cursor:
        used_ids = []
        for row in update_cursor:
            boolean = True
            while boolean == True:
                random_id = random.randint(0, value)
                if random_id not in used_ids:
                    row[0] = str(prefix) + str(random_id) + str(suffix)
                    update_cursor.updateRow(row)
                    used_ids.append(random_id)
                    boolean = False
            arcpy.SetProgressorPosition()

if choice == "INTERVAL":
    arcpy.AddMessage("Got here3")
    if use_group == "NO" and use_sort == "YES, ASCENDING":
        group_sort("NO", "YES", False, option)
    elif use_group == "NO" and use_sort == "YES, DESCENDING":
        group_sort("NO", "YES", True, option)
    elif use_group == "YES" and use_sort == "YES, ASCENDING":
        group_sort("YES", "YES", False, option)
    elif use_group == "YES" and use_sort == "YES, DESCENDING":
        group_sort("YES", "YES", True, option)
    elif use_group == "YES" and use_sort == "NO":
        arcpy.AddMessage("Got here4")
        group_sort("YES", "NO", True, option)
    else:
        with arcpy.da.UpdateCursor(input_feature, (fields)) as update_cursor:
            current_interval = value
            for row in update_cursor:
                row[0] = str(prefix) + str(current_interval) + str(suffix)
                update_cursor.updateRow(row)
                current_interval += value
                arcpy.SetProgressorPosition()

arcpy.ResetProgressor()
arcpy.GetMessages()

