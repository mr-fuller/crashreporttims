import arcpy, os, pandas, xlsxwriter
from datetime import datetime

def milocationsxy(mtcffile,outputGDB):
    #workspace= "Z:/fullerm/Safety Locations/Safety.gdb"

    # Input parameters
    # GCAT file/location
    '''GCATfile = arcpy.GetParameterAsText(0)  # user input for text file of crash location data from GCAT
    # Intermediate file/location?

    # Intersection polygon file/location
    # IntersectionFeatures = arcpy.GetParameterAsText(1)
    IntersectionFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/Fed_Aid_2010_LucWoo_Intersection_Buffer_Dissolve"
    psFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/CS_IP_Merge_copy_clip_2june2017_LMW_3857"
    psThreshold = str(0)
    countyFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/County_FLOWHHMS_Clipped_3857"
    # Segment polygon file/location
    # SegmentFeatures = arcpy.GetParameterAsText(2)
    SegmentFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/segment_buffer_70ft"
    # output file name/location for the spatial join
    GDBspot = arcpy.GetParameterAsText(1)  # user input location for gdb and result excel tables
    fatalwt = float(arcpy.GetParameterAsText(2))  # user input weight for fatal crashes
    seriouswt = float(arcpy.GetParameterAsText(3))  # user input weight for serious crashes
    nonseriouswt = float(arcpy.GetParameterAsText(4))  # user input weight for nonserious crashes
    possiblewt = float(arcpy.GetParameterAsText(5))  # user input weight for possible crashes
    # curious that these two thresholds don't have to be numbers
    IntersectionThreshold = arcpy.GetParameterAsText(6)  # user input number of crashes to qualify an intersection as high crash
    SegmentThreshold = arcpy.GetParameterAsText(7)  # User input number of crashes to qualify a segment as high crash
    #psThreshold = arcpy.GetParameterAsText(8)
    # output file name/location for excel table
    # TableFolder = arcpy.GetParameterAsText(4)

    # create geodatabase
    TimeDate = datetime.now()
    TimeDateStr = "CrashLocations" + TimeDate.strftime('%Y%m%d%H%M')+ "_XY"
    outputGDB = arcpy.CreateFileGDB_management(GDBspot, TimeDateStr)
    arcpy.env.workspace = str(outputGDB).replace('/','\\')
    # I kept getting errors because arcmap sets the field type based on first dozen features and some ids were numeric
    fldmppng = arcpy.FieldMappings()
    fldmppng.addTable(GCATfile)

    nameFI = fldmppng.findFieldMapIndex("LOCAL_REPORT_NUMBER_ID")
    fldmp = fldmppng.getFieldMap(nameFI)
    fld = fldmp.outputField

    fld.name = "LOCAL_REPORT_NUMBER_ID"
    fld.aliasName = "LOCAL_REPORT_NUMBER_ID"
    fld.type = "String"
    fldmp.outputField = fld
    fldmppng.replaceFieldMap(nameFI,fldmp)'''
    # convert GCAT txt file to gdb table and add to map
    NewTable = arcpy.TableToTable_conversion(mtcffile, outputGDB, "mtcf_mon")

    # display xy data and export to new feature class
    PointLayer = arcpy.MakeXYEventLayer_management(NewTable, "ODOT_LONGITUDE_NBR", "ODOT_LATITUDE_NBR", "mtcf_mon_xy",
                                                  arcpy.SpatialReference("NAD 1983"))

    # creating this extra feature class and working from it instead of the event layer
    # decreased script tool runtime from ~2 min to ~1 min
    pointfeatures = arcpy.CopyFeatures_management(PointLayer,"mon_xy_events_copy")

    '''dict = {'fatalities_count':"FATALITIES_NBR<>0",
            'incapac_inj_count': "Incapac_injuries_NBR<>0 and fatalities_nbr=0",
            'non_incapac_inj_count':"non_incapac_injuries_NBR<>0 and fatalities_nbr=0 and incapac_injuries_nbr=0",
            'possible_inj_count':"possible_injuries_nbr<>0 and FATALITIES_NBR=0 and non_incapac_injuries_nbr=0 and incapac_injuries_nbr=0"
            }
    fld_lst = ['fatalities_count',"FATALITIES_NBR", 'incapac_inj_count', "Incapac_injuries_NBR",
            'non_incapac_inj_count',"non_incapac_injuries_NBR", 'possible_inj_count',"possible_injuries_nbr"
            ]

    # add fields for point layer

    for key in dict:
        arcpy.AddField_management(pointcopy,key,"LONG")
        arcpy.SelectLayerByAttribute_management(PointFile, "NEW_SELECTION", dict[key])
        arcpy.CalculateField_management(PointFile, key, 1)
        arcpy.SelectLayerByAttribute_management(PointFile, "Switch_selection")
        arcpy.CalculateField_management(PointFile, key, 0)
    with arcpy.da.UpdateCursor(pointcopy,fld_lst) as cursor:
        for row in cursor:
            if row[1] != 0:
                row[0] = 1
                row[2] = 0
                row[4] = 0
                row[6] = 0
            elif (row[3] != 0 and row[1] == 0):
                row[0] = 0
                row[2] = 1
                row[4] = 0
                row[6] = 0
            elif (row[5] != 0 and row[3] == 0 and row[1] == 0):
                row[0] = 0
                row[2] = 0
                row[4] = 1
                row[6] = 0
            elif (row[7] != 0 and row[5] == 0 and row[3] == 0 and row[1] == 0):
                row[0] = 0
                row[2] = 0
                row[4] = 0
                row[6] = 1
            else:
                row[0] = 0
                row[2] = 0
                row[4] = 0
                row[6] = 0
            cursor.updateRow(row)

    # Clear Selected Features
    arcpy.SelectLayerByAttribute_management(PointFile, "clear_selection")

    PointFeatures = arcpy.FeatureClassToFeatureClass_conversion(pointcopy, outputGDB, "GCAT_LUCWOO_xy_points_" + TimeDateStr)
    #PointFeatures2 = arcpy.CopyFeatures_management(PointFeatures,os.path.join(GDBspot, TimeDateStr + ".gdb\PointFeatures2"))

    ftype = {
        'Intersection': [IntersectionThreshold, IntersectionFeatures],
        'Segment': [SegmentThreshold, SegmentFeatures],
        'Subdivision': [psThreshold, psFeatures]
    }
    # field map and merge rules
    attchmnt = []
    writer = pandas.ExcelWriter(os.path.join(GDBspot, TimeDateStr + ".xlsx"),engine='xlsxwriter')
    for f in ftype:

        # Create a new fieldmappings and add the two input feature classes.
        fieldmappings = arcpy.FieldMappings()
        fieldmappings.addTable(ftype[f][1])
        fieldmappings.addTable(PointFeatures)

        # First get the fieldmaps. POP1990 is a field in the cities feature class.
        # The output will have the states with the attributes of the cities. Setting the
        # field's merge rule to mean will aggregate the values for all of the cities for
        # each state into an average value. The field is also renamed to be more appropriate
        # for the output.
        flds = ["fatalities_count", "incapac_inj_count","non_incapac_inj_count","possible_inj_count"]
        for fld in flds:

            FieldIndex = fieldmappings.findFieldMapIndex(fld)
            fldmap = fieldmappings.getFieldMap(FieldIndex)
            # Get the output field's properties as a field object
            outputFld = fldmap.outputField
            # Rename the field and pass the updated field object back into the field map
            outputFld.name = "sum_" + fld
            outputFld.aliasName = "sum_" + fld
            fldmap.outputField = outputFld
            # Set the merge rule to sum and then replace the old fieldmaps in the mappings object
            # with the updated ones
            fldmap.mergeRule = "sum"
            fieldmappings.replaceFieldMap(FieldIndex, fldmap)

        # Run the Spatial Join tool, using the defaults for the join operation and join type
        loc = os.path.join(GDBspot, TimeDateStr + ".gdb\\" + f + "Join_XY")
        Join = arcpy.SpatialJoin_analysis(ftype[f][1], PointFeatures,loc,"Join_one_to_one", "keep_all", fieldmappings)

        arcpy.AddField_management(Join, "PDO_", "LONG")
        arcpy.AddField_management(Join, "EPDO_Index", "DOUBLE")

        CursorFlds = ['PDO_',
                      'EPDO_Index',
                      'Join_Count',
                      'sum_fatalities_count',
                      'sum_incapac_inj_count',
                      'sum_non_incapac_inj_count',
                      'sum_possible_inj_count'
                      ]

        # determine PDO  and EPDO Index/Rate
        with arcpy.da.UpdateCursor(Join, CursorFlds) as cursor:
            for row in cursor:
                try:
                    row[0] = row[2] - int(row[3]) - int(row[4]) - int(row[5]) - int(row[6])
                except:
                    row[0] = 0
                try:
                    row[1] = (float(row[3]) * fatalwt + float(row[4]) * seriouswt + float(row[5]) * nonseriouswt + float(
                        row[6]) * possiblewt + float(row[0])) / float(row[2])
                except:
                    row[1] = 0
                cursor.updateRow(row)

        # delete unnecessary fields
        keepFlds = ['OBJECTID',
                    'Shape',
                    'Shape_Area',
                    'Shape_Length',
                    'Name',
                    'NAMELSAD',
                    'COUNTY',
                    'Join_Count',
                    'sum_fatalities_count',
                    'sum_incapac_inj_count',
                    'sum_non_incapac_inj_count',
                    'sum_possible_inj_count',
                    'PDO_',
                    'EPDO_Index',
                    'Fed_Aid_Buffer_Segments_2_Name',
                    'Length_ft']
        #lstFlds = arcpy.ListFields(Join)

        dropFlds = [x.name for x in arcpy.ListFields(Join) if x.name not in keepFlds]
        # delete fields
        arcpy.DeleteField_management(Join, dropFlds)

        # select high crash locations
        JoinLayer = arcpy.MakeFeatureLayer_management(Join,os.path.join(GDBspot, TimeDateStr + ".gdb\\" + f + "JoinLayer"))
        arcpy.AddMessage("{}".format(type(JoinLayer)))
        #arcpy.SelectLayerByAttribute_management(JoinLayer, "NEW_SELECTION", "Join_Count >=" + ftype[f][0])
        fld_nmes = [fld.name for fld in arcpy.ListFields(JoinLayer)]
        fld_nmes.remove('Shape') # I think this field kept causing an exception: Data must be 1 dimensional
        arcpy.AddMessage("{}".format(fld_nmes))
        arcpy.AddMessage("{}".format(type(os.path.join(GDBspot, TimeDateStr + ".gdb\\" + f + "JoinLayer"))))
         # do this because political sud
        if f == 'Subdivision':
            n = arcpy.da.FeatureClassToNumPyArray(JoinLayer, fld_nmes,Join_Count > {}.format(-1),skip_nulls=False, null_value= 0)  # fields can be list or tuple, list works when 'Shape' field removed
            # whereClaus =
        else:
            n = arcpy.da.FeatureClassToNumPyArray(JoinLayer, fld_nmes, where_clause="Join_Count  >=" + ftype[f][0])  # fields can be list or tuple, list works when 'Shape' field removed
        n = arcpy.da.FeatureClassToNumPyArray(JoinLayer,fld_nmes, where_clause="Join_Count  >=" + ftype[f][0],skip_nulls=False,null_value=0) # fields can be list or tuple, list works when 'Shape' field removed
        df = pandas.DataFrame(n)
        if f == 'Subdivision':
            df.set_index('COUNTY',inplace=True)
            for county in df.index.get_level_values(0).unique():
                arcpy.AddMessage("{}".format(county))
                #  have to use something so that it remains a dataframe, either query or .xs().to_frame() because my
                #  arcgis pro environment doesn't have version 0.20 of pandas goddammit esri is the worst sometimes
                #  beginning in pandas 0.20, series have .to_excel() method
                temp_df = df.query('COUNTY == @county')
                temp_df.to_excel(writer,sheet_name=county  + "_" + f + "_Scores")
        else:
            df.to_excel(writer,sheet_name= f + "_Scores")

        # copy selected records and output to Excel
        # xltbl = arcpy.TableToExcel_conversion(JoinLayer, os.path.join(GDBspot, TimeDateStr + ".xls\\"+ f + "_Scores$"))
        # attchmnt.append(xltbl)
        #arcpy.SelectLayerByAttribute_management(JoinLayer, "CLEAR_SELECTION")'''
    #writer.save()  # save outside for loop to save both intersections and segments to the same excel file
    return pointfeatures
# email file as attachment to Marc/Lisa
