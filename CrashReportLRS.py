import arcpy, os, pandas, xlsxwriter
from datetime import datetime
from funktions import CRLRS_excel_export, addSumFlds, usr #CRLRS_EPDO_index,  fillCountFields
from MI_locations_XY import milocationsxy
from mtfcToGCAT import schemaCleaner

def crashReportLRS(GDBspot,csv,fatalwt, seriouswt,nonseriouswt,possiblewt,IntersectionThreshold,SegmentThreshold):
    # workspace= "Z:/fullerm/Safety Locations/Safety.gdb"

    # Input parameters
    # GCAT file/location
    GCATfile = csv  # csv created after mapping fields with schemaCleaner
    # Intermediate file/location?

    # Intersection polygon file/location
    # IntersectionFeatures = arcpy.GetParameterAsText(1)
    # IntersectionFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/Fed_Aid_2010_LucWoo_Intersection_Buffer_Dissolve"
    IntersectionFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/LMW_intersection_250ft_buffer_5Jul2017_3857"
    psFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/CS_IP_Merge_copy_clip_16june2017_LMW_3857"
    psThreshold = str(0)
    countyFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/County_FLOWHHMS_Clipped_3857"
    # Segment polygon file/location
    # SegmentFeatures = arcpy.GetParameterAsText(2)
    SegmentFeatures = "Z:/fullerm/Safety Locations/Crash_Report_Script_Tool.gdb/LMW_segments_70ft_buffer_5Jul2017_3857"
    # output file name/location for the spatial join
    # GDBspot = arcpy.GetParameterAsText(1)  # user input location for gdb and result excel tables

    # psThreshold = arcpy.GetParameterAsText(8)
    # output file name/location for excel table
    # TableFolder = arcpy.GetParameterAsText(4)
    # rdinv = arcpy.GetParameterAsText(8)
    rdinv = "C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/Safety Report Script.gdb/Road_Inventory_CopyFeatures"
    # create geodatabase
    TimeDate = datetime.now()
    TimeDateStr = "CrashLocations" + TimeDate.strftime('%Y%m%d%H%M') + "_LRS"

    outputGDB = arcpy.CreateFileGDB_management(GDBspot, TimeDateStr)
    arcpy.env.workspace = str(outputGDB).replace('/', '\\')
    # I kept getting errors because arcmap sets the field type based on first dozen features and some ids were numeric
    '''fldmppng = arcpy.FieldMappings()
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
    NewTable = arcpy.TableToTable_conversion(GCATfile, outputGDB, "OHMI_data")
    arcpy.TableSelect_analysis(NewTable, "oh_table", '"NLF_COUNTY_CD" <> \'Monroe\' ')
    ohTable = arcpy.CopyRows_management("oh_table", "ohTable")
    arcpy.TableSelect_analysis(NewTable, "mi_table", '"NLF_COUNTY_CD" = \'Monroe\' ')
    miTable = arcpy.CopyRows_management('mi_table', 'miTable')
    # arcpy.SelectLayerByAttribute_management(NewTable,"CLEAR_SELECTION")
    rdlyr = arcpy.MakeFeatureLayer_management(rdinv, "rdlyr" )
    rtloc = os.path.join(GDBspot, "Road_Inventory3456_CreateRoutes" + TimeDateStr + ".shp")
    lrs = arcpy.CreateRoutes_lr(rdlyr, "NLF_ID", rtloc, "TWO_FIELDS", "CTL_BEGIN", "CTL_END")
    event_props = "NLFID POINT COUNTY_LOG_NBR"
    PointFile = arcpy.MakeRouteEventLayer_lr(lrs, "NLF_ID", ohTable, event_props, "Crash_Events")
    # creating this extra feature class and working from it instead of the event layer
    # decreased script tool runtime from ~8 min to ~2 min
    arcpy.SelectLayerByAttribute_management(PointFile, "clear_selection")

    pointOH = arcpy.FeatureClassToFeatureClass_conversion(PointFile, outputGDB, "GCAT_LUCWOO_lrs_points_" + TimeDateStr)
    # pointOH = arcpy.CopyFeatures_management(PointFile, "LRS_Events_copy")
    mi_points = milocationsxy(miTable, outputGDB)
    pointcopy = arcpy.Merge_management([pointOH, mi_points], 'miohpointsmerge')

    dict = {'fatalities_count': "ODPS_TOTAL_FATALITIES_NBR<>0",
            'incapac_inj_count': "Incapac_injuries_NBR<>0 and ODPS_TOTAL_FATALITIES_NBR=0",
            'non_incapac_inj_count': "non_incapac_injuries_NBR<>0 and ODPS_TOTAL_FATALITIES_NBR=0 and incapac_injuries_nbr=0",
            'possible_inj_count': "possible_injuries_nbr<>0 and ODPS_TOTAL_FATALITIES_NBR=0 and non_incapac_injuries_nbr=0 and incapac_injuries_nbr=0"
            }
    fld_lst = ['SEVERITY_BY_TYPE_CD', 'fatalities_count', 'incapac_inj_count','non_incapac_inj_count',
               'possible_inj_count'
            ]

    # add fields for point layer

    for key in dict:
        arcpy.AddField_management(pointcopy, key, "LONG")
        '''arcpy.SelectLayerByAttribute_management(PointFile, "NEW_SELECTION", dict[key])
        arcpy.CalculateField_management(PointFile, key, 1)
        arcpy.SelectLayerByAttribute_management(PointFile, "Switch_selection")
        arcpy.CalculateField_management(PointFile, key, 0)'''
    # fillCountFields(pointcopy, fld_lst)
    with arcpy.da.UpdateCursor(pointcopy, fld_lst) as cursor:
        for row in cursor:
            if row[0] == 'Fatal Crashes':
                row[1] = 1
                row[2] = 0
                row[3] = 0
                row[4] = 0
            elif row[0] == 'Incapacitating Injury Crashes':
                row[1] = 0
                row[2] = 1
                row[3] = 0
                row[4] = 0
            elif row[0] == 'Non-Incapacitating Injury Crashes':
                row[1] = 0
                row[2] = 0
                row[3] = 1
                row[4] = 0
            elif row[0] == 'Possible Injury Crashes':
                row[1] = 0
                row[2] = 0
                row[3] = 0
                row[4] = 1
            else:
                row[1] = 0
                row[2] = 0
                row[3] = 0
                row[4] = 0
            cursor.updateRow(row)

    # Clear Selected Features
    arcpy.SelectLayerByAttribute_management(PointFile, "clear_selection")
    # PointFeatures2 = arcpy.CopyFeatures_management(PointFeatures,os.path.join(GDBspot, TimeDateStr + ".gdb\PointFeatures2"))
    PointFeatures = arcpy.FeatureClassToFeatureClass_conversion(pointcopy, outputGDB,
                                                                "ohmi_points_copy" + TimeDateStr)
    ftype = {
        'Intersection': [IntersectionThreshold, IntersectionFeatures],
        'Segment': [SegmentThreshold, SegmentFeatures],
        'Subdivision': [psThreshold, psFeatures]
    }
    # field map and merge rules
    attchmnt = []
    writer = pandas.ExcelWriter(os.path.join(GDBspot,  "Top_Locations.xlsx"),engine='xlsxwriter')
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
        addSumFlds(fieldmappings)

        # Run the Spatial Join tool, using the defaults for the join operation and join type
        loc = os.path.join(GDBspot, TimeDateStr + ".gdb\\" + f + "Join_LRS")
        Join = arcpy.SpatialJoin_analysis(ftype[f][1], PointFeatures,loc,"Join_one_to_one", "keep_all", fieldmappings)

        arcpy.AddField_management(Join, "PDO_", "LONG")
        arcpy.AddField_management(Join, "EPDO_Index", "DOUBLE")
        # CRLRS_EPDO_index(Join)
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
                    row[0] = 0  # null or divide by zero are the major exceptions we are handling here
                try:
                    row[1] = (float(row[3]) * fatalwt + float(row[4]) * seriouswt + float(row[5]) * nonseriouswt +
                              float(row[6]) * possiblewt + float(row[0])) / float(row[2])
                except:
                    row[1] = 0  # null or divide by zero are the major exceptions we are handling here
                cursor.updateRow(row)

        # delete unnecessary fields
        keepFlds = ['OBJECTID',
                    'Shape',
                    'Shape_Area',
                    'Shape_Length',
                    'Name',
                    'NAMELSAD',
                    'COUNTY',
                    'COUNTY_NME',
                    'Join_Count',
                    'sum_fatalities_count',
                    'sum_incapac_inj_count',
                    'sum_non_incapac_inj_count',
                    'sum_possible_inj_count',
                    'PDO_',
                    'EPDO_Index',
                    'Fed_Aid_Buffer_Segments_2_Name',
                    'Length_ft',
                    'County']
        # lstFlds = arcpy.ListFields(Join)

        dropFlds = [x.name for x in arcpy.ListFields(Join) if x.name not in keepFlds]
        # delete fields
        arcpy.DeleteField_management(Join, dropFlds)

        # select high crash locations
        JoinLayer = arcpy.MakeFeatureLayer_management(Join,os.path.join(GDBspot, TimeDateStr + ".gdb\\" + f + "JoinLayer"))
        arcpy.AddMessage("{}".format(type(JoinLayer)))
        # arcpy.SelectLayerByAttribute_management(JoinLayer, "NEW_SELECTION", "Join_Count >=" + ftype[f][0])
        fld_nmes = [fld.name for fld in arcpy.ListFields(JoinLayer)]

        fld_nmes.remove('Shape')  # I think this field kept causing an exception: Data must be 1 dimensional
        arcpy.AddMessage("{}".format(fld_nmes))
        arcpy.AddMessage("{}".format(type(os.path.join(GDBspot, TimeDateStr + ".gdb\\" + f + "JoinLayer"))))
        # do this because political sud
        # fields can be list or tuple, list works when 'Shape' field removed
        n = arcpy.da.FeatureClassToNumPyArray(JoinLayer,fld_nmes, where_clause="Join_Count  >=" + ftype[f][0],
                                              skip_nulls=False,null_value=0)

        df = pandas.DataFrame(n)

        CRLRS_excel_export(df, f, writer)

    writer.save()
    return os.path.join(GDBspot,  "Top_Locations.xlsx")
if __name__ == "__main__":
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    crashReportLRS(schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")[0],
                   schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")[1], 39.22, 39.22, 6.84, 4.63, '20', '30')
