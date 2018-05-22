import pandas
import pandas as pd
import arcpy
import numpy as np
#from mtfcToGCAT import dictofcolumns

usr = 'fullerm'
# usr = 'Michael'
ovrlp = -25
def pt_reindex_compare(pt1,pt2):
    # function to reindex both dataframes as the larger of the two because dropna=false wasn't working?
    if len(pt1) > len(pt2):
        pt_total = (pt2.reindex_like(pt1)).fillna(0) + pt1.fillna(0).fillna(0)  # do this, not .add()

    elif len(pt1) < len(pt2):
        pt_total = (pt1.reindex_like(pt2)).fillna(0) + pt2.fillna(0).fillna(0)  # do this, not .add()

    else:
        pt_total = pt1.add(pt2, fill_value=0)  # The dataframes should be the same in this case, right?

    return pt_total

def type_append_total(temp_df,clmn):
    # Function to sum columns and append a total row to the bottom of the table
    # because I don't like using the margins option of pandas for some reason
    sum_row = temp_df[temp_df.columns.values.tolist()].sum()
    pt_sum = pd.DataFrame(data=sum_row).T
    pt_sum = pt_sum.reindex(columns=temp_df.columns)
    temp_df = temp_df.append(pt_sum, ignore_index=False)

    if clmn == "U1_TYPE_OF_UNIT_CD":
        temp_df = temp_df.rename(index={0: "Total Vehicles"})
    elif clmn == "U1_AGE_NBR":
        temp_df = temp_df.rename(index={0: "Total Persons Involved"})
    else:
        temp_df = temp_df.rename(index={0: "Total Crashes"})
    #temp_df.index.name = dictofcolumns[clmn][0]
    return temp_df

county_fld = "NLF_COUNTY_CD"

def CRLRS_excel_export(df, f, writer):
    if f == 'Subdivision':
        df.rename(columns={  # 'NAMELSAD':'Municipality',
            'Join_Count': 'Total Crashes',
            'sum_non_incapac_inj_count': 'Non-Incapacitating Injury Crashes',
            'sum_possible_inj_count': 'Possible Injury Crashes',
            'sum_fatalities_count': 'Fatal Crashes',
            'sum_incapac_inj_count': 'Incapacitating Injury Crashes',
            'PDO_':'Property Damage Only Crashes'}, inplace=True)

        df['NAMELSAD'] = df['NAMELSAD'].str.title()
        df['Municipality'] = ' a'
        print(df)
        # splt = df['Municipality'].str.split()
        for index, row in df.iterrows():
            if row['NAMELSAD'].split()[-1:][0] != 'Township':
                df.loc[index, 'Municipality'] = row['NAMELSAD'].split()[-1:][0] + ' of ' + \
                                                ' '.join(row['NAMELSAD'].split()[:-1])
            else:
                df.loc[index, 'Municipality'] = row['NAMELSAD']
        df.reset_index(inplace=True)
        df.set_index('Municipality', inplace=True)
        #df.index.name = ""
        df.sort_index(inplace=True)
        df.drop(["NAMELSAD", "Shape_Length", "Shape_Area", "EPDO_Index", "OBJECTID","index"], axis=1, inplace=True)

        df = df
        #df.set_index('COUNTY', inplace=True)
        for county in np.unique(df['COUNTY'].tolist()):
            arcpy.AddMessage("{}".format(county))
            #  have to use something so that it remains a dataframe, either query or .xs().to_frame() because my
            #  arcgis pro environment doesn't have version 0.20 of pandas goddammit esri is the worst sometimes
            #  beginning in pandas 0.20, series have .to_excel() method
            temp_df = df.query('COUNTY == @county')
            temp_df = temp_df[["Fatal Crashes","Incapacitating Injury Crashes","Non-Incapacitating Injury Crashes",
                 "Possible Injury Crashes","Property Damage Only Crashes","Total Crashes"]]
            temp_df = type_append_total(temp_df,f)
            temp_df.index.name = 'Municipality'
            temp_df.to_excel(writer, sheet_name=county + "_" + f + "_Scores")
    else:
        for county in df['County'].unique():
            temp_df = df.query('County == @county')
            temp_df.to_excel(writer, sheet_name= county + '_' + f + "_Scores")

'''def CRLRS_EPDO_index(joinfile):
    CursorFlds = ['PDO_',
                  'EPDO_Index',
                  'Join_Count',
                  'sum_fatalities_count',
                  'sum_incapac_inj_count',
                  'sum_non_incapac_inj_count',
                  'sum_possible_inj_count'
                  ]

    # determine PDO  and EPDO Index/Rate
    with arcpy.da.UpdateCursor(joinfile, CursorFlds) as cursor:
        for row in cursor:
            try:
                row[0] = row[2] - int(row[3]) - int(row[4]) - int(row[5]) - int(row[6])
            except:
                row[0] = 0 # null to zero or divide by zero are the major exceptions we are handling here
            try:
                row[1] = (float(row[3]) * fatalwt + float(row[4]) * seriouswt + float(row[5]) * nonseriouswt + float(
                    row[6]) * possiblewt + float(row[0])) / float(row[2])
            except:
                row[1] = 0 # null to zero or divide by zero are the major exceptions we are handling here
            cursor.updateRow(row)'''

def addSumFlds(fldmppngs):
    # creates field map and merge rule for each field in new sum fields
    flds = ["fatalities_count", "incapac_inj_count", "non_incapac_inj_count", "possible_inj_count"]
    for fld in flds:
        FieldIndex = fldmppngs.findFieldMapIndex(fld)
        fldmap = fldmppngs.getFieldMap(FieldIndex)
        # Get the output field's properties as a field object
        outputFld = fldmap.outputField
        # Rename the field and pass the updated field object back into the field map
        outputFld.name = "sum_" + fld
        outputFld.aliasName = "sum_" + fld
        fldmap.outputField = outputFld
        # Set the merge rule to sum and then replace the old fieldmaps in the mappings object
        # with the updated ones
        fldmap.mergeRule = "sum"
        fldmppngs.replaceFieldMap(FieldIndex, fldmap)

'''def fillCountFields(pointcopy, fld_lst):
    # populates newly created count fields with 1 for crash or 0 for no crash
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
            cursor.updateRow(row)'''
def report_yrs(df):
    yrs = np.unique(df['CRASH_YR'].tolist())
    yrs_str = str(yrs[0]) + '-' + str(yrs[len(yrs) - 1])
    return yrs_str

