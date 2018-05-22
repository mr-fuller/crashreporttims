import pandas, os
import pandas as pd
#import numpy as np
#import string
from mtfcToGCAT import schemaCleaner
from funktions import usr, ovrlp,report_yrs, county_fld
#txt = "C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/GCAT_20Mar17_41337.txt"
#txt = "C:/Users/fullerm/Documents/ArcGIS/Projects/Work Computer Files/GCAT_20Mar17_41337.txt"
#txt = "C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/GCAT_12to14_LOSW_19May2017_47554.txt"
# This section cleans up GCAT file

def tblCrashYear(df,resultdir):

    yrs = report_yrs(df)
    severityCd = {
        1 : 'Property Damage Only Crashes',
        2 : 'Possible Injury Crashes',
        3 : 'Non-Incapacitating Injury Crashes',
        4 : 'Incapacitating Crashes',
        5 : 'Fatal Crashes'

    }
    lstofcolumns = {

        "ODPS_MOTORCYCLE_IND": "Motorcycle",
        "ODPS_BICYCLE_IND": "Bicycle", #TIMS didn't have this field
        "ODPS_PEDESTRIAN_IND": "Pedestrian", #TIMS didn't have this field
        "ODPS_SENIOR_DRIVER_IND": "Senior", # senior = over 65 according to odot; mdot, 59
        "DISTRACTED_DRIVER_IND": "Distracted",
        "ODOT_YOUNG_DRIVER_IND": "Young", # young = under 26 according to odot; mdot, 25
        "DOCUMENT_NBR": "Overall",
        "TRUCK_COUNT": "Truck"

    }
    writer = pandas.ExcelWriter(os.path.join(resultdir,  'CrashReportTableYear.xlsx'))
    #writer = pandas.ExcelWriter('C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/Script Tool Results/CrashReportTableYear.xlsx')
    dftxt = pandas.DataFrame(data=df)
    #dftxt.to_csv('C:/Users/fullerm/Desktop/emptyDF.csv')
    #dftxt.to_csv('C:/Users/Michael/Desktop/emptyDF.csv')
    print(type(dftxt), len(dftxt))
    print(len(dftxt["NLF_COUNTY_CD"]))
    for clmn in lstofcolumns:

        #dftxt = pandas.read_table(txt, sep=",",dtype={'ODPS_LOC_ROAD_NAME_TXT': str, 'ODPS_REF_GIVEN_TXT': str})
        #dftxt[clmn] = dftxt[clmn].map(lstofcolumns[clmn][0])
        #dftxt["SEVERITY_BY_TYPE_CD"]=dftxt["SEVERITY_BY_TYPE_CD"].map(severityCd)
        #dftxt["NLF_COUNTY_CD"]=dftxt["NLF_COUNTY_CD"].map(countyCD)
        #if clmn == "U1_TYPE_OF_UNIT_CD":
            #dftxt = df.query("U1_TYPE_OF_UNIT_CD == [13,14,15,16,17,18,19,20]") # What are the correct codes?

        indx=["CRASH_YR",county_fld]
        #vls=["f_count","ii_count","nii_count","pi_count","Total_Crashes"]
        vls=[clmn]
        if clmn == "DOCUMENT_NBR":
            aggfunk=len
        else:
            aggfunk=sum
        # I don't use margins=true for this pivot table so that the later append statement will create a single index
        # dataframe that is more similar to the format of the table in the report. Or it at least does what I want, so
        # I'll optimize it later.
        pt = pandas.pivot_table(dftxt,index=indx,values=vls,columns=["SEVERITY_BY_TYPE_CD"],aggfunc=aggfunk,fill_value=0)
                                #margins=True,margins_name="Total " + lstofcolumns[clmn] + " Crashes")
        print(pt)
        #pt.stack(level=0)
        #pt.fillna(0,inplace=True)

        pt.columns = pt.columns.droplevel(0)

        pandas.MultiIndex.set_names(pt.index,names=["CRASH_YR","County"],inplace=True)

        #print(pt)
        reindx = []
        s =  pt.sum(axis=0,level=0)
        s["year Totals"] = "sum(Year)"
        s.set_index("year Totals", append=True, inplace=True)
        print(s)

        pt = pt.append(s)
        pt.sort_index(inplace=True)
        pandas.MultiIndex.set_levels(pt.index, levels=['Lucas', 'Monroe', 'Wood', 'Totals'], level=1, inplace=True)
        pt["Total " + lstofcolumns[clmn] + " Crashes"] = pt.sum(axis=1)
        # This line of code turns index into two filled columns with years and counties
        pt.reset_index(inplace=True)
        # create a new column of labels by concatenating
        # had to use .map() to convert each individual value to string; using str(series) concatenated entire series to county
        pt["Year"]=pt["CRASH_YR"].map(str) + " " + pt["County"]
        pt.set_index("Year",inplace=True)
        # drop unnecessary columns
        pt.drop(["CRASH_YR","County"],axis=1,inplace=True)
        subTotal = [x for x in pt.index.values if x[-6:] == 'Totals']
        sum_row = pt.query("Year == @subTotal")[pt.columns.values.tolist()].sum()
        pt_sum = pd.DataFrame(data=sum_row).T
        pt_sum = pt_sum.reindex(columns=pt.columns)
        pt = pt.append(pt_sum, ignore_index=False)
        if clmn == "DOCUMENT_NBR":
            pt = pt.rename(index={0: yrs + ' Totals'})
        else:
            pt = pt.rename(index={0: 'Total Crashes'})
        pt.index.name = "Year"
        pt['Fatal & Incapacitating Injury'] = pt['Fatal Crashes'] + pt[
            'Incapacitating Injury Crashes']
        pt['Non-Incapacitating & Possible Injury'] = pt['Non-Incapacitating Injury Crashes'] + pt[
            'Possible Injury Crashes']
        pt.to_excel(writer,sheet_name=lstofcolumns[clmn])
        wrkbk = writer.book
        wrksht = writer.sheets[lstofcolumns[clmn]]
        chrt = wrkbk.add_chart({'type': 'column'})
        '''chrt.add_series({'categories': [lstofcolumns[clmn], 1, 0, len(temp_df) - 1, 0],
                         'values': [lstofcolumns[clmn], 1, 6, len(temp_df) - 1, 6],
                         'data_labels': {'percentage': True},
                         'name': 'All Crashes'})'''
        chrt.add_series({'categories': '=('+lstofcolumns[clmn]+'!$A$2:A4,'+lstofcolumns[clmn]+'!A6:A8,'+lstofcolumns[clmn]+'!A10:A12)',
                         'values': '=('+lstofcolumns[clmn]+'!$H$2:H4,'+lstofcolumns[clmn]+'!H6:H8,'+lstofcolumns[clmn]+'!H10:H12)',
                         'data_labels': {'value': True},
                         'name': 'Fatal & Incapacitating Injury',
                         'fill': {'color': 'red'}})
        chrt.add_series({'categories': '=('+lstofcolumns[clmn]+'!$A$2:A4,'+lstofcolumns[clmn]+'!A6:A8,'+lstofcolumns[clmn]+'!A10:A12)',
                         'values': '=('+lstofcolumns[clmn]+'!$I$2:I4,'+lstofcolumns[clmn]+'!I6:I8,'+lstofcolumns[clmn]+'!I10:I12)',
                         'data_labels': {'value': True},
                         'name': 'Non-Incapacitating & Possible Injury',
                         'fill':{'color':'yellow'}})
        '''chrt.add_series({'categories': '=('+lstofcolumns[clmn]+'!$A$2:A4,'+lstofcolumns[clmn]+'!A6:A8,'+lstofcolumns[clmn]+'!A10:A12)',
                         'values': '=('+lstofcolumns[clmn]+'!$D$2:D4,'+lstofcolumns[clmn]+'!D6:D8,'+lstofcolumns[clmn]+'!D10:D12)',
                         'data_labels': {'percentage': True},
                         'name': 'Non-Incapacitating Injury'})
        chrt.add_series({'categories': '=('+lstofcolumns[clmn]+'!$A$2:A4,'+lstofcolumns[clmn]+'!A6:A8,'+lstofcolumns[clmn]+'!A10:A12)',
                         'values': '=('+lstofcolumns[clmn]+'!$E$2:E4,'+lstofcolumns[clmn]+'!E6:E8,'+lstofcolumns[clmn]+'!E10:E12)',
                         'data_labels': {'percentage': True},
                         'name': 'Possible Injury'})'''
        chrt.add_series({'categories': '=('+lstofcolumns[clmn]+'!$A$2:A4,'+lstofcolumns[clmn]+'!A6:A8,'+lstofcolumns[clmn]+'!A10:A12)',
                         'values': '=('+lstofcolumns[clmn]+'!$F$2:F4,'+lstofcolumns[clmn]+'!F6:F8,'+lstofcolumns[clmn]+'!F10:F12)',
                         'data_labels': {'value': True},
                         'name': 'Property Damage',
                         'fill': {'color': '#00B050'},
                         'overlap': ovrlp})

        chrt.set_legend({'position': 'bottom'})
        chrt.set_x_axis({'major_tick_mark': 'none'})
        chrt.set_y_axis({'major_tick_mark': 'none',
                         'visible': False,
                         'major_gridlines': {'visible': False}})
        if clmn == 'ODPS_BICYCLE_IND' or clmn == 'ODPS_MOTORCYCLE_IND' or clmn == 'ODPS_PEDESTRIAN_IND' or clmn == "TRUCK_COUNT": # or truck eventually?
            chrt.set_title({'name': yrs + ' Crashes Involving ' + lstofcolumns[clmn] + 's by Year'})
        elif clmn == 'DOCUMENT_NBR':
            chrt.set_title({'name': yrs + ' Crashes ' + lstofcolumns[clmn] + ' by Year'})
        else:
            chrt.set_title({'name': yrs + ' Crashes Involving ' + lstofcolumns[clmn] + ' Drivers by Year'})
        wrksht.insert_chart('K2', chrt)

    writer.save()
    return os.path.join(resultdir,  'CrashReportTableYear.xlsx')
if __name__ == "__main__":
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    tblCrashYear(schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")[0], "C:/Users/" + usr + "/Desktop")