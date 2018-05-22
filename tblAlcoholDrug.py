import pandas,os
import pandas as pd
import numpy as np
from mtfcToGCAT import schemaCleaner
from funktions import usr, ovrlp, report_yrs, county_fld
#txt = "C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/GCAT_20Mar17_41337.txt"
#txt = "C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/GCAT_12to14_LOSW_19May2017_47554.txt"
# This section cleans up GCAT file
def alcoholdrugtbl(df,resultdir):

    yrs = report_yrs(df)
    alcoholDrugCd = {
        0:'Unknown',
        1:'None',
        2:'Yes - Alcohol Suspected',
        3:'Yes - HBD Not Impaired',
        4:'Yes - Drugs Suspected',
        5:'Yes - Alcohol and Drugs Suspected'
    }
    lstofcolumns = ["ODPS_ALCOHOL_IND", "ODPS_DRUG_IND"]



    writer = pandas.ExcelWriter(os.path.join(resultdir,'Alcohol_Drug.xlsx'))


    indx=[county_fld,"CRASH_YR"]
    #vls=["f_count","ii_count","nii_count","pi_count","Total_Crashes"]
    vls=["ODPS_ALCOHOL_IND","ODPS_DRUG_IND"]

    # I don't use margins=true for this pivot table so that the later append statement will create a single index
    # dataframe that is more similar to the format of the table in the report. Or it at least does what I want, so
    # I'll optimize it later.
    pt = pandas.pivot_table(df, index=indx, values=vls, columns=["SEVERITY_BY_TYPE_CD"],aggfunc=sum,fill_value=0)
                           # margins=True,margins_name="All Crashes")
    pandas.MultiIndex.set_names(pt.index, names=[None, None], inplace=True)
    #pt = pt[:-1]
    #print(pt)
    #pt.stack(level=0)
    #pt.fillna(0,inplace=True)
    print(pt)
    pt.columns = pt.columns.swaplevel(0,1)
    pt.sort_index(axis=1,level=0,inplace=True)
    print(pt.index)
    print(pt.columns)
    #pt["Year"] = [2012,2013,2014]
    #pt.set_index("Year",inplace=True)
    #print(pt)

    pandas.MultiIndex.set_levels(pt.index, ["Lucas", "Monroe", "Wood"], level= 0, inplace=True)
    #pandas.MultiIndex.set_names(pt.columns,names=["Alcohol","Drugs"],level=1, inplace=True)
    pt.columns.set_levels(["Alcohol", "Drugs"], level=1, inplace=True)
    pandas.MultiIndex.set_names(pt.columns, names=[None, None], inplace=True)
    print(pt.columns)

    s = pt.sum(axis=0, level=0)
    s["year Totals"] = "Total"
    s.set_index("year Totals", append=True, inplace=True)
    print(s)

    pt = pt.append(s)


    pt.reset_index(inplace=True)
    print(pt)
    #pt['Crash_Year'] = pt['Crash_Year'].map(str)
    #pt['Crash_Year'] = pt['Crash_Year'].str[:4]
    pt["Year"] = pt['level_0'] + " " + pt['level_1'].map(str)
    pt.set_index('Year', inplace=True)
    pt.drop(["level_0", "level_1"], axis=1, inplace=True)
    pt.sort_index(inplace=True)
    # pt.index.name = None
    #cols = pt.columns.tolist()
    #cols = cols[2:]+cols[:2]
    #pt = pt[cols]
    #rows = pt.index.values.tolist()
    #rows = rows[2:] + rows[:1]
    #pt = pt.reindex([rows])
    pt['Number of Crashes', 'Alcohol'] = pt['Fatal Crashes', 'Alcohol'] + pt['Incapacitating Injury Crashes', 'Alcohol'] + pt['Non-Incapacitating Injury Crashes','Alcohol'] + pt['Possible Injury Crashes', 'Alcohol'] + pt['Property Damage Only Crashes', 'Alcohol']
    pt['Number of Crashes', 'Drugs'] = pt['Fatal Crashes', 'Drugs'] + pt['Incapacitating Injury Crashes', 'Drugs'] + pt['Non-Incapacitating Injury Crashes', 'Drugs'] + pt['Possible Injury Crashes', 'Drugs'] + pt['Property Damage Only Crashes','Drugs']
    pt['Fatal & Incapacitating Injury', 'Alcohol'] = pt['Fatal Crashes', 'Alcohol'] + pt['Incapacitating Injury Crashes', 'Alcohol']
    pt['Fatal & Incapacitating Injury', 'Drugs'] = pt['Fatal Crashes', 'Drugs'] + pt['Incapacitating Injury Crashes', 'Drugs']
    pt['Non-Incapacitating & Possible Injury', 'Alcohol'] = pt['Non-Incapacitating Injury Crashes', 'Alcohol'] + pt['Possible Injury Crashes', 'Alcohol']
    pt['Non-Incapacitating & Possible Injury', 'Drugs'] = pt['Non-Incapacitating Injury Crashes', 'Drugs'] + pt['Possible Injury Crashes', 'Drugs']
    '''lst = [np.nan]*17
    newrow = pd.Series(lst)
    #pt = pt.append(newrow)
    #pandas.MultiIndex.append(pt.columns,newrow)
    pd.concat([pt], axis=1, keys=[np.nan], names=['Blah'])
    print(pt)
    pt.columns.swaplevel'''
    #pt_ = pt[:2].append(pt[4:])
    print(pt.columns)
    subTotal = [x for x in pt.index.values if x[-5:] == 'Total']
    sum_row = pt.query("Year == @subTotal")[pt.columns.values.tolist()].sum()
    pt_sum = pd.DataFrame(data=sum_row).T
    pt_sum = pt_sum.reindex(columns=pt.columns)
    pt = pt.append(pt_sum, ignore_index=False)
    pt = pt.rename(index={0: 'Total Crashes'})
    #pandas.MultiIndex.set_names(pt.columns, [None, None], inplace=True)
    #pt=pt.reset_index()
    #pt.rename(columns={'index': 'Year'}, inplace=True)
    print(pt)
    #pt.drop('Year',axis=0, inplace=True)
    pt.to_excel(writer, 'Drug_Alcohol')

    wrkbk = writer.book
    wrksht = writer.sheets['Drug_Alcohol']
    #for county in ["Lucas","Monroe","Wood"]:
    dict = {
        'Lucas': [3, 5],
        'Monroe': [7, 9],
        'Wood': [11, 13]
        }
    for key in dict:
        chrtsht = wrkbk.add_chartsheet(key +'Chart')
        chrt = wrkbk.add_chart({'type': 'column'})

        chrt.add_series({'categories': ['Drug_Alcohol', dict[key][0], 0, dict[key][1], 0],
                         'values': ['Drug_Alcohol', dict[key][0], 13, dict[key][1], 13],
                         'data_labels': {'value': True},
                         'name':  'Fatal & Incapacitating Injury Alcohol'})

        chrt.add_series({'categories': ['Drug_Alcohol', dict[key][0], 0, dict[key][1], 0],
                         'values': ['Drug_Alcohol', dict[key][0], 14, dict[key][1], 14],
                         'data_labels': {'value': True},
                         'name': 'Fatal & Incapacitating Injury Drug'})

        chrt.add_series({'categories': ['Drug_Alcohol', dict[key][0], 0, dict[key][1], 0],
                             'values': ['Drug_Alcohol', dict[key][0], 15, dict[key][1], 15],
                             'data_labels': {'value': True},
                             'name': 'Non-Incapacitating & Possible Injury Alcohol'})

        chrt.add_series({'categories': ['Drug_Alcohol', dict[key][0], 0, dict[key][1], 0],
                             'values': ['Drug_Alcohol', dict[key][0], 16, dict[key][1], 16],
                             'data_labels': {'value': True},
                             'name': 'Non-Incapacitating & Possible Injury Drug'})

        chrt.add_series({'categories': ['Drug_Alcohol', dict[key][0], 0, dict[key][1], 0],
                             'values': ['Drug_Alcohol', dict[key][0], 9, dict[key][1], 9],
                             'data_labels': {'value': True},
                             'name': 'Property Damage Alcohol'})

        chrt.add_series({'categories': ['Drug_Alcohol', dict[key][0], 0, dict[key][1], 0],
                             'values': ['Drug_Alcohol', dict[key][0], 10, dict[key][1], 10],
                             'data_labels': {'value': True},
                             'name': 'Property Damage Drug',
                             'overlap': ovrlp})
        chrt.set_x_axis({'major_tick_mark': 'none'})
        chrt.set_y_axis({'major_tick_mark': 'none',
                         'line': {'none':True},
                         'visible': False,
                         'major_gridlines':{'visible': False}})
        chrt.set_title({'name': yrs + ' ' + key + ' County Alcohol & Drug Related Crashes'})
        chrt.set_legend({'position': 'bottom'})
        chrtsht.set_chart(chrt)
    #writer.save() saving after each step seems to make subsequent charts not insert


    writer.save()
    return os.path.join(resultdir,"Alcohol_Drug.xlsx")
if __name__ == "__main__":
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    alcoholdrugtbl(schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")[0], "C:/Users/" + usr + "/Desktop")