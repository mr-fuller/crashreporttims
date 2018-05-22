import pandas, os
import pandas as pd
from funktions import pt_reindex_compare, type_append_total, report_yrs, usr, ovrlp, county_fld
from mtfcToGCAT import schemaCleaner, dictofcolumns
from chartcreator import typecharts
    #import os
    #import numpy as np
    #from mtcfToGCAT import schemaCleaner
    #from CrashReportLRS import TimeDateStr


def crashtypetbl(df, resultdir):

    #ovrlp = -25
    yrs = report_yrs(df)
    dftxt = df  # should return a dataframe
    writer = pandas.ExcelWriter(os.path.join(resultdir,  'Tables.xlsx'), engine='xlsxwriter')
    for clmn in dictofcolumns:

        pt_total = pandas.DataFrame()
        # county_fld = "COUNTY"
        # print(indx)
        # pt = pd.DataFrame(dftxt,index=pd.IntervalIndex.from_breaks)
        vls = ["DOCUMENT_NBR"]
        # indx=["NLF_COUNTY_CD", clmn]
        if clmn == "U1_CONT_CIR_PRIMARY_CD" or clmn == 'U1_TYPE_OF_UNIT_CD':
            pt1 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "1"],
                                     columns=['SEVERITY_BY_TYPE_CD'], aggfunc=len, fill_value=0, dropna=False)

            pt2 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "2"],
                                     columns=['SEVERITY_BY_TYPE_CD'], aggfunc=len, fill_value=0, dropna=False)
            pt3 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "3"],
                                     columns=['SEVERITY_BY_TYPE_CD'], aggfunc=len, fill_value=0, dropna=False)

            # God I hope that using the larger index covers everything
            pt_temp = pt_reindex_compare(pt1,pt2)
            pt_total = pt_reindex_compare(pt_temp,pt3)

        elif clmn == "U1_AGE_NBR":
            pt1 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "1"],
                                     columns=['SEVERITY_BY_TYPE_CD'], aggfunc=len, fill_value=0, dropna=False)

            pt2 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "2"],
                                     columns=['SEVERITY_BY_TYPE_CD'], aggfunc=len, fill_value=0, dropna=False)
            pt_total = pt_reindex_compare(pt1,pt2)
        elif clmn == "TIME_OF_CRASH":
            indx = [county_fld, "Hour_of_Crash"]
            pt_total = pandas.pivot_table(dftxt, index=indx, values=vls, columns=["SEVERITY_BY_TYPE_CD"],
                                      aggfunc=len, fill_value=0, dropna=False)
        # dropna = True for location and crash type codes because the schemas/values for ohio and michigan are different
        elif clmn == "ODOT_CRASH_LOCATION_CD" or clmn == "CRASH_TYPE_CD":
            indx = [county_fld, clmn]
            pt_total = pandas.pivot_table(dftxt, index=indx, values=vls, columns=["SEVERITY_BY_TYPE_CD"],
                                          aggfunc=len, fill_value=0, dropna=True)
        else:
            indx = [county_fld, clmn]
            pt_total = pandas.pivot_table(dftxt,index=indx, values=vls,columns=["SEVERITY_BY_TYPE_CD"],
                                          aggfunc=len,fill_value=0, dropna=False)
            # pt.fillna(0,inplace=True)

        pt_total.columns = pt_total.columns.droplevel(0)
        # print(pt_total.index)
        pandas.MultiIndex.set_names(pt_total.index,names=["County",dictofcolumns[clmn][0]],inplace=True)
        pandas.MultiIndex.set_levels(pt_total.index,["Lucas","Monroe","Wood"],level=0,inplace=True)

        pt_total["Total Crashes"] = pt_total.sum(axis=1)
        pt_total["Percent of " + dictofcolumns[clmn][0] + " of All Crashes"] = \
            pt_total["Total Crashes"] / pt_total["Total Crashes"].sum(level=0)
        if clmn == "DAY_IN_WEEK_CD" or clmn == "U1_AGE_NBR" or clmn == "TIME_OF_CRASH":
            # print(len(pt_total.index))
            pt_total = pt_total.reindex([dictofcolumns[clmn][1][k] for k in dictofcolumns[clmn][1]],
                                        level=dictofcolumns[clmn][0])

        elif clmn == "U1_TYPE_OF_UNIT_CD":
            # using level 0 ensures percentages are based on County totals

            pt_total = pt_total.reindex([dictofcolumns[clmn][1][k] for k in dictofcolumns[clmn][1]],
                                        level=dictofcolumns[clmn][0])
            pt_total = pt_total.sort_values(by="Total Crashes", ascending=False)
        elif clmn == "UNRESTRAIN_OCCUPANTS":
            pt_total = pt_total.rename(index={0: "None"}) # do this so there aren't 2 rows labeled 'Total Crashes'
            # potential to remove this row altogether
        else:
            pt_total = pt_total.sort_values(by="Total Crashes",ascending=False)
        pt_total['Fatal & Incapacitating Injury'] = pt_total['Fatal Crashes'] + pt_total['Incapacitating Injury Crashes']
        pt_total['Non-Incapacitating & Possible Injury'] = pt_total['Non-Incapacitating Injury Crashes'] + pt_total['Possible Injury Crashes']
        # pt_total.to_excel(writer,clmn)
        # pt_total[dictofcolumns[clmn][1]] = pt_total[dictofcolumns[clmn][1]].map(dictofcolumns[clmn][0])
        for county in pt_total.index.get_level_values(0).unique():
            temp_df = pt_total.xs(county,level=0)
            temp_df.to_excel(writer, sheet_name=county + dictofcolumns[clmn][0])
            wrkbk = writer.book
            wrksht = writer.sheets[county + dictofcolumns[clmn][0]]
            percent_fmt = wrkbk.add_format({'align': 'right', 'num_format': '0.00%'})
            wrksht.set_column('H:H', 12, percent_fmt)
            print(len(temp_df))
            temp_df.index.name = dictofcolumns[clmn][0]
            temp_df = type_append_total(temp_df,clmn)

            chrt = typecharts(clmn, wrkbk, county, temp_df, yrs)


            chrt.set_x_axis({'major_tick_mark': 'none'})
            chrt.set_title({'name': yrs + ' ' + county + ' County Crashes by ' + dictofcolumns[clmn][0]})
            chrtsht = wrkbk.add_chartsheet(county + dictofcolumns[clmn][0]+'Chart')
            chrtfii = wrkbk.add_chart({'type':'column'})
            chrtfii.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                                 'values': [county + dictofcolumns[clmn][0], 1, 1, len(temp_df) - 1, 1],
                                 'data_labels': {'value': False},
                                 'name': 'Fatal',
                                 'fill': {'color':'red'}})
            chrtfii.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                                 'values': [county + dictofcolumns[clmn][0], 1, 2, len(temp_df) - 1, 2],
                                 'data_labels': {'value': False},
                                 'name': 'Incapacitating Injury',
                                 'overlap': ovrlp,
                                 'fill':{'color': 'yellow'}})
            chrtfii.set_title(
                {'name': yrs + ' ' + county + ' County Fatal & Incapacitating Crashes by ' + dictofcolumns[clmn][0]})
            chrtfii.set_legend({'position':'bottom'})
            chrtfii.set_x_axis({'major_tick_mark': 'none'})
            chrtfii.set_y_axis({'major_tick_mark': 'none',
                                'line': {'none': True}})
            chrtsht.set_chart(chrtfii)

            wrksht.insert_chart('L2',chrt)
            temp_df.to_excel(writer,county + dictofcolumns[clmn][0])


        #pt.to_excel(writer)
    writer.save()
    writer.close()
    xlfile = 'C:/Users/fullerm/Desktop/Tables.xlsx'
    print(type(xlfile))
    return os.path.join(resultdir,  'Tables.xlsx')
if __name__ == "__main__":
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    crashtypetbl(schemaCleaner(txt,mtcfCsv,"C:/Users/" + usr + "/Desktop")[0],"C:/Users/" + usr + "/Desktop")
