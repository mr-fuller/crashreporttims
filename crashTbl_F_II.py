import pandas, os
import pandas as pd
from funktions import pt_reindex_compare, usr, report_yrs, county_fld
from mtfcToGCAT import schemaCleaner,dictofcolumns

    #import numpy as np
    #import string
    #from mtcfToGCAT import schemaCleaner
    #from CrashReportLRS import TimeDateStr

def tblCrashF_II(df,resultdir):

    yrs = report_yrs(df)
    dftxt = df
    writer = pandas.ExcelWriter(os.path.join(resultdir,  "Tables_F_II.xlsx"))
    for clmn in dictofcolumns:
        print(clmn)
        pt = pandas.DataFrame()
        #indx=["NLF_COUNTY_CD",clmn,"SEVERITY_BY_TYPE_CD"]
        vls=["DOCUMENT_NBR"]
        if clmn == "U1_CONT_CIR_PRIMARY_CD" or clmn == 'U1_TYPE_OF_UNIT_CD':
            pt1 = pandas.pivot_table(dftxt, values=["DOCUMENT_NBR"], index=[county_fld, dictofcolumns[clmn][0] + "1",
                                                               "SEVERITY_BY_TYPE_CD"], columns=["CRASH_YR"],
                                     aggfunc=len, fill_value=0, dropna=False)

            pt2 = pandas.pivot_table(dftxt, values=["DOCUMENT_NBR"], index=[county_fld, dictofcolumns[clmn][0] + "2",
                                                               "SEVERITY_BY_TYPE_CD"],columns=["CRASH_YR"],
                                     aggfunc=len, fill_value=0, dropna=False)
            pt3 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "3",'SEVERITY_BY_TYPE_CD'],
                                     columns=["CRASH_YR"], aggfunc=len, fill_value=0, dropna=False)

            # God I hope that using the larger index covers everything
            pt_temp = pt_reindex_compare(pt1, pt2)
            pt = pt_reindex_compare(pt_temp, pt3)
        elif clmn == "U1_AGE_NBR":
            pt1 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "1",'SEVERITY_BY_TYPE_CD'],
                                     columns=["CRASH_YR"], aggfunc=len, fill_value=0, dropna=False)

            pt2 = pandas.pivot_table(dftxt, values=vls, index=[county_fld, dictofcolumns[clmn][0] + "2",'SEVERITY_BY_TYPE_CD'],
                                     columns=["CRASH_YR"], aggfunc=len, fill_value=0, dropna=False)
            pt = pt_reindex_compare(pt1,pt2)

        elif clmn == "TIME_OF_CRASH":
            indx = [county_fld, "Hour_of_Crash","SEVERITY_BY_TYPE_CD"]
            pt = pandas.pivot_table(dftxt, index=indx, values=vls, columns=["CRASH_YR"],
                                      aggfunc=len, fill_value=0, dropna=False)

        else:
            pt = pandas.pivot_table(dftxt,values=["DOCUMENT_NBR"], index=[county_fld, clmn,"SEVERITY_BY_TYPE_CD"],
                                    columns=["CRASH_YR"],aggfunc=len,fill_value=0,dropna=False)
        #pt.unstack(level=2,fill_value=0).stack(level=2,dropna=False)
        #pt.fillna(0)


        pt.columns = pt.columns.droplevel(0)

        pandas.MultiIndex.set_names(pt.index,names=["County",dictofcolumns[clmn][0],"Severity"],inplace=True)
        pandas.MultiIndex.set_levels(pt.index,["Lucas","Monroe","Wood"],level=0,inplace=True)
        pt = pt.query('Severity == ["Fatal Crashes","Incapacitating Injury Crashes"]')
        #print(pt)
        #print(pt.index.levels[2])
        #pt.loc
        pt["Total"] = pt.sum(axis=1)
        #print(pt)
        #pandas.MultiIndex.reindex(pt,index="Data",)


        pt = pt.reindex([dictofcolumns[clmn][1][k] for k in dictofcolumns[clmn][1]], level=dictofcolumns[clmn][0])
        for county in pt.index.get_level_values(0).unique():
            temp_df = pt.xs(county,level=0)
            f_ii_t = pd.DataFrame()
            for item in ["Fatal Crashes","Incapacitating Injury Crashes"]:
                row_sum = temp_df.query('Severity ==  @item ')[temp_df.columns.values.tolist()].sum()
                #print(row_sum)
                row_sum_t = pd.DataFrame(data=row_sum).T
                row_sum_t = row_sum_t.rename(index={0: item})
                f_ii_t = f_ii_t.append(row_sum_t)

            #print(f_ii_t.index.name)
            f_ii_t["Severity"]=["Fatal Crash","Incapacitating Injury Crash"]
            f_ii_t.set_index("Severity",inplace=True)
            f_ii_t[dictofcolumns[clmn][0]]="Total"
            f_ii_t.set_index(dictofcolumns[clmn][0],append=True,inplace=True)
            f_ii_t = f_ii_t.reorder_levels([dictofcolumns[clmn][0],"Severity"])

            temp_df = temp_df.append(f_ii_t,ignore_index = False)

            temp_df.to_excel(writer,county + dictofcolumns[clmn][0] + 'F_II')
            wrkbk = writer.book
            wrksht = writer.sheets[county + dictofcolumns[clmn][0]+'F_II']
            lst_row = wrksht
            '''chartdf = temp_df
            chartdf.drop([2012,2013,2014],axis=1,inplace=True)
            #chartdf.columns = chartdf.columns.droplevel(0)
            chartdf = chartdf.unstack(1)
            print(chartdf)
            if clmn == "HOUR_OF_CRASH_NBR":
                chrt= wrkbk.add_chart({'type':'line'})
            else:
                chrt = wrkbk.add_chart({'type':'column'})
            chrt.set_title({'name': yrs + ' ' + county + ' County Fatal & Incapacitating Crashes by ' + dictofcolumns[clmn][0]})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0],1,0,len(temp_df)-2,1],
                             'values': [county + dictofcolumns[clmn][0], 1,5,len(temp_df)-2,5],
                             'data_labels': {'percentage':True}
                             })
            wrksht.insert_chart('I2', chrt)'''
            # make top row blue
            # add all borders to all cells
            # unbold text in columns a and b
            #bld_fmt = wrkbk.add_format({'border':1})
            #num_rows = len(dictofcolumns[1][0])*2
            #wrksht.set_column('B:F',None,bld_fmt)
            #writer.save()
            # make last two rows grey
        #pt.to_excel(writer)
    writer.save()

    return os.path.join(resultdir, "Tables_F_II.xlsx")
if __name__ == "__main__":
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    tblCrashF_II(schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")[0], "C:/Users/" + usr + "/Desktop")