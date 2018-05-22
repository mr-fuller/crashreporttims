import pandas
import pandas as pd
import os
from mtfcToGCAT import schemaCleaner
from funktions import report_yrs,ovrlp, county_fld


def drugalcoholfii(dftxt,resultdir):
    yrs = report_yrs(dftxt)
    ptbl = pandas.DataFrame()
    dict = {
        "ODPS_DRUG_IND":"Drugs", "ODPS_ALCOHOL_IND":"Alcohol"
    }
    writer = pandas.ExcelWriter(os.path.join(resultdir, "Drug_Alcohol_F_II.xlsx"))
    for clmn in dict:

        pt = pandas.pivot_table(dftxt,index=[county_fld,"SEVERITY_BY_TYPE_CD"],values=clmn,
                                            columns=["CRASH_YR"],aggfunc=sum,fill_value=0,dropna=False)


        #pt.columns = pt.columns.droplevel(0)

        pandas.MultiIndex.set_names(pt.index, names=["County", "Severity"], inplace=True)
        pandas.MultiIndex.set_levels(pt.index,["Lucas","Monroe","Wood"], level=0, inplace=True)
        pt["Substance"] = dict[clmn]
        pt.set_index("Substance",append=True,inplace=True)
        #pandas.MultiIndex.set_levels(pt.index,[["County","Substance","Data"]],inplace=True)
        pt = pt.reorder_levels(["County","Substance", "Severity"])

        pt = pt.query('Severity == ["Fatal Crashes","Incapacitating Injury Crashes"]')
        ptbl = ptbl.append(pt)
    # Doing these next things outside of the previous for loop is an important difference between this method and those
    # for other tables.
    ptbl["Total"] = ptbl.sum(axis=1)
    print(ptbl.sort_index(level=1, inplace=True))

    # Note the absence of any kind of total row in the following loop. Refer to design of tables in 2012-2014 report
    for county in ptbl.index.get_level_values(0).unique():
        #temper_df = ptbl.xs(county, level=0)
        temp_df = ptbl.xs(county, level=0)
        #temp_df = temper_df[0:2] + temper_df[4:]
        temp_df.to_excel(writer, county + 'DA_F_II')  #write it to excel so the data is there to make a chart
        wrkbk = writer.book
        wrksht = writer.sheets[county + 'DA_F_II']

        chrt = wrkbk.add_chart({'type': 'column'})
        chrt.add_series({'categories': [county + 'DA_F_II', 0, 2, 0, 4],
                             'values': [county + 'DA_F_II', 1, 2, 1, 4],
                         'name': 'Fatal Alcohol'
                         })
        chrt.add_series({'categories': [county + 'DA_F_II', 0, 2, 0, 4],
                         'values': [county + 'DA_F_II', 2, 2, 2, 4],
                         'name': 'Incapacitating Injury Alcohol'
                         })
        chrt.add_series({'categories': [county + 'DA_F_II', 0, 2, 0, 4],
                         'values': [county + 'DA_F_II', 3, 2, 3, 4],
                         'name': 'Fatal Drug'
                         })
        chrt.add_series({'categories': [county + 'DA_F_II', 0, 2, 0, 4],
                         'values': [county + 'DA_F_II', 4, 2, 4, 4],
                         'name': 'Incapacitating Injury Drug',
                         'overlap': ovrlp})
        chrt.set_x_axis({'major_tick_mark':'none'})
        chrt.set_y_axis({'major_tick_mark':'none',
                         'line': {'none': True}})
        chrt.set_legend({'position': 'bottom'})
        chrt.set_title({'name': yrs + ' ' + county + ' County Alcohol & Drug Related Fatal/Incapacitating Crashes'})
        wrksht.insert_chart('I2',chrt)
    writer.save()
    return os.path.join(resultdir, "Drug_Alcohol_F_II.xlsx")
if __name__ == "__main__":
    usr = 'Michael'
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    drugalcoholfii(schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")[0], "C:/Users/" + usr + "/Desktop")