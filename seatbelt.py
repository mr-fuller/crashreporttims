import pandas as pd
from funktions import type_append_total, usr, ovrlp, report_yrs, county_fld
import os
from mtfcToGCAT import schemaCleaner


def seatbelttbl(df,resultdir):
    yrs = report_yrs(df)
    pt_total = pd.pivot_table(df,"DOCUMENT_NBR",[county_fld,"UNRESTRAIN_OCCUPANTS"],"SEVERITY_BY_TYPE_CD",len,
                              fill_value=0,dropna=False)
    pt_total["Total Crashes"] = pt_total.sum(axis=1)

    pt_total['Fatal & Incapacitating Injury'] = pt_total['Fatal Crashes'] + pt_total[
        'Incapacitating Injury Crashes']
    pt_total['Non-Incapacitating & Possible Injury'] = pt_total['Non-Incapacitating Injury Crashes'] + pt_total[
        'Possible Injury Crashes']
    print(pt_total)
    pt_total.rename(index={0: 'None'},inplace=True)
    pd.MultiIndex.set_levels(pt_total.index,['Lucas','Monroe','Wood'],level=0,inplace=True)
    writer = pd.ExcelWriter(os.path.join(resultdir,"Seatbelt.xlsx"),engine='xlsxwriter')
    for county in pt_total.index.get_level_values(0).unique():
        temp_df = pt_total.xs(county, level=0)
        temp_df = type_append_total(temp_df,'')
        temp_df.to_excel(writer, sheet_name=county )
        wrkbk = writer.book
        wrksht = writer.sheets[county]
        #percent_fmt = wrkbk.add_format({'align': 'right', 'num_format': '0.00%'})
        #wrksht.set_column('H:H', 12, percent_fmt)
        print(len(temp_df))
        #temp_df = type_append_total(temp_df)

        temp_df.index.name = "Unrestrained Occupants"

        '''if clmn == "HOUR_OF_CRASH_NBR":
            chrt = wrkbk.add_chart({'type': 'line'})
        else:'''
        chrt = wrkbk.add_chart({'type': 'column'})
        '''chrt.add_series({'categories': [county, 1, 0, len(temp_df) - 1, 0],
                         'values': [county, 1, 6, len(temp_df) - 1, 6],
                         'data_labels': {'percentage': True},
                         'name': 'All Crashes'})'''
        chrt.add_series({'categories': [county, 2, 0, len(temp_df)-1 , 0], # start in row 2, don't include people that wore seatbelt
                         'values': [county, 2, 7, len(temp_df)-1 , 7],
                         'data_labels': {'value': True},
                         'name': 'Fatal & Incapacitating Injury',
                         'fill': {'color': 'red'}})
        chrt.add_series({'categories': [county, 2, 0, len(temp_df)-1 , 0],
                         'values': [county, 2, 8, len(temp_df)-1 , 8],
                         'data_labels': {'value': True},
                         'name': 'Non-Incapacitating & Possible Injury',
                         'fill': {'color': 'yellow'}})
        '''chrt.add_series({'categories': [county, 2, 0, len(temp_df) , 0],
                         'values': [county, 2, 3, len(temp_df), 3],
                         'data_labels': {'percentage': True},
                         'name': 'Non-Incapacitating Injury'})
        chrt.add_series({'categories': [county, 2, 0, len(temp_df) , 0],
                         'values': [county, 2, 4, len(temp_df) , 4],
                         'data_labels': {'percentage': True},
                         'name': 'Possible Injury'})'''
        chrt.add_series({'categories': [county, 2, 0, len(temp_df)-1 , 0],
                         'values': [county, 2, 5, len(temp_df)-1, 5],
                         'data_labels': {'value': True},
                         'name': 'Property Damage',
                         'fill':{'color':'#00B050'},
                         'overlap': ovrlp})  # light green
        chrt.set_legend({'position': 'bottom'})
        chrt.set_x_axis({'major_tick_mark': 'none'})
        chrt.set_y_axis({'major_tick_mark': 'none',
                         'visible': False,
                         'major_gridlines': {'visible': False}})
        chrt.set_title({'name': yrs + ' ' + county + ' County Crashes Involving Unrestrained Occupants'})
        wrksht.insert_chart('K2',chrt)
        temp_df.to_excel(writer,county)


        #pt.to_excel(writer)
    writer.save()
    writer.close()

    return os.path.join(resultdir,  'Seatbelt.xlsx')

if __name__ == "__main__":
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    seatbelttbl(schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")[0], "C:/Users/" + usr + "/Desktop")
