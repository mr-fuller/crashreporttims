from funktions import ovrlp
from mtfcToGCAT import dictofcolumns

def typecharts(clmn, wrkbk, county, temp_df, yrs):
    # takes GCAT/TIMS column name as input
    # dictofcolumn = dictofcolumns
    if clmn == "TIME_OF_CRASH":
        chrt = wrkbk.add_chart({'type': 'line'})
        chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                         'values': [county + dictofcolumns[clmn][0], 1, 8, len(temp_df) - 1, 8],
                         'data_labels': {'value': False},
                         'name': 'Fatal & Incapacitating Injury',
                         'line': {'color': 'red'}})
        chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                         'values': [county + dictofcolumns[clmn][0], 1, 9, len(temp_df) - 1, 9],
                         'data_labels': {'value': False},
                         'name': 'Non-Incapacitating & Possible Injury',
                         'line': {'color': 'yellow'}})
        chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                         'values': [county + dictofcolumns[clmn][0], 1, 5, len(temp_df) - 1, 5],
                         'data_labels': {'value': False},
                         'name': 'Property Damage',
                         'overlap': ovrlp,
                         'line': {'color': '#00B050'}})  # This is a green lighter than default 'green'
        chrt.set_y_axis({'visible': True,
                         'major_gridlines': {
                             'visible': False
                         }})
        # chrt.set_x_axis({'major_tick_mark': 'none'})
        chrt.set_legend({'position': 'bottom'})

    else:
        chrt = wrkbk.add_chart({'type': 'column'})
        if clmn == "U1_CONT_CIR_PRIMARY_CD" or clmn == "CRASH_TYPE_CD" or clmn == "ODOT_CRASH_LOCATION_CD":
            chrt.set_title({'name': yrs + ' ' + county + ' County Crashes by ' + dictofcolumns[clmn][0]})
            chrt.set_y_axis({'visible': False,
                             'major_gridlines': {
                                 'visible': False
                             }})
            chrt.set_legend({'none': True})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 7, len(temp_df) - 1, 7],
                             'data_labels': {'value': True,
                                             'font': {'rotation': -45}},
                             'name': 'All Crashes'})
        elif clmn == "U1_AGE_NBR":

            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 8, len(temp_df) - 1, 8],
                             'data_labels': {'value': True,
                                             'font': {'rotation': -90,
                                                      'color': 'red',
                                                      'size': 8}},
                             'name': 'Fatal & Incapacitating Injury',
                             'fill': {'color': 'red'}})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 9, len(temp_df) - 1, 9],
                             'data_labels': {'value': False},
                             'name': 'Non-Incapacitating & Possible Injury',
                             'fill': {'color': 'yellow'}})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 5, len(temp_df) - 1, 5],
                             'data_labels': {'value': False},
                             'name': 'Property Damage',
                             'overlap': ovrlp,
                             'fill': {'color': '#00B050'}})  # This is a green lighter than default 'green'
            chrt.set_plotarea({'fill': {'color': '#BFBFBF'}})
            chrt.set_y_axis({'major_gridlines':
                                 {'visible': True,
                                  'line': {'color': '#FFFFFF'}},
                             'major_tick_mark': 'none'
                             })
            # chrt.set_x_axis({'major_tick_mark': 'none'})
            chrt.set_legend({'position': 'bottom'})
        # el:
        elif clmn == 'DAY_IN_WEEK_CD':
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 8, len(temp_df) - 1, 8],
                             'data_labels': {'value': True},
                             'name': 'Fatal & Incapacitating Injury',
                             'fill': {'color': 'red'}})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 9, len(temp_df) - 1, 9],
                             'data_labels': {'value': True},
                             'name': 'Non-Incapacitating & Possible Injury',
                             'fill': {'color': 'yellow'}})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 5, len(temp_df) - 1, 5],
                             'data_labels': {'value': True},
                             'name': 'Property Damage',
                             'overlap': ovrlp,
                             'fill': {'color': '#00B050'}})  # This is a green lighter than default 'green'

            chrt.set_y_axis({'visible': False,
                             'major_gridlines': {
                                 'visible': False
                             }})
            # chrt.set_x_axis({'major_tick_mark': 'none'})
            chrt.set_legend({'position': 'bottom'})
        else:
            # elif clmn == 'TIME_OF_CRASH':
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 8, len(temp_df) - 1, 8],
                             'data_labels': {'value': False},
                             'name': 'Fatal & Incapacitating Injury',
                             'fill': {'color': 'red'}})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 9, len(temp_df) - 1, 9],
                             'data_labels': {'value': False},
                             'name': 'Non-Incapacitating & Possible Injury',
                             'fill': {'color': 'yellow'}})
            chrt.add_series({'categories': [county + dictofcolumns[clmn][0], 1, 0, len(temp_df) - 1, 0],
                             'values': [county + dictofcolumns[clmn][0], 1, 5, len(temp_df) - 1, 5],
                             'data_labels': {'value': False},
                             'name': 'Property Damage',
                             'overlap': ovrlp,
                             'fill': {'color': '#00B050'}})  # This is a green lighter than default 'green'
            chrt.set_y_axis({'visible': True,
                             'major_gridlines': {
                                 'visible': False
                             }})
            # chrt.set_x_axis({'major_tick_mark': 'none'})
            chrt.set_legend({'position': 'bottom'})

    return chrt
