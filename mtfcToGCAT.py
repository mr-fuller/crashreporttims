import pandas, os
import pandas as pd
import numpy as np
# from txtCleanUp import txtCleanUp
from funktions import usr
# usr = 'Michael'
#can I just read the time of day and slice[:8]?
timeOfDayMI = {
        '1:00 AM - 2:00 AM': '1:00 AM',
        '9:00 AM - 10:00 AM': '9:00 AM',
        '3:00 PM - 4:00 PM': '3:00 PM',
        '3:00 AM - 4:00 AM': '3:00 AM',
        '6:00 AM - 7:00 AM': '6:00 AM',
        '5:00 AM - 6:00 AM': '5:00 AM',
        '11:00 AM - 12:00 noon': '11:00 AM',
        '10:00 AM - 11:00 AM': '10:00 AM',
        '8:00 AM - 9:00 AM': '8:00 AM',
        '12:00 noon - 1:00 PM': 'Noon',
        '4:00 PM - 5:00 PM': '4:00 PM',
        '5:00 PM - 6:00 PM': '5:00 PM',
        '12:00 midnight - 1:00 AM': 'Midnight',
        '6:00 PM - 7:00 PM': '6:00 PM',
        '11:00 PM - 12:00 midnight': '11:00 PM',
        '7:00 AM - 8:00 AM': '7:00 AM',
        '7:00 PM - 8:00 PM': '7:00 PM',
        '9:00 PM - 10:00 PM': '9:00 PM',
        '2:00 AM - 3:00 AM': '2:00 AM',
        '2:00 PM - 3:00 PM': '2:00 PM',
        '1:00 PM - 2:00 PM': '1:00 PM',
        '10:00 PM - 11:00 PM': '10:00 PM',
        '4:00 AM - 5:00 AM': '4:00 AM',
        # 'Unknown':'Unknown',
        '8:00 PM - 9:00 PM': '8:00 PM'

    }
severityMI = {
        'Possible injury (C)': 'Possible Injury Crashes',
        'No injury (O)': 'Property Damage Only Crashes',
        'Suspected minor injury (B)': 'Non-Incapacitating Injury Crashes',
        'Fatal injury (K)': 'Fatal Crashes',
        'Suspected serious injury (A)': 'Incapacitating Injury Crashes'
    }
severityOH = {
        1: 'Property Damage Only Crashes',
        2: 'Possible Injury Crashes',
        3: 'Non-Incapacitating Injury Crashes',
        4: 'Incapacitating Injury Crashes',
        5: 'Fatal Crashes'
    }
typeCd = {
        0: 'Unknown',
        1: 'Head On',
        2: 'Rear End',
        3: 'Backing',
        4: 'Sideswipe – Meeting',
        5: 'Sideswipe – Passing',
        6: 'Angle',
        7: 'Parked Vehicle',
        8: 'Pedestrian',
        9: 'Animal',
        10: 'Train',
        11: 'Pedalcycles',
        12: 'Other Non-Vehicle',
        13: 'Fixed Object',
        14: 'Other Object',
        15: 'Falling From or In Vehicle',
        16: 'Overturning',
        17: 'Other Non-Collision',
        18: 'Left Turn',
        19: 'Right Turn'  # TIMS data had an additional type

    }
severityCd = {
        1: 'Property Damage Only Crashes',
        2: 'Possible Injury Crashes',
        3: 'Non-Incapacitating Injury Crashes',
        4: 'Incapacitating Crashes',
        5: 'Fatal Crashes'

    }
loctypeCd = {
        1: 'Not An Intersection',
        2: 'Four-Way Intersection',
        3: 'T-Intersection',
        4: 'Y-Intersection',
        5: 'Traffic Circle/Roundabout',
        6: '5 Or More Point Intersection',
        7: 'On Ramp',
        8: 'Off Ramp',
        9: 'Crossover',
        10: 'Driveway/Alley Access',
        11: 'Railroad Grade Crossing',
        12: 'Shared-Use Paths Or Trails',
        99: 'Unknown'

    }
contributingFactor = {
        1: 'None',
        2: 'Failure To Yield',
        3: 'Ran Red Light',
        4: 'Ran Stop Sign',
        5: 'Exceeded Speed Limit',
        6: 'Unsafe Speed',
        7: 'Improper Turn',
        8: 'Left Of Center',
        9: 'Followed Too Closely/ACDA',
        10: 'Improper Lane Change/Passing/Offroad',
        11: 'Improper Backing',
        12: 'Improper Start From Parked Position',
        13: 'Stopped Or Parked Illegally',
        14: 'Operating Vehicle In Negligent Manner',
        15: 'Swerving To Avoid',
        16: 'Wrong Side/Wrong Way',
        17: 'Failure To Control',
        18: 'Vision Obstruction',
        19: 'Operating Defective Equipment',
        20: 'Load Shifting/Falling/Spilling',
        21: 'Other Improper Action',
        22: 'None Non-Motorist',
        23: 'Improper Crossing',
        24: 'Darting',
        25: 'Lying And/Or Illegally In Roadway',
        26: 'Failure To Yield Right Of Way',
        27: 'Not Visible (Dark Clothing)',
        28: 'Inattentive',
        29: 'Failure To Obey Signs/Signals/Officer',
        30: 'Wrong Side Of The Road',
        31: 'Other Non-Motorist',
        99: 'Unknown'

    }
vehicleType = {
        1: 'Sub-Compact',
        2: 'Compact',
        3: 'Mid-Size',
        4: 'Full Size',
        5: 'Minivan',
        6: 'Sport Utility Vehicle',
        7: 'Pickup',
        8: 'Van',
        9: 'Motorcycle',
        10: 'Motorized Bicycle',
        11: 'Snowmobile/ATV',
        12: 'Other Passenger Vehicle',
        13: 'Single Unit Truck Or Van 2 Axle, 6 Tire',
        14: 'Single Unit Truck; 3+ Axles',
        15: 'Single Unit Truck/Trailer',
        16: 'Truck/Tractor (Bobtail)',
        17: 'Tractor/Semi-Trailer',
        18: 'Tractor/Double',
        19: 'Tractor/Triples',
        20: 'Other Med/Heavy Vehicle',
        21: 'Bus/Van (9-15 Seats Including Driver)',
        22: 'Bus (16+ Seats Including Driver)',
        23: 'Animal With Rider',
        24: 'Animal With Buggy, Wagon, Surrey',
        25: 'Bicycle/Pedalcyclist',
        26: 'Pedestrian/Skater',
        27: 'Other Non-Motorist',
        99: 'Unknown Or Hit/Skip'
    }
weekday = {
        1: 'Sunday',
        2: 'Monday',
        3: 'Tuesday',
        4: 'Wednesday',
        5: 'Thursday',
        6: 'Friday',
        7: 'Saturday'

    }
timeOfDayOH = {
        1: "Midnight",
        2: "1:00 AM",
        3: "2:00 AM",
        4: "3:00 AM",
        5: "4:00 AM",
        6: "5:00 AM",
        7: "6:00 AM",
        8: "7:00 AM",
        9: "8:00 AM",
        10: "9:00 AM",
        11: "10:00 AM",
        12: "11:00 AM",
        13: "Noon",
        14: "1:00 PM",
        15: "2:00 PM",
        16: "3:00 PM",
        17: "4:00 PM",
        18: "5:00 PM",
        19: "6:00 PM",
        20: "7:00 PM",
        21: "8:00 PM",
        22: "9:00 PM",
        23: "10:00 PM",
        24: "11:00 PM"

    }
truckbus = {
    13: 1,
    14: 1,
    15: 1,
    16: 1,
    17: 1,
    18: 1,
    19: 1,
    20: 1
    #21: 1,
    #22: 1
}
# These bins exclude all records where U1 or U2 was blank or 0
bins = [0, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 99]
timebins = [0, 100, 200, 300, 400, 500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700, 1800, 1900,
            2000, 2100, 2200, 2300, 2400]
print(timebins)
bin_lbls = {
        0: 'Under 16',
        1: '16-20',
        2: '21-25',
        3: '26-30',
        4: '31-35',
        5: '36-40',
        6: '41-45',
        7: '46-50',
        8: '51-55',
        9: '56-60',
        10: '61-65',
        11: '66-70',
        12: '71-75',
        13: '76-80',
        14: '81-85',
        15: 'Over 85'
    }  # make this a dict?
dictofcolumns = {
        "CRASH_TYPE_CD": ["Type of Crash", typeCd],
        "ODOT_CRASH_LOCATION_CD": ["Location", loctypeCd],
        "U1_CONT_CIR_PRIMARY_CD": ["Contributing Factor", contributingFactor],  # 3 columns to consolidate
        "U1_TYPE_OF_UNIT_CD": ["Vehicle Type", vehicleType],  # 3 columns to consolidate
        "DAY_IN_WEEK_CD": ["Day of Week", weekday],
        "TIME_OF_CRASH": ["Time of Day", timeOfDayOH],
        # "UNRESTRAIN_OCCUPANTS":["Unrestrained Occupants",{}],
        "U1_AGE_NBR": ["Age", bin_lbls]  # 3 columns to consolidate

    }
def schemaCleaner(txt,mtcfCsv,GDBspot):


    # txt = "C:/Users/Michael/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    # txt = "C:/Users/fullerm/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    # txt = txt
    # txt = "C:/Users/fullerm/Documents/ArcGIS/Projects/Work Computer Files/GCAT_20Mar17_41337.txt"
    # txt = "C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/GCAT_12to14_LOSW_19May2017_47554.txt"
    # This section cleans up GCAT file
    # txtCleanUp(txt)


    #mtcfCsv = "C:/Users/fullerm/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    #mtcfCsv = "C:/Users/Michael/OneDrive/Documents/Work Computer Files/MTCF_5June17_2909.csv"
    mtcfdf = pandas.read_csv(mtcfCsv)
    dftxt = pandas.read_csv(txt, dtype={'ODPS_LOC_ROAD_NAME_TXT': str, 'ODPS_REF_GIVEN_TXT': str})
    dftxt["SEVERITY_BY_TYPE_CD"] = dftxt["SEVERITY_BY_TYPE_CD"].map(severityOH)
    dftxt["DISTRACTED_DRIVER_IND"] = dftxt["DISTRACTED_DRIVER_IND"].map({'N':0,'Y':1})
    dftxt["ODOT_YOUNG_DRIVER_IND"] = dftxt["ODOT_YOUNG_DRIVER_IND"].map({'N':0,'Y':1})
    dftxt["ODPS_PEDESTRIAN_IND"] = dftxt["CRASH_TYPE_CD"].map({8:1})
    dftxt["ODPS_BICYCLE_IND"] = dftxt["CRASH_TYPE_CD"].map({11:1})

    dftxt["TRUCK_COUNT"] = dftxt["U1_TYPE_OF_UNIT_CD"].map(truckbus).fillna(0) + \
                           dftxt["U2_TYPE_OF_UNIT_CD"].map(truckbus).fillna(0) + \
                           dftxt["U3_TYPE_OF_UNIT_CD"].map(truckbus).fillna(0)

    print(dftxt["TRUCK_COUNT"])
    #dftxt["Hour_of_Crash"] = pd.cut(dftxt["TIME_OF_CRASH"],bins= 24,labels= [timeOfDayOH[x] for x in timeOfDayOH])
    #print(dftxt[["TIME_OF_CRASH","Hour_of_Crash"]])

    ## get the years of the crash report from the data
    #yrs = set(dftxt['CRASH_YR'])
    yrs = np.unique(dftxt['CRASH_YR'].tolist())
    for yr in yrs:
        print(yr)
    print(str(yrs[0]) + '-' +str(yrs[len(yrs)-1]) + ' county category of crashes')
    for clmn in dictofcolumns:
        if clmn == "U1_AGE_NBR":
            dftxt[dictofcolumns[clmn][0] + "1"] = pd.cut(dftxt[clmn], bins, labels=[bin_lbls[key] for key in bin_lbls])
            dftxt[dictofcolumns[clmn][0] + "2"] = pd.cut(dftxt['U2' + clmn[2:]], bins,
                                                         labels=[bin_lbls[key] for key in bin_lbls])

            # indx[1] = "Age Range"
        elif clmn == "U1_CONT_CIR_PRIMARY_CD" or clmn == "U1_TYPE_OF_UNIT_CD":
            dftxt[dictofcolumns[clmn][0] + "1"] = dftxt[clmn].map(dictofcolumns[clmn][1], na_action='ignore')
            dftxt[dictofcolumns[clmn][0] + "2"] = dftxt['U2' + clmn[2:]].map(dictofcolumns[clmn][1],
                                                                             na_action='ignore').replace('None', np.nan)
            dftxt[dictofcolumns[clmn][0] + "3"] = dftxt['U3' + clmn[2:]].map(dictofcolumns[clmn][1],
                                                                             na_action='ignore').replace('None',np.nan)
            # print(dftxt[dictofcolumns[clmn][0] + "2"])
        elif clmn == "TIME_OF_CRASH":
            dftxt["Hour_of_Crash"] = pd.cut(dftxt["TIME_OF_CRASH"], bins=timebins,
                                            labels=[timeOfDayOH[x] for x in timeOfDayOH], right=False) #use right=False so, for example, 1000 is counted as 10:00 AM.
            dftxt["Hour_of_Crash"] = dftxt["Hour_of_Crash"].astype('str')

        else:
            dftxt[clmn] = dftxt[clmn].map(dictofcolumns[clmn][1])
    dftxt.to_csv("C:/Users/"+usr+"/Desktop/test2df.csv")
    print(dftxt[['TIME_OF_CRASH','Hour_of_Crash']])
    emptyDF = pd.DataFrame(columns=dftxt.columns)
    emptyDF['SEVERITY_BY_TYPE_CD'] = mtcfdf['Worst Injury in Crash'].map(severityMI)
    emptyDF['CRASH_TYPE_CD'] = mtcfdf['Crash Type']  # Area of Road at Crash?
    emptyDF['DAY_IN_WEEK_CD'] = mtcfdf['Day of Week']  # Crash Day = date 1-31? Wednesday = 1?
    emptyDF['Hour_of_Crash'] = mtcfdf['Time of Day'].map(timeOfDayMI)
    emptyDF['ODPS_ALCOHOL_IND'] = mtcfdf['Crash: Drinking'].map({'No drinking involved': 0, 'Drinking involved': 1})
    emptyDF['ODPS_DRUG_IND'] = mtcfdf['Crash: Drug Use'].map({'No drugs involved': 0, 'Drugs involved': 1})
    emptyDF['ODPS_PEDESTRIAN_IND'] = mtcfdf['Crash: Pedestrian'].map(
        {'No pedestrian involved': 0, 'Pedestrian involved': 1})
    emptyDF['ODPS_BICYCLE_IND'] = mtcfdf['Crash: Bicyclist'].map({'No bicyclist involved': 0, 'Bicyclist involved': 1})
    emptyDF['CRASH_YR'] = mtcfdf['Crash Year']
    emptyDF['DOCUMENT_NBR'] = mtcfdf['Crash Instance']  # I am concerned about duplicates here
    emptyDF['ODPS_MOTORCYCLE_IND'] = mtcfdf['Crash: Motorcycle'].map(
        {'No motorcycle involved': 0, 'Motorcycle involved': 1})
    emptyDF['ODOT_LONGITUDE_NBR'] = mtcfdf['Crash Longitude']  # no linear referencing
    emptyDF['ODOT_LATITUDE_NBR'] = mtcfdf['Crash Latitude']
    emptyDF['NLF_COUNTY_CD'] = mtcfdf['County']  # can also get this with a spatial join
    emptyDF['U1_CONT_CIR_PRIMARY_CD'] = mtcfdf['Contributing Circumstances Road 1 (2016+)']
    emptyDF['U2_CONT_CIR_PRIMARY_CD'] = mtcfdf['Contributing Circumstances Road 2 (2016+)']
    emptyDF['ODOT_CRASH_LOCATION_CD'] = mtcfdf['Area of Road at Crash']
    emptyDF['TRUCK_COUNT'] = mtcfdf['Crash: Truck or Bus'].map({'Truck or bus involved':1})
    emptyDF['ODOT_CITY_VILLAGE_TWP_NAME'] = mtcfdf['City or Township']
    #emptyDF['ODPS_SENIOR_DRIVER_IND'] = mtcfdf['Crash: Senior Driver']
    # emptyDF.set_index("DOCUMENT_NBR",inplace=True)
    print(len(emptyDF))
    crashType = [
        'Single motor vehicle',
        'Other / unknown',
        'Head-on',
        'Rear-end',
        'Sideswipe same direction',
        'Angle',
        'Sideswipe opposite direction',
        'Rear-end left turn',
        'Head-on / left turn',
        'Other',
        'Rear-end right turn',
        'Unknown'

    ]

    # probably won't need contributing circumstances for the next update if they just started keeping track in 2016
    '''contcircumstances ={
    Uncoded & errors
    Glare
    Backup - regular congestion
    Other
    None
    Traffic control device
    Unknown
    Backup - other incident
    Prior crash
    Shoulders

    }'''

    dftxt = dftxt.append(emptyDF, ignore_index=True)
    dftxt.index.name = "GUID"
    dftxt.to_csv(os.path.join(GDBspot,"OHMI_data.csv"))
    return dftxt, os.path.join(GDBspot,"OHMI_data.csv")

if __name__ == "__main__":
    txt = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
    mtcfCsv = "C:/Users/" + usr + "/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
    schemaCleaner(txt, mtcfCsv, "C:/Users/" + usr + "/Desktop")