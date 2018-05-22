import pandas
from mtfcToGCAT import timeOfDayMI
from funktions import usr
from docx import Document
print(usr)
for key in timeOfDayMI:
    print(key, timeOfDayMI[key])
'''severityCd = {
        1: 'Property Damage Only Crashes',
        2: 'Possible Injury Crashes',
        3: 'Non-Incapacitating Injury Crashes',
        4: 'Incapacitating Crashes',
        5: 'Fatal Crashes'
}
txt = "C:/Users/Michael/OneDrive/Documents/Work Computer Files/GCAT_20Mar17_41337.txt"
dftxt = pandas.read_table(txt, sep=",", dtype={'ODPS_LOC_ROAD_NAME_TXT': str, 'ODPS_REF_GIVEN_TXT': str})
#dftxt["SEVERITY_BY_TYPE_CD"] = dftxt["SEVERITY_BY_TYPE_CD"].map(severityOH)
dftxt["SEVERITY_BY_TYPE_CD"] = dftxt["SEVERITY_BY_TYPE_CD"].map(severityCd)
dftxt["SEVERITY"] = dftxt["SEVERITY_BY_TYPE_CD"].astype('category', ordered=False)
dftxt["SEVERITY"] = dftxt["SEVERITY"].cat.set_categories([severityCd[x] for x in severityCd], ordered=False)
print(dftxt["SEVERITY"].sort_values())

print(pandas.pivot_table(dftxt,"DOCUMENT_NBR",["NLF_COUNTY_CD","CRASH_YR"],"SEVERITY_BY_TYPE_CD",aggfunc=len))'''
# link to MI allroads dataset

url = 'http://gis-michigan.opendata.arcgis.com/datasets/5fddb53b929b4b729b8d4282b4d23ade_20.geojson'
#df = pandas.read_(url)
#print(len(df))

df = pandas.DataFrame(data=[1,2,3,4,5],index=['a','b','c','d','e'])
print(df)
writer = pandas.ExcelWriter('C:/Users/fullerm/Desktop/sandbox.xlsx')
df.to_excel(writer)
wrkbk = writer.book
wrksht = writer.sheets['Sheet1']
chrt = wrkbk.add_chart({'type':'column'})

chrt.add_series({'categories': '=Sheet1!A2:A6',
                 'values': '=Sheet1!B2:B6',
                 'data_labels': {'value':True,
                                 'font': {'rotation':-45}},
                 })

wrksht.insert_chart("E1",chrt)
writer.save()