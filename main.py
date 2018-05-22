from mtfcToGCAT import schemaCleaner
from crashTblYear import tblCrashYear
from crashTblType import crashtypetbl
from crashTbl_F_II import tblCrashF_II
from electronicMail import composeAndSendEmail
from CrashReportLRS import crashReportLRS
from drugalcoholF_II import drugalcoholfii
import arcpy, os
from tblAlcoholDrug import alcoholdrugtbl
from seatbelt import seatbelttbl
# from overall_f_ii import overallfii
from exceltoword import exceltoword
import time
from datetime import datetime
TimeDate = datetime.now()
TimeDateStr = "CrashLocations" + TimeDate.strftime('%Y%m%d%H%M') + "_LRS"
# this version of the script tool is designed to take GCAT data from TIMS,
# which has a slightly different data structure/schema
start = time.time()
print('Started at ' + str(start))
GCATfile = arcpy.GetParameterAsText(0)
# GCATfile = "C:/Users/fullerm/OneDrive/Documents/Work Computer Files/20170609-1536-gcat-results48466.csv"
MTCFfile =arcpy.GetParameterAsText(1)
# MTCFfile = "C:/Users/fullerm/OneDrive/Documents/Work Computer Files/MTCF2012to2014_5June17_2897.csv"
# GDBspot = arcpy.GetParameterAsText(2) #this doesn't fucking change
GDBspot = 'C:/Users/fullerm/Documents/ArcGIS/Projects/Safety Report Script/Script Tool Results/' + TimeDateStr
if not os.path.exists(GDBspot):
    os.makedirs(GDBspot)
fatalwt = float(arcpy.GetParameterAsText(2))  # user input weight for fatal crashes
# fatalwt = 39.22  # user input weight for fatal crashes
seriouswt = float(arcpy.GetParameterAsText(3))  # user input weight for serious crashes
# seriouswt = 39.22  # user input weight for serious crashes
nonseriouswt = float(arcpy.GetParameterAsText(4))  # user input weight for nonserious crashes
# nonseriouswt = 6.84  # user input weight for nonserious crashes
possiblewt = float(arcpy.GetParameterAsText(5))  # user input weight for possible crashes
# possiblewt = 4.63  # user input weight for possible crashes
# curious that these two thresholds don't have to be numbers
IntersectionThreshold = arcpy.GetParameterAsText(6)  # user input number of crashes to qualify an intersection as high crash
# IntersectionThreshold = str(20)  # user input number of crashes to qualify an intersection as high crash
SegmentThreshold = arcpy.GetParameterAsText(7)
# SegmentThreshold = str(30)
lst = schemaCleaner(GCATfile,MTCFfile,GDBspot) # returns list containing df[0] and csv location[1]

# list of xlsx file with top locations[0], and new df with county/municipality based on location[1]
# spatialList =
xlFiles = [
    crashReportLRS(GDBspot, lst[1], fatalwt, seriouswt, nonseriouswt, possiblewt, IntersectionThreshold,
                   SegmentThreshold),
    tblCrashYear(lst[0], GDBspot),
    crashtypetbl(lst[0], GDBspot),
    tblCrashF_II(lst[0], GDBspot),
    # overallfii(lst[0],GDBspot),
    alcoholdrugtbl(lst[0],GDBspot),
    drugalcoholfii(lst[0],GDBspot),
    seatbelttbl(lst[0],GDBspot) # bundled into crashtypetbl; then unbundled to stop stuff from failing/breaking
           ]
worddoc = exceltoword(xlFiles,GDBspot)
xlFiles.append(worddoc)
print(xlFiles,type(xlFiles))
#print(emailAttachments, type(emailAttachments))
composeAndSendEmail(xlFiles)
print('Completed in {0:0.1f}'.format(time.time()-start))
