# crashreporttims
Scripts to automate crash analysis and report generation at TMACOG. Note that it is important for jurisdictions to identify high-crash 
locations in order to take advantage of certain funding programs/sources.

Program combines data from Ohio DOT's TIMS and [Michigan Traffic Crash Facts](https://www.michigantrafficcrashfacts.org/), counts the number of crashes within 250 
feet of an intersection or along a road segment, then prepares charts and tables in Excel to show how crashes correlate with certain variables.

The program is currently designed as a script tool that runs through ArcGIS Pro. If setting that up, 

1. use main.py as the script 
2. Set up input parameters based on main.py
  - GCATfile: text file of crash data
  - MTCFfile: text file of crash data
  - fatalwt: integer or float
  - seriouswt: integer or float
  - nonseriouswt: integer or float
  - possiblewt: integer or float
  - intersectionThreshold: integer or float
  - SegmentTheshold: integer or float
3. update directory to save files output by the program
4. Create buffered shapefiles/featureclasses/geopackages of the intersections and segments in your jurisdiction
5. Determine Equivalent Property Damage Only (EPDO) weights of more serious crashes
6. Decide on thresholds for intersections and segments (or just set the parameter to zero)
7. Run script
