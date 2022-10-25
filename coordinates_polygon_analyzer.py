import os
from os import path
import glob
from pykml import parser
from shapely.geometry import Point, Polygon
import ast
import pandas as pd
from datetime import datetime
import ctypes
import sys

#read coordinates points file
df = pd.read_excel("polygons_and_coordinates/coordinates.xlsx")
df.columns = df.columns.str.lower()

#rows number
#introws = len(df.index) -1 #-1 from header

#sort directory and quantity
lstdir = (glob.glob(os.path.join('polygons_and_coordinates', '*.kml')))
lstdirsorted = sorted(lstdir)
lendir = len(lstdirsorted)
print(str(lendir) + " .kml files detected")

for filepath in lstdirsorted:
    
    #open .kml
    with open(filepath) as f:
        
        root = parser.parse(f).getroot()
        
        print("Running... " + root.Document.name.text)
        
        #get coordinates tag
        strpolygon = root.Document.Placemark.LineString.coordinates.text
        
        #cleaning
        strpolygon = strpolygon.strip()
        strpolygon = "[("+ strpolygon + ")]"
        strpolygon = strpolygon.replace(" ","),(")
        strpolygon = strpolygon.replace(" ","),(")
        strpolygon = strpolygon.replace(",0)",")")
        
        #input
        lstpolygon = ast.literal_eval(strpolygon)
        polygon = Polygon(lstpolygon)
        
        #check coordinates points
        lstbool = []
        for row in df.itertuples():
            lstbool.append(Point(row.longitude,row.latitude).within(polygon))
            
        # datetime object containing current date and time + polygon name
        datenow = datetime.now()
        strdate = root.Document.name.text + " - " + datenow.strftime("%d/%m/%Y %H:%M ")
        
        #append column
        df[strdate] = lstbool

        print(root.Document.name.text + " successfully processed")
        
#save
df.to_excel (r"polygons_and_coordinates/coordinates.xlsx", index = False, header=True)
print("Finished!")

MessageBox = ctypes.windll.user32.MessageBoxW
MessageBox(None, 'Finished', 'Finished!', 0)

