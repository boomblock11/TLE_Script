# -*- coding: utf-8 -*-
#Full script
#Import basic utilities
import datetime as dt
import numpy as np
import os
import csv

#Support for COM
from comtypes.client import CreateObject
from comtypes.client import GetActiveObject



#Read in GPS data
with open(r'.\APID02310.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    next(readCSV)
    ephemLines = []
    for row in readCSV:
        time = row[1].replace(" ","T")
        x = row[34]
        y = row[35]
        z = row[36]
        xdot = row[37]
        ydot = row[38]
        zdot= row[39]
        ephemLines.append(f"{time} {x} {y} {z} {xdot} {ydot} {zdot}")
        
#Format data for ephemeris file
lineCount = len(ephemLines)
ephemLines.sort()
starttime=ephemLines[0]
starttime=starttime[0:23]
endtime=ephemLines[lineCount-1]
endtime=endtime[0:23]
ephemLines.insert(0,"stk.v.5.0")
ephemLines.insert(1,"BEGIN Ephemeris")
ephemLines.insert(2,f"NumberOfEphemerisPoints {lineCount}")
ephemLines.insert(3,"InterpolationMethod     Lagrange")
ephemLines.insert(4,"InterpolationOrder      5")
ephemLines.insert(5,"DistanceUnit           meters")
ephemLines.insert(6,"CentralBody             Earth")
ephemLines.insert(7,"TimeFormat ISO-YMD")
ephemLines.insert(8,"CoordinateSystem        Inertial")
ephemLines.insert(9,"EphemerisTimePosVel")
ephemLines.append("END Ephemeris")

#Write Data to ephemeris file        
ephem=open('ephemeris.e','w')
for line in ephemLines:
    ephem.write(line)
    ephem.write('\n')
ephem.close()


#Start STK11 Application
#app = GetActiveObject("STK11_x64.Application")
app = CreateObject("STK11_x64.Application")
app.Visible = True
app.UserControl = True

#Obtain reference to root object
root = app.Personality2

from comtypes.gen import STKObjects
from comtypes.gen import STKUtil

#Create new scenario
root.NewScenario("ScriptTesting")
scenario = root.CurrentScenario

#Access scenario


#Create satellite
satellite = scenario.Children.New(STKObjects.eSatellite, "TestSatellite")
satellite2 = satellite.QueryInterface(STKObjects.IAgSatellite)

#Set data
satellite2.SetPropagatorType(STKObjects.ePropagatorStkExternal)
exProp = satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorStkExternal)
exProp.Filename=r'.\ephemeris.e'
exProp.Propagate()

#Generate the tle
unitsCorr="Units_Set * Connect Date ISO-YMD"
root.ExecuteCommand(unitsCorr)
TLECmd = f"GenerateTLE */Satellite/TestSatellite Sampling \"{starttime}\" \"{endtime}\" 60.0 \"{endtime}\" 12345 20 .000000001"
#TLECmd = f"GenerateTLE */Satellite/TestSatellite Point \"{endtime}\" 12345 20 .0000000001"
root.ExecuteCommand(TLECmd)

#Display TLE
reportTLE = "ReportCreate */Satellite/TestSatellite Type Display Style \"TLE\""
root.ExecuteCommand(reportTLE)
#reportAccess = "ReportCreate */Satellite/Shuttle Type Display Style \"Access\" File \"c:\Data\shuttlewalreport.txt\" AccessObject */Facility/Wallops"