# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 14:42:06 2019

@author: Simon
"""

# Google_Calendars_HourInUSe

import os
import pandas as pd # Import pandas
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.dates as mdates
from datetime import datetime
import time 
from matplotlib.patches import Rectangle

#% %  Get Microscope Calendar
#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Levesque-ZeissLSM700"
#file = 'Levesque-ZeissLSM700_3_user_Final.xlsx'
#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\SlideScanner"
#file = 'SlideScanner_user_Final.xlsx'

#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\YDK_880"
#file = 'YDK Zeiss880_Scientifica_user.xlsx'
#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Parent"
#file = 'NouveauConfocalDrParent_user.xlsx'

#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\PDK-LSM510"
#file = 'PDK-LSM510_user.xlsx'

#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Inverse_7500"
#file = 'POM-Microscope_inverse_F-7500_user.xlsx'

directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\PDK-TIRF"
file = 'PDK_TIRF_user.xlsx'


os.chdir(directory)
data= pd.read_excel(file)

# Define lab membership
Labs = ["Levesque", "Parent","Deschenes", 
        "PDK" , "YDK", "Toth","SAG","JK","CFS",
        "JPJ","Labonte", "Menard","cicchetti",'Ethier',
        'Timofeev','Proulx','POM',float("NaN")]

#% %
#get data for march 2018
#MonthData = []
#for i in range(len(data)):
#   month = '03.2018'
#   if month in data.loc[i,'Start']:
#    MonthData.append(data.loc[i])      
#    
#    
#OwnerLab = "Levesque"
#% %
alltime = []
for findmonth in range(1,13):
#findmonth=1
    MonthData=[]
    for i in range(len(data)):
        month = "{:02d}".format(findmonth) + '.2018'
        if month in data.loc[i,'Start']:
            MonthData.append(data.loc[i])   
                   
    ReservationLab=[]
    ReservationHours=[]
    for i in range(len(MonthData)):               
        ReservationLab.append(MonthData[i]['Lab'])
        ReservationHours.append(MonthData[i]['Hours'])
    
    Labs_time = []
    for l in range(len(Labs)):
        Monthtime=float(0)  
        indexes = [i for i, e in enumerate(ReservationLab) if e == Labs[l]]
        for j in range(len(indexes)):
            Monthtime = (Monthtime + ReservationHours[indexes[j]])
        Labs_time.append(Monthtime)
    print(Labs_time)
    alltime.append(Labs_time)
    
# %
import calendar
monthname = []
for i in range(1,13):
    monthname.append(calendar.month_name[i])
#% %
ccmap = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231',
             '#911eb4', '#46f0f0', '#f032e6', '#bcf60c', '#fabebe', '#008080',
             '#e6beff', '#9a6324', '#fffac8', '#800000', '#aaffc3', '#808000', 
             '#ffd8b1', '#000075', '#808080', '#ffffff', '#000000']

#plt.close(1)
#fig = plt.figure(1,figsize=(12,8))
fig = plt.figure(1)
currentAxis = plt.gca()  
Totaltime = []
for labo in range(len(Labs)):
    timeused = []
    for month in range(len(alltime)):
        timeused.append(alltime[month][labo])
        if isinstance(Labs[labo], str) :
            Timelabel = Labs[labo] + '_Total_time=' + str(sum(timeused))
        else : 
            Timelabel = 'Unknown' + '_Total_time=' + str(sum(timeused))
    if sum(timeused) > 0 :
        plt.plot(monthname,timeused,label = Timelabel,color = ccmap[labo])

plt.legend()




