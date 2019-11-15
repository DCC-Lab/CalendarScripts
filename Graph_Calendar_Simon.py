# -*- coding: utf-8 -*-
"""
Created on Tue Mar 19 20:24:08 2019

@author: Simon
"""

import os
import pandas as pd # Import pandas
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.dates as mdates
from datetime import datetime
import time 
from matplotlib.patches import Rectangle

# % %  Get Microscope Calendar
directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Levesque-ZeissLSM700"
file = "Levesque-ZeissLSM700_3_user_Final.xlsx"
#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\SlideScanner"
#file = 'SlideScanner_user_Final.xlsx'

os.chdir(directory)
# Load spreadsheet
data= pd.read_excel(file)
# Define lab membership
Labs = ["Levesque", "Parent","Deschenes", 
        "PDK" , "YDK", "Toth","SAG","JK","CFS",
        "JPJ","Labonte", "Menard","cicchetti",'Ethier','Timofeev']
#% %

#get data for march 2018
MonthData = []
for i in range(len(data)):
   month = '03.2018'
   if month in data.loc[i,'Start']:
    MonthData.append(data.loc[i])       
    
timeStart = []
for i in range(len(MonthData)):
    timeStart.append(time.strptime(MonthData[i]['Start'],'%d.%m.%Y %H:%M'))
    
    
#% %
    plt.close(1)
fig = plt.figure(1,figsize=(12,8))
currentAxis = plt.gca()  
ccmap = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231',
             '#911eb4', '#46f0f0', '#f032e6', '#bcf60c', '#fabebe', '#008080',
             '#e6beff', '#9a6324', '#fffac8', '#800000', '#aaffc3', '#808000', 
             '#ffd8b1', '#000075', '#808080', '#ffffff', '#000000']

# Find Reservation on day 01

for day in range(1,33):
#    day = 25
    reservations=[]
    for i in range(len(MonthData)):
        if timeStart[i].tm_mday == day :
            reservations.append(MonthData[i])
#% %    
    for i in range(len(reservations)):
 #% %       
        x=day
        start = time.strptime(reservations[i]['Start'],'%d.%m.%Y %H:%M')
        y = start.tm_hour + start.tm_min / 60
        width = 0.75
        
        ending = time.strptime(reservations[i]['End'],'%d.%m.%Y %H:%M')
        ending2 = ending.tm_hour + ending.tm_min / 60   
        if ending.tm_mday == start.tm_mday:
            height = ending2-y # if hieght is negative should correct reservation overnight
        else:
            height=24 # do not correct for overnight....
        #% %        
        if reservations[i]['Lab']== 'Levesque' : 
            #colorrect= 'white'
            colorrect= ccmap[0]
            lablabel = 'Levesque'
        elif reservations[i]['Lab']== "Parent" :
            colorrect= ccmap[1]
            lablabel = 'Parent'
        elif reservations[i]['Lab']== "Deschenes" :
            colorrect= ccmap[2]
            lablabel = 'Deschenes'
        elif reservations[i]['Lab']== "PDK" :
            colorrect= ccmap[3]
            lablabel = 'PDK'
        elif reservations[i]['Lab']== "YDK" : 
            colorrect= ccmap[4]    
            lablabel = 'YDK'
        elif reservations[i]['Lab']== "Toth" :
            colorrect= ccmap[5]
            lablabel = 'Toth'
        elif reservations[i]['Lab']== "SAG":
            colorrect= ccmap[6]
            lablabel = 'SAG'
        elif reservations[i]['Lab']== "JK":
            colorrect= ccmap[7]
            lablabel = 'JK'
        elif reservations[i]['Lab']== "CFS":
            colorrect= ccmap[8]
            lablabel = 'CFS'
        elif reservations[i]['Lab']== "JPJ":
            colorrect= ccmap[9]   
            lablabel = 'JPJ'
        elif reservations[i]['Lab']== "Labonte":
            colorrect= ccmap[10]
            lablabel = 'Labonte'
        elif reservations[i]['Lab']== "Menard":
            colorrect= ccmap[11]
            lablabel = 'Menard'
        elif reservations[i]['Lab']== "cicchetti":
            colorrect= ccmap[12]
            lablabel = 'cicchetti'             
        else :
            colorrect= "black"
            lablabel = 'Unknown'
        
        currentAxis.add_patch(Rectangle((x, y), width,height,alpha=1,facecolor=colorrect,edgecolor='black',
                                        label=lablabel ))


plt.xticks(np.arange(1,33), np.arange(1,33))
plt.yticks(np.arange(0,25), np.arange(0,25))

handles, labels = currentAxis.get_legend_handles_labels()
handle_list, label_list = [], []
for handle, label in zip(handles, labels):
    if label not in label_list:
        handle_list.append(handle)
        label_list.append(label)
plt.legend(handle_list, label_list,loc='lower right')



plt.xlabel('Days of ' + month)
plt.ylabel('Time of Day (h)')
#currentcurrentAxisis.legend()

plt.show()
# % %
#plt.savefig('Levesque-ZeissLSM700_' + month + '.pdf')
plt.savefig('Levesque-ZeissLSM700_' + month + '.png')




