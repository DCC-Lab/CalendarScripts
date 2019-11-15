# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 09:10:01 2019

@author: Simon
"""

import calendar
cal= calendar.Calendar()
import os
import pandas as pd # Import pandas
import time
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import numpy as np
import easygui
import matplotlib.dates as mdates
#% %
file  = easygui.fileopenbox(msg='Get Google Calendar data', title=None, 
                    default='Y:\Simon\Platforme_Imagerie\GoogleCalendar_time' , 
                    filetypes='*.xlsx', multiple=False)
#% %   
#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\SlideScanner"
#file = 'SlideScanner_user_Final.xlsx'

# Read Microscope reservation file
#os.chdir(directory)
data= pd.read_excel(file)
#% %
alldates=[]
alldates_str=[]
for i in range(1,7):
    for x in cal.itermonthdates(2018, i):
        if x.month == i:
            alldates.append(x)
            alldates_str.append(x.strftime('%d.%m.%Y'))

#%%
# Read lab members
os.chdir("Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Lab_Members")
MemberName = 'Members_LabList.xlsx'
members = pd.read_excel(MemberName)
lablist = members.columns


Reserved = []
for singleday in alldates_str:
    print(singleday)
    for i in range(len(data)):
        if singleday in data.loc[i,'Start']:
            Reserved.append(data.loc[i]) 
            
# % %  Build reservations            
Reservation=[]
for i in range(len(Reserved)):
    start = time.strptime(Reserved[i]['Start'],'%d.%m.%Y %H:%M')
    End = time.strptime(Reserved[i]['End'],'%d.%m.%Y %H:%M')
    sameday = Reserved[i]['Start'][0:10] == Reserved[i]['End'][0:10]
    if sameday == True:  #If reservation start on sameday as end
            x= Reserved[i]['Start'][0:10]
            y = start.tm_hour + start.tm_min / 60
            width = 0.75
            height=Reserved[i]['Hours']
            Reservation.append([x,y,width,height,Reserved[i]['Lab']])
    elif sameday == False: #Overnight
        # make reservation on first day till midnight
            x= Reserved[i]['Start'][0:10]
            y = start.tm_hour + start.tm_min / 60
            width = 0.75
            height=24
            Reservation.append([x,y,width,height,Reserved[i]['Lab']])
        # compute for next day    
            dayindex = alldates_str.index(Reserved[i]['Start'][0:10])
            nextday = alldates_str[dayindex+1]  
            while nextday != Reserved[i]['End'][0:10]: # More than one day
                x= nextday
                y = 0
                width = 0.75
                height=24
                Reservation.append([x,y,width,height,Reserved[i]['Lab']])
                dayindex = alldates_str.index(nextday)
                nextday = alldates_str[dayindex+1]
        # finish Reservation    
            x= nextday
            y = 0
            width = 0.75
            height= End.tm_hour + End.tm_min / 60
            Reservation.append([x,y,width,height,Reserved[i]['Lab']])
#% %            
plt.close(2)
fig = plt.figure(2,figsize=(12,8))
currentAxis = plt.gca()  
ccmap = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231',
             '#911eb4', '#46f0f0', '#f032e6', '#bcf60c', '#fabebe', '#008080',
             '#e6beff', '#9a6324', '#fffac8', '#800000', '#aaffc3', '#808000', 
             '#ffd8b1', '#000075', '#808080', '#ffffff', '#000000']

day=-1
for i in range(0,len(alldates_str)):
    day += 1
    for j in range(len(Reservation)):
        if alldates_str[i] == Reservation[j][0]:
            x=day
            y = Reservation[j][1]
            width = 0.8
            height=Reservation[j][3]
            
            for lab in range(len(lablist)):
                if lablist[lab] == Reservation[j][-1]:
                 colorrect = ccmap[lab]   
                 lablabel =lablist[lab]
                 currentAxis.add_patch(Rectangle((x, y), width,height,alpha=1,
                                                 facecolor=colorrect,edgecolor='black', label=lablabel))


plt.yticks(np.arange(0,25), np.arange(0,25))
#plt.xticks(np.arange(1,len(alldates_str)), np.arange(1,len(alldates_str)))  


plt.xticks(np.arange(1,len(alldates_str)+1), alldates )  
handles, labels = currentAxis.get_legend_handles_labels()
handle_list, label_list = [], []
for handle, label in zip(handles, labels):
    if label not in label_list:
        handle_list.append(handle)
        label_list.append(label)
        
        
plt.legend(handle_list, label_list,loc='lower right')
#import matplotlib.dates as mdates
currentAxis.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday = 1, interval=1))
#currentAxis.xaxis.set_major_formatter(mdates.DateFormatter('%d'))
#currentAxis.get_xaxis().set_major_locator(mdates.MonthLocator(interval=2))
currentAxis.get_xaxis().set_major_formatter(mdates.DateFormatter('%m'))
#fig.autofmt_xdate()
plt.gcf().autofmt_xdate()
plt.show()


            
            
          

#%% Define all reservations

