# -*- coding: utf-8 -*-
"""
Created on Mon Mar 18 11:47:43 2019

@author: Simon
"""

import os
import pandas as pd # Import pandas
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.dates as mdates
from datetime import datetime
import time 
#from matplotlib.dates import date2num

#matplotlib

# % %  Get Microscope Calendar
#directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Levesque-ZeissLSM700"
#file = "Levesque-ZeissLSM700_3_user_Final.xlsx"

directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\SlideScanner"
file = 'SlideScanner_user_Final.xlsx'

os.chdir(directory)
# Load spreadsheet
data= pd.read_excel(file)
# Define lab membership
Labs = ["Levesque", "Parent","Deschenes", 
        "PDK" , "YDK", "Toth","SAG","JK","CFS",
        "JPJ","Labonte", "Menard","cicchetti",'Ethier','Timofeev']
#% %
#  Get data for year 2014
#Start time
d=[]
df = pd.DataFrame(data, columns= ['Start'])
for i in range(len(df)):
    d.append(df.at[i, 'Start'])
dates = [datetime.strptime(ii, '%d.%m.%Y %H:%M') for ii in d]    
timeStart = [time.strptime(ii, '%d.%m.%Y %H:%M') for ii in d]    
#%%
# Find Using time
dfUse = pd.DataFrame(data, columns= ['Hours'])
TimeinUse  = []
for i in range(len(dfUse)):
    TimeinUse.append(dfUse.at[i, 'Hours'])
    
#%%from matplotlib import pyplot as plt
from matplotlib.patches import Rectangle

fig = plt.figure(1)
currentAxis = plt.gca()

#for i in range(len(data)):
#for i in range(1929,len(data)):
#for i in range(2029,2095):
# 2018
for i in range(1929,2850):
    x=i
    y = timeStart[i].tm_hour + timeStart[i].tm_min / 60
    width = 0.75 ;
    height = TimeinUse[i]
    
    if i>0 :
        fd = [timeStart[i-1].tm_year, timeStart[i-1].tm_mon, timeStart[i-1].tm_mday]
        fd2 = [timeStart[i].tm_year, timeStart[i].tm_mon, timeStart[i].tm_mday]
        if fd == fd2:
            x = i-1

    # Define color based on lab
#% %    
    # Define Lab colors
#    cmap = plt.cm.coolwarm
    cmap = plt.cm.seismic
    ccmap = cmap(np.linspace(0, 1, len(Labs)))
    
    
    ccmap = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231',
             '#911eb4', '#46f0f0', '#f032e6', '#bcf60c', '#fabebe', '#008080',
             '#e6beff', '#9a6324', '#fffac8', '#800000', '#aaffc3', '#808000', 
             '#ffd8b1', '#000075', '#808080', '#ffffff', '#000000']
    if data.at[i,'Lab'] in Labs:# Found host lab
        lab = [x for x in Labs if data.at[i,'Lab'] in x]
        if lab[0] is 'Levesque' : 
#            colorrect= 'blue'
#            colorrect= 'white'
            colorrect= ccmap[0]
        elif lab[0] is "Parent" :
#            colorrect= 'red'
            colorrect= ccmap[1]
        elif lab[0] is "Deschenes" :
#            colorrect= 'yellow"
            colorrect= ccmap[2]
        elif lab[0] is "PDK" :
#            colorrect= 'yellow"
            colorrect= ccmap[3]
        elif lab[0] is "YDK" :
#            colorrect= 'blue" 
            colorrect= ccmap[4]    
        elif lab[0] is "Toth" :
            colorrect= ccmap[5]
        elif lab[0] is "SAG":
            colorrect= ccmap[6]
        elif lab[0] is "JK":
            colorrect= ccmap[7]
        elif lab[0] is "CFS":
            colorrect= ccmap[8]
        elif lab[0] is "JPJ":
            colorrect= ccmap[9]   
        elif lab[0] is "Labonte":
            colorrect= ccmap[10]
        elif lab[0] is "Menard":
            colorrect= ccmap[11] 
        elif lab[0] is "cicchetti":
            colorrect= ccmap[12]             
        else :
            colorrect= "black"
            print(lab[0])
        currentAxis.add_patch(Rectangle((x, y), width,height,alpha=1,color=colorrect))


    #% % currentAxis.add_patch(Rectangle((someX, someY2), width, height,alpha=1,color='red'))

start = dates[1929]
stop = dates[2850]

ax.plot((start, stop), (0, 0), 'k', alpha=.5)

plt.show()
plt.xlim([0,len(data)])
#plt.xlim([2029,2089])
plt.ylim([0, 25])
#plt.colorbar()
fig.savefig("foo.pdf", bbox_inches='tight')

#%%
fig, ax = plt.subplots(figsize=(8, 5))
start = min(dates)
stop = max(dates)
ax.plot((start, stop), (0, 0), 'k', alpha=.5)
plt.show()  


#%%

fig = plt.figure(figsize=(10,5))
ax = fig.add_subplot(111)
ax.set_title('Lab legend')


plt.plot(0, 1,color=ccmap[0],label="Levesque")
plt.plot(0, 2,color=ccmap[1],label="Parent")
plt.plot(0, 3,color=ccmap[2],label="Deschenes")
plt.plot(0, 4,color=ccmap[3],label="PDK")
plt.plot(0, 5,color=ccmap[4],label="YDK")
plt.plot(0, 6,color=ccmap[5],label="Toth")

plt.plot(0, 6,color=ccmap[6],label="SAG")
plt.plot(0, 7,color=ccmap[7],label="JK")
plt.plot(0, 8,color=ccmap[8],label="CFS")
plt.plot(0, 9,color=ccmap[9],label="JPJ")
plt.plot(0, 10,color=ccmap[10],label="Labonte")
plt.plot(0, 11,color=ccmap[11],label="Menard")
plt.plot(0, 12,color=ccmap[12],label="cicchetti")



ax.set_xlabel('ADR')
ax.set_ylabel('Rating')
ax.legend(loc='best')
plt.show()



