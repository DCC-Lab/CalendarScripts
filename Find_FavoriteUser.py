# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd
import os
import numpy as np
#% Load excel and display data
directory = 'Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Sted'
filename = '2018.xlsx'

os.chdir(directory) 
df = pd.read_excel(filename)

time = df['Hours']
users = df['Title']

Unique_Users = np.unique(users)

Unique_UsersList=[]
for i in range(len(Unique_Users)):
    Unique_UsersList.append(Unique_Users[i])

#% 