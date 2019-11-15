# -*- coding: utf-8 -*-
""
#Created on Thu Mar 14 13:05:13 2019

#@author: Simon
#""Time_Google_Calendar"
#from tkinter import filedialog
#import_file_path = filedialog.askopenfilename
        
import os
import pandas as pd # Import pandas
import numpy as np
# % %  Get Microscope Calendar
os.chdir("Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Test_Pyhton")
# Read excel file from microscope
file = 'Levesque-ZeissLSM700_3.xlsx'
# Load spreadsheet
data= pd.read_excel(file)
df = pd.DataFrame(data, columns= ['Title'])
# Change dataframe to list % Can i do better ?
userlist=[]
for i in range(len(df)):
    userlist.append(df.at[i, 'Title'])
    
#% % Get lab members
os.chdir("Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\Lab_Members")
MemberName = 'Members_LabList.xlsx'
members = pd.read_excel(MemberName)
LevesqueLab=[]
for i in range(len(members)):
    LevesqueLab.append(members.at[i, 'Levesque'])
# % %    
# Find Levesque's lab members based on Lab_members List.
L=[]
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in LevesqueLab):
            matching2 = [s for s in LevesqueLab if userlist[i] in s]
            L.append('Levesque')
        else:
            L.append("")
    else:
        L.append("Nan")
#data['Lab'] = pd.Series(L, index=data.index)        
           
#%%                
#ParentLab=[]
# Find Parent's lab members based on Lab_members List.
ParentLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'Parent'], str):
        ParentLab.append(members.at[i, 'Parent'])
    
#%% Data For Parent lab   
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in ParentLab):
            matching2 = [s for s in ParentLab if userlist[i] in s]
            L[i]=('Parent')
#data['Lab'] = pd.Series(L, index=data.index)   

#DeschenesLab
DeschenesLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'Deschenes'], str):
        DeschenesLab.append(members.at[i, 'Deschenes'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in DeschenesLab):
            matching2 = [s for s in DeschenesLab if userlist[i] in s]
            L[i]=('Deschenes')
#PDKLab
PDKLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'PDK'], str):
        PDKLab.append(members.at[i, 'PDK'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in PDKLab):
            matching2 = [s for s in PDKLab if userlist[i] in s]
            L[i]=('PDK')
#YDKLab
YDKLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'YDK'], str):
        YDKLab.append(members.at[i, 'YDK'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in YDKLab):
            matching2 = [s for s in YDKLab if userlist[i] in s]
            L[i]=('YDK')
#TothLab
TothLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'Toth'], str):
        TothLab.append(members.at[i, 'Toth'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in TothLab):
            matching2 = [s for s in TothLab if userlist[i] in s]
            L[i]=('Toth')
#SAGLab
SAGLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'SAG'], str):
        SAGLab.append(members.at[i, 'SAG'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in SAGLab):
            matching2 = [s for s in SAGLab if userlist[i] in s]
            L[i]=('SAG')
#JKLab
JKLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'JK'], str):
        JKLab.append(members.at[i, 'JK'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in JKLab):
            matching2 = [s for s in JKLab if userlist[i] in s]
            L[i]=('JK')
#CFSLab=[]
CFSLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'CFS'], str):
        CFSLab.append(members.at[i, 'CFS'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in CFSLab):
            matching2 = [s for s in CFSLab if userlist[i] in s]
            L[i]=('CFS')
#JPJLab
JPJLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'JPJ'], str):
        JPJLab.append(members.at[i, 'JPJ'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in JPJLab):
            matching2 = [s for s in JPJLab if userlist[i] in s]
            L[i]=('JPJ')
#LabonteLab
LabonteLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'Labonte'], str):
        LabonteLab.append(members.at[i, 'Labonte'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in LabonteLab):
            matching2 = [s for s in LabonteLab if userlist[i] in s]
            L[i]=('Labonte')
#MenardLab=[]
MenardLab=[]
for i in range(len(members)):
    if isinstance(members.at[i, 'Menard'], str):
        MenardLab.append(members.at[i, 'Menard'])
for i in range(len(userlist)) :
    matching2 =[]
    print( userlist[i])  
    if isinstance(userlist[i], str): #to avoid the NaN        
        if any(userlist[i] in s for s in MenardLab):
            matching2 = [s for s in MenardLab if userlist[i] in s]
            L[i]=('Menard')


data['Lab'] = pd.Series(L, index=data.index) 

with pd.ExcelWriter('Test.xlsx') as writer:
    data.to_excel(writer, sheet_name='Sheet1')
  
    
    
    
    
    
    
    
    
    

    