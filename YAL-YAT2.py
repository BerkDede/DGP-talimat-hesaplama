# -*- coding: utf-8 -*-
"""
Created on Mon Nov 15 09:31:52 2021

@author: Berk Dede 
"""

import pandas as pd
import sys
import numpy as np
import os
from pathlib import Path




df = pd.read_excel('Talimat Listeleme.xlsx')
df=df.filter(items=['Yön', 'BaşlangıçSaat','BitişSaat',
                    'Enerji','Fiyat', 'YerineGetirememe  Enerji'])
df['YerineGetirememe  Enerji'] = df['YerineGetirememe  Enerji'].fillna(0)



Enerji=df.iloc[:,3:4].values
YerineGetirememe=df.iloc[:,5:].values
Enerji2=Enerji-YerineGetirememe
Enerji2=Enerji2.astype(float)
Enerji3=pd.DataFrame(Enerji2)
frame2=[df,Enerji3]
df2=pd.concat(frame2,axis=1)


BaşlangıçSaat=df2.iloc[:,1:2]
BitişSaat=df2.iloc[:,2:3]
BaşlangıçSaat=BaşlangıçSaat.squeeze()
BaşlangıçSaat=BaşlangıçSaat.str.split(':',expand=True)
BaşlangıçSaat=BaşlangıçSaat.iloc[:,0:1]

BitişSaat=BitişSaat.squeeze()
BitişSaat=BitişSaat.str.split(':',expand=True)
BitişSaat=BitişSaat.iloc[:,0:1]
df2_1=df2.iloc[:,3:]
df2_2=df2.iloc[:,0:1]
frame3=[BaşlangıçSaat,BitişSaat,df2_1,df2_2]
df3=pd.concat(frame3, axis=1)
df3=df3.set_index(['Yön'])
df3 = df3.set_axis(['BaşlangıçSaat','BitişSaat',
                      'Enerji','Fiyat','YerineGetirememe Enerji','YerineGetirilen'], axis=1, inplace=False)

df_YAT=df3.filter(regex='YAT', axis=0)
df_YAT = df_YAT[df_YAT['YerineGetirilen']!=0]
length=len(df_YAT)

if length >=1:
    
   
    YAT_w_avg=df_YAT.groupby('BaşlangıçSaat').apply( lambda gp: np.average(gp.Fiyat,weights=gp.YerineGetirilen)).rename('w_avg_fiyat')
    YAT_w_avg=YAT_w_avg.to_frame()
  
 
    
    
    df_YAT=df_YAT.drop(['Fiyat'],axis=1)
    df_YAT=df_YAT.set_index(['BaşlangıçSaat'])
    df_YAT=df_YAT.groupby(level=0).sum()
   
    
    frame5=[df_YAT,YAT_w_avg]
    
    df_YAT2=pd.concat(frame5,axis=1)
    
   
    df_YAT2=df_YAT2.drop(['Enerji','YerineGetirememe Enerji'],axis=1)
    df_YAT2.to_excel("YAT.xlsx",
                        sheet_name='Sheet_name_1')
    absolutePath = Path('C:\\Users\\berkd\\.spyder-py3\\YAT.xlsx').resolve()
    os.system(f'start excel.exe "{absolutePath}"')
else:
    df_empty=pd.DataFrame()
    df_empty.to_excel("YAT.xlsx",
                        sheet_name='Sheet_name_1')
    absolutePath = Path('C:\\Users\\berkd\\.spyder-py3\\YAT.xlsx').resolve()
    os.system(f'start excel.exe "{absolutePath}"')
    pass
        
df_YAL=df3.filter(regex='YAL', axis=0)
df_YAL = df_YAL[df_YAL['YerineGetirilen']!=0]
length2=len(df_YAL)
if length2 >=1:
    YAL_w_avg=df_YAL.groupby('BaşlangıçSaat').apply( 
        lambda gp: np.average(gp.Fiyat,weights=gp.YerineGetirilen)).rename('w_avg_fiyat')
    YAL_w_avg=YAL_w_avg.to_frame()
    
 
    
    
    df_YAL=df_YAL.drop(['Fiyat'],axis=1)
    df_YAL=df_YAL.set_index(['BaşlangıçSaat'])
    df_YAL=df_YAL.groupby(level=0).sum()
   
    
    frame5=[df_YAL,YAL_w_avg]
    
    df_YAL2=pd.concat(frame5,axis=1)
    
   
    df_YAL2=df_YAL2.drop(['Enerji','YerineGetirememe Enerji'],axis=1)
    df_YAL2.to_excel("YAL.xlsx",
                        sheet_name='Sheet_name_1')
    absolutePath = Path('C:\\Users\\berkd\\.spyder-py3\\YAL.xlsx').resolve()
    os.system(f'start excel.exe "{absolutePath}"')
else:
    df_empty=pd.DataFrame()
    df_empty.to_excel("YAL.xlsx",
                        sheet_name='Sheet_name_1')
    absolutePath = Path('C:\\Users\\berkd\\.spyder-py3\\YAL.xlsx').resolve()
    os.system(f'start excel.exe "{absolutePath}"')
    sys.exit()
    