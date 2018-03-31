# -*- coding: utf-8 -*-
"""
Created on Fri Mar 30 09:57:39 2018
author: yunjinqi
E-mail:yunjinqi@qq.com
signature:The world of investment is a place without free lunches.
"""
import pandas as pd
import numpy as np
import os
file_list=os.listdir('C:/Users/Administrator/Desktop/data')
for file in file_list[2:]:
   
    df=pd.read_excel('C:/Users/Administrator/Desktop/标的企业.xlsx',dtype={'证券代码':str})
    df.index=range(len(df))
    file_path='C:/Users/Administrator/Desktop/data'+'/'+file
    data=pd.read_excel(file_path)
    columns=data.columns
    
    for i in range(len(df)):
        print(i)
        code=df.loc[i,'证券代码']
        year=df.loc[i,'年份']
        try:
            df1=data[(data['Stkcd']==code)&(data['Accper']==year)&(data['Typrep']=='A')]
        except:
            df1=data[(data['Stkcd']==code)&(data['Accper']==year)]
        for column in columns[3:]:
            #column='C001001000'
            new_column=data.loc[0,column]
            try:
                df.loc[i,new_column]=list(df1[column])[0]
            except:
                df.loc[i,new_column]=np.NAN
    df.to_excel(file[:-5]+'_数据的整理.xlsx')            

