#! -*- coding utf-8 -*-
#! @Time  :2019/3/20 9:44
#! Author :Frank Zhang
#! @File  :Pandas_ReadExcelV1.1.py
#! Python Version 3.7

"""
模块功能：读取当前文件夹下的Source里的Excel文件，显示其相关信息

说明：数据从第1行开始，不设列名，不把第1行作为列名
      这里获取的最大行就是是Excel的最大行

"""

import pandas as pd
 
sExcelFile="./Source/Book1.xlsx"

df = pd.read_excel(sExcelFile,sheet_name='Sheet1',header=None)

#获取最大行和最大列数
nrows=df.shape[0]
ncols=df.columns.size
 

print("=====================================================")
print('Max Rows: '+str(nrows))
print('Max Columns: '+str(ncols))

#显示某特定单元格的值
print(df.iloc[0,0])
print(df.iloc[0,1])
print("=====================================================")

#查看某行的内容
print("====================显示某一行=======================")
#iRow=1
print("请输入行号(1-"+str(nrows)+"):")
iRow=int(input())-1
for iCol in range(ncols):
    print(df.iloc[iRow,iCol])

print("====================显示某一列=======================")
#iCol=1
print("请输入列号(1-"+str(ncols)+"):")
iCol=int(input())-1
if iCol>=0 and iCol<=ncols:
    for iRow in range(nrows):
        print(df.iloc[iRow,iCol])
else:
    print('输入了错误的列号')
 
 
#遍历逐行逐列
print("\n逐行逐列显示：")
for iRow in range(nrows):
    for iCol in range(ncols):
        print(df.iloc[iRow,iCol])
    
print('=========================End=========================')



        
 


