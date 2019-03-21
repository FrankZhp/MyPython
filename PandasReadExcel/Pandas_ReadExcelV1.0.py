#! -*- coding utf-8 -*-
#! @Time  :2019/3/20 22:00
#! Author :Frank Zhang
#! @File  :Pandas_ReadExcelV1.0.py
#! Python Version 3.7

"""
模块功能：读取当前文件夹下的Source里的Excel文件，显示其相关信息

说明：默认把Excel的第一行当做列名，数据的第1行是从Excel的第2行开始
      这里获取的最大行是Excel的最大行减去作为列名的第1行

"""

import pandas as pd
 
sExcelFile="./Source/Book1.xlsx"
df = pd.read_excel(sExcelFile,sheet_name='Sheet1')

#获取最大行，最大列
nrows=df.shape[0]
ncols=df.columns.size


print("=========================================================================")
print('Max Rows:'+str(nrows))
print('Max Columns'+str(ncols))

#显示列名，以列表形式显示
print(df.columns)

#显示列名，并显示列名的序号
for iCol in range(ncols):
    print(str(iCol)+':'+df.columns[iCol])

#列出特定行列，单元格的值
print(df.iloc[0,0])
print(df.iloc[0,1])

print("=========================================================================")



#查看某列内容
#sColumnName='fd1'
print(df[sColumnName])

#查看第3列的内容，列的序号从0开始
sColumnName=df.columns[2]
print(df[sColumnName])

 

#查看某行的内容
iRow=1
for iCol in range(ncols):
    print(df.iloc[iRow,iCol])


#遍历逐行逐列
for iRow in range(nrows):
    for iCol in range(ncols):
        print(df.iloc[iRow,iCol])

print('=====================================End==================================')

 

        
 


