# -*- coding: utf-8 -*-

"""
Module implementing MakeReport.
#三个文件：1、已经合并汇总好的Excel文件 2、当期价格表 3、待生成报表模板样例文件
#模块功能：计算价格

#模块功能：按照地区别、季度的价格，计算Ticket的价格
#第1个文件，已经汇总Mapping好的Excel文件
#第2个文件，价格表
#第3个文件，待生成报表的模板文件

"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QFileDialog
from PyQt5 import  QtWidgets

from Ui_SubForm9_MakeReport import Ui_Form

import os
import time
import xlrd
import xlwt
from xlutils.copy import copy
import csv
import numpy as np
from collections import Counter

countryList=['AU','NZ','CN','HK','TW','SG','TH','ID','JP','KR','MY','VN','PH']
periodList=[]

class MakeReport(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(MakeReport, self).__init__(parent)
        self.setupUi(self)
        self.cbo_Country.addItem('')
        self.cbo_Country.addItems(countryList)
        
        periodList = self.InitPeriod()
        self.cbo_Period.addItem('')
        self.cbo_Period.addItems(periodList)
    
    @pyqtSlot()
    def on_btn_Process_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        reply = QMessageBox.information(self,                         #使用infomation信息框
                                    "提示信息",
                                    "点Yes按钮,开始处理......",
                                    QMessageBox.Yes | QMessageBox.No)
        if reply ==QMessageBox.No:
            return
        
        #dictArea={0:'AU',1:'NZ',2:'CN',3:'HK',4:'TW',5:'SG',6:'TH',7:'ID',8:'MY',9:'VN',10:'PH'}
        #dictPeriod={0:'2018Q1',1:'2018Q2',2:'2018Q3',3:'2018Q4',4:'2019Q1',5:'2019Q2',6:'2019Q3',7:'2019Q4'}
        dictCurrency={'AU':'AUD','CN':'CNY','HK':'HKD','ID':'IDR','IN':'INR','JP':'JPY','KR':'KRW','MY':'MYR','NZ':'NZD','PH':'PHP','TH':'THB','TW':'TWD','VN':'VND','SG':'SGD'}

        sSummaryFile=self.txt_SummaryFile.toPlainText()
        sPriceFile=self.txt_PriceFile.toPlainText()
        sModFile=os.getcwd()+"\\Mod\\xlsMod\\FESTO Summary via USU & CS XXX_Mod.xls"

        print(sModFile)
        wb1 = xlrd.open_workbook(filename=sSummaryFile)
        wb3 = xlrd.open_workbook(filename=sPriceFile)
        #csv_file = csv.reader(open(sPriceFile))
        #print(csv_file)
        
        wb2 = xlrd.open_workbook(filename=sModFile,formatting_info=True)                                 

        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows                                               #Summary File max row's number
        ncols1 = sheet1.ncols                                               #Summary File max col's number

        sheet3 = wb3.sheet_by_index(0)   
        nrows3 = sheet3.nrows                                               #Summary File max row's number
        ncols3 = sheet3.ncols                                               #Summary File max col's number
        rows = sheet3.row_values(0)                                         #获取行内容,第1行，列名
        cols = sheet3.col_values(1)                                         #获取列内容,第2列，SI Code
        
    
        #print(rows)
        #print(cols)
       
     
        wb = copy(wb2)
        ws = wb.get_sheet(1)

        #sArea1=dictArea.get(cbo_Country.currentText())
        sArea1=self.cbo_Country.currentText()
        sINV_Currency = dictCurrency.get(sArea1)                            #由货币的2个字符的简称转化成三个字符的简称
        #判断价格所在列
        if sArea1 in rows:
            iPriceCol=rows.index(sArea1)
            #ws.write(0,48,sArea1)

        #sPeriod1=dictPeriod.get(cbo_Period.currentText())
        sPeriod1=self.cbo_Period.currentText()
      
        
        iRow2=0
        lst_SICode=[]
        dict_SICode={}
        
        for iRow in range(1,nrows1):
            sTicketNo=sheet1.cell_value(iRow,0)
            sArea2=sheet1.cell_value(iRow,34).strip()                       #AI列Country
            #print('行号:'+str(iRow)+'Area:'+sArea2)
            #print("TicketNo:"+sTicketNo)
            sPeriod2=sheet1.cell_value(iRow,1).strip()
            
            if sTicketNo!="" and sArea1==sArea2 and sPeriod1==sPeriod2:
                #print('Area1'+sArea1+" Period1:" +sPeriod1)
                iRow2=iRow2+1
                for iCol in range(ncols1):
                    ws.write(iRow2,iCol,sheet1.cell_value(iRow,iCol))

                #填写AQ列，Month Reported,在4.Summay文件里没有此列，以O列Date reported，转化成yyyymm样式
                sCellValue = sheet1.cell_value(iRow,14)
                sYMTemp=self.getYearMonth(sheet1.cell(iRow, 14).ctype,sCellValue)
                #print('YM:'+sYMTemp)
                ws.write(iRow2,42,sYMTemp.replace("-",""))
                    
                #计算价格
                sServiceType=sheet1.cell_value(iRow,28).strip()
                sServiceSubType=sheet1.cell_value(iRow,29).strip()
                sServiceLevel=sheet1.cell_value(iRow,31).strip()
                sTicketType=sheet1.cell_value(iRow,30).strip()

                sSICode=sServiceType+"_"+sServiceSubType+"_"+sServiceLevel+"_"+sTicketType
                print('SICode：'+sSICode)
                lst_SICode.append(sSICode)
                

                
                if sSICode in cols:
                    i=cols.index(sSICode)
                    if i>=0:
                        tempRows=sheet3.row_values(i)
                        #print("行："+str(i))
                        #print(tempRows)
                        sServiceItemNo=tempRows[0]
                        priceItem=round(tempRows[iPriceCol],2)
                        ws.write(iRow2,46,sServiceItemNo)
                        ws.write(iRow2,47,sSICode)
                        ws.write(iRow2,48,priceItem)
                        ws.write(iRow2,4,priceItem)                     #第E列Invoice Account
                        ws.write(iRow2,5,sINV_Currency)                 #第F列INV_Currenyc，填入三个字符表示的币种,如：CNY、HKD等
                else:
                    print('Not in Cols')
            else:
               # print('Row'+str(iRow)+'TicketNo:'+sTicketNo +'Area:'+sArea2)
                pass   

        #记录不同SI Code的Ticket的条数
        print('记录数：'+str(len(lst_SICode)))
        #print(all_np(lst_SICode))
        dict = {}
        for key in lst_SICode:
            dict[key] = dict.get(key, 0) + 1
        #print(dict)
        iRow2=iRow2+3
        ws.write(iRow2,0,"Service Item No")
        ws.write(iRow2,1,"SI Code")
        ws.write(iRow2,2,"Unit Price")
        ws.write(iRow2,3,"Ticket")
        ws.write(iRow2,4,"Total")
        totalPrice=0
        for k,v in dict.items():
            #print(k,v)
            if k in cols:
                sSICode=k
                i=cols.index(sSICode)
                if i>=0:
                    iRow2=iRow2+1
                    tempRows=sheet3.row_values(i)
                    #print("行："+str(i))
                    #print(tempRows)
                    sServiceItemNo=tempRows[0]
                    priceItem=round(tempRows[iPriceCol],2)
                    ws.write(iRow2,0,sServiceItemNo)
                    ws.write(iRow2,1,k)
                    ws.write(iRow2,2,priceItem)
                    ws.write(iRow2,3,v)
                    subtotal=int(v)*float(priceItem)
                    totalPrice=totalPrice+subtotal
                    ws.write(iRow2,4,subtotal)
                else:
                    print('Not in Cols')
        iRow2=iRow2+1
        ws.write(iRow2,0,"Total")
        ws.write(iRow2,4,totalPrice)
              
        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="9.ComputePriceStyle1_"+sArea1+'_'+strTime+".xls"
        wb.save(sPath+sReportName)  
        
        print("Work is over.")
        QMessageBox.information(self,"提示信息","处理完毕，存放位置：当前目录\\Result,\n文件名称:"+sReportName,QMessageBox.Yes)
         
    
    @pyqtSlot()
    def on_btn_Exit_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        reply = QMessageBox.question(self, "标题", "确定要退出吗?", 
                   QMessageBox.Yes|QMessageBox.No)
         
        if reply == QMessageBox.Yes:
            self.close()
    
    @pyqtSlot()
    def on_btn_Browse_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "Excel Files (*.xls)")
        print(fileName1,filetype)
        #self.txt_File.setText=fileName1
        self.txt_SummaryFile.setPlainText(fileName1)
    
    @pyqtSlot()
    def on_btn_Browse_2_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "Excel Files (*.xls)")
        print(fileName1,filetype)
        #self.txt_File.setText=fileName1
        self.txt_PriceFile.setPlainText(fileName1)
        
    def InitForm(self):
        self.txt_SummaryFile.setPlainText("")
        self.txt_PriceFile.setPlainText("")
        
    def getModColumn(self):
        #list，把列名放在其中
        modlist=['Ticket No','Period','Invoice	FESTO PO','Invoice Amount','INV_Currency','Remarks']		  
            
    def all_np(arr):
        #"""获取每个元素的出现次数，使用Numpy"""
        arr = np.array(arr)
        key = np.unique(arr)
        result = {}
        for k in key:
            mask = (arr == k)
            arr_new = arr[mask]
            v = arr_new.size
            result[k] = v
        return result
        
    """
       名称:InitPeriod
        功能：初始化PeriodList
        配置文件：./Doc/config.xls
    """
    def InitPeriod(self):
        sPath = os.getcwd()+"\\Config\\"
        sConfigFile = "config.xls"
        wb1 = xlrd.open_workbook(filename=sPath+sConfigFile)    
        sheet1 = wb1.sheet_by_index(0)
        col0=sheet1.col_values(0)
        del col0[0]                               #去掉第1行的列名
        return col0
    '''
       函数功能：格式化日期
           参数：
               cType:0 empty;1 string;2 number;3 date
    '''
    def getYearMonth(self, ctype,sCellValue):
        sYM=""
        if ctype==3:                                                   
            date = xlrd.xldate_as_tuple(sCellValue, 0)                   #转化为元组形式，如2019-02-01转化为：(2019, 2, 1, 0, 0, 0)
     
            sYear=str(date[0])
            sMonth=str(date[1])
            if len(sMonth)==1:
                sMonth='0'+sMonth
            sYM=sYear+'-'+sMonth
        elif ctype==1:                                                   #ctype =1,字符串
            sCellValue = sCellValue.strip()
            if sCellValue !="":
                list1=sCellValue.split(' ')
                s1=list1[0]
                list2=s1.split('/')
                sYear=list2[2]
                sMonth=list2[0]
                sYM=sYear + '-'+sMonth.zfill(2)
            
            
        return sYM
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    subForm9 = MakeReport()
    subForm9.show()
    sys.exit(app.exec_())
