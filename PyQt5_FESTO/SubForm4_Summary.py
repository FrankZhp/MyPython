# -*- coding: utf-8 -*-

"""
Module implementing Summary.

功能模块：从已Mapping的Excel File转换格式输出文件
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QFileDialog

from Ui_SubForm4_Summary import Ui_Form

import xlrd
#import xlwt
from xlutils.copy import copy
import os
import time
#import datetime

class Summary(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(Summary, self).__init__(parent)
        self.setupUi(self)
    
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
        sMappingFile=self.txt_MappingFile.toPlainText()
        #sClosedFile=text2.get()
        sModFile=os.getcwd()+"\\Source\\Mod\\Summary_mod.xls"

        print(sModFile)
        wb1 = xlrd.open_workbook(filename=sMappingFile)
        wb2 = xlrd.open_workbook(filename=sModFile, formatting_info=True)
        #wb2 = xlrd.open_workbook(filename=sModFile)                                 #Summary Mod Excel

        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows

        sheet2 = wb2.sheet_by_index(0)   
        nrows2 = sheet2.nrows
        ncols2 = sheet2.ncols

        wb3 = copy(wb2)
        ws3 = wb3.get_sheet(0)

        #填写生成文件从AA列开始的台头
        #for i in range(ncols2):
        #     ws3.write(0,i+26,sheet2.cell_value(0,i))

        for i in range(1,nrows1):
            #sTicketNo1=sheet1.cell_value(i,0)
            ws3.write(i,0,sheet1.cell_value(i,0))
            #TicketNo
            #B列Period,格式：yyyyQX,来自closedDate
            sClosedDate=sheet1.cell_value(i,32)
            sPeriod=self.getPeriod(sClosedDate)
            sPriceTable=self.getPriceTable(sClosedDate)
            ws3.write(i,1,sPeriod)
            ws3.write(i,9,sPriceTable)

            ws3.write(i,10,sheet1.cell_value(i,26))                       #Priority
            ws3.write(i,11,sheet1.cell_value(i,27))                       #Ticket No
            ws3.write(i,12,sheet1.cell_value(i,28))                       #  
            ws3.write(i,13,sheet1.cell_value(i,29))                        
            ws3.write(i,14,sheet1.cell_value(i,30))
            ws3.write(i,15,sheet1.cell_value(i,31))
            ws3.write(i,16,sheet1.cell_value(i,32))                       #closedDate
            ws3.write(i,17,sheet1.cell_value(i,33))                       #Ticket Type             
            ws3.write(i,18,sheet1.cell_value(i,34))                       #Shorttext       
            #用户要求增加1列Main Description,在第T列，第20列
            ws3.write(i,19,sheet1.cell_value(i,35))                       #Main Description      
            ws3.write(i,20,sheet1.cell_value(i,36))                       #Parent category			 		  
            ws3.write(i,21,sheet1.cell_value(i,37))                       #Category  
            ws3.write(i,22,sheet1.cell_value(i,38))                       #User ID 
            ws3.write(i,23,sheet1.cell_value(i,39))                       #Name(reported for) 
            ws3.write(i,24,sheet1.cell_value(i,40))                       #First Name
            ws3.write(i,25,sheet1.cell_value(i,41))                       #Last Name
            ws3.write(i,26,sheet1.cell_value(i,42))                       #Country

            for iCol1 in range(15):
                iCol2=iCol1+27
                ws3.write(i,iCol2,sheet1.cell_value(i,iCol1))                        #TicketNo

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="4.Summary_"+strTime+".xls"
        wb3.save(sPath+sReportName)  
     
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
        self.txt_MappingFile.setPlainText(fileName1)
    
    def InitForm(self):
        self.txt_MappingFile.setPlainText("")   
    
    #自定义方法
    def getPriceTable(self, sClosedDate):
        sPriceTable=""
        sYear=""
        sMonth=""
        #sDay=""
        if sClosedDate!="":
            list1=sClosedDate.split(" ")
            sDate=list1[0]
            list2=sDate.split("/")
            i=0
            for row in list2:
                if i==0:
                    sMonth=row
                #if i==1:
                #    sDay=row
                if i==2:
                    sYear=row
                i=i+1
            sPriceTable=sYear+"-"+sMonth
        return sPriceTable
    def getPeriod(self, sClosedDate):
        sPeriod=""
        sYear=""
        sMonth=""
        #sDay=""
        try:
            if sClosedDate!="":
                list1=sClosedDate.split(" ")
                sDate=list1[0]
                list2=sDate.split("/")
                i=0
                for row in list2:
                    if i==0:
                        sMonth=row
                    if i==1:
                        #sDay=row
                        pass
                    if i==2:
                        sYear=row
                    i=i+1
            dict={1:'Q1',2:'Q1',3:'Q1',4:'Q2',5:'Q2',6:'Q2',7:'Q3',8:'Q3',9:'Q3',10:'Q4',11:'Q4',12:'Q4'}
            sPeriod=sYear+dict.get(int(sMonth))
        except:
            print(sClosedDa)
        return sPeriod 

    def is_valid_date(self, str):
        '''判断是否是一个有效的日期字符串'''
        try:
            time.strptime(str, "%Y-%m-%d")
            return True
        except:
            return False
            
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    subForm4 = Summary()
    subForm4.show()
    sys.exit(app.exec_())
