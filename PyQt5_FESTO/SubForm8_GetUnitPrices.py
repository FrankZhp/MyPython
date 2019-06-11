# -*- coding: utf-8 -*-

"""
Module implementing GetUnitPricesStyle2.
功能模块：进行换算，由汇率资费表得到Unit Price


   #第一个文件是ECB Rate File
    #第二个文件是APJ_FESTO_Services_Pricing_Mod.xls
    '''
       业务逻辑：
       A、Global Accounts Quarterly页
       1、step1 首先根据所选的年月，在第3行找到待汇总年月所在的列
       2、第1列Currency Name、第2列 ISO Code，这两列有一个为空则不处理，说明这一行不是数据
       3、找到1欧元=XXX美元，作为后面计算的基准
       4、汇率换算方式分两种：Currency Name里带base字样和不带base字样的
       5、只处理ASIA PACIFIC的即可，14个Country
       
       6、从第7步生成的文件读取数据，根据第2行的汇率和第6-68行的数据，从新汇总核算

    '''
    style2格式的报表，是选择生成的style1的结果文件，.xlsx格式
    '''
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QFileDialog
from PyQt5 import  QtWidgets

from Ui_SubForm8_GetUnitPrices import Ui_Form

import xlrd
import xlwt
from xlutils.copy import copy
import os
import time
import datetime
import types
import openpyxl
from xlrd import xldate_as_tuple

#monthList=[] #['2018-10','2018-11','2018-12','2019-01','2019-02','2019-03','2019-04','2019-05','2019-06','2019-07','2019-08','2019-09','2019-10','2019-11','2019-12']
countryList=['AUD','CNY','HKD','IDR','INR','JPY','KRW','MYR','NZD','PHP','THB','TWD','VND','SGD']

class GetUnitPricesStyle2(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(GetUnitPricesStyle2, self).__init__(parent)
        self.setupUi(self)
        #monthList = self.InitMonth()
        #self.cbo_Month.addItem('')
        #self.cbo_Month.addItems(monthList)
    
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
        
        sECBRateFile=self.txt_ECBRateFile.toPlainText()
  
        sModFile=os.getcwd()+"\\Mod\\xlsMod\\APJ_FESTO_Services_Pricing_Mod.xls"

        wb1 = xlrd.open_workbook(filename=sECBRateFile)
        wb2 = xlrd.open_workbook(filename=sModFile, formatting_info=True)
        
        #ECB Rate File
        sheet1 = wb1.sheet_by_index(1)
        #nrows1 = sheet1.nrows
        ncols1 = sheet1.ncols
        
        sMonth=sheet1.cell_value(1,1)
        reply = QMessageBox.question(self,                         #使用infomation信息框
                                    "提示信息",
                                    "是对"+sMonth+'的数据进行处理吗?',                         
                                    QMessageBox.Yes | QMessageBox.No)
        if reply ==QMessageBox.No:
            return

        #APJ_FESTO_Services_Pricing_Mod.xls 
        sheet2 = wb2.sheet_by_index(0)
        #nrows2 = sheet2.nrows
        #ncols2 = sheet2.ncols
        serviceItemNoList=sheet2.col_values(0)                                                               #取模板文件里第1列的Service Item No
   
        #使用xlwt写新的excel,文件名后缀.xls
        wb = copy(wb2)
        ws = wb.get_sheet(0)

        for iRow1 in range(6,68):
            sConditionCell=sheet1.cell_value(iRow1,3)
            if  sConditionCell!="" and sConditionCell !="AU":
                sCellValue=sheet1.cell_value(iRow1,0).rstrip(".")                                           #取第1列的Service Item No,如果所取字符串最右侧带“.”，则把“.”去掉
                sServiceItemNo1 = sCellValue.strip()
               
                if sServiceItemNo1 in serviceItemNoList:
                    iRow2 = serviceItemNoList.index(sServiceItemNo1)
                     
                    for iCol1 in range(3,ncols1):
                        iCol2 = iCol1 -1
                        sUnitPrice = sheet1.cell_value(iRow1,iCol1) * sheet1.cell_value(1,iCol1)             #按照汇率以及基准单价计算得到需要的结果
                        ws.write(iRow2,iCol2,sUnitPrice)
                         
                else:
                    print('Service Item No: '+sServiceItemNo1 +' Not found')
            else:
                print("Check:" + sConditionCell)

        sYM=sheet1.cell_value(1,1).replace("-","")

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time()))
        sPath=os.getcwd()+"\\Result\\"
        sReportName="8.ECBRate_" + sYM + "_"+ strTime + "_Style2.xls"
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
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "Excel Files (*.xlsx)")
        print(fileName1,filetype)
        #self.txt_File.setText=fileName1
        self.txt_ECBRateFile.setPlainText(fileName1)
    
    def InitForm(self):
        self.txt_ECBRateFile.setPlainText("")
    '''
        ===================================函数=========================================
        名称：getYearMonth
        功能：xlrd从Excel读出的日期单元格的数据，不是以日期形式显示的，转化为yyyy-mm格式
        参数：ctype 单元格类型，3 为日期类型
              sCellValue 从单元格传过来的值
    '''
    def getYearMonth(ctype,sCellValue):
        sYM=""
        if ctype==3: 
            date = xlrd.xldate_as_tuple(sCellValue, 0)                   #转化为元组形式，如2019-02-01转化为：(2019, 2, 1, 0, 0, 0)
     
            sYear=str(date[0])
            sMonth=str(date[1])
            if len(sMonth)==1:
                sMonth='0'+sMonth
            sYM=sYear+'-'+sMonth
        return sYM
    '''
       =====================================函数========================================
       名称：checkForm
       功能：检查界面输入条件，在点Process按钮以前，是否已给出相应条件
    '''
    def checkForm(self):
        #sYM = self.cbo_Month.currentText()
        sECBRateFile = self.txt_ECBRateFile.toPlainText()

        if sECBRateFile == "":
            return False
        #if   sYM == "":
        #    return False
        
        return True
    
    '''
       名称:InitMonth
        功能：初始化MonthList
        配置文件：./Doc/config.xls
    '''
    def InitMonth(self):
        sPath = os.getcwd()+"\\Config\\"
        sConfigFile = "config.xls"
        wb1 = xlrd.open_workbook(filename=sPath+sConfigFile)    
        sheet1 = wb1.sheet_by_index(0)
        col1=sheet1.col_values(1)
        del col1[0]                               #去掉第1行的列名
        return col1
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    subForm8 = GetUnitPricesStyle2()
    subForm8.show()
    sys.exit(app.exec_())
