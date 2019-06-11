# -*- coding: utf-8 -*-

"""
Module implementing AppendInvoice.
功能模块：把合并好的发票信息追加到整理好的Summary文件里
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QFileDialog
from PyQt5 import  QtWidgets

from Ui_SubForm6_AppendInvoice import Ui_Form

import xlrd
import xlwt
from xlutils.copy import copy
import os
import time
import datetime

class AppendInvoice(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(AppendInvoice, self).__init__(parent)
        self.setupUi(self)
    
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
    def on_btn_Browse1_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "Excel Files (*.xls)")
        print(fileName1,filetype)
        #self.txt_File.setText=fileName1
        self.txt_File1.setPlainText(fileName1)
    
    @pyqtSlot()
    def on_btn_Browse2_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "Excel Files (*.xls)")
        print(fileName1,filetype)
        #self.txt_File.setText=fileName1
        self.txt_File2.setPlainText(fileName1)
    
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
        sInvoiceFile=self.txt_InvoiceFile.get()
        sSummaryFile=self.txt_SummaryFile.get()
      
        wb1 = xlrd.open_workbook(filename=sInvoiceFile)
        wb2 = xlrd.open_workbook(filename=sSummaryFile, formatting_info=True)
   
        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows

        sheet2 = wb2.sheet_by_index(0)   
        nrows2 = sheet2.nrows
        ncols2 = sheet2.ncols
        cols2 = sheet2.col_values(0)    #获取第1列TicketNo内容,放入list

        wb3 = copy(wb2)
        ws3 = wb3.get_sheet(0)

        for i in range(1,nrows1):
            
            sTicketNo=sheet1.cell_value(i,0)
            
            if sTicketNo in cols2:                                          #判断TicketNo 是否在List中  
                iRow=cols2.index(sTicketNo)
                ws3.write(iRow,1,sheet1.cell_value(i,1))                    #B列：Period                   
                ws3.write(iRow,2,sheet1.cell_value(i,2))                    #C列：Invoice 
                ws3.write(iRow,3,sheet1.cell_value(i,3))                    #D列：FESTO PO
                ws3.write(iRow,4,sheet1.cell_value(i,4))                    #E列：Invoice Amount
                ws3.write(iRow,5,sheet1.cell_value(i,5))                    #F列：INV_Currency
                ws3.write(iRow,6,sheet1.cell_value(i,6))                    #G列：Remarks
                
        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="6.Summary_"+strTime+".xls"
        wb3.save(sPath+sReportName)         
        
        print("Work is over.")
        QMessageBox.information(self,"提示信息","处理完毕，存放位置：当前目录\\Result,\n文件名称:"+sReportName,QMessageBox.Yes)
    
    def InitForm(self):
        self.txt_InvoiceFile.setPlainText("")
        self.txt_SummaryFile.setPlainText("")
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    subForm6 = AppendInvoice()
    subForm6.show()
    sys.exit(app.exec_())
    

