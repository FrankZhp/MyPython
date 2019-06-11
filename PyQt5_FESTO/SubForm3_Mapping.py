# -*- coding: utf-8 -*-

"""
Module implementing Mapping.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QFileDialog

from Ui_SubForm3_Mapping import Ui_Form

import xlrd
import xlwt
from xlutils.copy import copy
import os
import time


class Mapping(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(Mapping, self).__init__(parent)
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
                                    "点Yes按钮,开始转换CSV文件为Excel",
                                    QMessageBox.Yes | QMessageBox.No)
        if reply ==QMessageBox.No:
            return
            
        sMergedFile=self.txt_MergedExcel.toPlainText()
        sClosedFile=self.txt_ClosedTicketFile.toPlainText()
        wb1 = xlrd.open_workbook(filename=sMergedFile, formatting_info=True)
        wb2 = xlrd.open_workbook(filename=sClosedFile)

        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows

        sheet2 = wb2.sheet_by_index(0)   
        nrows2 = sheet2.nrows                                #Max Rows
        ncols2 = sheet2.ncols                                #Max Cols  
        cols2 = sheet2.col_values(1)                         #TicketNo

        wb3 = copy(wb1)
        ws3 = wb3.get_sheet(0)

        #填写生成文件从AA列开始的台头
        for i in range(ncols2):
            ws3.write(0,i+26,sheet2.cell_value(0,i))

        for i in range(1,nrows1):
            sTicketNo1=sheet1.cell_value(i,0)
            if sTicketNo1.strip() in cols2:
                j=cols2.index(sTicketNo1.strip())
                for k in range(ncols2):
                    ws3.write(i,k+26,sheet2.cell_value(j,k))
            else:
                ws3.write(i,1+26,"No Found")
                

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="3.MappingReport_"+strTime+".xls"
        wb3.save(sPath+sReportName)        
       

        print("Work is over.")
        QMessageBox.information(self,"提示信息","处理完毕，存放位置：当前目录\\Result,文件名称：\n"+sReportName,QMessageBox.Yes)
        
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
        self.txt_MergedExcel.setPlainText(fileName1)
    
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
        self.txt_ClosedTicketFile.setPlainText(fileName1)
    
    def InitForm(self):
        self.txt_MergedExcel.setPlainText("")
        self.txt_ClosedTicketFile.setPlainText("")
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    subForm3 = Mapping()
    subForm3.show()
    sys.exit(app.exec_())
