# -*- coding: utf-8 -*-

"""
Module implementing CSVToExcel.
"""
import time
import pandas as pd
import xlrd
import xlwt
import os

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QFileDialog
from PyQt5 import  QtWidgets
 
from Ui_SubForm1_CSVToExcel import Ui_Form

'''
   使用pandas 转换为Excel
   1.先用Pandas读取CSV文件，然后生成xlsx文件
   2.再用xlrd读取第1步生成的xlsx文件，使用xlwt转化生成xls文件，去掉xlsx文件第1列的序号

   说明：
   1.之所以使用Pandas读取CSV，生成Excel文件，是因为使用csv.reader(sFile)，无法解决MainDes里有逗号，导致读取到的数据乱列
   2.之所以再用xlrd读取一次，用xlwt生成文件，是因为使用Pandas生成的Excel里第1列是序号，而且是xlsx格式，后面步骤需要用到的是xls文件格式
'''

class CSVToExcel(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(CSVToExcel, self).__init__(parent)
        self.setupUi(self)
        
    @pyqtSlot()
    def on_btn_Process_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Process Button")
        #self.txt_File.setPlainText("aaaa")
        reply = QMessageBox.information(self,                         #使用infomation信息框
                                    "提示信息",
                                    "点Yes按钮,开始转换CSV文件为Excel",
                                    QMessageBox.Yes | QMessageBox.No)
        if reply ==QMessageBox.No:
            return
        
    
        sCsvFile=self.txt_File.toPlainText()
     
        csv = pd.read_csv(sCsvFile, encoding='utf-8')  
       
        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="1.ClosedTickets_"+strTime+".xlsx"            

        csv.to_excel(sPath+sReportName,sheet_name='Sheet1')


        #把生成的.xlsx文件格式转化为.xls文件格式,同时去掉因Pandas生成的第1列的序号
        sSourceFile = sPath+sReportName
        sTargetFile = sPath + "1.ClosedTickets_"+strTime+".xls" 
        self.ChangeXlsxToXls(sSourceFile,sTargetFile)

        self.RemoveMidFile(sSourceFile)
        
        Print("Work is over.")
        QMessageBox.information(self,"提示信息","已经转换生成Excel文件。\n"+sTargetFile,QMessageBox.Yes)
    
     
    def ChangeXlsxToXls(self, sSourceFile,sTargetFile):
        wb1 = xlrd.open_workbook(filename=sSourceFile)
        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows
        ncols2 = sheet1.ncols

        wb2 = xlwt.Workbook()
        ws = wb2.add_sheet('Sheet1',cell_overwrite_ok=True)

        for iRow in range(nrows1):
            for iCol in range(1,ncols2):                                     #第1列是使用Pandas生成的Excel,自动生成序号，去掉
                ws.write(iRow,iCol-1,sheet1.cell_value(iRow,iCol))
        wb2.save(sTargetFile)
        
    def RemoveMidFile(self, sTempFile):
        if os.path.exists(sTempFile):
            os.remove(sTempFile)         

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
        print("Click Browse Button")
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "CSV Files (*.csv)")
        print(fileName1,filetype)
        #self.txt_File.setText=fileName1
        self.txt_File.setPlainText(fileName1)
        
    def InitForm(self):
        self.txt_File.setPlainText("")
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    subForm1 = CSVToExcel()
    subForm1.show()
    sys.exit(app.exec_())

