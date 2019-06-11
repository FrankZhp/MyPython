# -*- coding: utf-8 -*-

"""
Module implementing MergeFESTO.

模块功能：读取某个文件夹下的所有FESTO Excel文件，读取每个文件里的所有Sheet，按照模板文件的样例，读取相应列，把数据合并到一个Excel文件中
业务逻辑说明
1.把模板文件的第一列读入list中,读取的FESTO文件，个别的可能与模板文件的列以及顺序不一样，那么按照模板文件的列及顺序，合并到一个新的文件中。
2.模板文件里没有的列则不读取

"""


from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, QFileDialog

from Ui_SubForm2_MergeFESTO import Ui_Form

import xlrd
import xlwt
import os
#import openpyxl
#import datetime
import time

Dict_ModExcelCol={}
modlist=[]

class MergeFESTO(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(MergeFESTO, self).__init__(parent)
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
                                    "点Yes按钮,开始合并FESTO Excel文件......",
                                    QMessageBox.Yes | QMessageBox.No)
        if reply ==QMessageBox.No:
            return
            
        sPath=self.txt_Directory.toPlainText()
        #print(sPath)
        
        filenames = os.listdir(sPath)
        #print(filenames)
        
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet(u'Sheet1')
        #为去重，定义一个list，遍历时，判断遍历到的TicketNo是否在此list中
        #如果不在，则增加到这个list中，并填写到结果文件中
        last_List=[]                            

        #iRow记录写的Excel的行号，第一行是从0开始，所以，初始化为-1
        iRow2=-1
        iCol2=-1

        #读取Excel模板文件，把模板文件里的的第一行的列名作为结果文件的列名
        self.read_ModExcel()

        #填写结果文件的第一行
        iRow2=0
        for i in range(len(Dict_ModExcelCol)):
            worksheet.write(iRow2,i,modlist[i])
     
        for iFile, filename in enumerate(filenames):
            #print('第:'+str(iFile+1)+'个文件： '+filename)
            sFileName=filename
                
            #wb = openpyxl.load_workbook(sPath+'\\'+sFileName)
            wb = xlrd.open_workbook(filename=sPath+'\\'+sFileName)
            # 获取workbook中所有的表格
            #sheets = wb.sheetnames
            sheets=wb.sheet_names()
            
            # 循环遍历所有sheet
            for i in range(len(sheets)):
                #sheet = wb[sheets[i]]
                sheet = wb.sheet_by_index(i)
                #print('第' + str(i + 1) + '个sheet Name: ' + sheet.title)
                #print('第' + str(i + 1) + '个sheet Name: ' + sheet.name)
                rows = sheet.row_values(0)   # 获取第一行内容
                cols = sheet.col_values(0)  #获取第1列的内容
                max_row=len(cols)
                max_column=len(rows)

                #判断这个Sheet的第一行第一个单元格是否是Ticket No.，如果不是，则不读取这一页
                sColumn1=sheet.cell(0,0).value
                #print("第一个单元格："+sColumn1)
                if sColumn1=="Ticket No":
                    pass
                else:
                    #print("退出这个sheet"+sheet.name+"File Name:"+sFileName)
                    continue
                   
                #print("Write Excel"+sheet.name)
                #第一列关键字，如果重复则去掉
                old_List=sheet.col_values(0)
                
                #第一行是列名，从第二行开始
                for iRow1 in range(1,max_row):
                   
                    for iCol1 in range(max_column):
                        if iCol1==0:
                            if old_List[iRow1] in last_List:                     #如果已有，则退出for循环，不增加重复数据
                                break                                   
                            else:
                                iRow2=iRow2+1
                                last_List.append(old_List[iRow1])                   #把没有关键字增加到列表中

                                #判断应该填写在哪一列，根据模板列来填写
                                iCol2=iCol1
                                worksheet.write(iRow2,iCol2,sheet.cell(iRow1,iCol1).value)
                        else:
                            #判断应该填写在哪一列，根据模板列来填写
                            if rows[iCol1] in Dict_ModExcelCol.keys():
                                iCol2=Dict_ModExcelCol.get(rows[iCol1])
                                worksheet.write(iRow2,iCol2,sheet.cell(iRow1,iCol1).value)

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="2.FESTOReport_"+strTime+".xls"
        workbook.save(sPath+sReportName)

        print("Work is over.")
        QMessageBox.information(self,"提示信息","处理完毕，存放位置：当前工作目录\\Result \n文件名："+sReportName,QMessageBox.Yes)
    
    @pyqtSlot()
    def on_btn_Exit_clicked(self):
        """
        Slot documentation goes here.
        """
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
        directory1  = QFileDialog.getExistingDirectory(self, "选择文件", "./")
        
        #self.txt_File.setText=fileName1
        self.txt_Directory.setPlainText(directory1)
    
    @pyqtSlot()
    def on_btn_Browse2_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选择文件", "./", "Excel Files (*.xlsx)")
        print(fileName1,filetype)
        #self.txt_File.setText=fileName1
        self.txt_ModFile.setPlainText(fileName1)
    
    def InitForm(self):
        self.txt_Directory.setPlainText("")
        self.txt_ModFile.setPlainText("")
    
    #自定义函数
    def read_ModExcel(self):
        sModExcel=self.txt_ModFile.toPlainText()
        wb = xlrd.open_workbook(filename=sModExcel)    #打开文件
        sheet1 = wb.sheet_by_index(0)                  #通过索引获取表格
        rows = sheet1.row_values(0)                    #获取行内容
        max_cols=sheet1.ncols

        #建立一个Dict，把列名和所在列的列号存入Dict中
        for i in range(max_cols):
            Dict_ModExcelCol[rows[i]]=i
            modlist.append(rows[i])
            
        #print(Dict_ModExcelCol)
    
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    subForm2 = MergeFESTO()
    subForm2.show()
    sys.exit(app.exec_())
