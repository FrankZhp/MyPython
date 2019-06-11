# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QAction, QMessageBox 
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon

from Ui_QMainWindow import Ui_MainWindow

from SubForm1_CSVToExcel import CSVToExcel
from SubForm2_MergeFESTO import MergeFESTO
from SubForm3_Mapping import Mapping
from SubForm4_Summary import  Summary
from SubForm5_MergeInvoice import MergeInvoice
from SubForm6_AppendInvoice import AppendInvoice
from SubForm7_GetUnitPrices import GetUnitPricesStyle1
from SubForm8_GetUnitPrices import GetUnitPricesStyle2
from SubForm9_MakeReport import MakeReport

from SubForm_Help import HelpForm
from SubForm_About import AboutForm


class MainWindow(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        
        #增加部分
        self.setWindowTitle('Compute FESTO Tickets prices')
        
        #工具栏
        # 添加打开tool 在pyqt5里面是一个action
        tb1  = QAction(QIcon("./images/CSV_Excel_ToolBar.png"), "1.CSV To Excel", self)
        self.toolBar.addAction(tb1)
        tb2  = QAction(QIcon("./images/MergeExcel_ToolBar.png"), "2.Merge FESTO Files", self)
        self.toolBar.addAction(tb2)
        tb3  = QAction(QIcon("./images/Mapping_ToolBar.png"), "3.Mapping", self)
        self.toolBar.addAction(tb3)
        tb4  = QAction(QIcon("./images/Summary_ToolBar.png"), "4.Summary", self)
        self.toolBar.addAction(tb4)
        
        tb5  = QAction(QIcon("./images/MergeInvoice_ToolBar.png"), "5.Merge Invoice", self)
        self.toolBar.addAction(tb5)
        
        tb6  = QAction(QIcon("./images/AppendInvoice_ToolBar.png"), "6.Append Invoice To Summary", self)
        self.toolBar.addAction(tb6)
        
        tb7  = QAction(QIcon("./images/UnitPriceStyle1_ToolBar.png"), "7.Get Unit Prices Style 1", self)
        self.toolBar.addAction(tb7)
        
        tb8  = QAction(QIcon("./images/UnitPriceStyle2_ToolBar.png"), "8.Get Unit Prices Style 2", self)
        self.toolBar.addAction(tb8)
        
        tb9  = QAction(QIcon("./images/MakeReport_ToolBar.png"), "9.Make Report", self)
        self.toolBar.addAction(tb9)
        self.toolBar.actionTriggered[QAction].connect(self.toolBtnPressed)
        
        #状态栏
        self.statusBar.showMessage('Frank Zhang Support')

    def toolBtnPressed(self,qaction):
        print("pressed too btn is",qaction.text())
        if  qaction.text()=="1.CSV To Excel":
            print("Click 1")
            subForm1.show()
            subForm1.InitForm()
        elif qaction.text()=='2.Merge FESTO Files':
            print("Click 2")
            subForm2.show()
            subForm2.InitForm()
        elif qaction.text()=='3.Mapping':
            print("Click 3")
            subForm3.show()
            subForm3.InitForm()
        elif qaction.text()=='4.Summary':
            print("Click 4")
            subForm4.show()
            subForm4.InitForm()
        elif qaction.text()=='5.Merge Invoice':
            print("Click 5")
            subForm5.show()
            subForm5.InitForm()
        elif qaction.text()=='6.Append Invoice To Summary':
            print("Click 6")
            subForm6.show()
            subForm6.InitForm()
        elif qaction.text()=='7.Get Unit Prices Style 1':
            print("Click 7")
            subForm7.show()
            subForm7.InitForm()
        elif qaction.text()=='8.Get Unit Prices Style 2':
            print("Click 8")
            subForm8.show()
            subForm8.InitForm()
        elif qaction.text()=='9.Make Report':
            print("Click 9")
            subForm9.show()
            subForm9.InitForm()
        

    @pyqtSlot()
    def on_action11_CSV_File_To_Excel_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 1.CSV To Excel")
        subForm1.show()
        subForm1.InitForm()
        
    @pyqtSlot()
    def on_action12_Merge_FESTO_Files_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 2.Merge FESTO Files")
        subForm2.show()
        subForm2.InitForm()
    
    
    @pyqtSlot()
    def on_action13_Mapping_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 3.Mapping")
        subForm3.show()
        subForm3.InitForm()
    
    @pyqtSlot()
    def on_action14_Summary_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 4.Summary")
        subForm4.show()
        subForm4.InitForm()
    
    @pyqtSlot()
    def on_action15_Merge_Invoice_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 5.Merge Invoice")
        subForm5.show()
        subForm5.InitForm()
    
    @pyqtSlot()
    def on_action16_Append_Invoice_to_Summary_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 6.Append Invoice To Summary")
        subForm6.show()
        subForm6.InitForm()
    
    @pyqtSlot()
    def on_action17_Get_Unit_Prices_Style_1_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 7.Get Unit Prices Style 1")
        subForm7.show()
        subForm7.InitForm()
    
    @pyqtSlot()
    def on_action18_Get_Unit_Prices_Style_2_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 8.Get Unit Prices Style 2")
        subForm8.show()
        subForm8.InitForm()
    
    @pyqtSlot()
    def on_action19_Compute_Tickets_Price_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu 9.Compute Tickets Price")
        subForm9.show()
        subForm9.InitForm()

    
    @pyqtSlot()
    def on_actionExit_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu Exit")
        reply = QMessageBox.question(self, "标题", "确定要退出系统吗?", 
                   QMessageBox.Yes|QMessageBox.No)
         
        if reply == QMessageBox.Yes:
            sys.exit(0)
        
    @pyqtSlot()
    def on_action21_Help_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu Help")
        helpForm.show()
    
    @pyqtSlot()
    def on_action22_About_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu About")
        aboutForm.show()
    
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    mainForm = MainWindow()
    mainForm.show()
    
      
    #实例化子窗体
    #子菜单1
    subForm1 = CSVToExcel()
    subForm2 = MergeFESTO()
    subForm3 = Mapping()
    subForm4 = Summary()
    subForm5 = MergeInvoice()
    subForm6 = AppendInvoice()
    subForm7 = GetUnitPricesStyle1()
    subForm8 = GetUnitPricesStyle2()
    subForm9 = MakeReport()
    
    #子菜单2
    aboutForm = AboutForm()
    helpForm    = HelpForm()
    
    sys.exit(app.exec_())
    
   
