# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow,  QMessageBox , QAction
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon

from NewForm import NewForm
from OpenForm import OpenForm
from SaveForm import SaveForm
from AboutForm import AboutForm
from HelpForm  import HelpForm 

from Ui_QMainWindow import Ui_MainWindow


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
        
        #工具栏
        # 添加打开tool 在pyqt5里面是一个action
        tb1  = QAction(QIcon("./images/menuNew.png"), "New", self)
        self.toolBar.addAction(tb1)
        tb2  = QAction(QIcon("./images/MenuOpen.png"), "Open", self)
        self.toolBar.addAction(tb2)
        tb3  = QAction(QIcon("./images/MenuSave.png"), "Save", self)
        self.toolBar.addAction(tb3)
        
        #连接槽函数
        self.toolBar.actionTriggered[QAction].connect(self.toolBtnPressed)
        
        #状态栏
        self.statusBar.showMessage('Frank Zhang Support')
      
    
    @pyqtSlot()
    def on_actionNew_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        subForm1.show()
    
    @pyqtSlot()
    def on_actionOpen_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        subForm2.show()
    
    @pyqtSlot()
    def on_actionSave_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise No#ImplementedError
        subForm3.show()
    
    @pyqtSlot()
    def on_actionHelp_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        subForm4.show()
    
    @pyqtSlot()
    def on_actionAbout_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        subForm5.show()
    
    @pyqtSlot()
    def on_actionExit_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        reply = QMessageBox.question(self, "标题", "确定要退出系统吗?", 
                   QMessageBox.Yes|QMessageBox.No)
         
        if reply == QMessageBox.Yes:
            sys.exit(0)
    
    def toolBtnPressed(self,qaction):
        print("pressed too btn is",qaction.text())
        if  qaction.text()=="New":
            print("Click ToolBar New Button")
            subForm1.show()
            subForm1.InitForm()
        elif qaction.text()=='Open':
            print("Click ToolBar Open Button")
            subForm2.show()
            subForm2.InitForm()
        elif qaction.text()=='Save':
            print("Click ToolBar Save Button")
            subForm3.show()
            subForm3.InitForm()
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    mainForm = MainWindow()
    mainForm.show()
    
    #实例化子窗体
    subForm1 = NewForm()
    subForm2 = OpenForm()
    subForm3 = SaveForm()
    subForm4 = HelpForm()
    subForm5 = AboutForm()
 
    sys.exit(app.exec_())
