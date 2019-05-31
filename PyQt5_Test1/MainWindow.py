# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QWidget 
from PyQt5 import QtWidgets

from Ui_MainWindow import Ui_MainWindow
from Ui_ChildForm import Ui_ChildForm


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
    
    @pyqtSlot()
    def on_btn_Open_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        ch.show()
    
    @pyqtSlot()
    def on_actionOpen_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        ch.show()
    
    @pyqtSlot()
    def on_actionExit_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        sys.exit(0)
        
class ChildForm(QWidget, Ui_ChildForm):
    def __init__(self):
        #super(QWidget, self).__init__(parent)
        super(QWidget, self).__init__()
        self.setupUi(self)
        self.btn_Exit.clicked.connect(self.close)
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    
    #实例化主窗体
    mainForm = MainWindow()
    
    #显示主窗体
    mainForm.show()
    
    #实例化子窗体
    ch=ChildForm()
    
    sys.exit(app.exec_())
