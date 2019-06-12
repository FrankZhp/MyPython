# -*- coding: utf-8 -*-

"""
Module implementing OpenForm.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox

from Ui_OpenForm import Ui_Form


class OpenForm(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(OpenForm, self).__init__(parent)
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
     
    #初始化窗体界面
    def InitForm(self):
        pass
