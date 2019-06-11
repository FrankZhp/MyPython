# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Work\Python\PyQt5\PyQt5_FESTO\SubForm7_GetUnitPrices.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(755, 249)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        Form.setFont(font)
        self.btn_Process = QtWidgets.QPushButton(Form)
        self.btn_Process.setGeometry(QtCore.QRect(230, 140, 101, 41))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.btn_Process.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("images/btn_ok.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Process.setIcon(icon)
        self.btn_Process.setObjectName("btn_Process")
        self.btn_Exit = QtWidgets.QPushButton(Form)
        self.btn_Exit.setGeometry(QtCore.QRect(390, 140, 101, 41))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.btn_Exit.setFont(font)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("images/btn_Exit.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Exit.setIcon(icon1)
        self.btn_Exit.setObjectName("btn_Exit")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(40, 70, 91, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.btn_Browse = QtWidgets.QPushButton(Form)
        self.btn_Browse.setGeometry(QtCore.QRect(640, 70, 101, 31))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.btn_Browse.setFont(font)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("images/btn_File_Browser.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Browse.setIcon(icon2)
        self.btn_Browse.setObjectName("btn_Browse")
        self.txt_ECBRateFile = QtWidgets.QPlainTextEdit(Form)
        self.txt_ECBRateFile.setGeometry(QtCore.QRect(140, 70, 491, 28))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.txt_ECBRateFile.setFont(font)
        self.txt_ECBRateFile.setPlainText("")
        self.txt_ECBRateFile.setObjectName("txt_ECBRateFile")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(44, 20, 91, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.cbo_Month = QtWidgets.QComboBox(Form)
        self.cbo_Month.setGeometry(QtCore.QRect(140, 20, 191, 28))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.cbo_Month.setFont(font)
        self.cbo_Month.setObjectName("cbo_Month")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Step7 Get Unit Prices Style 1"))
        self.btn_Process.setText(_translate("Form", "Process"))
        self.btn_Exit.setText(_translate("Form", "Exit"))
        self.label.setText(_translate("Form", "ECB Rate File:"))
        self.btn_Browse.setText(_translate("Form", " Browse"))
        self.label_2.setText(_translate("Form", "Select Month:"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

