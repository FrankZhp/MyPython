# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Work\Python\PyQt5\PyQt5_FESTO\SubForm2_MergeFESTO.ui'
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
        self.btn_Process.setGeometry(QtCore.QRect(240, 170, 101, 41))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.btn_Process.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("images/btn_ok.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Process.setIcon(icon)
        self.btn_Process.setObjectName("btn_Process")
        self.btn_Exit = QtWidgets.QPushButton(Form)
        self.btn_Exit.setGeometry(QtCore.QRect(400, 170, 101, 41))
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
        self.label.setGeometry(QtCore.QRect(30, 30, 151, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.btn_Browse1 = QtWidgets.QPushButton(Form)
        self.btn_Browse1.setGeometry(QtCore.QRect(640, 30, 101, 28))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.btn_Browse1.setFont(font)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("images/btn_File_Browser.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Browse1.setIcon(icon2)
        self.btn_Browse1.setObjectName("btn_Browse1")
        self.txt_Directory = QtWidgets.QPlainTextEdit(Form)
        self.txt_Directory.setGeometry(QtCore.QRect(190, 30, 441, 28))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.txt_Directory.setFont(font)
        self.txt_Directory.setPlainText("")
        self.txt_Directory.setObjectName("txt_Directory")
        self.lbl_ModFile = QtWidgets.QLabel(Form)
        self.lbl_ModFile.setGeometry(QtCore.QRect(80, 80, 101, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lbl_ModFile.setFont(font)
        self.lbl_ModFile.setObjectName("lbl_ModFile")
        self.txt_ModFile = QtWidgets.QPlainTextEdit(Form)
        self.txt_ModFile.setGeometry(QtCore.QRect(190, 80, 441, 28))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.txt_ModFile.setFont(font)
        self.txt_ModFile.setPlainText("")
        self.txt_ModFile.setObjectName("txt_ModFile")
        self.btn_Browse2 = QtWidgets.QPushButton(Form)
        self.btn_Browse2.setGeometry(QtCore.QRect(640, 80, 101, 28))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.btn_Browse2.setFont(font)
        self.btn_Browse2.setIcon(icon2)
        self.btn_Browse2.setObjectName("btn_Browse2")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Step 2 Merge FESTO Files"))
        self.btn_Process.setText(_translate("Form", "Process"))
        self.btn_Exit.setText(_translate("Form", "Exit"))
        self.label.setText(_translate("Form", "Select FESTO Files Path:"))
        self.btn_Browse1.setText(_translate("Form", " Browse"))
        self.lbl_ModFile.setText(_translate("Form", "Select Mod File:"))
        self.btn_Browse2.setText(_translate("Form", " Browse"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

