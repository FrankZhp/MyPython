# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Work\Python\PyQt5\PyQt5_FESTO\SubForm6_AppendInvoice.ui'
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
        self.btn_Process.setGeometry(QtCore.QRect(240, 160, 101, 41))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.btn_Process.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("images/btn_ok.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Process.setIcon(icon)
        self.btn_Process.setObjectName("btn_Process")
        self.btn_Exit = QtWidgets.QPushButton(Form)
        self.btn_Exit.setGeometry(QtCore.QRect(400, 160, 101, 41))
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
        self.label.setGeometry(QtCore.QRect(28, 30, 121, 31))
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
        self.txt_InvoiceFile = QtWidgets.QPlainTextEdit(Form)
        self.txt_InvoiceFile.setGeometry(QtCore.QRect(140, 30, 491, 28))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.txt_InvoiceFile.setFont(font)
        self.txt_InvoiceFile.setPlainText("")
        self.txt_InvoiceFile.setObjectName("txt_InvoiceFile")
        self.txt_SummaryFile = QtWidgets.QPlainTextEdit(Form)
        self.txt_SummaryFile.setGeometry(QtCore.QRect(140, 80, 491, 28))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.txt_SummaryFile.setFont(font)
        self.txt_SummaryFile.setPlainText("")
        self.txt_SummaryFile.setObjectName("txt_SummaryFile")
        self.btn_Browse2 = QtWidgets.QPushButton(Form)
        self.btn_Browse2.setGeometry(QtCore.QRect(640, 80, 101, 28))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.btn_Browse2.setFont(font)
        self.btn_Browse2.setIcon(icon2)
        self.btn_Browse2.setObjectName("btn_Browse2")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(10, 80, 121, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Step 6 Append Invoice"))
        self.btn_Process.setText(_translate("Form", "Process"))
        self.btn_Exit.setText(_translate("Form", "Exit"))
        self.label.setText(_translate("Form", "Invoice Excel File:"))
        self.btn_Browse1.setText(_translate("Form", " Browse"))
        self.btn_Browse2.setText(_translate("Form", " Browse"))
        self.label_2.setText(_translate("Form", "Summary Excel File:"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

