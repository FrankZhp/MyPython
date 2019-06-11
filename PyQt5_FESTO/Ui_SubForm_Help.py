# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Work\Python\PyQt5\PyQt5_FESTO\SubForm_Help.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        self.btn_Exit = QtWidgets.QPushButton(Form)
        self.btn_Exit.setGeometry(QtCore.QRect(171, 180, 75, 28))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.btn_Exit.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("images/btn_Exit.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Exit.setIcon(icon)
        self.btn_Exit.setObjectName("btn_Exit")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(162, 54, 73, 16))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(36, 90, 298, 28))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.btn_Exit.setText(_translate("Form", "Exit"))
        self.label.setText(_translate("Form", "帮助信息"))
        self.label_2.setText(_translate("Form", "帮助文档存放在本项目当前目录下的Doc文件夹里"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

