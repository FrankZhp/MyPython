# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Work\Python\PyQt5\PyQt5_Test1\NewForm.ui'
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
        self.btn_Exit.setGeometry(QtCore.QRect(144, 135, 82, 37))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("images/btn_Exit.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_Exit.setIcon(icon)
        self.btn_Exit.setObjectName("btn_Exit")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "NewForm"))
        self.btn_Exit.setText(_translate("Form", "Exit"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

