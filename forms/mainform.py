# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'forms/ui/mainform.ui'
#
# Created by: PyQt5 UI code generator 5.8.1
#
# WARNING! All changes made in this file will be lost!
import os

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from lib import PugachevMethod


class Ui_MainWindow(object):
    def __init__(self):
        self.method = PugachevMethod()

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowModality(QtCore.Qt.NonModal)
        MainWindow.resize(640, 471)
        MainWindow.setMinimumSize(QtCore.QSize(400, 400))
        MainWindow.setMaximumSize(QtCore.QSize(1000, 1000))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 13, 111, 16))
        self.label.setObjectName("label")
        self.btn_choose_file = QtWidgets.QPushButton(self.centralwidget)
        self.btn_choose_file.setGeometry(QtCore.QRect(140, 10, 101, 23))
        self.btn_choose_file.setObjectName("btn_choose_file")
        self.groupBox_data = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_data.setEnabled(False)
        self.groupBox_data.setGeometry(QtCore.QRect(20, 50, 561, 291))
        self.groupBox_data.setObjectName("groupBox_data")
        self.label_2 = QtWidgets.QLabel(self.groupBox_data)
        self.label_2.setGeometry(QtCore.QRect(10, 23, 31, 16))
        self.label_2.setObjectName("label_2")
        self.comboBox_list = QtWidgets.QComboBox(self.groupBox_data)
        self.comboBox_list.setGeometry(QtCore.QRect(70, 20, 121, 22))
        self.comboBox_list.setObjectName("comboBox_list")
        self.groupBox_large_sample = QtWidgets.QGroupBox(self.groupBox_data)
        self.groupBox_large_sample.setGeometry(QtCore.QRect(40, 70, 211, 101))
        self.groupBox_large_sample.setObjectName("groupBox_large_sample")
        self.label_4 = QtWidgets.QLabel(self.groupBox_large_sample)
        self.label_4.setGeometry(QtCore.QRect(10, 61, 71, 16))
        self.label_4.setObjectName("label_4")
        self.label_3 = QtWidgets.QLabel(self.groupBox_large_sample)
        self.label_3.setGeometry(QtCore.QRect(30, 33, 47, 13))
        self.label_3.setObjectName("label_3")
        self.lineEdit_top_left = QtWidgets.QLineEdit(self.groupBox_large_sample)
        self.lineEdit_top_left.setGeometry(QtCore.QRect(90, 30, 113, 20))
        self.lineEdit_top_left.setObjectName("lineEdit_top_left")
        self.lineEdit_right_bottom = QtWidgets.QLineEdit(self.groupBox_large_sample)
        self.lineEdit_right_bottom.setGeometry(QtCore.QRect(90, 60, 113, 20))
        self.lineEdit_right_bottom.setObjectName("lineEdit_right_bottom")
        self.groupBox_small_sample = QtWidgets.QGroupBox(self.groupBox_data)
        self.groupBox_small_sample.setGeometry(QtCore.QRect(300, 70, 211, 101))
        self.groupBox_small_sample.setObjectName("groupBox_small_sample")
        self.label_5 = QtWidgets.QLabel(self.groupBox_small_sample)
        self.label_5.setGeometry(QtCore.QRect(10, 61, 71, 16))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.groupBox_small_sample)
        self.label_6.setGeometry(QtCore.QRect(30, 33, 47, 13))
        self.label_6.setObjectName("label_6")
        self.lineEdit_top_left_2 = QtWidgets.QLineEdit(self.groupBox_small_sample)
        self.lineEdit_top_left_2.setGeometry(QtCore.QRect(90, 30, 113, 20))
        self.lineEdit_top_left_2.setObjectName("lineEdit_top_left_2")
        self.lineEdit__right_bottom = QtWidgets.QLineEdit(self.groupBox_small_sample)
        self.lineEdit__right_bottom.setGeometry(QtCore.QRect(90, 60, 113, 20))
        self.lineEdit__right_bottom.setObjectName("lineEdit__right_bottom")
        self.groupBox_4 = QtWidgets.QGroupBox(self.groupBox_data)
        self.groupBox_4.setGeometry(QtCore.QRect(170, 190, 211, 80))
        self.groupBox_4.setObjectName("groupBox_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(self.groupBox_4)
        self.lineEdit_5.setGeometry(QtCore.QRect(40, 30, 141, 20))
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.btn_calc = QtWidgets.QPushButton(self.centralwidget)
        self.btn_calc.setEnabled(False)
        self.btn_calc.setGeometry(QtCore.QRect(240, 360, 141, 41))
        self.btn_calc.setObjectName("btn_calc")
        self.label_filename = QtWidgets.QLabel(self.centralwidget)
        self.label_filename.setGeometry(QtCore.QRect(260, 15, 107, 13))
        self.label_filename.setText("")
        self.label_filename.setObjectName("label_filename")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 640, 21))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        self.menu_2 = QtWidgets.QMenu(self.menubar)
        self.menu_2.setObjectName("menu_2")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action = QtWidgets.QAction(MainWindow)
        self.action.setObjectName("action")
        self.action_2 = QtWidgets.QAction(MainWindow)
        self.action_2.setObjectName("action_2")
        self.action_4 = QtWidgets.QAction(MainWindow)
        self.action_4.setObjectName("action_4")
        self.action_5 = QtWidgets.QAction(MainWindow)
        self.action_5.setObjectName("action_5")
        self.action_6 = QtWidgets.QAction(MainWindow)
        self.action_6.setObjectName("action_6")
        self.menu.addAction(self.action)
        self.menu.addAction(self.action_2)
        self.menu.addSeparator()
        self.menu.addAction(self.action_4)
        self.menu_2.addAction(self.action_5)
        self.menu_2.addAction(self.action_6)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menu_2.menuAction())

        self.retranslateUi(MainWindow)
        self.action_4.triggered.connect(MainWindow.close)
        self.btn_choose_file.clicked.connect(self.action.trigger)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.action.triggered.connect(self.open_file)
        self.action_2.triggered.connect(self.close_file)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Метод Пугачева"))
        self.label.setText(_translate("MainWindow", "Выберите имя файла:"))
        self.btn_choose_file.setText(_translate("MainWindow", "Выбрать файл"))
        self.groupBox_data.setTitle(_translate("MainWindow", "Данные для обработки:"))
        self.label_2.setText(_translate("MainWindow", "Лист:"))
        self.groupBox_large_sample.setTitle(_translate("MainWindow", "Большая выборка:"))
        self.label_4.setText(_translate("MainWindow", "RightBottom:"))
        self.label_3.setText(_translate("MainWindow", "TopLeft:"))
        self.groupBox_small_sample.setTitle(_translate("MainWindow", "Малая выборка:"))
        self.label_5.setText(_translate("MainWindow", "RightBottom:"))
        self.label_6.setText(_translate("MainWindow", "TopLeft:"))
        self.groupBox_4.setTitle(_translate("MainWindow", "Ячейка результата:"))
        self.btn_calc.setText(_translate("MainWindow", "Расчитать"))
        self.menu.setTitle(_translate("MainWindow", "Файл"))
        self.menu_2.setTitle(_translate("MainWindow", "Справка"))
        self.action.setText(_translate("MainWindow", "Открыть"))
        self.action.setShortcut(_translate("MainWindow", "Ctrl+O"))
        self.action_2.setText(_translate("MainWindow", "Закрыть"))
        self.action_2.setShortcut(_translate("MainWindow", "Ctrl+F"))
        self.action_4.setText(_translate("MainWindow", "Выход"))
        self.action_5.setText(_translate("MainWindow", "Помощь"))
        self.action_5.setShortcut(_translate("MainWindow", "Ctrl+F1"))
        self.action_6.setText(_translate("MainWindow", "О программе"))

    def open_file(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Открыть файл", "", "xls-файлы (*.xls)")

        if filename:
            sheets = self.method.open_file(filename)

            self.label_filename.setText(os.path.basename(filename))
            self.comboBox_list.addItems(sheets)
            self.groupBox_data.setEnabled(True)

    def close_file(self):
        self.method = PugachevMethod()
        self.comboBox_list.clear()
        self.groupBox_data.setEnabled(False)
        self.label_filename.setText('')

