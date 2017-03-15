# coding: utf-8
import os

from PyQt5.QtWidgets import QFileDialog, QMessageBox, QDialog

from forms.mainform import Ui_MainWindow
from forms.about import Ui_Dialog
from lib import PugachevMethod


class UiMainWindowRealization(Ui_MainWindow):
    def __init__(self):
        self.method = PugachevMethod()

    def setupUi(self, MainWindow):
        super(UiMainWindowRealization, self).setupUi(MainWindow)

        self.action.triggered.connect(self.open_file)
        self.action_2.triggered.connect(self.close_file)
        self.action_6.triggered.connect(self.about)
        self.btn_calc.clicked.connect(self.calc)

        self.lineEdit_5.setText('B1478')
        self.lineEdit_top_left.setText('B6')
        self.lineEdit_right_bottom.setText('AL1387')
        self.lineEdit_top_left_2.setText('B1397')
        self.lineEdit__right_bottom.setText('AL1474')

    def open_file(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Открыть файл", "", "xls-файлы (*.xls)")

        if filename:
            sheets = self.method.open_file(filename)

            self.label_filename.setText(os.path.basename(filename))
            self.comboBox_list.addItems(sheets)
            self.groupBox_data.setEnabled(True)
            self.btn_calc.setEnabled(True)

            self.statusbar.showMessage("Открытие файла прошло успешно!")

    def close_file(self):
        self.method = PugachevMethod()
        self.comboBox_list.clear()
        self.groupBox_data.setEnabled(False)
        self.btn_calc.setEnabled(False)
        self.label_filename.setText('')

    def calc(self):
        large_sample = self.lineEdit_top_left.text(), self.lineEdit_right_bottom.text()
        small_sample = self.lineEdit_top_left_2.text(), self.lineEdit__right_bottom.text()
        worksheet = self.comboBox_list.currentText()
        result_cell = self.lineEdit_5.text()

        self.method.calculate(large_sample, small_sample,
                              worksheet=worksheet,
                              result_start_cell=result_cell)

        filename, _ = QFileDialog.getSaveFileName(None, "Сохранить как", "", "xls-файлы (*.xls)")
        self.method.excel_mgr.save(filename)
        self.statusbar.showMessage("Файл {} был успешно сохранен!".format(filename))

    def about(self):
        self.statusbar.showMessage('Вызвано диалоговое окно "О программе"')
        about_window = QDialog()
        ui = Ui_Dialog()
        ui.setupUi(about_window)

        about_window.setModal(True)
        about_window.exec()
        self.statusbar.showMessage('')

