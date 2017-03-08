# coding: utf-8
import os

from PyQt5.QtWidgets import QFileDialog, QMessageBox

from forms.mainform import Ui_MainWindow
from lib import PugachevMethod


class UiMainWindowRealization(Ui_MainWindow):
    def __init__(self):
        self.method = PugachevMethod()

    def setupUi(self, MainWindow):
        super(UiMainWindowRealization, self).setupUi(MainWindow)

        self.action.triggered.connect(self.open_file)
        self.action_2.triggered.connect(self.close_file)

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

