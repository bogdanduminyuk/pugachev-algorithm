# coding: utf-8

import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow
from forms.ui_mainwindow_realization import UiMainWindowRealization

if __name__ == "__main__":
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = \
        "E:\Cloud.Mail.ru\Учеба\WorkSpace\Python\Pugachev\env\Lib\site-packages\PyQt5\Qt\plugins\platforms"

    try:
        app = QApplication(sys.argv)
        window = QMainWindow()

        ui = UiMainWindowRealization()
        ui.setupUi(window)
        window.show()

        sys.exit(app.exec_())

    except Exception as e:
        print(type(e))
        print(e)
