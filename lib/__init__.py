# coding: utf-8
from .io import ExcelManager


class PugachevMethod:
    def __init__(self):
        self.filename = ''
        self.large_sample, self.small_sample = None, None
        self.worksheet_index, self.result_start_cell = None, None

        self.excel_mgr = ExcelManager()

    def open_file(self, filename):
        """
        Opens file and returns count of sheets

        :param filename: input file name
        :return: count of sheets
        """
        self.filename = filename
        self.excel_mgr.open(filename)

        return self.excel_mgr.get_sheets_count()

    def set_sheet(self, index):
        self.excel_mgr.set_sheet(index)

    def calculate(self):
        pass
