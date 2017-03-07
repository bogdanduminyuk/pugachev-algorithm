# coding: utf-8
from .io import ExcelManager
import numpy as np

class PugachevMethod:
    def __init__(self):
        self.filename = ''
        self.excel_mgr = ExcelManager()

    def open_file(self, filename):
        """
        Opens file and returns count of sheets

        :param filename: input file name
        :return: count of sheets
        """
        self.filename = filename
        self.excel_mgr.open(filename)

        return self.excel_mgr.get_sheets()

    def set_sheet(self, index):
        self.excel_mgr.set_sheet(index)

    def calculate(self, large_coords, small_coords, worksheet_index, result_start_cell):
        self.excel_mgr.set_sheet(worksheet_index)

        large_matrix = self.excel_mgr.get_array(*large_coords)
        small_matrinx = self.excel_mgr.get_array(*small_coords)
        result_lambda_i0 = np.matrix([i for i in range(0, 500, 10)])

        # do smth
        pass

        self.excel_mgr.append_matrix(result_lambda_i0, result_start_cell)
        self.excel_mgr.save('test.xls')
