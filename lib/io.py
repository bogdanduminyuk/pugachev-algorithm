# coding: utf-8
import numpy as np
import xlrd
import xlwt
from xlutils import copy

from . import functions


class ExcelManager:
    def __init__(self):
        self.rb = None
        self.wb = None
        self.r_sheet = None
        self.w_sheet = None
        self.format = None
        self.sheets = None

    def open(self, filename, sheet_index):
        self.rb = xlrd.open_workbook(filename, formatting_info=True)
        self.sheets = self.rb.sheets()
        self.wb = copy.copy(self.rb)
        self.r_sheet = self.rb.sheet_by_index(sheet_index)
        self.w_sheet = self.wb.get_sheet(sheet_index)

    def get_array(self, start_cell, end_cell):
        start_row, start_col = functions.xl_cell_to_row_col(start_cell)
        end_row, end_col = functions.xl_cell_to_row_col(end_cell)

        end_row += 1
        end_col += 1

        # TODO: try cell.type

        data = []
        for row_num in range(start_row, end_row):
            row = self.r_sheet.row_values(row_num, start_col, end_col)
            data.append(row)

        return np.matrix(data)

    def append_matrix(self, matrix, start_cell):
        start_row, start_col = functions.xl_cell_to_row_col(start_cell)
        style = xlwt.XFStyle()
        style.num_format_str = '0.00'

        for row_num, row in enumerate(matrix.tolist(), start=start_row):
            for col_num, col_element in enumerate(row, start=start_col):
                self.w_sheet.write(row_num, col_num, float(col_element), self.format)

    def save(self, filename):
        self.wb.save(filename)
