# coding: utf-8
import numpy as np
import xlrd

from . import functions


class ExcelReader:
    def __init__(self):
        self.sheet = None

    def open(self, filename, sheet_index):
        rb = xlrd.open_workbook(filename)
        self.sheet = rb.sheet_by_index(sheet_index)

    def get_array(self, start_cell, end_cell):
        start_row, start_col = functions.xl_cell_to_row_col(start_cell)
        end_row, end_col = functions.xl_cell_to_row_col(end_cell)

        end_row += 1
        end_col += 1

        data = []
        for row_num in range(start_row, end_row):
            row = self.sheet.row_values(row_num, start_col, end_col)
            data.append(row)

        return np.matrix(data)
