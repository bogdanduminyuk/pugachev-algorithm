# coding: utf-8
import numpy as np
import xlrd
import xlwt
from xlutils import copy

from . import functions


class ExcelManager:
    def __init__(self):
        # rb - read book, wb - write book
        # r_sheet - read sheet, w_sheet - write sheet
        self.rb, self.wb = None, None
        self.sheets, self.r_sheet, self.w_sheet = None, None, None

        self.format = xlwt.XFStyle()
        self.format.num_format_str = '0.00'

    def open(self, filename):
        """
        Opens filename. Gets necessary info and keeps it in self.vars.

        :param filename: name of the excel book
        :return: None
        """
        self.rb = xlrd.open_workbook(filename, formatting_info=True)
        self.sheets = self.rb.sheets()
        self.wb = copy.copy(self.rb)

    def set_sheet(self, sheet_index):
        self.r_sheet = self.rb.sheet_by_index(sheet_index)
        self.w_sheet = self.wb.get_sheet(sheet_index)

    def get_array(self, start_cell, end_cell):
        """
        Returns numpy matrix of the given cells.

        Cells should be Excel cells (e.g. A1, B3... etc).

        :param start_cell: left top matrix cell
        :param end_cell: right bottom matrix cell
        :return: numpy matrix
        """
        start_row, start_col = functions.xl_cell_to_row_col(start_cell)
        end_row, end_col = functions.xl_cell_to_row_col(end_cell)

        end_row += 1
        end_col += 1

        data = []
        for row_num in range(start_row, end_row):
            row = self.r_sheet.row_values(row_num, start_col, end_col)
            data.append(row)

        return np.matrix(data)

    def append_matrix(self, matrix, start_cell):
        """
        Appends matrix to the input book.

        :param matrix: matrix that will be appended
        :param start_cell: left top cell position in given sheet
        :return: None
        """
        start_row, start_col = functions.xl_cell_to_row_col(start_cell)

        for row_num, row in enumerate(matrix.tolist(), start=start_row):
            for col_num, col_element in enumerate(row, start=start_col):
                self.w_sheet.write(row_num, col_num, float(col_element), self.format)

    def save(self, filename):
        """
        Saves the current workbook to the given filename.

        :param filename: path where workbook will be saved
        :return: None
        """
        self.wb.save(filename)
