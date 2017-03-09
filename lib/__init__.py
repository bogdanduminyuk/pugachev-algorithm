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

    def calculate(self, large_coords, small_coords, worksheet, result_start_cell):
        self.excel_mgr.set_sheet(worksheet)

        large_matrix = self.excel_mgr.get_array(*large_coords)
        small_matrix = self.excel_mgr.get_array(*small_coords)
        result_lambda_i0 = np.matrix([i for i in range(0, 500, 10)])

        print(self.get_kss(large_matrix, small_matrix))
        # do smth
        pass

        self.excel_mgr.append_matrix(result_lambda_i0, result_start_cell)

    @staticmethod
    def get_m(large_matrix, row_idx, col_idx):
        m_ast = np.mean(large_matrix[0:row_idx, col_idx])
        m = np.mean(large_matrix[row_idx:, col_idx])
        return m, m_ast

    @staticmethod
    def get_kss(large_matrix, small_matrix):
        count = 0
        sum = 0.0

        # TODO: debug it
        for i in range(small_matrix.shape[1]):
            m, m_ast = PugachevMethod.get_m(large_matrix, 7, i)
            small_matrix_col = small_matrix[:, i]
            first_multiplier = small_matrix_col - m
            second_multiplier = first_multiplier.getT()
            mult = first_multiplier * second_multiplier
            count += 1

        return sum/count

