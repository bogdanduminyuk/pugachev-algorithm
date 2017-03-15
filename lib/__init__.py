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

    def calculate(self, large_coords, small_coords, worksheet, result_start_cell):
        self.excel_mgr.set_sheet(worksheet)

        large_matrix = self.excel_mgr.get_array(*large_coords)
        small_matrix = self.excel_mgr.get_array(*small_coords)
        result_lambda_i0 = []

        for col in range(large_matrix.shape[1]):
            small_col = small_matrix[:, col]
            large_col = large_matrix[:, col]
            curr_lambda_i0 = self.get_lambda_i(large_col,small_col)
            result_lambda_i0.append(curr_lambda_i0)

        result_lambda_i0 = np.matrix(result_lambda_i0)
        self.excel_mgr.append_matrix(result_lambda_i0, result_start_cell)

    @staticmethod
    def get_lambda_i(large_col, small_col):
        large_col = PugachevMethod.filtrate_column(large_col, 0)
        small_col = PugachevMethod.filtrate_column(small_col, 0)
        m = np.mean(large_col[:7])
        curr_lambda = np.mean(large_col)
        kss = PugachevMethod.get_K(small_col, small_col, m)
        krs = PugachevMethod.get_K(large_col, small_col, m)

        return curr_lambda - krs*kss

    @staticmethod
    def get_K(first_col, second_col, m):
        # first = first_col.tolist()

        first = first_col - m
        """second = second_col - m
        matrix = first * second.getT()
        diagonal = matrix.diagonal()
        return np.mean(diagonal)"""
        return np.mean(first)

    @staticmethod
    def filtrate_column(column, value):
        indices = []
        for i in range(len(column)):
            if column[i][0] == value:
                indices.append(i)

        narray = np.delete(column, indices)

        return np.asmatrix(narray)


