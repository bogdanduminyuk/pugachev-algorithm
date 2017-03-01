# coding: utf-8

from lib import io


# input vars
filename = 'data/pugachev.xls'
start_cell = 'B6'
end_cell = 'D11'
sheet_index = 0
result_start_cell = 'B1478'


# program
if __name__ == "__main__":
    e_mgr = io.ExcelManager()
    e_mgr.open(filename, sheet_index)
    matrix = e_mgr.get_array(start_cell, end_cell)
    matrix_m = matrix * matrix.T

    e_mgr.append_matrix(matrix_m, result_start_cell)
    e_mgr.save('test.xls')
