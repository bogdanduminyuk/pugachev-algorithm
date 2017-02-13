# coding: utf-8
import xlrd
import numpy as np

# input vars
filename = 'data/pugachev.xls'
start_row = 6
start_col = 2
width = 2
height = 4

# local vars
start_row -= 1
start_col -= 1
end_row = start_row + height
end_col = start_col + width
matrix = []

# program
rb = xlrd.open_workbook(filename)
sheet = rb.sheet_by_index(0)

for row_num in range(start_row, start_row + height):
    row = sheet.row_values(row_num)

    curr = []
    for i in range(start_col, start_col + width):
        curr.append(int(row[i]))

    matrix.append(curr)

matrix = np.matrix(matrix)
print(matrix)
print('---------')
print(matrix.getT())
print('---------')
print(matrix * matrix.T)
