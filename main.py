# coding: utf-8

from lib import PugachevMethod


# input vars
filename = 'data/pugachev.xls'
start_cell = 'B6'
end_cell = 'D11'
sheet_index = 0
result_start_cell = 'B1478'


# program
if __name__ == "__main__":
    method = PugachevMethod()
    sheets_count = method.open_file(filename)

    # here chosen 0 sheet from range(0, sheet_count)
    current_sheet_index = 0
    large_sample = 'B6', 'AL1387'
    small_sample = 'B1397', 'AL1474'
    result_start_cell = 'B1478'

    method.calculate(large_sample, small_sample,
                     worksheet_index=current_sheet_index,
                     result_start_cell=result_start_cell)

    # matrix_m = matrix * matrix.T
