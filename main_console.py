# coding: utf-8
from lib import PugachevMethod

res_cell = 'B1478'
large_sample = 'B6', 'AL1387'
small_sample = 'B1397', 'AL1474'
filename = 'data/pugachev.xls'
sheet = 'Лист1'

method = PugachevMethod()
method.open_file(filename)
method.calculate(large_sample, small_sample, sheet, res_cell)
