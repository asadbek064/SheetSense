import pandas as pd
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active

formulas = [
    ['=A1/0', '=B1*C1', '=UNKNOWNFUNC()', '=D1+E1'],
    ['=VLOOKUP("x",A1:A2,3,FALSE)', '=SUM(#REF!)', '=VALUE("abc")', '=NULL'],
    ['=1/0', '=NA()', '=NAME?', '=VALUE!'],
    ['=IF(A1="",,)', '=INDIRECT("invalid")', '=1+"text"', '=SUM()']
]

for row_idx, row in enumerate(formulas, start=1):
    for col_idx, formula in enumerate(row, start=1):
        ws.cell(row=row_idx, column=col_idx, value=formula)

wb.save('formula_errors.xlsx')
