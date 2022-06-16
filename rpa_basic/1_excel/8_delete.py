from openpyxl import load_workbook
from yaml import load
wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.delete_rows(8)  # deleting the 8th row
# ws.delete_rows(8, 3)  # deleting 3 rows from the 8th row

# ws.delete_cols(2)  # deleting the B column
ws.delete_cols(2, 2)  # deleting 2 columns from the 2nd column

wb.save("sample_delete_col.xlsx")
