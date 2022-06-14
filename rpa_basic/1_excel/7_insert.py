from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8) # emptying the 8th row
# ws.insert_rows(8, 5)  # inserting 5 rows at the 8th row
# wb.save('sample_insert_rows.xlsx')

ws.insert_cols(2)
wb.save("sample_insert_columns.xlsx")
