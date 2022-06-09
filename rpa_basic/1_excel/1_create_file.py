from openpyxl import Workbook
wb = Workbook()  # Creating a new workbook in excel
ws = wb.active  # Bringing in an activated sheet
ws.title = 'MySheet'  # Changing the name of the sheet
wb.save("sample.xlsx")
wb.close()
