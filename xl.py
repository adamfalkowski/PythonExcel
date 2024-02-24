from openpyxl import Workbook, load_workbook

wb = load_workbook('xl.xlsx')
ws = wb.active

wb.create_sheet("Testing")
print(wb.sheetnames)
wb.save('xl.xlsx')