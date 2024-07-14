import openpyxl
from openpyxl.styles import Border, Side 

wb = openpyxl.load_workbook("example.xlsx")

ws = wb['balance']

top = Side(border_style='dashed',color='1818F9')
bottom = Side(border_style='double',color='1FF918')


border = Border(top=top,bottom=bottom)


ws['B6'].border = border

wb.save("example.xlsx")



