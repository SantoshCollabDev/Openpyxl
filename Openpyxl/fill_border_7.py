import openpyxl
from openpyxl.styles import Border, Side

wb = openpyxl.load_workbook('Example.xlsx')

ws = wb['Balance']

top_border    = Side(border_style='dashed', color='FF0627')
bottom_border = Side(border_style='double', color='06FF59')
fill_border = Border(top=top_border, bottom=bottom_border)

# apply the border to cell 'B7'
ws['B7'].border = fill_border

wb.save("Example.xlsx")