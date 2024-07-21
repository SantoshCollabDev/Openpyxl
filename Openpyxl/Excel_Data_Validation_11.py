import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.load_workbook("Validate_data.xlsx")
sheet = wb.active

valid_options = '"Yet to Start, In Progress, Completed"'

rule = DataValidation(type='list', formula1=valid_options, allow_blank=True)

rule.error = 'Your entry is not valid'
rule.errorTitle = 'Invalid entry'

rule.prompt = "Please select from the list."
rule.promptTitle = 'Select option'

sheet.add_data_validation(rule)   # add a rule to the sheet
rule.add('C2:C100')               # Exact column & range on which the rule to be applied


wb.save("output_valdiated.xlsx")


