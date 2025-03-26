import openpyxl

# Load the workbook
wb = openpyxl.load_workbook('output/Intersight_Foundation.xlsx')

# Select the Template sheet
template_sheet = wb['Template']

# Check data validations in column D
for dv in template_sheet.data_validations.dataValidation:
    if 'D2:D1000' in dv.sqref:
        print("Dropdown found in column D with options:", dv.formula1)
        break
else:
    print("No dropdown found in column D.")
