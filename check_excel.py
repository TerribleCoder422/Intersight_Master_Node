import pandas as pd

# Load the Excel file
excel_file = 'output/Intersight_Foundation.xlsx'
xlsx = pd.ExcelFile(excel_file)

# Check available sheets
print(f"Available sheets in the Excel file: {xlsx.sheet_names}")

# Check Template sheet
if 'Template' in xlsx.sheet_names:
    template_df = pd.read_excel(excel_file, sheet_name='Template')
    print("\nTemplate sheet content:")
    print(f"Number of rows: {len(template_df)}")
    if len(template_df) > 0:
        print(template_df.head())
    else:
        print("Template sheet is empty")

# Check Profiles sheet
if 'Profiles' in xlsx.sheet_names:
    profiles_df = pd.read_excel(excel_file, sheet_name='Profiles')
    print("\nProfiles sheet content:")
    print(f"Number of rows: {len(profiles_df)}")
    if len(profiles_df) > 0:
        print(profiles_df.head())
    else:
        print("Profiles sheet is empty")
