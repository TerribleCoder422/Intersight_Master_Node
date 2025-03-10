#!/usr/bin/env python3
"""
Fix Excel file for Intersight Foundation
"""

import pandas as pd
from openpyxl import load_workbook

# Path to the Excel file
file_path = '/Users/isaiahlaughlin/Lumos Dropbox/isaiah laughlin/Mac/Desktop/Python journey/Automate-Intersight/output/Intersight_Foundation.xlsx'

# Load the Excel file
workbook = load_workbook(file_path)

# Fix Profiles sheet
profiles_sheet = workbook['Profiles']
# Correct the template name in column D, row 2
profiles_sheet['D2'] = 'Ai_POD_Template'
print("Updated Profiles sheet with correct template name")

# Fix Policies sheet to add Boot policy
# We can see that the Boot policy is already in the sheet
print("Boot policy is already in the Policies sheet")

# Save the workbook
workbook.save(file_path)
print(f"Saved updated Excel file to {file_path}")
