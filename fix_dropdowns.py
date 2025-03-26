#!/usr/bin/env python3
"""
Script to fix the get_intersight_info function in create_intersight_foundation.py
"""

# Replace the code in get_intersight_info function in create_intersight_foundation.py
# with the code below (for the Policies and Template sheet sections):

'''
        # Policies sheet dropdown
        if 'Policies' in workbook.sheetnames:
            policies_sheet = workbook['Policies']
            
            # Update or add policy types dropdown
            policy_types = ['vNIC', 'BIOS', 'BOOT', 'QoS', 'Storage']
            policy_type_updated = False
            
            # Check for existing validations
            for dv in list(policies_sheet.data_validations.dataValidation):
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('A'):
                        # Update existing policy type dropdown
                        dv.formula1 = f'"{",".join(policy_types)}"'
                        policy_type_updated = True
                        break
            
            # Add policy types dropdown if it doesn't exist
            if not policy_type_updated:
                policy_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(policy_types)}"',
                    allow_blank=True
                )
                policy_validation.add('A2:A1000')  # Apply to Policy Types column
                policies_sheet.add_data_validation(policy_validation)
            
            # Update or add organization dropdown
            org_updated = False
            for dv in list(policies_sheet.data_validations.dataValidation):
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('D'):
                        # Update existing organization dropdown
                        dv.formula1 = f'"{",".join(org_names)}"'
                        org_updated = True
                        break
            
            # Add organization dropdown if it doesn't exist
            if not org_updated:
                org_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(org_names)}"',
                    allow_blank=True
                )
                org_validation.add('D2:D1000')  # Apply to Organizations columns
                policies_sheet.add_data_validation(org_validation)
                
            print("Added/Updated dropdowns for Policy Types and Organizations in Policies sheet")
            
        # Template sheet dropdowns
        if 'Template' in workbook.sheetnames:
            template_sheet = workbook['Template']
            
            # Update or add platform types dropdown
            platform_types = ['FIAttached', 'Standalone']
            platform_updated = False
            
            # Check for existing validations
            for dv in list(template_sheet.data_validations.dataValidation):
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('D'):
                        # Update existing platform dropdown
                        dv.formula1 = f'"{",".join(platform_types)}"'
                        platform_updated = True
                        break
            
            # Add platform types dropdown if it doesn't exist
            if not platform_updated:
                platform_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(platform_types)}"',
                    allow_blank=True
                )
                platform_validation.add('D2:D1000')  # Apply to Platform Types column
                template_sheet.add_data_validation(platform_validation)
                print("Added dropdown for Platform Types in Template sheet")

            # Update or add organization dropdown
            org_updated = False
            for dv in list(template_sheet.data_validations.dataValidation):
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('B'):
                        # Update existing organization dropdown
                        dv.formula1 = f'"{",".join(org_names)}"'
                        org_updated = True
                        break
            
            # Add organization dropdown if it doesn't exist
            if not org_updated:
                org_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(org_names)}"',
                    allow_blank=True
                )
                org_validation.add('B2:B1000')  # Apply to Organizations column
                template_sheet.add_data_validation(org_validation)
                
            print("Added/Updated dropdowns for Platform Types and Organizations in Template sheet")
'''

# Also make sure to update the Profiles sheet section too:

'''
            # Update or add organization dropdown
            org_updated = False
            for dv in list(profiles_sheet.data_validations.dataValidation):
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('C'):
                        # Update existing organization dropdown
                        dv.formula1 = f'"{",".join(org_names)}"'
                        org_updated = True
                        break
            
            # Add organization dropdown if it doesn't exist
            if not org_updated:
                org_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(org_names)}"',
                    allow_blank=True
                )
                org_validation.add('C2:C1000')  # Apply to Organization column range
                profiles_sheet.add_data_validation(org_validation)
'''

print("This is a reference file showing how to fix the dropdowns in create_intersight_foundation.py.")
print("Please modify the get_intersight_info function with these code snippets.")
