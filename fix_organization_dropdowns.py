#!/usr/bin/env python3
"""
Script to modify create_intersight_foundation.py to handle organization dropdowns correctly
"""
import os
import re
import sys

def fix_create_intersight_foundation():
    # Path to the file to modify
    file_path = './create_intersight_foundation.py'
    
    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # Read the file content
    with open(file_path, 'r') as file:
        content = file.read()

    # Replace the profiles sheet section (lines 1863-1899)
    profiles_pattern = r"""        # Set up Profiles sheet dropdowns
        if 'Profiles' in workbook.sheetnames:
            profiles_sheet = workbook\['Profiles'\]
            
            # Remove all existing data validations
            profiles_sheet.data_validations.dataValidation = \[\]
            
            # Add server dropdown
            server_options = \[f"{server.name} \| SN: {server.serial}" for server in servers.results\]
            server_formula = '"' \+ ','.join\(server_options\) \+ '"'
            server_validation = DataValidation\(
                type='list',
                formula1=server_formula,
                allow_blank=True
            \)
            server_validation.add\('E2:E1000'\)  # Apply to Server column
            profiles_sheet.add_data_validation\(server_validation\)
            
            # Add deploy dropdown
            deploy_validation = DataValidation\(
                type='list',
                formula1='"Yes,No"',
                allow_blank=True
            \)
            deploy_validation.add\('G2:G1000'\)  # Apply to Deploy column
            profiles_sheet.add_data_validation\(deploy_validation\)
            
            # Add organization dropdown
            org_validation = DataValidation\(
                type='list',
                formula1=f'"{",".join\(org_names\)}"',
                allow_blank=True
            \)
            org_validation.add\('C2:C1000'\)  # Apply to Organization column range
            profiles_sheet.add_data_validation\(org_validation\)
            
            print\("Added dropdowns for Server, Deploy, and Organization columns"\)"""

    profiles_replacement = """        # Set up Profiles sheet dropdowns
        if 'Profiles' in workbook.sheetnames:
            profiles_sheet = workbook['Profiles']
            
            # Store existing dropdowns
            existing_dv = []
            org_dv_found = False
            
            for dv in list(profiles_sheet.data_validations.dataValidation):
                # Check if this is an organization dropdown
                is_org_dv = False
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('C'):
                        # This is the org dropdown, update it
                        dv.formula1 = f'"{",".join(org_names)}"'
                        org_dv_found = True
                        is_org_dv = True
                        break
                
                # Keep all non-org dropdowns
                if not is_org_dv:
                    existing_dv.append(dv)
            
            # Clear and re-add all validations
            profiles_sheet.data_validations.dataValidation = []
            
            # Re-add existing non-org validations
            for dv in existing_dv:
                profiles_sheet.add_data_validation(dv)
            
            # Add server dropdown if not found
            if not any(any(str(cell).startswith('E') for cell in dv.sqref) for dv in existing_dv):
                server_options = [f"{server.name} | SN: {server.serial}" for server in servers.results]
                server_formula = '"' + ','.join(server_options) + '"'
                server_validation = DataValidation(
                    type='list',
                    formula1=server_formula,
                    allow_blank=True
                )
                server_validation.add('E2:E1000')  # Apply to Server column
                profiles_sheet.add_data_validation(server_validation)
            
            # Add deploy dropdown if not found
            if not any(any(str(cell).startswith('G') for cell in dv.sqref) for dv in existing_dv):
                deploy_validation = DataValidation(
                    type='list',
                    formula1='"Yes,No"',
                    allow_blank=True
                )
                deploy_validation.add('G2:G1000')  # Apply to Deploy column
                profiles_sheet.add_data_validation(deploy_validation)
            
            # Add organization dropdown if not found
            if not org_dv_found:
                org_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(org_names)}"',
                    allow_blank=True
                )
                org_validation.add('C2:C1000')  # Apply to Organization column range
                profiles_sheet.add_data_validation(org_validation)
            
            print("Added/Updated dropdowns for Server, Deploy, and Organization columns")"""

    # Replace the Policies sheet section (lines 1916-1933)
    policies_pattern = r"""        # Policies sheet dropdown
        if 'Policies' in workbook.sheetnames:
            policies_sheet = workbook\['Policies'\]
            policy_types = \['vNIC', 'BIOS', 'BOOT', 'QoS', 'Storage'\]
            policy_validation = DataValidation\(
                type='list',
                formula1=f'"{",".join\(policy_types\)}"',
                allow_blank=True
            \)
            policy_validation.add\('A2:A1000'\)  # Apply to Policy Types column
            policies_sheet.add_data_validation\(policy_validation\)
            org_validation = DataValidation\(
                type='list',
                formula1=f'"{",".join\(org_names\)}"',
                allow_blank=True
            \)
            org_validation.add\('D2:D1000'\)  # Apply to Organizations columns
            policies_sheet.add_data_validation\(org_validation\)
            print\("Added dropdowns for Policy Types and Organizations in Policies sheet"\)"""

    policies_replacement = """        # Policies sheet dropdown
        if 'Policies' in workbook.sheetnames:
            policies_sheet = workbook['Policies']
            
            # Store existing dropdowns
            existing_dv = []
            org_dv_found = False
            
            for dv in list(policies_sheet.data_validations.dataValidation):
                # Check if this is an organization dropdown
                is_org_dv = False
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('D'):
                        # This is the org dropdown, update it
                        dv.formula1 = f'"{",".join(org_names)}"'
                        org_dv_found = True
                        is_org_dv = True
                        break
                
                # Keep all non-org dropdowns
                if not is_org_dv:
                    existing_dv.append(dv)
            
            # Clear and re-add all validations
            policies_sheet.data_validations.dataValidation = []
            
            # Re-add existing non-org validations
            for dv in existing_dv:
                policies_sheet.add_data_validation(dv)
            
            # Add policy types dropdown if not found
            if not any(any(str(cell).startswith('A') for cell in dv.sqref) for dv in existing_dv):
                policy_types = ['vNIC', 'BIOS', 'BOOT', 'QoS', 'Storage']
                policy_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(policy_types)}"',
                    allow_blank=True
                )
                policy_validation.add('A2:A1000')  # Apply to Policy Types column
                policies_sheet.add_data_validation(policy_validation)
            
            # Add organization dropdown if not found
            if not org_dv_found:
                org_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(org_names)}"',
                    allow_blank=True
                )
                org_validation.add('D2:D1000')  # Apply to Organizations columns
                policies_sheet.add_data_validation(org_validation)
            
            print("Added/Updated dropdowns for Policy Types and Organizations in Policies sheet")"""

    # Replace the Template sheet section (lines 1935-1956)
    template_pattern = r"""        # Template sheet dropdowns
        if 'Template' in workbook.sheetnames:
            template_sheet = workbook\['Template'\]
            platform_types = \['FIAttached', 'Standalone'\]
            platform_validation = DataValidation\(
                type='list',
                formula1=f'"{",".join\(platform_types\)}"',
                allow_blank=True
            \)
            platform_validation.add\('D2:D1000'\)  # Apply to Platform Types column
            template_sheet.add_data_validation\(platform_validation\)
            print\("Added dropdown for Platform Types in Template sheet"\)

            # Add organization dropdown to column B
            org_validation = DataValidation\(
                type='list',
                formula1=f'"{",".join\(org_names\)}"',
                allow_blank=True
            \)
            org_validation.add\('B2:B1000'\)  # Apply to Organizations column
            template_sheet.add_data_validation\(org_validation\)
            print\("Added dropdowns for Platform Types and Organizations in Template sheet"\)"""

    template_replacement = """        # Template sheet dropdowns
        if 'Template' in workbook.sheetnames:
            template_sheet = workbook['Template']
            
            # Store existing dropdowns
            existing_dv = []
            org_dv_found = False
            platform_dv_found = False
            
            for dv in list(template_sheet.data_validations.dataValidation):
                is_special_dv = False
                
                # Check if organization dropdown
                for cell_range in dv.sqref:
                    if str(cell_range).startswith('B'):
                        # This is the org dropdown, update it
                        dv.formula1 = f'"{",".join(org_names)}"'
                        org_dv_found = True
                        is_special_dv = True
                        break
                
                # Check if platform dropdown
                if not is_special_dv:
                    for cell_range in dv.sqref:
                        if str(cell_range).startswith('D'):
                            # This is the platform dropdown, note it exists
                            platform_dv_found = True
                            is_special_dv = False  # Keep this one
                            break
                
                # Keep all non-org dropdowns
                if not is_special_dv:
                    existing_dv.append(dv)
            
            # Clear and re-add all validations
            template_sheet.data_validations.dataValidation = []
            
            # Re-add existing non-org validations
            for dv in existing_dv:
                template_sheet.add_data_validation(dv)
            
            # Add platform types dropdown if not found
            if not platform_dv_found:
                platform_types = ['FIAttached', 'Standalone']
                platform_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(platform_types)}"',
                    allow_blank=True
                )
                platform_validation.add('D2:D1000')  # Apply to Platform Types column
                template_sheet.add_data_validation(platform_validation)
                print("Added dropdown for Platform Types in Template sheet")
            
            # Add organization dropdown if not found
            if not org_dv_found:
                org_validation = DataValidation(
                    type='list',
                    formula1=f'"{",".join(org_names)}"',
                    allow_blank=True
                )
                org_validation.add('B2:B1000')  # Apply to Organizations column
                template_sheet.add_data_validation(org_validation)
            
            print("Added/Updated dropdowns for Platform Types and Organizations in Template sheet")"""

    # Replace all the sections in the content
    content = re.sub(profiles_pattern, profiles_replacement, content)
    content = re.sub(policies_pattern, policies_replacement, content)
    content = re.sub(template_pattern, template_replacement, content)

    # Write back the modified content
    with open(file_path, 'w') as file:
        file.write(content)
    
    print(f"Successfully modified {file_path}")
    print("Organization dropdowns will now be properly updated with Intersight data")

if __name__ == "__main__":
    fix_create_intersight_foundation()
