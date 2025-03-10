#!/usr/bin/env python3
"""
Create a structured Excel template for Cisco Intersight Pools and Policies.
This template follows Intersight's configuration structure and includes dropdowns.
"""

import pandas as pd
import os
import json
import intersight
from intersight.api_client import ApiClient
from intersight.configuration import Configuration
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import uuid
from intersight.api import (
    bios_api,
    boot_api,
    compute_api,
    fabric_api,
    firmware_api,
    ippool_api,
    macpool_api,
    organization_api,
    resource_api,
    server_api,
    storage_api,
    uuidpool_api,
    vnic_api
)
from intersight.model.vnic_lan_connectivity_policy import VnicLanConnectivityPolicy
from intersight.model.vnic_eth_if import VnicEthIf
from intersight.model.vnic_eth_adapter_policy import VnicEthAdapterPolicy
from intersight.model.vnic_eth_qos_policy import VnicEthQosPolicy
from intersight.model.fabric_eth_network_group_policy import FabricEthNetworkGroupPolicy
from intersight.model.server_profile_template import ServerProfileTemplate
import time
import argparse
import sys
import re
from copy import copy

def get_api_client():
    """
    Create an Intersight API client using the API key file
    """
    try:
        # Get API key details from environment variables
        api_key_id = os.getenv('INTERSIGHT_API_KEY_ID')
        api_key_file = os.getenv('INTERSIGHT_PRIVATE_KEY_FILE', './SecretKey.txt')
        
        if not api_key_id or not os.path.exists(api_key_file):
            print("Error: API key configuration not found")
            return None
            
        # Create configuration
        config = Configuration(
            host = os.getenv('INTERSIGHT_BASE_URL', 'https://intersight.com'),
            signing_info = intersight.signing.HttpSigningConfiguration(
                key_id = api_key_id,
                private_key_path = api_key_file,
                signing_scheme = intersight.signing.SCHEME_HS2019,
                signing_algorithm = intersight.signing.ALGORITHM_ECDSA_MODE_FIPS_186_3,
                hash_algorithm = intersight.signing.HASH_SHA256,
                signed_headers = [
                    intersight.signing.HEADER_REQUEST_TARGET,
                    intersight.signing.HEADER_HOST,
                    intersight.signing.HEADER_DATE,
                    intersight.signing.HEADER_DIGEST,
                ]
            )
        )
        
        # Create API client
        api_client = ApiClient(configuration=config)
        return api_client
        
    except Exception as e:
        print(f"Error creating API client: {str(e)}")
        return None

def get_organizations(api_client):
    """
    Get list of organizations from Intersight
    """
    try:
        # Create API instance for organizations
        api_instance = organization_api.OrganizationApi(api_client)
        
        # Get list of organizations
        orgs = api_instance.get_organization_organization_list()
        
        # Extract organization names
        org_names = [org['Name'] for org in orgs.results]
        return org_names
    except Exception as e:
        print(f"Error getting organizations: {str(e)}")
        return ['default']  # Return default as fallback

def get_organizations(api_client):
    """
    Get list of organizations from Intersight
    """
    if not api_client:
        print("Debug: No API client available, defaulting to 'default' organization")
        return ["default"]
        
    try:
        # Import here to avoid circular imports
        from intersight.api import organization_api
        org_api = organization_api.OrganizationApi(api_client)
        print("Debug: Successfully created organization API client")
        
        orgs = org_api.get_organization_organization_list()
        print(f"Debug: Found organizations: {[org.name for org in orgs.results]}")
        
        return [org.name for org in orgs.results] if orgs.results else ["default"]
    except Exception as e:
        print(f"Debug: Error fetching organizations: {str(e)}")
        return ["default"]

def create_mac_pool(api_client, pool_data):
    """
    Create a MAC Pool in Intersight
    """
    from intersight.api import macpool_api
    from intersight.model.macpool_pool import MacpoolPool
    from intersight.model.macpool_block import MacpoolBlock
    from intersight.model.mo_mo_ref import MoMoRef
    
    try:
        # Get organization MOID
        org_moid = get_org_moid(api_client, "Gruve")
        if not org_moid:
            print("Error: Gruve organization not found")
            return False

        # Create organization reference
        org_ref = MoMoRef(
            class_id="mo.MoRef",
            object_type="organization.Organization",
            moid=org_moid
        )

        # Create MAC pool block
        block = MacpoolBlock(
            class_id="macpool.Block",
            object_type="macpool.Block",
            _from=pool_data['Start Address'],
            size=int(pool_data['Size'])
        )
        
        # Create MAC pool
        pool = MacpoolPool(
            class_id="macpool.Pool",
            object_type="macpool.Pool",
            name=pool_data['Pool Name'],
            description=pool_data['Description'] if pd.notna(pool_data['Description']) else "",
            organization=org_ref,
            assignment_order="sequential",
            MacBlocks=[block]
        )
        
        # Create API instance
        api_instance = macpool_api.MacpoolApi(api_client)
        result = api_instance.create_macpool_pool(macpool_pool=pool)
        print(f"Successfully created MAC Pool: {result.name}")
        return True
        
    except Exception as e:
        print(f"Error creating MAC Pool: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def create_uuid_pool(api_client, pool_data):
    """
    Create a UUID Pool in Intersight
    """
    from intersight.api import uuidpool_api
    from intersight.model.uuidpool_pool import UuidpoolPool
    from intersight.model.uuidpool_uuid_block import UuidpoolUuidBlock
    from intersight.model.mo_mo_ref import MoMoRef
    
    try:
        # Get organization MOID
        org_moid = get_org_moid(api_client, "Gruve")
        if not org_moid:
            print("Error: Gruve organization not found")
            return False

        # Create organization reference
        org_ref = MoMoRef(
            class_id="mo.MoRef",
            object_type="organization.Organization",
            moid=org_moid
        )

        # Create UUID pool block
        block = UuidpoolUuidBlock(
            class_id="uuidpool.UuidBlock",
            object_type="uuidpool.UuidBlock",
            _from=pool_data['Start Address'],
            size=int(pool_data['Size'])
        )
        
        # Create UUID pool
        pool = UuidpoolPool(
            class_id="uuidpool.Pool",
            object_type="uuidpool.Pool",
            name=pool_data['Pool Name'],
            description=pool_data['Description'] if pd.notna(pool_data['Description']) else "",
            organization=org_ref,
            assignment_order="sequential",
            prefix="000025B5-0000-0000",  # Standard prefix for UUIDs
            UuidSuffixBlocks=[block]
        )
        
        # Create API instance
        api_instance = uuidpool_api.UuidpoolApi(api_client)
        result = api_instance.create_uuidpool_pool(uuidpool_pool=pool)
        print(f"Successfully created UUID Pool: {result.name}")
        return True
        
    except Exception as e:
        print(f"Error creating UUID Pool: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def format_uuid_suffix(uuid_str):
    """Format a UUID suffix to match Intersight's expected pattern: XXXX-XXXXXXXXXXXX"""
    # Remove any non-hex characters and pad to 16 characters
    clean_uuid = ''.join(c for c in uuid_str if c.isalnum()).zfill(16)
    return f"{clean_uuid[:4]}-{clean_uuid[4:]}"

def pool_exists(api_client, pool_type, pool_name):
    """
    Check if a pool already exists in Intersight
    """
    try:
        # Get organization MOID
        org_moid = get_org_moid(api_client, "Gruve")
        if not org_moid:
            return False

        # Create API instance based on pool type
        if pool_type == 'MAC Pool':
            from intersight.api import macpool_api
            api_instance = macpool_api.MacpoolApi(api_client)
            filter_str = f"Name eq '{pool_name}' and Organization.Moid eq '{org_moid}'"
            api_response = api_instance.get_macpool_pool_list(filter=filter_str)
        elif pool_type == 'UUID Pool':
            from intersight.api import uuidpool_api
            api_instance = uuidpool_api.UuidpoolApi(api_client)
            filter_str = f"Name eq '{pool_name}' and Organization.Moid eq '{org_moid}'"
            api_response = api_instance.get_uuidpool_pool_list(filter=filter_str)
        else:
            print(f"Unsupported pool type: {pool_type}")
            return False

        # Check if any pools were found
        return len(api_response.results) > 0

    except Exception as e:
        print(f"Error checking if pool exists: {str(e)}")
        return False

def create_pool(api_client, pool_data):
    """
    Create a pool in Intersight based on pool type
    """
    try:
        pool_type = pool_data['Pool Type']
        pool_name = pool_data['Pool Name']
        
        # Check if pool already exists
        if pool_exists(api_client, pool_type, pool_name):
            print(f"\nPool {pool_name} already exists, skipping creation")
            return True
            
        print(f"\nCreating {pool_type}: {pool_name}")
        print(f"Description: {pool_data['Description'] if pd.notna(pool_data['Description']) else 'None'}")
        print(f"Start Address: {pool_data['Start Address']}")
        print(f"Size: {pool_data['Size']}")
        
        if pool_type == 'MAC Pool':
            return create_mac_pool(api_client, pool_data)
        elif pool_type == 'UUID Pool':
            return create_uuid_pool(api_client, pool_data)
        else:
            print(f"Unsupported pool type: {pool_type}")
            return False
            
    except Exception as e:
        print(f"Error creating pool: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def get_mac_pool_moid(api_client, pool_name, org_moid):
    """
    Get the MOID of a MAC Pool by name and organization MOID
    """
    from intersight.api import macpool_api
    
    api_instance = macpool_api.MacpoolApi(api_client)
    pools = api_instance.get_macpool_pool_list()
    for pool in pools.results:
        if pool.name == pool_name and pool.organization.moid == org_moid:
            return pool.moid
    return None

def get_pool_moid(api_client, pool_name):
    """
    Get the MOID of a pool by name
    """
    from intersight.api import macpool_api
    
    api_instance = macpool_api.MacpoolApi(api_client)
    pools = api_instance.get_macpool_pool_list(filter=f"Name eq '{pool_name}'").results
    
    if pools:
        return pools[0].moid
    else:
        raise Exception(f"Pool '{pool_name}' not found")

def get_policy_moid(api_client, policy_type, policy_name):
    """Get the MOID of a policy by name"""
    try:
        if policy_type == "bios.Policy":
            api_instance = bios_api.BiosApi(api_client)
            policies = api_instance.get_bios_policy_list()
        elif policy_type == "vnic.LanConnectivityPolicy":
            api_instance = vnic_api.VnicApi(api_client)
            policies = api_instance.get_vnic_lan_connectivity_policy_list()
        elif policy_type == "vnic.EthQosPolicy":
            api_instance = vnic_api.VnicApi(api_client)
            policies = api_instance.get_vnic_eth_qos_policy_list()
        elif policy_type == "storage.StoragePolicy":
            api_instance = storage_api.StorageApi(api_client)
            policies = api_instance.get_storage_storage_policy_list()
        elif policy_type == "macpool.Pool":
            api_instance = macpool_api.MacpoolApi(api_client)
            policies = api_instance.get_macpool_pool_list()
        elif policy_type == "boot.PrecisionPolicy":
            api_instance = boot_api.BootApi(api_client)
            policies = api_instance.get_boot_precision_policy_list()
        elif policy_type == "storage.StoragePolicies":
            api_instance = storage_api.StorageApi(api_client)
            policies = api_instance.get_storage_storage_policy_list()
        else:
            raise Exception(f"Unsupported policy type: {policy_type}")
        
        # Find the policy by name
        for policy in policies.results:
            if policy.name == policy_name:
                return policy.moid
                
        print(f"Policy {policy_name} not found")
        return None
        
    except Exception as e:
        print(f"Error getting policy MOID: {str(e)}")
        return None

def policy_exists(api_client, policy_type, policy_name):
    """
    Check if a policy already exists in Intersight
    """
    try:
        # Get organization MOID
        org_moid = get_org_moid(api_client)
        
        # Create API instance based on policy type
        if policy_type == "bios.Policy":
            api_instance = bios_api.BiosApi(api_client)
            response = api_instance.get_bios_policy_list(filter=f"Name eq '{policy_name}'")
        elif policy_type == "vnic.EthQosPolicy":
            api_instance = vnic_api.VnicApi(api_client)
            response = api_instance.get_vnic_eth_qos_policy_list(filter=f"Name eq '{policy_name}'")
        elif policy_type == "vnic.EthAdapterPolicy":
            api_instance = vnic_api.VnicApi(api_client)
            response = api_instance.get_vnic_eth_adapter_policy_list(filter=f"Name eq '{policy_name}'")
        elif policy_type == "fabric.EthNetworkGroupPolicy":
            api_instance = fabric_api.FabricApi(api_client)
            response = api_instance.get_fabric_eth_network_group_policy_list(filter=f"Name eq '{policy_name}'")
        elif policy_type == "vnic.LanConnectivityPolicy":
            api_instance = vnic_api.VnicApi(api_client)
            response = api_instance.get_vnic_lan_connectivity_policy_list(filter=f"Name eq '{policy_name}'")
        elif policy_type == "boot.PrecisionPolicy":
            api_instance = boot_api.BootApi(api_client)
            response = api_instance.get_boot_precision_policy_list(filter=f"Name eq '{policy_name}'")
        elif policy_type == "storage.StoragePolicy":
            api_instance = storage_api.StorageApi(api_client)
            response = api_instance.get_storage_storage_policy_list(filter=f"Name eq '{policy_name}'")
        else:
            return False

        return len(response.results) > 0

    except Exception as e:
        print(f"Error checking if policy exists: {str(e)}")
        return False

def check_vnic_exists(api_client, vnic_name, lan_connectivity_moid):
    """
    Check if a vNIC already exists in the LAN Connectivity Policy
    """
    try:
        vnic_instance = vnic_api.VnicApi(api_client)
        vnic_list = vnic_instance.get_vnic_eth_if_list(filter=f"Name eq '{vnic_name}'")
        for vnic in vnic_list.results:
            if hasattr(vnic, 'lan_connectivity_policy') and vnic.lan_connectivity_policy.moid == lan_connectivity_moid:
                return True
        return False
    except Exception as e:
        print(f"Error checking vNIC existence: {str(e)}")
        return False

def move_sheet_after(workbook, sheet_to_move, target_sheet):
    """Move a worksheet to be right after another worksheet"""
    if target_sheet not in workbook.sheetnames or sheet_to_move not in workbook.sheetnames:
        return False
        
    # Get the indices
    target_index = workbook.sheetnames.index(target_sheet)
    current_index = workbook.sheetnames.index(sheet_to_move)
    
    # If sheet is already after target, do nothing
    if current_index == target_index + 1:
        return True
        
    # Remove and insert at new position
    sheet = workbook[sheet_to_move]
    workbook.remove(sheet)
    workbook.create_sheet(sheet_to_move, target_index + 1)
    
    # Copy the removed sheet to the new position
    new_sheet = workbook[sheet_to_move]
    for row in sheet.iter_rows():
        for cell in row:
            new_sheet[cell.coordinate].value = cell.value
            if cell.has_style:
                new_sheet[cell.coordinate].font = copy(cell.font)
                new_sheet[cell.coordinate].border = copy(cell.border)
                new_sheet[cell.coordinate].fill = copy(cell.fill)
                new_sheet[cell.coordinate].number_format = copy(cell.number_format)
                new_sheet[cell.coordinate].protection = copy(cell.protection)
                new_sheet[cell.coordinate].alignment = copy(cell.alignment)
    
    # Copy sheet properties
    new_sheet.sheet_format = copy(sheet.sheet_format)
    new_sheet.sheet_properties = copy(sheet.sheet_properties)
    new_sheet.merged_cells = copy(sheet.merged_cells)
    new_sheet.page_margins = copy(sheet.page_margins)
    new_sheet.page_setup = copy(sheet.page_setup)
    
    # Copy column dimensions
    for key, value in sheet.column_dimensions.items():
        new_sheet.column_dimensions[key] = copy(value)
    
    # Copy row dimensions
    for key, value in sheet.row_dimensions.items():
        new_sheet.row_dimensions[key] = copy(value)

def process_foundation_template(excel_file):
    """
    Read the Excel template and create pools and policies in Intersight
    """
    try:
        # Read Excel file
        df = pd.read_excel(excel_file, sheet_name=None)
        
        # Get API client
        api_client = get_api_client()
        if not api_client:
            print("Error: Failed to get API client")
            return False
            
        # Process Pools sheet first
        if 'Pools' in df:
            pools_df = df['Pools']
            # Rename columns to remove asterisks
            pools_df.columns = pools_df.columns.str.replace('*', '')
            
            # Track pool creation success
            pools_created = True
            failed_pools = []
            
            # Create or verify each pool
            for _, row in pools_df.iterrows():
                pool_name = row['Pool Name']
                pool_type = row['Pool Type']
                
                # Check if pool exists
                if pool_exists(api_client, pool_type, pool_name):
                    print(f"\nPool {pool_name} already exists, skipping creation")
                    continue
                    
                # Try to create the pool
                if not create_pool(api_client, row):
                    pools_created = False
                    failed_pools.append(pool_name)
            
            # If any pools failed to create, stop here
            if not pools_created:
                print("\nError: Failed to create the following pools:")
                for pool in failed_pools:
                    print(f"  - {pool}")
                print("\nAborting further processing until pool creation issues are resolved.")
                return False
                
            print("\nAll pools created or verified successfully.")
                
        # Only proceed with policies if pools were successful
        if 'Policies' in df:
            policies_df = df['Policies']
            # Rename columns to remove asterisks
            policies_df.columns = policies_df.columns.str.replace('*', '')
            
            # Create policies in order: BIOS, QoS, vNIC, Boot, Storage
            policy_order = ['BIOS', 'QoS', 'vNIC', 'Boot', 'Storage']
            
            # Track policy creation success
            policies_created = True
            failed_policies = []
            
            for policy_type in policy_order:
                print(f"\nProcessing {policy_type} policies...")
                policy_rows = policies_df[policies_df['Policy Type'] == policy_type]
                
                for _, row in policy_rows.iterrows():
                    policy_name = row['Policy Name']
                    
                    # Check if policy exists
                    if policy_exists(api_client, get_policy_class_id(policy_type), policy_name):
                        print(f"Policy {policy_name} already exists, skipping creation")
                        continue
                        
                    # Try to create the policy
                    if not create_policy(api_client, row):
                        policies_created = False
                        failed_policies.append(f"{policy_type}: {policy_name}")
                        break  # Stop processing this policy type if one fails
                
                # If any policies failed, stop processing
                if not policies_created:
                    print("\nError: Failed to create the following policies:")
                    for policy in failed_policies:
                        print(f"  - {policy}")
                    print("\nAborting further processing until policy creation issues are resolved.")
                    return False
                    
                print(f"All {policy_type} policies created or verified successfully.")
                
                # Add a small delay between policy types
                if policy_type != policy_order[-1]:
                    print(f"Waiting for {policy_type} policies to be fully created...")
                    time.sleep(5)
            
            print("\nAll policies created or verified successfully.")
            
        print("\nCompleted processing the Foundation template")
        return True
        
    except Exception as e:
        print(f"\nError processing Foundation template: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def create_and_push_configuration(api_client, excel_file):
    """
    Read the Excel template and create pools and policies in Intersight
    """
    try:
        # Read Excel file
        df = pd.read_excel(excel_file, sheet_name=None)
        
        # Process Pools sheet
        if 'Pools' in df:
            pools_df = df['Pools']
            for _, row in pools_df.iterrows():
                create_pool(api_client, row)
                
        # Process Policies sheet in specific order
        if 'Policies' in df:
            policies_df = df['Policies']
            
            # Create policies in order: BIOS, QoS, vNIC, Boot, Storage
            policy_order = ['BIOS', 'QoS', 'vNIC', 'Boot', 'Storage']
            
            for policy_type in policy_order:
                policy_rows = policies_df[policies_df['Policy Type'] == policy_type]
                for _, row in policy_rows.iterrows():
                    if policy_exists(api_client, get_policy_class_id(policy_type), row['Name']):
                        print(f"Policy {row['Name']} already exists, skipping creation")
                    else:
                        create_policy(api_client, row)
                    
        print("Completed processing the Foundation template")
        return True
        
    except Exception as e:
        print(f"Error processing Foundation template: {str(e)}")
        return False

def get_policy_class_id(policy_type):
    """Get the class ID for a policy type"""
    policy_map = {
        'BIOS': 'bios.Policy',
        'QoS': 'vnic.EthQosPolicy',
        'vNIC': 'vnic.LanConnectivityPolicy',
        'Storage': 'storage.StoragePolicy',
        'Boot': 'boot.PrecisionPolicy'
    }
    return policy_map.get(policy_type)

def add_template_sheet(excel_file, api_client):
    """Add or update the Template sheet with dropdowns"""
    try:
        # Load workbook
        workbook = load_workbook(excel_file)
        
        # Store existing values if sheet exists
        existing_values = []
        if 'Template' in workbook.sheetnames:
            template_sheet = workbook['Template']
            existing_values = list(template_sheet.iter_rows(min_row=2, values_only=True))
            workbook.remove(template_sheet)
        
        # Create new sheet
        template_sheet = workbook.create_sheet(title='Template')
        
        # Add headers
        headers = [
            "Template Name*", 
            "Organization*", 
            "Description",
            "Target Platform*",
            "BIOS Policy*",
            "Boot Policy*",
            "LAN Connectivity Policy*",
            "Storage Policy*"
        ]
        
        # Define styles
        header_fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True)
        required_font = Font(color='FF0000', bold=True)
        
        # Add headers with styling
        for col, header in enumerate(headers, 1):
            cell = template_sheet.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            if '*' in header:
                cell.font = required_font
            cell.alignment = Alignment(horizontal='center')
        
        # Example template data
        template_example = [
            "Ai_POD_Template",
            "Gruve",
            "Server template for AI POD workloads",
            "FIAttached",
            "Ai_POD-BIOS",
            "Ai_POD-BOOT",
            "Ai_POD-vNIC-A",
            "Ai_POD-Storage"
        ]
        
        # Add example data
        for col, value in enumerate(template_example, 1):
            template_sheet.cell(row=2, column=col, value=value)
        
        # Create named range for Target Platform options
        platform_options = ['FIAttached', 'Standalone']
        platform_range_name = 'PlatformOptions'
        platform_range = f'"{",".join(platform_options)}"'
        
        # Add data validation for Target Platform using the named range
        platform_validation = DataValidation(
            type='list',
            formula1=platform_range,
            allow_blank=False,
            showDropDown=True
        )
        template_sheet.add_data_validation(platform_validation)
        platform_validation.add('D2:D1000')  # Apply to Target Platform column
        
        # Get available organizations
        org_api = organization_api.OrganizationApi(api_client)
        orgs = org_api.get_organization_organization_list()
        org_names = [org.name for org in orgs.results]
        
        # Add data validation for Organization
        org_validation = DataValidation(
            type='list',
            formula1=f'"{",".join(org_names)}"',
            allow_blank=False,
            showDropDown=True
        )
        org_validation.add('B2:B1000')
        
        # Adjust column widths
        min_widths = {
            'A': 20,  # Template Name
            'B': 15,  # Organization
            'C': 30,  # Description
            'D': 15,  # Target Platform
            'E': 20,  # BIOS Policy
            'F': 20,  # Boot Policy
            'G': 25,  # LAN Connectivity Policy
            'H': 20   # Storage Policy
        }
        
        for column in template_sheet.columns:
            col_letter = get_column_letter(column[0].column)
            max_length = max((len(str(cell.value or "")) for cell in column))
            min_width = min_widths.get(col_letter, 15)
            adjusted_width = max(max_length + 2, min_width)
            template_sheet.column_dimensions[col_letter].width = adjusted_width
        
        # Save the workbook
        workbook.save(excel_file)
        print("\nTemplate sheet updated with:")
        print("- Target Platform dropdown (FIAttached/Standalone)")
        print(f"- {len(org_names)} organizations in dropdown")
        return True
        
    except Exception as e:
        print(f"Error adding template sheet: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def create_server_template_from_excel(api_client, excel_file):
    """Create a server template in Intersight from Excel configuration"""
    try:
        # Read the template sheet
        template_df = pd.read_excel(excel_file, sheet_name='Template')
        print("Template DataFrame:")
        print(template_df)
        
        # Get the first row of data (we only support one template for now)
        if len(template_df) == 0:
            print("No template data found in Excel file")
            return False
            
        template_data = template_df.iloc[0].to_dict()
        print("\nTemplate Data:")
        print(template_data)
        
        # Get organization MOID
        org_moid = get_org_moid(api_client)
        
        # Create Server Profile Template API instance
        api_instance = server_api.ServerApi(api_client)
        
        # Get policy MOIDs
        bios_policy_moid = get_policy_moid(api_client, 'bios.Policy', template_data.get('BIOS Policy*', ''))
        boot_policy_moid = get_policy_moid(api_client, 'boot.PrecisionPolicy', template_data.get('Boot Policy*', ''))
        lan_policy_moid = get_policy_moid(api_client, 'vnic.LanConnectivityPolicy', template_data.get('LAN Connectivity Policy*', ''))
        storage_policy_moid = get_policy_moid(api_client, 'storage.StoragePolicy', template_data.get('Storage Policy*', ''))
        
        # Create the template body
        template_body = {
            'Name': template_data.get('Template Name*', ''),
            'Description': template_data.get('Description', ''),
            'Organization': {
                'ObjectType': 'organization.Organization',
                'Moid': org_moid
            },
            'TargetPlatform': template_data.get('Target Platform*', 'FIAttached'),
            'PolicyBucket': []
        }
        
        # Add policies to the template
        if bios_policy_moid:
            template_body['PolicyBucket'].append({
                'ObjectType': 'bios.Policy',
                'Moid': bios_policy_moid
            })
            
        if boot_policy_moid:
            template_body['PolicyBucket'].append({
                'ObjectType': 'boot.PrecisionPolicy',
                'Moid': boot_policy_moid
            })
            
        if lan_policy_moid:
            template_body['PolicyBucket'].append({
                'ObjectType': 'vnic.LanConnectivityPolicy',
                'Moid': lan_policy_moid
            })
            
        if storage_policy_moid:
            template_body['PolicyBucket'].append({
                'ObjectType': 'storage.StoragePolicy',
                'Moid': storage_policy_moid
            })
            
        # Create the template
        print(f"\nCreating server template '{template_data.get('Template Name*', '')}'...")
        print("Template body:")
        print(template_body)
        api_instance.create_server_profile_template(template_body)
        print(f"Successfully created server template: {template_data.get('Template Name*', '')}")
        return True
        
    except Exception as e:
        print(f"Error creating server template: {str(e)}")
        print("Full error details:")
        import traceback
        traceback.print_exc()
        return False

def get_available_servers(api_client):
    """Get list of available servers from Intersight"""
    try:
        # Create API instance for compute servers
        api_instance = compute_api.ComputeApi(api_client)
        
        # Get list of physical servers
        servers = api_instance.get_compute_rack_unit_list()
        
        # Extract server details
        server_list = []
        for server in servers.results:
            # Get organization name if available
            org_name = 'default'
            if hasattr(server, 'organization') and server.organization:
                org_api = organization_api.OrganizationApi(api_client)
                org = org_api.get_organization_organization_by_moid(server.organization.moid)
                org_name = org.name if hasattr(org, 'name') else 'default'
            
            # Build server info
            server_info = {
                'Name': server.name if hasattr(server, 'name') else 'Unknown',
                'Model': server.model if hasattr(server, 'model') else 'Unknown',
                'Serial': server.serial if hasattr(server, 'serial') else 'Unknown',
                'Organization': org_name,
                'PowerState': server.oper_power_state if hasattr(server, 'oper_power_state') else 'Unknown',
                'ConnectionState': server.connection_status if hasattr(server, 'connection_status') else 'Unknown',
                'IP': server.ip_address if hasattr(server, 'ip_address') else 'Unknown',
                'Firmware': server.running_firmware if hasattr(server, 'running_firmware') else 'Unknown',
                'Moid': server.moid if hasattr(server, 'moid') else None
            }
            server_list.append(server_info)
            
        # Sort servers by name
        server_list.sort(key=lambda x: x['Name'])
        
        # Print server details for debugging
        print("\nAvailable Servers:")
        for server in server_list:
            print(f"- {server['Name']} ({server['Model']}) | SN: {server['Serial']} | State: {server['PowerState']} | Connection: {server['ConnectionState']}")
        
        return server_list
    except Exception as e:
        print(f"Error getting servers: {str(e)}")
        return []

def get_available_servers(api_client):
    """Get list of available servers from Intersight"""
    try:
        compute_api_instance = compute_api.ComputeApi(api_client)
        servers = compute_api_instance.get_compute_rack_unit_list()
        server_options = []
        print("\nAvailable Servers:")
        for server in servers.results:
            # Format: Name (Model) | Serial | State
            server_info = f"{server.name} ({server.model}) | SN: {server.serial} | State: {server.oper_state}"
            print(f"- {server_info}")
            server_options.append(server_info)
        return server_options
    except Exception as e:
        print(f"Error getting available servers: {str(e)}")
        return []

def add_profiles_sheet(excel_file, api_client):
    """Add or update the Profiles sheet with dropdowns"""
    try:
        # Load workbook
        workbook = load_workbook(excel_file)
        
        # Store existing values if sheet exists
        existing_values = []
        if 'Profiles' in workbook.sheetnames:
            profiles_sheet = workbook['Profiles']
            existing_values = list(profiles_sheet.iter_rows(min_row=2, values_only=True))
            workbook.remove(profiles_sheet)
        
        # Create new sheet
        profiles_sheet = workbook.create_sheet(title='Profiles')
        
        # Add headers
        headers = ["Profile Name*", "Description", "Organization*", "Template Name*", "Server*", "Description", "Deploy*"]
        for col, header in enumerate(headers, 1):
            cell = profiles_sheet.cell(row=1, column=col, value=header)
            cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
            cell.font = Font(color='FFFFFF', bold=True)
            if '*' in header:
                cell.font = Font(color='FF0000', bold=True)
        
        # Example profile data
        profiles_sheet.append(['AI-Server-01', 'AI POD Host Profile', 'default', 'Ai_POD_Template', '', 'Production AI POD Host', 'No'])
        
        # Get available servers from Intersight
        compute_api_instance = compute_api.ComputeApi(api_client)
        servers = compute_api_instance.get_compute_rack_unit_list()
        
        # Create server list with details
        server_options = []
        for server in servers.results:
            server_info = f"{server.name} ({server.model}) | SN: {server.serial} | State: {server.oper_state}"
            server_options.append(server_info)
        
        # Add server dropdown validation
        server_formula = '"' + ','.join(server_options) + '"'
        server_validation = DataValidation(
            type='list',
            formula1=server_formula,
            allow_blank=True
        )
        server_validation.add('E2:E1000')  # Apply to Server column
        profiles_sheet.add_data_validation(server_validation)
        
        # Add deploy dropdown validation
        deploy_validation = DataValidation(
            type='list',
            formula1='"Yes,No"',
            allow_blank=True
        )
        deploy_validation.add('G2:G1000')  # Apply to Deploy column
        profiles_sheet.add_data_validation(deploy_validation)
        
        # Add organization dropdown
        org_api = organization_api.OrganizationApi(api_client)
        orgs = org_api.get_organization_organization_list()
        org_names = [org.name for org in orgs.results]
        org_validation = DataValidation(
            type='list',
            formula1=f'"{",".join(org_names)}"',
            allow_blank=True
        )
        org_validation.add('C2:C7')  # Apply to Organization column
        
        # Adjust column widths
        min_widths = {
            'A': 20,  # Profile Name
            'B': 30,  # Description
            'C': 15,  # Organization
            'D': 20,  # Template Name
            'E': 60,  # Server Assignment (wider for server details)
            'F': 30,  # Description
            'G': 10   # Deploy
        }
        
        for column in profiles_sheet.columns:
            col_letter = get_column_letter(column[0].column)
            max_length = max((len(str(cell.value or "")) for cell in column))
            min_width = min_widths.get(col_letter, 15)
            adjusted_width = max(max_length + 2, min_width)
            profiles_sheet.column_dimensions[col_letter].width = adjusted_width
        
        return True
        
    except Exception as e:
        print(f"Error adding profiles sheet: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def get_server_templates(api_client):
    """Get list of server profile templates from Intersight"""
    try:
        # Create API instance for server
        api_instance = server_api.ServerApi(api_client)
        
        # Get list of server profile templates
        templates = api_instance.get_server_profile_template_list()
        
        # Extract template details
        template_list = []
        for template in templates.results:
            # Get organization name from organization reference
            org_name = 'default'
            if hasattr(template, 'organization') and template.organization:
                org_api = organization_api.OrganizationApi(api_client)
                org = org_api.get_organization_organization_by_moid(template.organization.moid)
                org_name = org.name if hasattr(org, 'name') else 'default'
            
            template_list.append({
                'Name': template.name if hasattr(template, 'name') else 'Unknown',
                'Description': template.description if hasattr(template, 'description') else '',
                'Organization': org_name,
                'Moid': template.moid if hasattr(template, 'moid') else None
            })
        return template_list
    except Exception as e:
        print(f"Error getting templates: {str(e)}")
        return []

def get_available_templates(api_client):
    """Get list of available server profile templates"""
    try:
        server_api_instance = server_api.ServerApi(api_client)
        templates = server_api_instance.get_server_profile_template_list()
        return templates.results
    except Exception as e:
        print(f"Error getting templates: {str(e)}")
        return []

def create_server_profiles_from_excel(api_client, excel_file):
    """Create server profiles from the Profiles sheet"""
    try:
        # Load workbook
        workbook = load_workbook(excel_file)
        if 'Profiles' not in workbook.sheetnames:
            print("No Profiles sheet found in Excel file")
            return False
        
        worksheet = workbook['Profiles']
        
        # Get all rows except header
        rows = list(worksheet.iter_rows(min_row=2, values_only=True))
        if not rows:
            print("No profile configurations found in Profiles sheet")
            return False
        
        # Create API instances
        server_api_instance = server_api.ServerApi(api_client)
        compute_api_instance = compute_api.ComputeApi(api_client)
        
        profiles_created = 0
        for row in rows:
            if not any(row):  # Skip empty rows
                continue
                
            # Adjusted column mapping
            name_pattern = row[0]
            num_profiles = int(row[1]) if isinstance(row[1], int) else 1
            org_name = row[2]
            template_name = row[3]
            server_info = row[4]
            description = row[5] if len(row) > 5 else ''
            deploy = row[6] if len(row) > 6 else 'No'
            
            print(f"Processing row: {row}")
            
            if not all([name_pattern, org_name, template_name]):
                print(f"Skipping row due to missing required fields: {row}")
                continue
                
            if deploy.lower() != 'yes':
                print(f"Skipping profile creation for {name_pattern} as Deploy is set to No")
                continue
            
            # Get organization
            org_api = organization_api.OrganizationApi(api_client)
            orgs = org_api.get_organization_organization_list(filter=f"Name eq '{org_name}'")
            if not orgs.results:
                print(f"Organization not found: {org_name}")
                continue
            org_moid = orgs.results[0].moid
            
            # Get template
            templates = server_api_instance.get_server_profile_template_list(filter=f"Name eq '{template_name}'")
            if not templates.results:
                print(f"Template not found: {template_name}")
                continue
            template = templates.results[0]
            
            # Get server if specified
            server_moid = None
            if server_info:
                # Extract serial number from server info string
                serial_match = re.search(r'SN: (\w+)', server_info)
                if serial_match:
                    serial = serial_match.group(1)
                    servers = compute_api_instance.get_compute_rack_unit_list(filter=f"Serial eq '{serial}'")
                    if servers.results:
                        server_moid = servers.results[0].moid
                    else:
                        print(f"Server not found with serial: {serial}")
                        continue
            
            # Create profiles
            for i in range(num_profiles):
                profile_name = f"{name_pattern}{i+1}" if num_profiles > 1 else name_pattern
                
                # Create profile from template
                profile_body = {
                    "Name": profile_name,
                    "Description": description,
                    "Organization": {"ObjectType": "organization.Organization", "Moid": org_moid},
                    "SrcTemplate": {"ObjectType": "server.ProfileTemplate", "Moid": template.moid}
                }
                
                if server_moid:
                    profile_body["AssignedServer"] = {"ObjectType": "compute.RackUnit", "Moid": server_moid}
                
                try:
                    print(f"\nCreating profile: {profile_name}")
                    print(f"Profile body: {profile_body}")
                    profile = server_api_instance.create_server_profile(profile_body)
                    print(f"Successfully created profile: {profile_name}")
                    profiles_created += 1
                except Exception as e:
                    print(f"Error creating profile {profile_name}: {str(e)}")
                    continue
        
        print(f"\nCreated {profiles_created} server profiles")
        return True
        
    except Exception as e:
        print(f"Error creating server profiles: {str(e)}")
        print(f"Error details: {str(e.__class__.__name__)}")
        import traceback
        traceback.print_exc()
        return False

def reorder_sheets(workbook):
    """Reorder sheets to match the desired order"""
    desired_order = ['Pools', 'Policies', 'Template', 'Profiles', 'Templates', 'Organizations', 'Servers']
    current_sheets = workbook.sheetnames
    
    # Create missing sheets if needed
    for sheet_name in desired_order:
        if sheet_name not in current_sheets:
            workbook.create_sheet(sheet_name)
    
    # Reorder sheets
    for i, sheet_name in enumerate(desired_order):
        # Get the current index of the sheet
        current_index = workbook.sheetnames.index(sheet_name)
        # If it's not in the right position, move it
        if current_index != i:
            sheet = workbook[sheet_name]
            workbook.remove(sheet)
            workbook.create_sheet(sheet_name, i)
            
            # Copy the removed sheet to the new position
            new_sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    new_sheet[cell.coordinate].value = cell.value
                    if cell.has_style:
                        new_sheet[cell.coordinate].font = copy(cell.font)
                        new_sheet[cell.coordinate].border = copy(cell.border)
                        new_sheet[cell.coordinate].fill = copy(cell.fill)
                        new_sheet[cell.coordinate].number_format = copy(cell.number_format)
                        new_sheet[cell.coordinate].protection = copy(cell.protection)
                        new_sheet[cell.coordinate].alignment = copy(cell.alignment)
            
            # Copy sheet properties
            new_sheet.sheet_format = copy(sheet.sheet_format)
            new_sheet.sheet_properties = copy(sheet.sheet_properties)
            new_sheet.merged_cells = copy(sheet.merged_cells)
            new_sheet.page_margins = copy(sheet.page_margins)
            new_sheet.page_setup = copy(sheet.page_setup)
            
            # Copy column dimensions
            for key, value in sheet.column_dimensions.items():
                new_sheet.column_dimensions[key] = copy(value)
            
            # Copy row dimensions
            for key, value in sheet.row_dimensions.items():
                new_sheet.row_dimensions[key] = copy(value)

def setup_excel_file(api_client, excel_file):
    """Set up a new Excel file with the correct structure"""
    try:
        workbook = Workbook()
        
        # Create sheets in the correct order
        sheets = ['Pools', 'Policies', 'Template', 'Profiles']
        for sheet_name in sheets:
            if sheet_name in workbook.sheetnames:
                workbook.remove(workbook[sheet_name])
            workbook.create_sheet(sheet_name)
        
        # Set up Pools sheet
        pools_sheet = workbook.active
        pools_sheet.title = 'Pools'
        headers = ["Pool Type*", "Pool Name*", "Description", "Start Address*", "Size*"]
        for col, header in enumerate(headers, 1):
            cell = pools_sheet.cell(row=1, column=col, value=header)
            cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
            cell.font = Font(color='FFFFFF', bold=True)
        
        # Add sample pool data
        sample_pools = [
            ("MAC Pool", "Ai_POD-MAC-A", "MAC Pool for AI POD Fabric A", "00:25:B5:A0:00:00", "256"),
            ("MAC Pool", "Ai_POD-MAC-B", "MAC Pool for AI POD Fabric B", "00:25:B5:B0:00:00", "256"),
            ("UUID Pool", "Ai_POD-UUID-Pool", "UUID Pool for AI POD Servers", "0000-000000000001", "100")
        ]
        for idx, example in enumerate(sample_pools, 2):
            for col, value in enumerate(example, 1):
                pools_sheet.cell(row=idx, column=col, value=value)
        
        # Set up Policies sheet
        policies_sheet = workbook.create_sheet('Policies')
        policies_headers = ["Policy Type*", "Policy Name*", "Description", "Organization*"]
        for col, header in enumerate(policies_headers, 1):
            cell = policies_sheet.cell(row=1, column=col, value=header)
            cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
            cell.font = Font(color='FFFFFF', bold=True)
        
        # Add sample policy data
        sample_policies = [
            ('vNIC', 'Ai_POD-vNIC-A', 'vNIC Policy for AI POD Fabric A', 'default'),
            ('vNIC', 'Ai_POD-vNIC-B', 'vNIC Policy for AI POD Fabric B', 'default'),
            ('BIOS', 'Ai_POD-BIOS', 'BIOS Policy for AI POD', 'default'),
            ('BOOT', 'Ai_POD-BOOT', 'Boot Policy for AI POD', 'default'),
            ('QoS', 'Ai_POD-QoS', 'QoS Policy for AI POD', 'default'),
            ('Storage', 'Ai_POD-Storage', 'Storage Policy for AI POD', 'default')
        ]
        for row_idx, row in enumerate(sample_policies, 2):
            for col_idx, value in enumerate(row, 1):
                policies_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Set up Template sheet
        template_sheet = workbook.create_sheet('Template')
        template_headers = [
            "Template Name*", 
            "Organization*", 
            "Description",
            "Target Platform*",
            "BIOS Policy*",
            "Boot Policy*",
            "LAN Connectivity Policy*",
            "Storage Policy*"
        ]
        for col, header in enumerate(template_headers, 1):
            cell = template_sheet.cell(row=1, column=col, value=header)
            cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
            cell.font = Font(color='FFFFFF', bold=True)
            if '*' in header:
                template_sheet.cell(row=1, column=col).font = Font(color='FF0000', bold=True)
        
        # Add sample template data
        template_example = [
            "Ai_POD_Template",
            "Gruve",
            "Server template for AI POD workloads",
            "FIAttached",
            "Ai_POD-BIOS",
            "Ai_POD-BOOT",
            "Ai_POD-vNIC-A",
            "Ai_POD-Storage"
        ]
        for col, value in enumerate(template_example, 1):
            template_sheet.cell(row=2, column=col, value=value)
        
        # Set up Profiles sheet
        profiles_sheet = workbook.create_sheet('Profiles')
        profile_headers = ["Profile Name*", "Description", "Organization*", "Template Name*", "Server*", "Description", "Deploy*"]
        for col, header in enumerate(profile_headers, 1):
            cell = profiles_sheet.cell(row=1, column=col, value=header)
            cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
            cell.font = Font(color='FFFFFF', bold=True)
            if '*' in header:
                cell.font = Font(color='FF0000', bold=True)
        
        # Add example profile data
        profiles_sheet.append(['AI-Server-01', 'AI POD Host Profile', 'default', 'Ai_POD_Template', '', 'Production AI POD Host', 'No'])
        
        # Get servers from Intersight
        compute_api_instance = compute_api.ComputeApi(api_client)
        servers = compute_api_instance.get_compute_rack_unit_list()
        
        # Create server options list
        server_options = []
        for server in servers.results:
            server_info = f"{server.name} ({server.model}) | SN: {server.serial} | State: {server.oper_state}"
            server_options.append(server_info)
        
        # Add server dropdown validation
        server_list = '","'.join(server_options)
        server_validation = DataValidation(type='list', formula1=f'"{server_list}"', allow_blank=True)
        profiles_sheet.add_data_validation(server_validation)
        server_validation.add('E2:E1000')
        
        # Add deploy dropdown validation
        deploy_validation = DataValidation(type='list', formula1='"Yes,No"', allow_blank=True)
        profiles_sheet.add_data_validation(deploy_validation)
        deploy_validation.add('G2:G1000')
        
        # Set up dropdowns for all sheets
        # Pools sheet dropdown
        if 'Pools' in workbook.sheetnames:
            pools_sheet = workbook['Pools']
            pool_types = ['MAC Pool', 'UUID Pool']
            pool_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(pool_types)}"',
                allow_blank=True
            )
            pool_validation.add('A2:A1000')
            pools_sheet.add_data_validation(pool_validation)
            print("Added dropdown for Pool Types in Pools sheet")

        # Policies sheet dropdown
        if 'Policies' in workbook.sheetnames:
            policies_sheet = workbook['Policies']
            policy_types = ['vNIC', 'BIOS', 'BOOT', 'QoS', 'Storage']
            policy_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(policy_types)}"',
                allow_blank=True
            )
            policy_validation.add('A2:A1000')
            policies_sheet.add_data_validation(policy_validation)
            print("Added dropdown for Policy Types in Policies sheet")

        # Template sheet dropdowns
        if 'Template' in workbook.sheetnames:
            template_sheet = workbook['Template']
            platform_types = ['FIAttached', 'Standalone']
            platform_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(platform_types)}"',
                allow_blank=True
            )
            platform_validation.add('D2:D1000')
            template_sheet.add_data_validation(platform_validation)
            print("Added dropdown for Platform Types in Template sheet")

        # Save the workbook
        workbook.save(excel_file)
        print("Excel file has been set up with correct sheet order and structure")
        return True
        
    except Exception as e:
        print(f"Error setting up Excel file: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def create_template_excel(excel_file):
    """Create a fresh template Excel file with the original structure"""
    workbook = Workbook()
    
    # Remove default sheet
    default = workbook.active
    workbook.remove(default)
    
    # Create sheets in the correct order
    sheets = [
        'Pools',
        'Policies',
        'Template',
        'Profiles',
        'Templates',  # Info sheet
        'Organizations',  # Info sheet
        'Servers'  # Info sheet
    ]
    
    for sheet_name in sheets:
        workbook.create_sheet(sheet_name)
    
    # Set up Pools sheet
    pools_sheet = workbook.active
    pools_sheet.title = 'Pools'
    headers = ["Pool Type*", "Pool Name*", "Description", "Start Address*", "Size*"]
    for col, header in enumerate(headers, 1):
        cell = pools_sheet.cell(row=1, column=col, value=header)
        cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        cell.font = Font(color='FFFFFF', bold=True)
    
    # Add sample pool data
    sample_pools = [
        ("MAC Pool", "Ai_POD-MAC-A", "MAC Pool for AI POD Fabric A", "00:25:B5:A0:00:00", "256"),
        ("MAC Pool", "Ai_POD-MAC-B", "MAC Pool for AI POD Fabric B", "00:25:B5:B0:00:00", "256"),
        ("UUID Pool", "Ai_POD-UUID-Pool", "UUID Pool for AI POD Servers", "0000-000000000001", "100")
    ]
    for idx, example in enumerate(sample_pools, 2):
        for col, value in enumerate(example, 1):
            pools_sheet.cell(row=idx, column=col, value=value)
    
    # Set up Policies sheet
    policies_sheet = workbook.create_sheet('Policies')
    policies_headers = ["Policy Type*", "Policy Name*", "Description", "Organization*"]
    for col, header in enumerate(policies_headers, 1):
        cell = policies_sheet.cell(row=1, column=col, value=header)
        cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        cell.font = Font(color='FFFFFF', bold=True)
    
    # Add sample policy data
    sample_policies = [
        ('vNIC', 'Ai_POD-vNIC-A', 'vNIC Policy for AI POD Fabric A', 'default'),
        ('vNIC', 'Ai_POD-vNIC-B', 'vNIC Policy for AI POD Fabric B', 'default'),
        ('BIOS', 'Ai_POD-BIOS', 'BIOS Policy for AI POD', 'default'),
        ('BOOT', 'Ai_POD-BOOT', 'Boot Policy for AI POD', 'default'),
        ('QoS', 'Ai_POD-QoS', 'QoS Policy for AI POD', 'default'),
        ('Storage', 'Ai_POD-Storage', 'Storage Policy for AI POD', 'default')
    ]
    for row_idx, row in enumerate(sample_policies, 2):
        for col_idx, value in enumerate(row, 1):
            policies_sheet.cell(row=row_idx, column=col_idx, value=value)
    
    # Set up Template sheet
    template_sheet = workbook.create_sheet('Template')
    template_headers = [
        "Template Name*", 
        "Organization*", 
        "Description",
        "Target Platform*",
        "BIOS Policy*",
        "Boot Policy*",
        "LAN Connectivity Policy*",
        "Storage Policy*"
    ]
    for col, header in enumerate(template_headers, 1):
        cell = template_sheet.cell(row=1, column=col, value=header)
        cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        cell.font = Font(color='FFFFFF', bold=True)
        if '*' in header:
            template_sheet.cell(row=1, column=col).font = Font(color='FF0000', bold=True)
    
    # Add sample template data
    template_example = [
        "Ai_POD_Template",
        "Gruve",
        "Server template for AI POD workloads",
        "FIAttached",
        "Ai_POD-BIOS",
        "Ai_POD-BOOT",
        "Ai_POD-vNIC-A",
        "Ai_POD-Storage"
    ]
    for col, value in enumerate(template_example, 1):
        template_sheet.cell(row=2, column=col, value=value)
    
    # Set up Profiles sheet
    profiles_sheet = workbook.create_sheet('Profiles')
    profile_headers = ["Profile Name*", "Description", "Organization*", "Template Name*", "Server*", "Description", "Deploy*"]
    for col, header in enumerate(profile_headers, 1):
        cell = profiles_sheet.cell(row=1, column=col, value=header)
        cell.fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        cell.font = Font(color='FFFFFF', bold=True)
    
    # Add sample profile data with Deploy set to No
    profiles_sheet.append(['AI-Server-', '1', 'Gruve', 'Ai_Pod_Template', '', 'AI POD Server Profile', 'No'])
    
    # Add data validation for Deploy column
    deploy_validation = DataValidation(type='list', formula1='"Yes,No"', allow_blank=True)
    profiles_sheet.add_data_validation(deploy_validation)
    deploy_validation.add('G2:G1000')
    
    # Set column widths for all sheets
    for sheet in workbook.sheetnames:
        ws = workbook[sheet]
        for column in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            adjusted_width = max(max_length + 2, 15)
            ws.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width
    
    # Save the workbook
    workbook.save(excel_file)
    print(f"Created template Excel file: {excel_file}")
    return True

def add_data_validation(worksheet, column, start_row, end_row, formula):
    """Helper function to add data validation to a worksheet"""
    validation = DataValidation(type='list', formula1=formula)
    worksheet.add_data_validation(validation)
    validation.add(f'{column}{start_row}:{column}{end_row}')

def get_intersight_info(api_client, excel_file):
    """Get information from Intersight and update the Excel file"""
    try:
        # Load existing workbook
        workbook = load_workbook(excel_file)
        
        # Correct sheet naming and order
        sheet_renames = {
            'Pools': 'Pools1',
            'Policies': 'Policies1',
            'Template': 'Template1',
            'Profiles': 'Profiles1',
            'Pools1': 'Pools',
            'Policies1': 'Policies',
            'Template1': 'Template',
            'Profiles1': 'Profiles'
        }
        
        for old_name, new_name in sheet_renames.items():
            if old_name in workbook.sheetnames:
                workbook[old_name].title = new_name
        
        # Ensure correct sheet order
        desired_order = ['Pools', 'Policies', 'Template', 'Profiles']
        for sheet_name in desired_order:
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                workbook.remove(sheet)
                workbook._add_sheet(sheet, desired_order.index(sheet_name))

        # Get organizations
        print("\nGetting organizations from Intersight...")
        org_api = organization_api.OrganizationApi(api_client)
        orgs = org_api.get_organization_organization_list()
        org_names = [org.name for org in orgs.results]
        print(f"Found {len(org_names)} organizations: {org_names}")

        # Get servers
        print("\nGetting servers from Intersight...")
        compute_api_instance = compute_api.ComputeApi(api_client)
        servers = compute_api_instance.get_compute_rack_unit_list()
        server_names = [server.name for server in servers.results]
        print(f"Found {len(server_names)} servers: {server_names}")
        
        # Set up Profiles sheet dropdowns
        if 'Profiles' in workbook.sheetnames:
            profiles_sheet = workbook['Profiles']
            
            # Remove all existing data validations
            profiles_sheet.data_validations.dataValidation = []
            
            # Add server dropdown
            server_options = [f"{server.name} | SN: {server.serial}" for server in servers.results]
            server_formula = '"' + ','.join(server_options) + '"'
            server_validation = DataValidation(
                type='list',
                formula1=server_formula,
                allow_blank=True
            )
            server_validation.add('E2:E1000')  # Apply to Server column
            profiles_sheet.add_data_validation(server_validation)
            
            # Add deploy dropdown
            deploy_validation = DataValidation(
                type='list',
                formula1='"Yes,No"',
                allow_blank=True
            )
            deploy_validation.add('G2:G1000')  # Apply to Deploy column
            profiles_sheet.add_data_validation(deploy_validation)
            
            # Add organization dropdown
            org_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(org_names)}"',
                allow_blank=True
            )
            org_validation.add('C2:C1000')  # Apply to Organization column range
            profiles_sheet.add_data_validation(org_validation)
            
            print("Added dropdowns for Server, Deploy, and Organization columns")
        
        # Set up dropdowns for all sheets
        # Pools sheet dropdown
        if 'Pools' in workbook.sheetnames:
            pools_sheet = workbook['Pools']
            pool_types = ['MAC Pool', 'UUID Pool']
            pool_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(pool_types)}"',
                allow_blank=True
            )
            pool_validation.add('A2:A1000')  # Apply to Pool Types column
            pools_sheet.add_data_validation(pool_validation)
            print("Added dropdown for Pool Types in Pools sheet")

        # Policies sheet dropdown
        if 'Policies' in workbook.sheetnames:
            policies_sheet = workbook['Policies']
            policy_types = ['vNIC', 'BIOS', 'BOOT', 'QoS', 'Storage']
            policy_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(policy_types)}"',
                allow_blank=True
            )
            policy_validation.add('A2:A1000')  # Apply to Policy Types column
            policies_sheet.add_data_validation(policy_validation)
            org_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(org_names)}"',
                allow_blank=True
            )
            org_validation.add('D2:D1000')  # Apply to Organizations columns
            policies_sheet.add_data_validation(org_validation)
            print("Added dropdowns for Policy Types and Organizations in Policies sheet")

        # Template sheet dropdowns
        if 'Template' in workbook.sheetnames:
            template_sheet = workbook['Template']
            platform_types = ['FIAttached', 'Standalone']
            platform_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(platform_types)}"',
                allow_blank=True
            )
            platform_validation.add('D2:D1000')  # Apply to Platform Types column
            template_sheet.add_data_validation(platform_validation)
            print("Added dropdown for Platform Types in Template sheet")

            # Add organization dropdown to column B
            org_validation = DataValidation(
                type='list',
                formula1=f'"{",".join(org_names)}"',
                allow_blank=True
            )
            org_validation.add('B2:B1000')  # Apply to Organizations column
            template_sheet.add_data_validation(org_validation)
            print("Added dropdowns for Platform Types and Organizations in Template sheet")
        
        # Save workbook
        print("\nSaving Excel file...")
        workbook.save(excel_file)
        print("Successfully updated Excel file")
        return True
        
    except Exception as e:
        print(f"Error updating Excel file: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def get_org_moid(api_client, org_name="Gruve"):  # Set default to Gruve
    """
    Get the MOID (Managed Object ID) for an organization by name
    """
    from intersight.api import organization_api
    
    try:
        # Create Organization API instance
        api_instance = organization_api.OrganizationApi(api_client)
        
        # Get list of organizations
        orgs = api_instance.get_organization_organization_list(filter=f"Name eq '{org_name}'")
        
        if orgs.results and len(orgs.results) > 0:
            return orgs.results[0].moid
        else:
            raise Exception(f"Organization '{org_name}' not found")
            
    except Exception as e:
        raise Exception(f"Error getting organization MOID: {str(e)}")

def create_policy(api_client, policy_data):
    """
    Create a policy in Intersight based on the provided data
    """
    policy_type = policy_data['Policy Type']
    policy_name = policy_data['Policy Name']  # Updated from 'Name' to 'Policy Name'
    
    try:
        # Get Gruve organization MOID
        org_moid = get_org_moid(api_client, "Gruve")
        if not org_moid:
            print("Error: Gruve organization not found")
            return False

        org_ref = {
            "class_id": "mo.MoRef",
            "object_type": "organization.Organization",
            "moid": org_moid
        }
        
        print(f"\nCreating {policy_type} policy: {policy_name}")
        
        if policy_type == 'BIOS':
            from intersight.api import bios_api
            from intersight.model.bios_policy import BiosPolicy
            
            api_instance = bios_api.BiosApi(api_client)
            
            # Create BIOS policy with performance settings
            policy = BiosPolicy(
                class_id="bios.Policy",
                object_type="bios.Policy",
                name=policy_name,
                organization=org_ref,
                cpu_performance="enterprise",
                cpu_power_management="performance",
                cpu_energy_performance="performance",
                intel_virtualization_technology="enabled"
            )
            
            result = api_instance.create_bios_policy(policy)
            print(f"Successfully created BIOS Policy: {result.name}")
            return True
            
        elif policy_type == 'QoS':
            from intersight.api import vnic_api
            
            api_instance = vnic_api.VnicApi(api_client)
            
            # Create QoS policy
            qos = {
                "name": policy_name,
                "organization": org_ref,
                "mtu": 9000,
                "rate_limit": 0,
                "cos": 5,
                "burst": 1024,
                "priority": "Best Effort",
                "class_id": "vnic.EthQosPolicy",
                "object_type": "vnic.EthQosPolicy"
            }
            
            result = api_instance.create_vnic_eth_qos_policy(qos)
            print(f"Successfully created QoS Policy: {result.name}")
            return True
            
        elif policy_type == 'vNIC':
            from intersight.api import vnic_api, fabric_api
            from intersight.model.vnic_lan_connectivity_policy import VnicLanConnectivityPolicy
            from intersight.model.vnic_eth_if import VnicEthIf
            
            # Create API instances
            vnic_instance = vnic_api.VnicApi(api_client)
            fabric_instance = fabric_api.FabricApi(api_client)
            
            # Create Ethernet Adapter Policy
            eth_adapter = {
                "class_id": "vnic.EthAdapterPolicy",
                "object_type": "vnic.EthAdapterPolicy",
                "name": f"{policy_name}-eth-adapter",
                "organization": org_ref,
                "rss_settings": True,
                "uplink_failback_timeout": 5,
                "interrupt_settings": {
                    "coalescing_time": 125,
                    "coalescing_type": "MIN",
                    "count": 4,
                    "mode": "MSIx"
                },
                "rx_queue_settings": {
                    "count": 1,
                    "ring_size": 512
                },
                "tx_queue_settings": {
                    "count": 1,
                    "ring_size": 256
                },
                "completion_queue_settings": {
                    "count": 2,
                    "ring_size": 1
                },
                "tcp_offload_settings": {
                    "large_receive": True,
                    "large_send": True,
                    "rx_checksum": True,
                    "tx_checksum": True
                },
                "advanced_filter": True
            }
            
            eth_adapter_result = vnic_instance.create_vnic_eth_adapter_policy(eth_adapter)
            print(f"Successfully created Ethernet Adapter Policy: {eth_adapter_result.name}")

            # Create Network Group Policies for Fabric A and B
            network_group_a = {
                "class_id": "fabric.EthNetworkGroupPolicy",
                "object_type": "fabric.EthNetworkGroupPolicy",
                "name": f"{policy_name}-network-group-A",
                "organization": org_ref,
                "vlan_settings": {
                    "native_vlan": 1,
                    "allowed_vlans": "2-100"
                }
            }
            
            network_group_b = {
                "class_id": "fabric.EthNetworkGroupPolicy",
                "object_type": "fabric.EthNetworkGroupPolicy",
                "name": f"{policy_name}-network-group-B",
                "organization": org_ref,
                "vlan_settings": {
                    "native_vlan": 1,
                    "allowed_vlans": "2-100"
                }
            }
            
            group_a_result = fabric_instance.create_fabric_eth_network_group_policy(network_group_a)
            print(f"Successfully created Network Group Policy A: {group_a_result.name}")
            
            group_b_result = fabric_instance.create_fabric_eth_network_group_policy(network_group_b)
            print(f"Successfully created Network Group Policy B: {group_b_result.name}")

            # Create vNIC Policy
            lan_connectivity = {
                "class_id": "vnic.LanConnectivityPolicy",
                "object_type": "vnic.LanConnectivityPolicy",
                "name": policy_name,
                "organization": org_ref,
                "target_platform": "FIAttached"
            }
            
            lan_policy = vnic_instance.create_vnic_lan_connectivity_policy(lan_connectivity)
            print(f"Successfully created vNIC LAN Connectivity Policy: {lan_policy.name}")

            # Create vNIC eth0 for Fabric A
            eth0 = {
                "class_id": "vnic.EthIf",
                "object_type": "vnic.EthIf",
                "name": f"eth0_{lan_policy.name}",  # Make the name unique
                "order": 0,
                "placement": {
                    "class_id": "vnic.PlacementSettings",
                    "object_type": "vnic.PlacementSettings",
                    "id": "MLOM",
                    "pci_link": 0,
                    "switch_id": "A",
                    "uplink": 0
                },
                "cdn": {
                    "class_id": "vnic.Cdn",
                    "object_type": "vnic.Cdn",
                    "source": "vnic",
                    "value": "eth0"
                },
                "eth_qos_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.EthQosPolicy",
                    "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", "Ai_POD-QoS")
                },
                "eth_adapter_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.EthAdapterPolicy",
                    "moid": eth_adapter_result.moid
                },
                "fabric_eth_network_group_policy": [{
                    "class_id": "mo.MoRef",
                    "object_type": "fabric.EthNetworkGroupPolicy",
                    "moid": group_a_result.moid
                }],
                "lan_connectivity_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.LanConnectivityPolicy",
                    "moid": lan_policy.moid
                },
                "mac_pool": {
                    "class_id": "mo.MoRef",
                    "object_type": "macpool.Pool",
                    "moid": get_mac_pool_moid(api_client, "Ai_POD-MAC-A", org_moid)
                }
            }

            # Create vNIC eth1 for Fabric B
            eth1 = {
                "class_id": "vnic.EthIf",
                "object_type": "vnic.EthIf",
                "name": f"eth1_{lan_policy.name}",  # Make the name unique
                "order": 1,
                "placement": {
                    "class_id": "vnic.PlacementSettings",
                    "object_type": "vnic.PlacementSettings",
                    "id": "MLOM",
                    "pci_link": 1,  # Different PCI link for eth1
                    "switch_id": "B",
                    "uplink": 0
                },
                "cdn": {
                    "class_id": "vnic.Cdn",
                    "object_type": "vnic.Cdn",
                    "source": "vnic",
                    "value": "eth1"
                },
                "eth_qos_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.EthQosPolicy",
                    "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", "Ai_POD-QoS")
                },
                "eth_adapter_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.EthAdapterPolicy",
                    "moid": eth_adapter_result.moid
                },
                "fabric_eth_network_group_policy": [{
                    "class_id": "mo.MoRef",
                    "object_type": "fabric.EthNetworkGroupPolicy",
                    "moid": group_b_result.moid
                }],
                "lan_connectivity_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.LanConnectivityPolicy",
                    "moid": lan_policy.moid
                },
                "mac_pool": {
                    "class_id": "mo.MoRef",
                    "object_type": "macpool.Pool",
                    "moid": get_mac_pool_moid(api_client, "Ai_POD-MAC-B", org_moid)
                }
            }

            # Create the vNICs
            eth0_name = f"eth0_{lan_policy.name}"
            eth1_name = f"eth1_{lan_policy.name}"

            if not check_vnic_exists(api_client, eth0_name, lan_policy.moid):
                print("\nCreating vNIC eth0 for Fabric A...")
                eth0_result = vnic_instance.create_vnic_eth_if(eth0)
                print(f"Successfully created vNIC eth0 for Fabric A")
            else:
                print(f"\nvNIC {eth0_name} already exists, skipping creation")

            if not check_vnic_exists(api_client, eth1_name, lan_policy.moid):
                print("\nCreating vNIC eth1 for Fabric B...")
                eth1_result = vnic_instance.create_vnic_eth_if(eth1)
                print(f"Successfully created vNIC eth1 for Fabric B")
            else:
                print(f"\nvNIC {eth1_name} already exists, skipping creation")
            
            return True
            
        elif policy_type == 'Storage':
            from intersight.api import storage_api
            from intersight.model.storage_storage_policy import StorageStoragePolicy
            from intersight.model.storage_virtual_drive_policy import StorageVirtualDrivePolicy
            from intersight.model.storage_r0_drive import StorageR0Drive
            
            api_instance = storage_api.StorageApi(api_client)
            
            # Create virtual drive policy
            virtual_drive_policy = StorageVirtualDrivePolicy(
                drive_cache="Default",
                read_policy="Default",
                strip_size=512,
                access_policy="Default"
            )
            
            # Create RAID0 drive configuration
            raid0_drive = StorageR0Drive(
                enable=True,
                virtual_drive_policy=virtual_drive_policy
            )
            
            # Create storage policy
            storage_pol = StorageStoragePolicy(
                name=policy_name,
                description=policy_data.get('Description', ''),
                organization=org_ref,
                default_drive_mode="RAID0",
                raid0_drive=raid0_drive,
                unused_disks_state="NoChange",
                use_jbod_for_vd_creation=False
            )
            
            try:
                result = api_instance.create_storage_storage_policy(storage_storage_policy=storage_pol)
                print(f"Successfully created Storage Policy: {result.name}")
                return True
            except Exception as e:
                print(f"Error creating Storage policy: {str(e)}")
                raise
            
        elif policy_type == 'Boot':
            from intersight.api import boot_api
            from intersight.model.boot_precision_policy import BootPrecisionPolicy
            from intersight.model.boot_device_base import BootDeviceBase
            from intersight.model.boot_uefi_shell import BootUefiShell
            from intersight.model.boot_pxe import BootPxe
            
            api_instance = boot_api.BootApi(api_client)
            
            # Create UEFI Shell boot device
            boot_uefi = BootUefiShell(
                class_id="boot.UefiShell",
                object_type="boot.UefiShell",
                name="uefi1",
                enabled=True
            )
            
            # Create PXE boot device
            boot_pxe = BootPxe(
                class_id="boot.Pxe",
                object_type="boot.Pxe",
                name="pxe1",
                interface_name="eth0",
                ip_type="IPv4",
                enabled=True
            )
            
            # Create local disk boot device
            boot_local_disk = BootDeviceBase(
                class_id="boot.LocalDisk",
                object_type="boot.LocalDisk",
                name="local_disk1",
                enabled=True
            )
            
            # Create boot devices list
            boot_devices = [
                boot_local_disk,
                boot_uefi,
                boot_pxe
            ]
            
            # Create boot policy with the boot devices
            boot_pol = BootPrecisionPolicy(
                name=policy_name,
                description=policy_data.get('Description', ''),
                organization=org_ref,
                configured_boot_mode="Uefi",
                boot_devices=boot_devices
            )
            
            try:
                result = api_instance.create_boot_precision_policy(boot_precision_policy=boot_pol)
                print(f"Successfully created Boot Policy: {result.name}")
                return True
            except Exception as e:
                print(f"Error creating Boot policy: {str(e)}")
                raise
            
        else:
            print(f"Unsupported policy type: {policy_type}")
            return False
            
    except Exception as e:
        print(f"Error creating {policy_type} policy: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def update_profiles_with_server_info(api_client, excel_file):
    """Update the Profiles sheet with server information from Intersight"""
    try:
        # Load workbook
        workbook = load_workbook(excel_file)
        if 'Profiles' not in workbook.sheetnames:
            print("No Profiles sheet found in Excel file")
            return False
        
        profiles_sheet = workbook['Profiles']
        
        # Get servers from Intersight
        compute_api_instance = compute_api.ComputeApi(api_client)
        servers = compute_api_instance.get_compute_rack_unit_list()
        
        # Collect server info for dropdown
        server_options = [f"{server.name} | SN: {server.serial}" for server in servers.results]
        server_list = ','.join(server_options)
        
        # Add server dropdown to row 2
        server_validation = DataValidation(
            type='list',
            formula1=f'"{server_list}"',
            allow_blank=True
        )
        profiles_sheet.add_data_validation(server_validation)
        server_validation.add('E2')
        print("Added server dropdown to row 2 in Profiles sheet")
        
        # Save workbook
        try:
            workbook.save(excel_file)
            print("Successfully saved Excel file")
        except Exception as e:
            print(f"Failed to save Excel file: {str(e)}")
        
        return True
        
    except Exception as e:
        print(f"Error updating Profiles sheet: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Create and push Intersight Foundation configuration')
    parser.add_argument('--action', choices=['push', 'template', 'profiles', 'all', 'setup', 'create-template', 'get-info', 'update-servers'], required=True,
                      help='Action to perform: push (create pools and policies), template (create server template), profiles (create server profiles), all (do everything), setup (just set up Excel file), create-template (create fresh template), get-info (get current Intersight information), update-servers (update server info in Profiles sheet)')
    parser.add_argument('--file', required=True, help='Path to Excel file')
    args = parser.parse_args()
    
    if args.action == 'update-servers':
        api_client = get_api_client()
        if not api_client:
            sys.exit(1)
        update_profiles_with_server_info(api_client, args.file)
    elif args.action == 'create-template':
        create_template_excel(args.file)
    elif args.action == 'setup':
        api_client = get_api_client()
        if not api_client:
            sys.exit(1)
        setup_excel_file(api_client, args.file)
    elif args.action == 'get-info':
        api_client = get_api_client()
        if not api_client:
            sys.exit(1)
        get_intersight_info(api_client, args.file)
    else:
        api_client = get_api_client()
        if not api_client:
            sys.exit(1)
        
        if args.action in ['push', 'all']:
            process_foundation_template(args.file)
        
        if args.action in ['template', 'all']:
            create_server_template_from_excel(api_client, args.file)
        
        if args.action in ['profiles', 'all']:
            create_server_profiles_from_excel(api_client, args.file)
