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
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import uuid
from intersight.api import uuidpool_api
from intersight.api import bios_api
from intersight.api import boot_api
from intersight.api import fabric_api
from intersight.api import macpool_api
from intersight.api import storage_api
from intersight.api import vnic_api
from intersight.model.vnic_lan_connectivity_policy import VnicLanConnectivityPolicy
from intersight.model.vnic_eth_if import VnicEthIf
from intersight.model.vnic_placement_settings import VnicPlacementSettings
from intersight.model.vnic_san_connectivity_policy import VnicSanConnectivityPolicy
from intersight.model.vnic_fc_if import VnicFcIf
from intersight.model.vnic_eth_qos_policy import VnicEthQosPolicy
from intersight.model.bios_policy import BiosPolicy
from intersight.model.mo_mo_ref import MoMoRef
from intersight.model.boot_precision_policy import BootPrecisionPolicy
import time

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
    if not api_client:
        return ["default"]
        
    try:
        # Import here to avoid circular imports
        from intersight.api import organization_api
        org_api = organization_api.OrganizationApi(api_client)
        orgs = org_api.get_organization_organization_list()
        return [org.name for org in orgs.results] if orgs.results else ["default"]
    except Exception as e:
        print(f"Error fetching organizations: {str(e)}")
        return ["default"]

def create_mac_pool(api_client, pool_data):
    """
    Create a MAC Pool in Intersight
    """
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

        from intersight.api import macpool_api
        api_instance = macpool_api.MacpoolApi(api_client)
        
        # Parse MAC blocks from the input
        mac_blocks_str = pool_data['ID Blocks']
        mac_blocks = []
        
        # Format should be: start-end,start-end (e.g., "00:25:B5:00:00:00-00:25:B5:00:00:FF")
        if mac_blocks_str and isinstance(mac_blocks_str, str):
            for block in mac_blocks_str.split(','):
                start, end = block.strip().split('-')
                mac_blocks.append({
                    "ClassId": "macpool.Block",
                    "ObjectType": "macpool.Block",
                    "From": start.strip(),
                    "To": end.strip()
                })
        
        # Create the pool
        pool = {
            "ClassId": "macpool.Pool",
            "ObjectType": "macpool.Pool",
            "Name": pool_data['Name'],
            "Description": pool_data['Description'] if pd.notna(pool_data['Description']) else "",
            "AssignmentOrder": pool_data['Assignment Order'].lower() if pd.notna(pool_data['Assignment Order']) else "sequential",
            "MacBlocks": mac_blocks,
            "Organization": org_ref
        }
        
        # Create the pool in Intersight
        result = api_instance.create_macpool_pool(macpool_pool=pool)
        print(f"Successfully created MAC Pool: {result.name}")
        return True
        
    except Exception as e:
        print(f"Error creating MAC Pool: {str(e)}")
        return False

def create_uuid_pool(api_client, pool_data):
    """
    Create a UUID Pool in Intersight
    """
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

        from intersight.api import uuidpool_api
        api_instance = uuidpool_api.UuidpoolApi(api_client)
        
        # Create UUID pool block
        uuid_blocks = []
        if 'From' in pool_data and 'Size' in pool_data:
            block = {
                "class_id": "uuidpool.Block",
                "object_type": "uuidpool.Block",
                "from_uuid": pool_data['From'],
                "size": pool_data['Size'],
                "to": None  # This is optional and can be calculated by Intersight
            }
            uuid_blocks.append(block)
        
        # Create UUID pool
        pool = {
            "class_id": "uuidpool.Pool",
            "object_type": "uuidpool.Pool",
            "name": pool_data['Name'],
            "description": "UUID Pool for UCS Servers",
            "organization": org_ref,
            "assignment_order": "sequential",  # Must be lowercase
            "prefix": "000025B5-0000-0000",
            "uuid_suffix_blocks": uuid_blocks
        }
        
        result = api_instance.create_uuidpool_pool(pool)
        print(f"Successfully created UUID Pool: {result.name}")
        return True
        
    except Exception as e:
        print(f"Error creating UUID Pool: {str(e)}")
        uuid_blocks_str = json.dumps(uuid_blocks, indent=1) if 'uuid_blocks' in locals() else "No blocks defined"
        print(f"UUID blocks: {uuid_blocks_str}")
        return False

def format_uuid_suffix(uuid_str):
    """Format a UUID suffix to match Intersight's expected pattern: XXXX-XXXXXXXXXXXX"""
    # Remove any non-hex characters and pad to 16 characters
    clean_uuid = ''.join(c for c in uuid_str if c.isalnum()).zfill(16)
    return f"{clean_uuid[:4]}-{clean_uuid[4:]}"

def create_pool(api_client, pool_data):
    """
    Create a pool in Intersight based on the provided data
    """
    pool_type = pool_data['Pool Type']
    pool_name = pool_data['Name']
    
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

        if pool_type == 'MAC Pool':
            return create_mac_pool(api_client, pool_data)
        elif pool_type == 'UUID Pool':
            return create_uuid_pool(api_client, pool_data)
        elif pool_type in ['WWNN Pool', 'WWPN Pool']:
            from intersight.api import fcpool_api
            from intersight.model.fcpool_pool import FcpoolPool
            from intersight.model.fcpool_block import FcpoolBlock
            
            api_instance = fcpool_api.FcpoolApi(api_client)
            
            # Create FC pool block
            block = FcpoolBlock(
                class_id="fcpool.Block",
                object_type="fcpool.Block",
                From="20:00:00:25:B5:00:00:00",  # Using a valid WWN format
                To="20:00:00:25:B5:00:00:FF"     # Using a valid WWN format
            )
            
            # Create FC pool
            pool = FcpoolPool(
                class_id="fcpool.Pool",
                object_type="fcpool.Pool",
                name=pool_name,
                organization=org_ref,
                pool_purpose=pool_type.split()[0],  
                assignment_order="sequential",
                id_blocks=[block]
            )
            
            result = api_instance.create_fcpool_pool(pool)
            print(f"Successfully created {pool_type}: {result.name}")
            return True
            
        elif pool_type in ['IP Pool', 'IQN Pool', 'vNIC', 'vHBA']:
            print(f"Pool type {pool_type} implementation coming soon")
            return False
        else:
            print(f"Unknown pool type: {pool_type}")
            return False
            
    except Exception as e:
        print(f"Error creating {pool_type} {pool_name}: {str(e)}")
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
    from intersight.api import vnic_api, bios_api, fabric_api, boot_api
    from intersight.model.vnic_lan_connectivity_policy import VnicLanConnectivityPolicy
    from intersight.model.vnic_eth_if import VnicEthIf
    from intersight.model.vnic_placement_settings import VnicPlacementSettings
    from intersight.model.vnic_san_connectivity_policy import VnicSanConnectivityPolicy
    from intersight.model.vnic_fc_if import VnicFcIf
    from intersight.model.vnic_eth_qos_policy import VnicEthQosPolicy
    from intersight.model.bios_policy import BiosPolicy
    from intersight.model.mo_mo_ref import MoMoRef
    from intersight.model.boot_precision_policy import BootPrecisionPolicy
    import json
    
    policy_type = policy_data['Policy Type']
    policy_name = policy_data['Name']
    
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

        if policy_type == 'BIOS':
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
            
            print("\nBIOS Policy JSON:")
            print(json.dumps(policy.to_dict(), indent=2))  # Show the JSON being sent
            
            result = api_instance.create_bios_policy(policy)
            print(f"Successfully created BIOS Policy: {result.name}")
            return True
            
        elif policy_type == 'QoS':
            api_instance = vnic_api.VnicApi(api_client)
            
            # Create QoS policy with high priority for control plane traffic
            policy = VnicEthQosPolicy(
                class_id="vnic.EthQosPolicy",
                object_type="vnic.EthQosPolicy",
                name=policy_name,
                organization=org_ref,
                mtu=9000,  
                rate_limit=0,
                cos=5,
                burst=1024,
                priority="Best Effort"  # Fixed the priority value
            )
            
            print("\nQoS Policy JSON:")
            print(json.dumps(policy.to_dict(), indent=2))
            
            result = api_instance.create_vnic_eth_qos_policy(policy)
            print(f"Successfully created QoS Policy: {result.name}")
            return True
            
        elif policy_type == 'vNIC':
            return create_vnic_policy(api_client, policy_data)
            
        elif policy_type == 'Storage':
            from intersight.api import storage_api
            
            try:
                # Create Storage Policy with manual drive configuration
                storage_policy = {
                    "class_id": "storage.StoragePolicy",
                    "object_type": "storage.StoragePolicy",
                    "name": f"{policy_name}",
                    "description": "Storage Policy for OS and ETCD",
                    "organization": org_ref,
                    "use_jbod_for_vd_creation": True,
                    "unused_disks_state": "UnconfiguredGood",
                    "m2_virtual_drive": {
                        "enable": False
                    },
                    "global_hot_spares": "",
                    "raid0_drive": {
                        "drive_slots": "1,2",
                        "enable": True,
                        "virtual_drive_policy": {
                            "access_policy": "ReadWrite",
                            "drive_cache": "Enable",
                            "read_policy": "ReadAhead",
                            "strip_size": 64,
                            "write_policy": "WriteBackGoodBbu"
                        }
                    }
                }
                
                storage_instance = storage_api.StorageApi(api_client)
                storage_result = storage_instance.create_storage_storage_policy(storage_policy)
                print(f"\nCreated Storage Policy: {storage_result.name}")
                
            except Exception as e:
                print(f"Error creating Storage policy {policy_name}: {str(e)}")
                return False
            return True
            
        # Other policy types will go here...
        
    except Exception as e:
        print(f"Error creating {policy_type} policy {policy_name}: {str(e)}")
        return False

def create_vnic_policy(api_client, policy_data):
    """
    Create a vNIC Policy in Intersight
    """
    try:
        # Get organization MOID
        org_moid = get_org_moid(api_client, "Gruve")
        if not org_moid:
            print("Error: Organization not found")
            return False

        org_ref = {
            "class_id": "mo.MoRef",
            "object_type": "organization.Organization",
            "moid": org_moid
        }

        # Create LAN Connectivity Policy
        vnic_instance = vnic_api.VnicApi(api_client)
        lan_connectivity = {
            "class_id": "vnic.LanConnectivityPolicy",
            "object_type": "vnic.LanConnectivityPolicy",
            "name": policy_data['Name'],
            "description": policy_data['Description'] if pd.notna(policy_data['Description']) else "",
            "organization": org_ref,
            "target_platform": "FIAttached"
        }
        
        lan_policy = vnic_instance.create_vnic_lan_connectivity_policy(lan_connectivity)
        print(f"\nCreated LAN Connectivity Policy: {lan_policy.name}")

        # Add a small delay to ensure the LAN policy is ready
        time.sleep(2)

        # Create vNIC eth0 for Fabric A
        eth0 = {
            "class_id": "vnic.EthIf",
            "object_type": "vnic.EthIf",
            "name": "eth0",
            "order": 0,
            "placement": {
                "class_id": "vnic.PlacementSettings",
                "object_type": "vnic.PlacementSettings",
                "id": "1",
                "pci_link": 0,
                "uplink": 0,
                "switch_id": "A"
            },
            "cdn": {
                "class_id": "vnic.Cdn",
                "object_type": "vnic.Cdn",
                "source": "vnic"
            },
            "eth_adapter_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthAdapterPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthAdapterPolicy", "vNIC-Default-eth-adapter")
            },
            "eth_qos_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthQosPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", "vNIC-Default-qos")
            },
            "fabric_eth_network_group_policy": {
                "class_id": "mo.MoRef",
                "object_type": "fabric.EthNetworkGroupPolicy",
                "moid": get_policy_moid(api_client, "fabric.EthNetworkGroupPolicy", "vNIC-Default-network-group-A")
            },
            "lan_connectivity_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.LanConnectivityPolicy",
                "moid": lan_policy.moid
            },
            "mac_pool": {
                "class_id": "mo.MoRef",
                "object_type": "macpool.Pool",
                "moid": get_policy_moid(api_client, "macpool.Pool", "MAC-Pool-A")
            }
        }
        
        eth0_if = vnic_instance.create_vnic_eth_if(eth0)
        print(f"\nCreated vNIC eth0 for Fabric A")

        # Add another small delay before creating the second vNIC
        time.sleep(2)

        # Create vNIC eth1 for Fabric B
        eth1 = {
            "class_id": "vnic.EthIf",
            "object_type": "vnic.EthIf",
            "name": "eth1",
            "order": 1,
            "placement": {
                "class_id": "vnic.PlacementSettings",
                "object_type": "vnic.PlacementSettings",
                "id": "2",
                "pci_link": 0,
                "uplink": 0,
                "switch_id": "B"
            },
            "cdn": {
                "class_id": "vnic.Cdn",
                "object_type": "vnic.Cdn",
                "source": "vnic"
            },
            "eth_adapter_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthAdapterPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthAdapterPolicy", "vNIC-Default-eth-adapter")
            },
            "eth_qos_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthQosPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", "vNIC-Default-qos")
            },
            "fabric_eth_network_group_policy": {
                "class_id": "mo.MoRef",
                "object_type": "fabric.EthNetworkGroupPolicy",
                "moid": get_policy_moid(api_client, "fabric.EthNetworkGroupPolicy", "vNIC-Default-network-group-B")
            },
            "lan_connectivity_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.LanConnectivityPolicy",
                "moid": lan_policy.moid
            },
            "mac_pool": {
                "class_id": "mo.MoRef",
                "object_type": "macpool.Pool",
                "moid": get_policy_moid(api_client, "macpool.Pool", "MAC-Pool-B")
            }
        }
        
        eth1_if = vnic_instance.create_vnic_eth_if(eth1)
        print(f"\nCreated vNIC eth1 for Fabric B")

        print("\nUpdated LAN Connectivity Policy with vNIC references")
        return True

    except Exception as e:
        print(f"Error creating vNIC policy {policy_data['Name']}: {str(e)}")
        if hasattr(e, 'status') and hasattr(e, 'reason'):
            print(f"Status Code: {e.status}")
            print(f"Reason: {e.reason}")
        if hasattr(e, 'headers'):
            print(f"HTTP response headers: {e.headers}")
        if hasattr(e, 'body'):
            print(f"HTTP response body: {e.body}")
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
    """
    Get the MOID of a policy by name and type
    """
    # Create the appropriate API instance based on policy type
    if policy_type == "vnic.EthAdapterPolicy":
        api_instance = vnic_api.VnicApi(api_client)
        policies = api_instance.get_vnic_eth_adapter_policy_list()
    elif policy_type == "vnic.EthQosPolicy":
        api_instance = vnic_api.VnicApi(api_client)
        policies = api_instance.get_vnic_eth_qos_policy_list()
    elif policy_type == "fabric.EthNetworkControlPolicy":
        api_instance = fabric_api.FabricApi(api_client)
        policies = api_instance.get_fabric_eth_network_control_policy_list()
    elif policy_type == "fabric.EthNetworkGroupPolicy":
        api_instance = fabric_api.FabricApi(api_client)
        policies = api_instance.get_fabric_eth_network_group_policy_list()
    elif policy_type == "macpool.Pool":
        api_instance = macpool_api.MacpoolApi(api_client)
        policies = api_instance.get_macpool_pool_list()
    else:
        raise Exception(f"Unsupported policy type: {policy_type}")
    
    # Search for the policy by name
    for policy in policies.results:
        if policy.name == policy_name:
            return policy.moid
    
    print(f"Warning: Policy '{policy_name}' of type '{policy_type}' not found")
    return None

def process_foundation_template(excel_file):
    """
    Read the Excel template and create pools and policies in Intersight
    """
    try:
        # Get API client
        api_client = get_api_client()
        
        # Read Excel file
        pools_df = pd.read_excel(excel_file, sheet_name='Pools')
        policies_df = pd.read_excel(excel_file, sheet_name='Policies')
        
        # Process pools
        for _, pool in pools_df.iterrows():
            if pd.notna(pool['Name']):  # Only process rows with a name
                create_pool(api_client, pool)
        
        # Process policies
        for _, policy in policies_df.iterrows():
            if pd.notna(policy['Name']):  # Only process rows with a name
                create_policy(api_client, policy)
                
        print("Completed processing the Foundation template")
        
    except Exception as e:
        print(f"Error processing template: {str(e)}")

# Create output directory if it doesn't exist
os.makedirs('output', exist_ok=True)

# Define the output file path
output_file = 'output/Intersight_Foundation.xlsx'

# Create a Pandas Excel writer using openpyxl as the engine
writer = pd.ExcelWriter(output_file, engine='openpyxl')

# Get organizations from Intersight
api_client = get_api_client()
organizations = get_organizations(api_client)
organizations_str = ','.join(organizations) if organizations else "default"

# Pools Sheet Data
pools_data = {
    'Pool Type': [
        'MAC Pool',
        'MAC Pool',
        'UUID Pool'
    ],
    'Name': [
        'MAC-Pool-A',
        'MAC-Pool-B',
        'UUID-Pool'
    ],
    'Description': [
        'MAC Pool for Fabric A',
        'MAC Pool for Fabric B',
        'UUID Pool for UCS Servers'
    ],
    'Organization': ['default'] * 3,
    'Assignment Order': ['sequential'] * 3,
    'ID Blocks': [
        '00:25:B5:A1:00:00-00:25:B5:A1:00:FF',  
        '00:25:B5:A2:00:00-00:25:B5:A2:00:FF',  
        '0000-000000000010'  
    ],
    'First Address': [
        '', '', ''
    ],
    'Size': [
        '', '', ''
    ]
}

# Policies Sheet Data
policies_data = {
    'Policy Type': [
        'BIOS',
        'QoS',
        'vNIC',
        'Storage'
    ],
    'Name': [
        'BIOS-Default',
        'QoS-Default',
        'vNIC-Default',
        'Storage-Default'
    ],
    'Description': [
        'Default BIOS policy',
        'Default QoS policy',
        'Default vNIC policy',
        'Default Storage policy'
    ],
    'Organization': ['default'] * 4,
    'MAC Pool A': ['MAC-Pool-A'] * 4,
    'MAC Pool B': ['MAC-Pool-B'] * 4,
    'WWNN Pool': [''] * 4,
    'WWPN Pool A': [''] * 4,
    'WWPN Pool B': [''] * 4
}

# Create DataFrames
pools_df = pd.DataFrame(pools_data)
policies_df = pd.DataFrame(policies_data)

# Write DataFrames to Excel
pools_df.to_excel(writer, sheet_name='Pools', index=False)
policies_df.to_excel(writer, sheet_name='Policies', index=False)

# Get workbook and sheets
workbook = writer.book
pools_sheet = writer.sheets['Pools']
policies_sheet = writer.sheets['Policies']

# Define styles
header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
header_font = Font(bold=True, color='FFFFFF')
border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Apply styles to Pools sheet
for col in range(1, len(pools_data.keys()) + 1):
    cell = pools_sheet.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font
    column_letter = get_column_letter(col)
    pools_sheet.column_dimensions[column_letter].width = 20

# Apply styles to Policies sheet
for col in range(1, len(policies_data.keys()) + 1):
    cell = policies_sheet.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font
    column_letter = get_column_letter(col)
    policies_sheet.column_dimensions[column_letter].width = 20

# Add data validation for Assignment Order
assignment_order_dv = DataValidation(
    type="list",
    formula1='"SEQUENTIAL,RANDOM"',
    allow_blank=True
)
pools_sheet.add_data_validation(assignment_order_dv)
assignment_order_col = 4  
for row in range(2, len(pools_data['Pool Type']) + 2):
    assignment_order_dv.add(f'D{row}')

# Add data validation for Organizations in Pools sheet
org_dv = DataValidation(
    type="list",
    formula1=f'"{organizations_str}"',
    allow_blank=True
)
pools_sheet.add_data_validation(org_dv)
for row in range(2, len(pools_data['Pool Type']) + 2):
    org_dv.add(f'E{row}')  

# Add data validation for Organizations in Policies sheet
policies_org_dv = DataValidation(
    type="list",
    formula1=f'"{organizations_str}"',
    allow_blank=True
)
policies_sheet.add_data_validation(policies_org_dv)
for row in range(2, len(policies_data['Policy Type']) + 2):
    policies_org_dv.add(f'E{row}')  

# Add data validation for Policy Type
policy_type_dv = DataValidation(
    type="list",
    formula1='"BIOS,QoS,vNIC,Storage"',
    allow_blank=True
)
policies_sheet.add_data_validation(policy_type_dv)
for row in range(2, len(policies_data['Policy Type']) + 2):
    policy_type_dv.add(f'A{row}')

# Add data validation for MAC Pool
mac_pool_dv = DataValidation(
    type="list",
    formula1='"' + ','.join(pools_data['Name']) + '"',
    allow_blank=True
)
policies_sheet.add_data_validation(mac_pool_dv)
for row in range(2, len(policies_data['Policy Type']) + 2):
    mac_pool_dv.add(f'G{row}')  

# Add data validation for WWNN Pool
wwnn_pool_dv = DataValidation(
    type="list",
    formula1='"' + '"' + '"',
    allow_blank=True
)
policies_sheet.add_data_validation(wwnn_pool_dv)
for row in range(2, len(policies_data['Policy Type']) + 2):
    wwnn_pool_dv.add(f'H{row}')  

# Add data validation for WWPN Pool A
wwpn_pool_a_dv = DataValidation(
    type="list",
    formula1='"' + '"' + '"',
    allow_blank=True
)
policies_sheet.add_data_validation(wwpn_pool_a_dv)
for row in range(2, len(policies_data['Policy Type']) + 2):
    wwpn_pool_a_dv.add(f'I{row}')  

# Add data validation for WWPN Pool B
wwpn_pool_b_dv = DataValidation(
    type="list",
    formula1='"' + '"' + '"',
    allow_blank=True
)
policies_sheet.add_data_validation(wwpn_pool_b_dv)
for row in range(2, len(policies_data['Policy Type']) + 2):
    wwpn_pool_b_dv.add(f'J{row}')  

# Add data validation for Pool Type
pool_type_dv = DataValidation(
    type="list",
    formula1='"MAC Pool,UUID Pool"',
    allow_blank=True
)
pools_sheet.add_data_validation(pool_type_dv)
for row in range(2, len(pools_data['Pool Type']) + 2):
    pool_type_dv.add(f'A{row}')

# Save the workbook
writer.close()

print(f"Excel template has been created at: {output_file}")

if __name__ == "__main__":
    output_file = 'output/Intersight_Foundation.xlsx'
    
    # Create output directory if it doesn't exist
    os.makedirs('output', exist_ok=True)
    
    # Create Excel writer
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # Create DataFrames
    pools_df = pd.DataFrame(pools_data)
    policies_df = pd.DataFrame(policies_data)
    
    # Write DataFrames to Excel
    pools_df.to_excel(writer, sheet_name='Pools', index=False)
    policies_df.to_excel(writer, sheet_name='Policies', index=False)
    
    # Get workbook and sheets
    workbook = writer.book
    pools_sheet = writer.sheets['Pools']
    policies_sheet = writer.sheets['Policies']
    
    # Add data validation for Pool Type
    pool_type_dv = DataValidation(
        type="list",
        formula1='"MAC Pool,UUID Pool"',
        allow_blank=True
    )
    pools_sheet.add_data_validation(pool_type_dv)
    for row in range(2, len(pools_data['Pool Type']) + 2):
        pool_type_dv.add(f'A{row}')
    
    # Add data validation for Assignment Order
    assignment_order_dv = DataValidation(
        type="list",
        formula1='"SEQUENTIAL,RANDOM"',
        allow_blank=True
    )
    pools_sheet.add_data_validation(assignment_order_dv)
    for row in range(2, len(pools_data['Assignment Order']) + 2):
        assignment_order_dv.add(f'C{row}')
    
    # Add data validation for Organizations
    org_dv = DataValidation(
        type="list",
        formula1='"default"',
        allow_blank=True
    )
    pools_sheet.add_data_validation(org_dv)
    for row in range(2, len(pools_data['Pool Type']) + 2):
        org_dv.add(f'E{row}')
    
    # Add data validation for Policy Type
    policy_type_dv = DataValidation(
        type="list",
        formula1='"BIOS,QoS,vNIC,Storage"',
        allow_blank=True
    )
    policies_sheet.add_data_validation(policy_type_dv)
    for row in range(2, len(policies_data['Policy Type']) + 2):
        policy_type_dv.add(f'A{row}')
    
    # Add data validation for Organizations in Policies sheet
    policies_org_dv = DataValidation(
        type="list",
        formula1='"default"',
        allow_blank=True
    )
    policies_sheet.add_data_validation(policies_org_dv)
    for row in range(2, len(policies_data['Policy Type']) + 2):
        policies_org_dv.add(f'E{row}')
    
    # Add data validation for MAC Pool
    mac_pool_dv = DataValidation(
        type="list",
        formula1='"' + ','.join(pools_data['Name']) + '"',
        allow_blank=True
    )
    policies_sheet.add_data_validation(mac_pool_dv)
    for row in range(2, len(policies_data['Policy Type']) + 2):
        mac_pool_dv.add(f'G{row}')
    
    # Add data validation for WWNN Pool
    wwnn_pool_dv = DataValidation(
        type="list",
        formula1='"' + '"' + '"',
        allow_blank=True
    )
    policies_sheet.add_data_validation(wwnn_pool_dv)
    for row in range(2, len(policies_data['Policy Type']) + 2):
        wwnn_pool_dv.add(f'H{row}')
    
    # Add data validation for WWPN Pool A
    wwpn_pool_a_dv = DataValidation(
        type="list",
        formula1='"' + '"' + '"',
        allow_blank=True
    )
    policies_sheet.add_data_validation(wwpn_pool_a_dv)
    for row in range(2, len(policies_data['Policy Type']) + 2):
        wwpn_pool_a_dv.add(f'I{row}')
    
    # Add data validation for WWPN Pool B
    wwpn_pool_b_dv = DataValidation(
        type="list",
        formula1='"' + '"' + '"',
        allow_blank=True
    )
    policies_sheet.add_data_validation(wwpn_pool_b_dv)
    for row in range(2, len(policies_data['Policy Type']) + 2):
        wwpn_pool_b_dv.add(f'J{row}')
    
    # Save the workbook
    writer.close()
    
    print(f"Excel template has been created at: {output_file}")
    
    # Ask if user wants to process the template
    response = input("Would you like to process the template and create the pools/policies in Intersight? (y/n): ")
    if response.lower() == 'y':
        process_foundation_template(output_file)
