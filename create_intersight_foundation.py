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
from intersight.api import (
    bios_api,
    boot_api,
    compute_api,
    fabric_api,
    iam_api,
    macpool_api,
    organization_api,
    server_api,
    storage_api,
    uuidpool_api,
    vnic_api
)
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
import argparse
import sys

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

def pool_exists(api_client, pool_type, pool_name):
    """
    Check if a pool already exists in Intersight
    """
    try:
        # Get organization MOID
        org_moid = get_org_moid(api_client)
        
        # Create API instance based on pool type
        if pool_type == "MAC Pool":
            api_instance = macpool_api.MacpoolApi(api_client)
            response = api_instance.get_macpool_pool_list(filter=f"Name eq '{pool_name}'")
        elif pool_type == "UUID Pool":
            api_instance = uuidpool_api.UuidpoolApi(api_client)
            response = api_instance.get_uuidpool_pool_list(filter=f"Name eq '{pool_name}'")
        else:
            return False

        return len(response.results) > 0

    except Exception as e:
        print(f"Error checking if pool exists: {str(e)}")
        return False

def create_pool(api_client, pool_data):
    """
    Create a pool in Intersight based on the provided data
    """
    pool_type = pool_data['Pool Type']
    pool_name = pool_data['Name']
    
    try:
        # Check if pool already exists
        if pool_exists(api_client, pool_type, pool_name):
            print(f"Pool {pool_name} already exists, skipping creation")
            return True

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
            print(json.dumps(policy.to_dict(), indent=2))
            
            result = api_instance.create_bios_policy(policy)
            print(f"Successfully created BIOS Policy: {result.name}")
            return True
            
        elif policy_type == 'QoS':
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
            
            print("\nQoS Policy JSON:")
            print(json.dumps(qos, indent=2))
            
            result = api_instance.create_vnic_eth_qos_policy(qos)
            print(f"Successfully created QoS Policy: {result.name}")
            return True
            
        elif policy_type == 'vNIC':
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
            
            print("\nEthernet Adapter Policy JSON:")
            print(json.dumps(eth_adapter, indent=2))
            
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
            
            print("\nNetwork Group Policy A JSON:")
            print(json.dumps(network_group_a, indent=2))
            
            group_a_result = fabric_instance.create_fabric_eth_network_group_policy(network_group_a)
            print(f"Successfully created Network Group Policy A: {group_a_result.name}")
            
            print("\nNetwork Group Policy B JSON:")
            print(json.dumps(network_group_b, indent=2))
            
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
            
            print("\nvNIC LAN Connectivity Policy JSON:")
            print(json.dumps(lan_connectivity, indent=2))
            
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
                "mac_address_type": "POOL",
                "mac_pool": {
                    "class_id": "mo.MoRef",
                    "object_type": "macpool.Pool",
                    "moid": get_mac_pool_moid(api_client, "MAC-Pool-A", org_moid)
                },
                "eth_qos_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.EthQosPolicy",
                    "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", "QoS-Ai_pod")
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
                "mac_address_type": "POOL",
                "mac_pool": {
                    "class_id": "mo.MoRef",
                    "object_type": "macpool.Pool",
                    "moid": get_mac_pool_moid(api_client, "MAC-Pool-B", org_moid)
                },
                "eth_qos_policy": {
                    "class_id": "mo.MoRef",
                    "object_type": "vnic.EthQosPolicy",
                    "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", "QoS-Ai_pod")
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
            
    except Exception as e:
        print(f"Error creating {policy_type} policy {policy_name}: {str(e)}")
        return False

def create_vnic_policy(api_client, policy_data):
    """
    Create a vNIC Policy in Intersight
    """
    try:
        # Get organization MOID
        org_name = policy_data['Organization'] if pd.notna(policy_data['Organization']) else "default"
        org_moid = get_org_moid(api_client, org_name)
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
                "source": "vnic"
            },
            "eth_adapter_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthAdapterPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthAdapterPolicy", f"{policy_data['Name']}-eth-adapter")
            },
            "eth_qos_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthQosPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", policy_data['Name'])
            },
            "fabric_eth_network_group_policy": [
                {
                    "class_id": "mo.MoRef",
                    "object_type": "fabric.EthNetworkGroupPolicy",
                    "moid": get_policy_moid(api_client, "fabric.EthNetworkGroupPolicy", f"{policy_data['Name']}-network-group-A")
                }
            ],
            "lan_connectivity_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.LanConnectivityPolicy",
                "moid": lan_policy.moid
            },
            "mac_pool": {
                "class_id": "mo.MoRef",
                "object_type": "macpool.Pool",
                "moid": get_mac_pool_moid(api_client, "MAC-Pool-A", org_moid)
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
                "source": "vnic"
            },
            "eth_adapter_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthAdapterPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthAdapterPolicy", f"{policy_data['Name']}-eth-adapter")
            },
            "eth_qos_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.EthQosPolicy",
                "moid": get_policy_moid(api_client, "vnic.EthQosPolicy", policy_data['Name'])
            },
            "fabric_eth_network_group_policy": [
                {
                    "class_id": "mo.MoRef",
                    "object_type": "fabric.EthNetworkGroupPolicy",
                    "moid": get_policy_moid(api_client, "fabric.EthNetworkGroupPolicy", f"{policy_data['Name']}-network-group-B")
                }
            ],
            "lan_connectivity_policy": {
                "class_id": "mo.MoRef",
                "object_type": "vnic.LanConnectivityPolicy",
                "moid": lan_policy.moid
            },
            "mac_pool": {
                "class_id": "mo.MoRef",
                "object_type": "macpool.Pool",
                "moid": get_mac_pool_moid(api_client, "MAC-Pool-B", org_moid)
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

        print("\nUpdated LAN Connectivity Policy with vNIC references")
        return True

    except Exception as e:
        print(f"Error creating vNIC policy {policy_data['Name']}: {str(e)}")
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
        if policy_type == "bios/Policies":
            api_instance = bios_api.BiosApi(api_client)
            policies = api_instance.get_bios_policy_list()
        elif policy_type == "vnic/LanConnectivityPolicies":
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
            return False
            
        # Process Pools sheet
        if 'Pools' in df:
            pools_df = df['Pools']
            for _, row in pools_df.iterrows():
                create_pool(api_client, row)
                
        # Process Policies sheet in specific order
        if 'Policies' in df:
            policies_df = df['Policies']
            
            # Create policies in order: BIOS, QoS, vNIC, Storage
            policy_order = ['BIOS', 'QoS', 'vNIC', 'Storage']
            
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
            
            # Create policies in order: BIOS, QoS, vNIC, Storage
            policy_order = ['BIOS', 'QoS', 'vNIC', 'Storage']
            
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
    policy_class_ids = {
        'BIOS': 'bios/Policies',
        'QoS': 'vnic.EthQosPolicy',
        'vNIC': 'vnic/LanConnectivityPolicies',
        'Storage': 'storage.StoragePolicy'
    }
    return policy_class_ids.get(policy_type, None)

def add_template_sheet(excel_file, api_client):
    """Add a template sheet to the existing Excel file with consistent formatting"""
    try:
        # First, read the Pools sheet to get the exact color
        workbook = load_workbook(excel_file)
        pools_sheet = workbook['Pools']
        header_cell = pools_sheet['A1']
        header_color = header_cell.fill.start_color.rgb
        
        # Get organizations from Intersight
        organizations = get_organizations(api_client)
        
        # Create a list of rows for the template
        data = [
            ['Configuration Type', 'Value', 'Description'],
            ['Template Name', 'Ai_Pod_Template', 'Name of the server profile template'],
            ['Template Description', 'Server template for AI Pod configuration', 'Description of the template'],
            ['Organization', organizations[0] if organizations else 'default', 'Organization to create the template in'],
            ['BIOS Policy', 'BIOS-Ai_pod', 'BIOS policy to use for server configuration'],
            ['UUID Pool', 'UUID-Pool', 'UUID pool for server identification'],
            ['LAN Connectivity Policy', 'vNIC-Ai_pod', 'LAN connectivity policy for network configuration'],
            ['QoS Policy', 'QoS-Ai_pod', 'Quality of Service policy for network traffic']
        ]

        # Remove existing Template sheet if it exists
        if 'Template' in workbook.sheetnames:
            del workbook['Template']
        
        # Create new Template sheet
        worksheet = workbook.create_sheet('Template')
        
        # Write data to worksheet
        for row in data:
            worksheet.append(row)
        
        # Define styles - using the exact color from Pools sheet
        header_fill = PatternFill(start_color=header_color[2:], end_color=header_color[2:], fill_type='solid')  # Remove 'FF' prefix from RGB
        header_font = Font(color='FFFFFF', bold=True)
        cell_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Apply header formatting
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = cell_border
            
        # Apply cell formatting to all cells
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.border = cell_border
                cell.alignment = Alignment(horizontal='left', vertical='center')
                
        # Adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Add data validation for Organization with actual organizations from Intersight
        org_validation = DataValidation(
            type='list',
            formula1=f'"{",".join(organizations)}"',
            allow_blank=True
        )
        worksheet.add_data_validation(org_validation)
        org_row = [i+1 for i, row in enumerate(data) if row[0] == 'Organization'][0]
        org_validation.add(f'B{org_row}')
        
        # Add data validation for policies
        try:
            policies_df = pd.read_excel(excel_file, sheet_name='Policies')
            for policy_type in ['BIOS Policy', 'UUID Pool', 'LAN Connectivity Policy', 'QoS Policy']:
                row_num = [i+1 for i, row in enumerate(data) if row[0] == policy_type][0]
                if policy_type == 'UUID Pool':
                    validation = DataValidation(
                        type='list',
                        formula1='"UUID-Pool"',
                        allow_blank=True
                    )
                else:
                    policy_name = policy_type.split()[0]
                    policy_values = policies_df[policies_df['Policy Type'] == policy_name]['Name'].tolist()
                    if policy_values:
                        validation = DataValidation(
                            type='list',
                            formula1=f'"{",".join(policy_values)}"',
                            allow_blank=True
                        )
                    else:
                        continue
                        
                worksheet.add_data_validation(validation)
                validation.add(f'B{row_num}')
        except Exception as e:
            print(f"Warning: Could not add policy validations: {str(e)}")
        
        # Save the workbook
        workbook.save(excel_file)
        print(f"\nUpdated Template sheet with organizations: {organizations}")
        return True
        
    except Exception as e:
        print(f"Error adding template sheet: {str(e)}")
        return False

def create_server_template_from_excel(api_client, excel_file):
    """Create a server template in Intersight from Excel configuration"""
    try:
        # Read the template sheet
        template_df = pd.read_excel(excel_file, sheet_name='Template')
        template_data = dict(zip(template_df['Configuration Type'], template_df['Value']))
        
        # Get organization MOID
        org_moid = get_org_moid(api_client)
        
        # Create Server Profile Template API instance
        api_instance = server_api.ServerApi(api_client)
        
        # Get policy MOIDs
        bios_policy_moid = get_policy_moid(api_client, 'bios/Policies', template_data.get('BIOS Policy', ''))
        lan_policy_moid = get_policy_moid(api_client, 'vnic/LanConnectivityPolicies', template_data.get('LAN Connectivity Policy', ''))
        
        # Create the template body
        template_body = {
            'Name': template_data.get('Template Name', ''),
            'Description': template_data.get('Template Description', ''),
            'Organization': {
                'ObjectType': 'organization.Organization',
                'Moid': org_moid
            },
            'TargetPlatform': 'FIAttached',
            'PolicyBucket': []
        }
        
        # Add policies to the template
        if bios_policy_moid:
            template_body['PolicyBucket'].append({
                'ObjectType': 'bios.Policy',
                'Moid': bios_policy_moid
            })
            
        if lan_policy_moid:
            template_body['PolicyBucket'].append({
                'ObjectType': 'vnic.LanConnectivityPolicy',
                'Moid': lan_policy_moid
            })
            
        # Create the template
        print(f"\nCreating server template '{template_data.get('Template Name', '')}'...")
        api_instance.create_server_profile_template(template_body)
        print(f"Successfully created server template: {template_data.get('Template Name', '')}")
        return True
        
    except Exception as e:
        print(f"Error creating server template: {str(e)}")
        return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Create and push Intersight Foundation configuration')
    parser.add_argument('--action', choices=['push', 'template', 'both'], required=True,
                      help='Action to perform: push (create policies), template (apply template), or both')
    parser.add_argument('--file', required=True, help='Path to the Excel file')
    args = parser.parse_args()

    # Get API client
    api_client = get_api_client()
    if not api_client:
        sys.exit(1)

    if args.action in ['push', 'both']:
        print(f"Pushing configuration from {args.file} to Intersight...")
        create_and_push_configuration(api_client, args.file)

    if args.action in ['template', 'both']:
        print(f"\nApplying template configuration from {args.file}...")
        add_template_sheet(args.file, api_client)
        
        # Create the server template in Intersight
        create_server_template_from_excel(api_client, args.file)
