#!/usr/bin/env python3
"""
Standalone module for mapping servers to resource groups in Intersight
Uses the combined_selector technique from the sample script
"""
import logging
import json
import os
import traceback
from dotenv import load_dotenv

# Place script specific intersight api imports here
from intersight.api import resource_api
from intersight.api import asset_api
from intersight.api_client import ApiClient
from intersight.configuration import Configuration
import intersight

# Configure logging
FORMAT = '%(asctime)-15s [%(levelname)s] [%(filename)s:%(lineno)s] %(message)s'
logging.basicConfig(format=FORMAT, level=logging.INFO)
logger = logging.getLogger('intersight_rg_mapper')

def get_api_client():
    """Get Intersight API client with proper authentication"""
    try:
        # Load environment variables from .env file
        load_dotenv()
        
        # Get API key details from environment variables
        api_key_id = os.getenv('INTERSIGHT_API_KEY_ID')
        api_key_file = os.getenv('INTERSIGHT_PRIVATE_KEY_FILE', './SecretKey.txt')
        
        if not api_key_id or not os.path.exists(api_key_file):
            logger.error("Error: API key configuration not found")
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
        logger.error(f"Error getting API client: {str(e)}")
        return None

def map_servers_to_resource_groups(server_details):
    """
    Maps servers to resource groups using Intersight API
    
    Key improvements in this version:
    1. Uses multiple methods to find resource group memberships
    2. Tries direct MOID matching in addition to hostname/serial
    3. Falls back to querying resource group memberships directly
    4. Enhanced error handling and debug logging
    """
    """
    Maps servers to resource groups using Intersight API
    
    Args:
        server_details: Dictionary of server details with serial as key
        
    Returns:
        Tuple of (server_resource_groups, success_flag)
        - server_resource_groups: Dictionary with resource group names as keys and server lists as values
        - success_flag: Boolean indicating if real mappings were found
    """
    # Get the API client
    client = get_api_client()
    if not client:
        logger.error("Failed to get API client")
        return {}, False
    
    # Initialize API instances
    resource_instance = resource_api.ResourceApi(client)
    asset_instance = asset_api.AssetApi(client)
    
    # Dictionary to hold resource group to server mappings
    server_resource_groups = {}
    real_mappings_found = False
    
    try:
        # Get all resource groups (excluding License groups)
        logger.info("Querying Intersight API for resource groups...")
        query_filter = "not startsWith(Name,'License')"
        api_response = resource_instance.get_resource_group_list(filter=query_filter)
        logger.info(f"Found {len(api_response.results)} resource groups")
        
        # Initialize empty lists for each resource group
        for result in api_response.results:
            rg_name = result.name
            server_resource_groups[rg_name] = []
            
        # For each resource group, find its server members using combined_selector
        for result in api_response.results:
            rg_name = result.name
            logger.info(f"Checking resource group: {rg_name}")
            
            if hasattr(result, 'per_type_combined_selector') and result.per_type_combined_selector:
                # This is the key technique from the sample script
                combined_selector = result.per_type_combined_selector[0].combined_selector
                logger.info(f"Using combined selector: {combined_selector}")
                
                try:
                    # Query devices using the combined selector
                    reg_response = asset_instance.get_asset_device_registration_list(filter=combined_selector)
                    servers_found = 0
                    
                    # Process each device in this resource group
                    for reg in reg_response.results:
                        # Get the device MOID which is crucial for matching
                        device_moid = getattr(reg, 'moid', None)
                        if not device_moid:
                            continue
                            
                        # Output device details for debugging
                        raw_hostname = getattr(reg, 'device_hostname', 'Unknown')
                        raw_serial = getattr(reg, 'serial', 'Unknown')
                        
                        # Format hostname properly for display
                        hostname_display = raw_hostname
                        if isinstance(raw_hostname, list) and raw_hostname:
                            hostname_display = raw_hostname[0]
                            
                        # Format serial properly for display
                        serial_display = raw_serial
                        if isinstance(raw_serial, list) and raw_serial:
                            serial_display = raw_serial[0]
                            
                        logger.info(f"    Found device in resource group: {hostname_display} / {serial_display} / MOID: {device_moid}")
                        
                        # Try to match by hostname, serial, or MOID to our server list
                        for serial, server in server_details.items():
                            # Various ways to match servers
                            hostname_match = False
                            serial_match = False
                            moid_match = False
                            
                            # Check hostname match (case insensitive)
                            if hasattr(reg, 'device_hostname') and reg.device_hostname:
                                # Handle both string and list formats
                                device_hostname = reg.device_hostname
                                if isinstance(device_hostname, list) and device_hostname:
                                    device_hostname = device_hostname[0]  # Take first item if it's a list
                                
                                if isinstance(device_hostname, str) and isinstance(server['name'], str):
                                    hostname_match = server['name'].lower() == device_hostname.lower()
                                    logger.info(f"       Hostname comparison: '{server['name'].lower()}' vs '{device_hostname.lower()}' = {hostname_match}")
                            
                            # Check serial match (exact match)
                            if hasattr(reg, 'serial') and reg.serial:
                                # Handle both string and list formats
                                device_serial = reg.serial
                                if isinstance(device_serial, list) and device_serial:
                                    device_serial = device_serial[0]  # Take first item if it's a list
                                
                                if isinstance(device_serial, str):
                                    serial_match = server.get('serial') == device_serial
                                    logger.info(f"       Serial comparison: '{server.get('serial')}' vs '{device_serial}' = {serial_match}")
                            
                            # Check MOID match if server has MOID
                            if server.get('moid') and device_moid:
                                moid_match = server['moid'] == device_moid
                            
                            # Match if any of our matching criteria are met
                            if hostname_match or serial_match or moid_match:
                                # Add to resource group mapping
                                if 'resource_groups' not in server:
                                    server['resource_groups'] = []
                                    
                                if rg_name not in server['resource_groups']:
                                    server['resource_groups'].append(rg_name)
                                    server_entry = f"{server['serial']} | {server['name']}"
                                    server_resource_groups[rg_name].append(server_entry)
                                    real_mappings_found = True
                                    servers_found += 1
                                    logger.info(f"✓ Mapped server {server['name']} to resource group {rg_name}")
                    
                    logger.info(f"Found {servers_found} servers in resource group {rg_name}")
                    
                except Exception as e:
                    logger.error(f"Error querying devices for resource group {rg_name}: {str(e)}")
            else:
                logger.warning(f"No per_type_combined_selector found for resource group {rg_name}")
                
        # If no mappings found via combined_selector approach, try an alternative approach 
        # directly querying the 'resource/GroupMembers' endpoint
        if not real_mappings_found:
            logger.info("\nNo mappings found via combined_selector approach, trying direct group membership query...")
            
            try:
                # Get all resource groups first
                resource_groups_list = resource_instance.get_resource_group_list()
                
                # For each group, try to get its direct members
                for group in resource_groups_list.results:
                    rg_name = group.name
                    rg_moid = group.moid
                    
                    # Skip if not in our filtered list
                    if rg_name not in server_resource_groups.keys():
                        continue
                    
                    logger.info(f"Checking direct members for resource group: {rg_name} (MOID: {rg_moid})")
                    
                    try:
                        # Query the resource group members directly
                        endpoint = f"/api/v1/resource/GroupMembers?$filter=Resource.ObjectType eq 'compute.RackUnit' and Group.Moid eq '{rg_moid}'"
                        members_response = client.call_api(endpoint, 'GET')
                        
                        if members_response.status_code == 200:
                            members_data = members_response.json()
                            logger.info(f"Found {len(members_data.get('Results', []))} compute rack units in group {rg_name}")
                            
                            # Process each member
                            servers_found = 0
                            for member in members_data.get('Results', []):
                                resource = member.get('Resource', {})
                                server_moid = resource.get('Moid')
                                
                                if server_moid:
                                    # Try to match this MOID to our server list
                                    for serial, server in server_details.items():
                                        if server.get('moid') == server_moid:
                                            # Add to resource group mapping
                                            if rg_name not in server['resource_groups']:
                                                server['resource_groups'].append(rg_name)
                                                server_entry = f"{server['serial']} | {server['name']}"
                                                server_resource_groups[rg_name].append(server_entry)
                                                real_mappings_found = True
                                                servers_found += 1
                                                logger.info(f"✓ Mapped server {server['name']} to resource group {rg_name} (direct method)")
                            
                            if servers_found > 0:
                                logger.info(f"Found {servers_found} servers in resource group {rg_name} via direct method")
                        else:
                            logger.warning(f"Failed to get members for resource group {rg_name}: {members_response.status_code}")
                    
                    except Exception as e:
                        logger.error(f"Error querying direct membership for resource group {rg_name}: {str(e)}")
            
            except Exception as e:
                logger.error(f"Error in direct membership approach: {str(e)}")
        
        # If no mappings were found, notify the user
        if not real_mappings_found:
            logger.warning("⚠ WARNING: No server-to-resource-group mappings found!")
        
        # Overall status
        if real_mappings_found:
            logger.info("Successfully mapped servers to resource groups using Intersight SDK!")
        else:
            logger.warning("No server-to-resource-group mappings found via API")
            
    except Exception as e:
        logger.error(f"Error mapping servers to resource groups: {str(e)}")
        traceback.print_exc()
        
    return server_resource_groups, real_mappings_found

def fallback_to_mapping_file(server_details, server_resource_groups):
    """
    Uses a fallback mapping file when API-based mapping fails
    
    Args:
        server_details: Dictionary of server details with serial as key
        server_resource_groups: Dictionary with resource group names as keys and server lists as values
        
    Returns:
        Boolean indicating if mappings were found in the file
    """
    real_mappings_found = False
    rg_mapping_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resource_group_mappings.json')
    
    try:
        if os.path.exists(rg_mapping_file):
            logger.info(f"Found mapping file: {rg_mapping_file}")
            with open(rg_mapping_file, 'r') as f:
                mappings = json.load(f)
            
            # Process the mappings from the file
            logger.info("Loading mappings from file...")
            for rg_name, server_list in mappings.items():
                # Skip if resource group not in our list
                if rg_name not in server_resource_groups.keys():
                    logger.warning(f"Skipping unknown resource group in mapping file: {rg_name}")
                    continue
                
                logger.info(f"Processing fallback mappings for: {rg_name}")
                servers_found = 0
                
                for server_name in server_list:
                    # Find this server in our list
                    for serial, server in server_details.items():
                        if server['name'].lower() == server_name.lower():
                            # Add to resource group mapping
                            if 'resource_groups' not in server:
                                server['resource_groups'] = []
                                
                            if rg_name not in server['resource_groups']:
                                server['resource_groups'].append(rg_name)
                                server_entry = f"{server['serial']} | {server['name']}"
                                server_resource_groups[rg_name].append(server_entry)
                                real_mappings_found = True
                                servers_found += 1
                                logger.info(f"✓ Mapped server {server['name']} to resource group {rg_name} (from file)")
                
                if servers_found > 0:
                    logger.info(f"Found {servers_found} servers for resource group {rg_name} in mapping file")
            
            if real_mappings_found:
                logger.info("Successfully mapped servers to resource groups using fallback file!")
        else:
            logger.warning(f"No mapping file found at: {rg_mapping_file}")
            logger.info("Creating template mapping file for future use...")
            
            # Create a template mapping file with all resource groups
            template_mappings = {rg_name: [] for rg_name in server_resource_groups.keys()}
            
            with open(rg_mapping_file, 'w') as f:
                json.dump(template_mappings, f, indent=2)
            
            logger.info(f"Created template mapping file: {rg_mapping_file}")
            logger.info("You can edit this file to add server-to-resource-group mappings")
    except Exception as e:
        logger.error(f"Error processing mapping file: {str(e)}")
    
    return real_mappings_found

if __name__ == "__main__":
    # This is a demo to show how this module works
    # It's meant to be imported by update_intersight_data.py
    
    # Sample server data
    sample_servers = {
        "FCH123456789": {
            "name": "C220M5-Hosting-Server1",
            "serial": "FCH123456789",
            "moid": "abcdef1234567890"
        },
        "FCH123456788": {
            "name": "C220M5-Hosting-Server2",
            "serial": "FCH123456788",
            "moid": "abcdef1234567891"
        }
    }
    
    # Map servers to resource groups
    server_resource_groups, api_success = map_servers_to_resource_groups(sample_servers)
    
    # If API mapping fails, try fallback file
    if not api_success:
        fallback_success = fallback_to_mapping_file(sample_servers, server_resource_groups)
        if not fallback_success:
            logger.warning("⚠ WARNING: No server-to-resource-group mappings found!")
    
    # Print unassigned servers
    unassigned_servers = [server['name'] for server in sample_servers.values() 
                         if 'resource_groups' not in server or not server['resource_groups']]
    if unassigned_servers:
        logger.info(f"{len(unassigned_servers)} servers are not assigned to any resource group:")
        for name in unassigned_servers:
            logger.info(f"- {name}")
    
    # Print resource group to server mapping summary
    logger.info("Resource Group to Server Mapping Summary:")
    for rg_name, servers_in_group in server_resource_groups.items():
        logger.info(f"{rg_name}: {len(servers_in_group)} servers")
