"""
Utility functions for Intersight automation script
"""

import os
import time
import logging
import functools
import pandas as pd
from tqdm import tqdm
from colorama import Fore, Style, init
from datetime import datetime

# Initialize colorama for colored terminal output
init(autoreset=True)

# Set up logging
log_dir = "logs"
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"intersight_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Maximum number of retries for API calls
MAX_RETRIES = 3
# Delay between retries in seconds
RETRY_DELAY = 2


def retry_api_call(max_retries=MAX_RETRIES, delay=RETRY_DELAY):
    """Decorator to retry API calls with exponential backoff"""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            retries = 0
            current_delay = delay
            while retries < max_retries:
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    retries += 1
                    if retries >= max_retries:
                        logger.error(f"{Fore.RED}API call failed after {max_retries} attempts: {str(e)}{Style.RESET_ALL}")
                        raise
                    
                    logger.warning(f"{Fore.YELLOW}API call failed. Retrying in {current_delay}s... ({retries}/{max_retries}){Style.RESET_ALL}")
                    time.sleep(current_delay)
                    current_delay *= 1.5  # Exponential backoff
            return wrapper
        return wrapper
    return decorator


def validate_pools_data(pools_df):
    """Validate pools data before creating in Intersight"""
    invalid_pools = []
    
    for idx, row in pools_df.iterrows():
        pool_type = row.get('Pool Type')
        pool_name = row.get('Pool Name')
        
        # Check for missing required fields
        if not pool_type or pd.isna(pool_type):
            invalid_pools.append(f"Row {idx+2}: Missing Pool Type")
            continue
        
        if not pool_name or pd.isna(pool_name):
            invalid_pools.append(f"Row {idx+2}: Missing Pool Name")
            continue
            
        # Validate pool type specific fields
        if pool_type == 'MAC Pool':
            start_addr = row.get('Start Address')
            size = row.get('Size')
            
            if not start_addr or pd.isna(start_addr):
                invalid_pools.append(f"Row {idx+2}: Missing Start Address for MAC Pool '{pool_name}'")
            elif not isinstance(start_addr, str):
                invalid_pools.append(f"Row {idx+2}: Invalid Start Address format for MAC Pool '{pool_name}'")
            
            if not size or pd.isna(size):
                invalid_pools.append(f"Row {idx+2}: Missing Size for MAC Pool '{pool_name}'")
            elif not str(size).isdigit():
                invalid_pools.append(f"Row {idx+2}: Size must be a number for MAC Pool '{pool_name}'")
        
        elif pool_type == 'UUID Pool':
            prefix = row.get('Start Address')
            size = row.get('Size')
            
            if not prefix or pd.isna(prefix):
                invalid_pools.append(f"Row {idx+2}: Missing Prefix for UUID Pool '{pool_name}'")
            
            if not size or pd.isna(size):
                invalid_pools.append(f"Row {idx+2}: Missing Size for UUID Pool '{pool_name}'")
            elif not str(size).isdigit():
                invalid_pools.append(f"Row {idx+2}: Size must be a number for UUID Pool '{pool_name}'")
    
    return invalid_pools


def validate_policies_data(policies_df):
    """Validate policies data before creating in Intersight"""
    invalid_policies = []
    
    for idx, row in policies_df.iterrows():
        policy_type = row.get('Policy Type')
        policy_name = row.get('Name')
        
        # Check for missing required fields
        if not policy_type or pd.isna(policy_type):
            invalid_policies.append(f"Row {idx+2}: Missing Policy Type")
            continue
        
        if not policy_name or pd.isna(policy_name):
            invalid_policies.append(f"Row {idx+2}: Missing Policy Name")
            continue
            
        # Validate organization
        org_name = row.get('Organization')
        if not org_name or pd.isna(org_name):
            invalid_policies.append(f"Row {idx+2}: Missing Organization for Policy '{policy_name}'")
            
    return invalid_policies


def validate_templates_data(templates_df):
    """Validate templates data before creating in Intersight"""
    invalid_templates = []
    
    for idx, row in templates_df.iterrows():
        template_name = row.get('Template Name')
        
        # Check for missing required fields
        if not template_name or pd.isna(template_name):
            invalid_templates.append(f"Row {idx+2}: Missing Template Name")
            continue
            
        # Validate organization
        org_name = row.get('Organization')
        if not org_name or pd.isna(org_name):
            invalid_templates.append(f"Row {idx+2}: Missing Organization for Template '{template_name}'")
            
        # Validate target platform
        platform = row.get('Target Platform')
        if not platform or pd.isna(platform):
            invalid_templates.append(f"Row {idx+2}: Missing Target Platform for Template '{template_name}'")
        elif platform not in ['FIAttached', 'Standalone']:
            invalid_templates.append(f"Row {idx+2}: Invalid Target Platform '{platform}' for Template '{template_name}'. Must be 'FIAttached' or 'Standalone'")
    
    return invalid_templates


def validate_profiles_data(profiles_df):
    """Validate profiles data before creating in Intersight"""
    invalid_profiles = []
    
    for idx, row in profiles_df.iterrows():
        profile_name = row.get('Profile Name')
        
        # Check for missing required fields
        if not profile_name or pd.isna(profile_name):
            invalid_profiles.append(f"Row {idx+2}: Missing Profile Name")
            continue
            
        # Validate organization
        org_name = row.get('Organization')
        if not org_name or pd.isna(org_name):
            invalid_profiles.append(f"Row {idx+2}: Missing Organization for Profile '{profile_name}'")
            
        # Validate template name
        template_name = row.get('Template Name')
        if not template_name or pd.isna(template_name):
            invalid_profiles.append(f"Row {idx+2}: Missing Template Name for Profile '{profile_name}'")
    
    return invalid_profiles


def progress_bar(iterable, desc="Processing", total=None):
    """Create a progress bar for iterables"""
    return tqdm(
        iterable, 
        desc=desc, 
        total=total, 
        bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]"
    )


def print_success(message):
    """Print a success message with green color"""
    logger.info(f"{Fore.GREEN}{message}{Style.RESET_ALL}")


def print_warning(message):
    """Print a warning message with yellow color"""
    logger.warning(f"{Fore.YELLOW}{message}{Style.RESET_ALL}")


def print_error(message):
    """Print an error message with red color"""
    logger.error(f"{Fore.RED}{message}{Style.RESET_ALL}")


def print_info(message):
    """Print an info message with cyan color"""
    logger.info(f"{Fore.CYAN}{message}{Style.RESET_ALL}")


def print_summary(title, success_items, failed_items):
    """Print a summary of successful and failed items"""
    print_info(f"\n{'=' * 80}")
    print_info(f" {title} Summary ")
    print_info(f"{'=' * 80}")
    
    if success_items:
        print_success(f"\nSuccessfully processed {len(success_items)} items:")
        for item in success_items:
            print_success(f"  ✓ {item}")
    
    if failed_items:
        print_error(f"\nFailed to process {len(failed_items)} items:")
        for item in failed_items:
            print_error(f"  ✗ {item}")
            
    print_info(f"{'=' * 80}\n")
