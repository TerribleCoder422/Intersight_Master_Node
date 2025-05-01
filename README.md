# Automate-Intersight

A Python project for automating Cisco Intersight operations, including pools, policies, templates, and server profiles via standardized Excel templates with dynamic resource group filtering.

## Overview

This project provides tools to interact with Cisco Intersight, offering a full-cycle workflow to:

1. Create an Excel template with pre-populated data
2. Retrieve and populate the template with current Intersight data
3. Allow customization of the template
4. Push the configuration to Intersight, creating all necessary components

## Features

- **Authentication**: Secure authentication with Intersight using API keys
- **Standardized Templates**: Automatically generate Excel templates with consistent naming
- **Dynamic Resource Group Filtering**: Server dropdowns that automatically filter based on resource group selection
- **Excel-Compatible Named Ranges**: Ensures Excel compatibility while maintaining dynamic filtering capabilities
- **Idempotent Operations**: Pools, policies and templates are created only if they don't already exist
- **Pool Management**: Create and manage MAC, UUID, and other pools
- **Policy Management**: Create and manage BIOS, Boot, vNIC, Storage and other policies
- **Template Management**: Create and manage UCS Server Profile templates
- **Profile Management**: Create profiles from templates and assign servers WITHOUT automatic deployment
- **Server Assignment**: Automatically assign servers to profiles with clear verification messages
- **Progress Tracking**: Visual progress indicators and color-coded outputs
- **Smart Template Matching**: Flexible name matching to find templates in Intersight
- **Dynamic Dropdowns**: Automatically populate Excel dropdowns with data from your Intersight instance
- **Data Validation**: Validate required fields before pushing to Intersight
- **Error Handling**: Comprehensive error handling with clear messages
- **Multi-Organization Support**: Work with resources across different organizations

## Setup

### Prerequisites

1. Python 3.7 or later installed on your system
2. Access to Cisco Intersight with API keys
3. Basic knowledge of Intersight resources (organizations, pools, policies)

### Dependencies

This tool requires the following key Python packages:

- **intersight**: Cisco Intersight API client
- **pandas**: For data manipulation and Excel handling
- **openpyxl**: For Excel file creation and modification
- **requests**: For API communication
- **cryptography**: For Intersight authentication
- **tqdm**: For progress bars
- **colorama/termcolor**: For colored console output

All dependencies are listed in the `requirements.txt` file and will be installed automatically with the installation commands below.

### Installation for Windows PC

1. Clone this repository to your PC:
   ```bash
   git clone https://github.com/TerribleCoder422/Intersight_Master_Node.git
   cd Intersight_Master_Node
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up your Intersight API credentials:
   - Create a `.env` file in the root directory by copying `.env.example`
   - In Intersight, generate an API key and download the private key file
   - Update the `.env` file with your API key ID and path to your private key file

   Example `.env` file:
   ```
   INTERSIGHT_API_KEY_ID=12345abcde12345abcde
   INTERSIGHT_PRIVATE_KEY_FILE=.\secret.key
   INTERSIGHT_BASE_URL=https://intersight.com
   ```

4. Verify connectivity by running the following commands:
   ```bash
   python create_standard_excel.py
   python update_intersight_data.py output/Create_Intersight_Template.xlsx
   ```
   
   If successful, you should see a new file created at `output/Create_Intersight_Template.xlsx` with data from your Intersight instance.

### Installation for macOS

1. Clone this repository to your Mac:
   ```bash
   git clone https://github.com/TerribleCoder422/Intersight_Master_Node.git
   cd Intersight_Master_Node
   ```

2. Install dependencies:
   ```bash
   pip3 install -r requirements.txt
   ```

3. Set up your Intersight API credentials:
   - Create a `.env` file in the root directory by copying `.env.example`
   - In Intersight, generate an API key and download the private key file
   - Update the `.env` file with your API key ID and path to your private key file

   Example `.env` file:
   ```
   INTERSIGHT_API_KEY_ID=12345abcde12345abcde
   INTERSIGHT_PRIVATE_KEY_FILE=./secret.key
   INTERSIGHT_BASE_URL=https://intersight.com
   ```

4. Verify connectivity by running the following commands:
   ```bash
   python3 create_standard_excel.py
   python3 update_intersight_data.py output/Create_Intersight_Template.xlsx
   ```
   
   If successful, you should see a new file created at `output/Create_Intersight_Template.xlsx` with data from your Intersight instance.

## Usage

### Complete Workflow

The recommended workflow consists of the following steps:

1. Create a standardized template
2. Populate it with Intersight data (including dynamic resource group filtering)
3. Modify the template as needed
4. Push the configuration to Intersight

### Setup Excel Template

Generate and set up the standardized Excel template for Intersight configurations:

```bash
python3 create_standard_excel.py
```

This creates a standard template at `output/Create_Intersight_Template.xlsx` with the basic structure needed.

The standardized template includes:

- **Pool Management**:
  - MAC Pools for both fabrics
  - UUID Pool for server identification

- **Policy Management**:
  - BIOS Policy for performance optimization
  - vNIC Policies for network connectivity
  - QoS Policy for traffic prioritization
  - Storage Policy for disk configuration

- **Template Configuration**:
  - Default organization setting
  - Profile template definitions
  - Server assignment options

### Update Excel with Intersight Data

Fetch data from Intersight and update the Excel template with dynamic dropdowns and resource group filtering:

```bash
python3 update_intersight_data.py output/Create_Intersight_Template.xlsx
```

This command:
- Retrieves organizations, resource groups, and servers from Intersight
- Creates dynamic server dropdowns that filter based on resource group selection
- Maps servers to their respective resource groups
- Creates Excel-compatible named ranges for better compatibility
- Updates all available dropdown options with current Intersight data

### Push Configuration to Intersight

Create pools, policies, templates and profiles in Intersight based on the Excel template:

```bash
python3 push_intersight_template.py --action all --file output/Create_Intersight_Template.xlsx
```

You can also run specific actions:
- `--action push`: Only create pools and policies
- `--action template`: Only create server templates
- `--action profiles`: Only create server profiles

### Update Server Information

Refresh the list of available servers in your Excel template:

```bash
python3 update_intersight_data.py output/Create_Intersight_Template.xlsx
```

## Working with the Excel Template

After generating the template, you can manually edit it to customize your configurations before pushing to Intersight.

## Excel Template Structure

### Template Tab

The Template tab allows you to define UCS Server Profile templates with:

1. **Basic Information**:
   - Name
   - Description
   - Organization
   - Tags
   - Target Platform

2. **Policies**:
   - BIOS Policy
   - Boot Order Policy
   - Virtual Media Policy
   - And many more...

3. **LAN Connectivity**:
   - vNIC configuration
   - MAC addresses
   - VLAN IDs

4. **SAN Connectivity**:
   - vHBA configuration
   - WWPN addresses
   - VSAN IDs

### Profiles Tab

The Profiles Tab displays UCS Server Profiles with:
- Profile Name
- Description
- Organization
- Resource Group
- Template
- Target Server (dynamically filtered by Resource Group)
- Status
- Deploy Option

### ServerMap Sheet

The ServerMap sheet is a hidden sheet that maps servers to their resource groups. This enables the dynamic filtering functionality:

- Each resource group has its own named range of servers
- The named ranges follow the pattern: `ResourceGroupName_Servers`
- These named ranges are referenced via INDIRECT formulas in the Profiles sheet
- When you select a resource group in a row, the Server dropdown for that row automatically filters to only show servers belonging to that resource group

### Documentation Tab

For pre-configured templates, a Documentation tab is included with detailed specifications and configuration notes.

### Excel Compatibility Improvements

This solution includes several enhancements to ensure compatibility with Excel while maintaining dynamic filtering functionality:

- Uses Excel-compatible named ranges with proper formatting
- Implements CONCATENATE() instead of the & operator in Excel formulas for better compatibility
- Ensures named ranges comply with Excel's 31-character limitation
- Creates proper scope for named ranges to avoid repair warnings
- Handles empty resource groups gracefully to prevent Excel errors

These improvements help eliminate Excel repair warnings while preserving the dynamic resource group filtering capability.

## Requirements

- Python 3.6+
- pandas
- openpyxl
- intersight SDK

## License

This project is licensed under the MIT License - see the LICENSE file for details.
