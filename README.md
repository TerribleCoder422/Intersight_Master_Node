# Automate-Intersight

A Python project for automating Cisco Intersight operations, including UCS Server Profile template management via Excel.

## Overview

This project provides tools to interact with Cisco Intersight, including:

1. Authentication with Intersight API
2. Creating and managing UCS Server Profile templates
3. Excel-based template creation and management
4. Pre-configured templates for specific use cases (e.g., OpenShift masternodes)

## Features

- **Authentication**: Secure authentication with Intersight using API keys
- **Excel Template**: Create UCS Server Profile templates using Excel
- **Profile Management**: View and manage UCS Server Profiles
- **Pre-configured Templates**: Ready-to-use templates for common deployment scenarios
- Create Excel templates for UCS Server Profile templates
- Create pre-configured templates for specific use cases (e.g., OpenShift masternodes)
- Update existing Excel files with pre-configured settings
- **NEW: Update Excel templates with dynamic dropdowns populated from your Intersight instance**
- Import templates from Excel to Intersight

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Configure your Intersight API credentials:
   - API Key ID
   - Private Key

## Usage

### Authentication Test

Run the authentication test script to verify connectivity:

```bash
python3 src/main.py auth
```

### Excel Template Creation

Generate the Excel template for UCS Server Profile templates:

```bash
python3 src/main.py create-template
```

This will create an Excel file in the `output` directory with the following tabs:
- **Template**: For creating UCS Server Profile templates
- **Profiles**: For viewing and managing UCS Server Profiles

### Create a blank Excel template

```bash
python3 src/main.py create-template --output template.xlsx
```

### Create a pre-configured masternode template

```bash
python3 src/main.py create-masternode --output masternode_template.xlsx
```

### Pre-configured Templates

Generate pre-configured templates for specific use cases:

#### OpenShift Masternode Template

```bash
python3 src/main.py create-masternode
```

This creates a template with the following OpenShift masternode configuration:

- **Boot Order**: PXE → Local Disk
- **vNICs**:
  - 2 vNICs (one per UCS fabric interconnect)
  - VLANs for OpenShift Management, API, and etcd traffic
  - MAC Address Pool assigned
- **vHBAs** (If using SAN storage):
  - 2 vHBAs (one per fabric)
  - WWPN pool assigned
- **BIOS Policy**:
  - Performance Mode enabled
  - CPU C-States disabled (for stability)
  - VT-x enabled for virtualization
- **Boot Policy**:
  - PXE for initial deployment (Ignition)
  - Local disk boot after installation
- **vMedia Policy**:
  - Optional OpenShift boot ISO for manual deployment
- **QoS Policy**:
  - High priority for control plane traffic
- **Storage**:
  - 1 x SSD/NVMe for OS
  - 1 x Disk for etcd data (if running etcd on nodes)

### Update Existing Excel File with Masternode Configuration

You can also update an existing Excel file with the masternode configuration:

```bash
python3 src/main.py update-with-masternode --input path/to/your/excel/file.xlsx
```

This will add the masternode configuration to the specified Excel file, including:
- Setting the template name to "MasterNode-Template"
- Configuring all policies for OpenShift deployment
- Setting up vNICs and vHBAs
- Adding a Documentation tab with detailed specifications

### Update Excel with dynamic dropdowns from Intersight

```bash
python3 src/main.py update-with-intersight-data --input path/to/your/excel/file.xlsx
```

This command will:
1. Connect to your Intersight instance
2. Retrieve all available organizations and policies
3. Update your Excel template with dropdowns containing these options
4. Allow you to select from valid options when configuring templates

### Excel to Intersight Integration

Use the Excel to Intersight integration script to:
- Read template data from Excel
- Create templates in Intersight
- Update the Excel with data from Intersight

```bash
python3 src/main.py import-template --input output/masternode_template.xlsx
```

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
- Template
- Target Server
- Status
- Organization
- Created Date
- Actions

### Documentation Tab

For pre-configured templates, a Documentation tab is included with detailed specifications and configuration notes.

## Requirements

- Python 3.6+
- pandas
- openpyxl
- intersight SDK

## License

This project is licensed under the MIT License - see the LICENSE file for details.
