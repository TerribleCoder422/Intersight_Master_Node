# Intersight Automation CLI Runbook

This runbook provides step-by-step instructions for using the Intersight Automation tool via CLI commands.

## Prerequisites

Before you begin, ensure you have:

1. Python 3.6 or later installed
2. Required Python packages installed (`pip install -r requirements.txt`)
3. Intersight API credentials configured:
   - API Key ID in environment variable `INTERSIGHT_API_KEY_ID`
   - Path to private key file in environment variable `INTERSIGHT_PRIVATE_KEY_FILE`

## Environment Setup

```bash
# Set environment variables
export INTERSIGHT_API_KEY_ID="your-api-key-id"
export INTERSIGHT_PRIVATE_KEY_FILE="/path/to/your/secret.key"
```

## Workflow Steps

### 1. Create the Standard Template

This creates the standard AI_POD_master_Template.xlsx file:

```bash
python3 create_intersight_foundation.py --action setup --file dummy.xlsx
```

The output will be saved as: `output/AI_POD_master_Template.xlsx`

### 2. Retrieve Intersight Data

This populates the template with current data from your Intersight instance:

```bash
python3 create_intersight_foundation.py --action get-info --file output/AI_POD_master_Template.xlsx
```

### 3. Customize the Excel Template

At this point, you should:
1. Open the Excel file
2. Modify settings as needed
3. Save the file

### 4. Push the Configuration to Intersight

This creates all the pools, policies, templates and profiles defined in your Excel file:

```bash
python3 create_intersight_foundation.py --action push --file output/AI_POD_master_Template.xlsx
```

## Additional Commands

### Update Server Information Only

```bash
python3 create_intersight_foundation.py --action update-servers --file output/AI_POD_master_Template.xlsx
```

### Create Template Only

```bash
python3 create_intersight_foundation.py --action template --file output/AI_POD_master_Template.xlsx
```

### Create Profiles Only

```bash
python3 create_intersight_foundation.py --action profiles --file output/AI_POD_master_Template.xlsx
```

### Run Complete Workflow

This runs the entire workflow (templates + profiles):

```bash
python3 create_intersight_foundation.py --action all --file output/AI_POD_master_Template.xlsx
```

## Troubleshooting

### Connection Issues

If you encounter connection issues:
1. Verify your API key ID is correct
2. Ensure your private key file is valid and readable
3. Check your internet connection
4. Verify your Intersight account has the necessary privileges

### Template Processing Issues

If template processing fails:
1. Verify all required fields are filled out in the Excel file
2. Check for typos in policy and template names
3. Ensure organization names match exactly with those in Intersight
4. Verify the Excel file format hasn't been corrupted

## Example Workflow

Here's a complete example workflow:

```bash
# 1. Set up environment variables
export INTERSIGHT_API_KEY_ID="your-api-key-id"
export INTERSIGHT_PRIVATE_KEY_FILE="/path/to/your/secret.key"

# 2. Create standard template
python3 create_intersight_foundation.py --action setup --file dummy.xlsx

# 3. Update with Intersight data
python3 create_intersight_foundation.py --action get-info --file output/AI_POD_master_Template.xlsx

# 4. [MANUAL STEP] Modify the Excel file as needed

# 5. Push configuration to Intersight
python3 create_intersight_foundation.py --action push --file output/AI_POD_master_Template.xlsx
```
