#!/bin/bash
# Set up environment variables for Intersight API access

# Your Intersight API Key ID
export INTERSIGHT_API_KEY_ID="6148e05e7564612d33ec43fa/677c321075646132014cbd02/67c22a0f75646132019eff55"

# Path to your private key file
export INTERSIGHT_PRIVATE_KEY_FILE="./SecretKey.txt"

# Optional: Set Intersight base URL if different from default
# export INTERSIGHT_BASE_URL="https://intersight.com"

# Run the Intersight workflow with all actions
python3 create_intersight_foundation.py --action all --file output/Intersight_Foundation.xlsx

# Done
echo "Workflow completed"
