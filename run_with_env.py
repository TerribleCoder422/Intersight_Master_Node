#!/usr/bin/env python3
"""
Wrapper script to load .env file and run create_intersight_foundation.py
"""
import os
import sys
from dotenv import load_dotenv
import subprocess

# Load environment variables from .env file
load_dotenv()

# Get command line arguments
args = sys.argv[1:]

# Construct the command to run
cmd = ["python3", "create_intersight_foundation.py"] + args

# Run the script with the loaded environment variables
print("Running:", " ".join(cmd))
subprocess.run(cmd, env=os.environ)
