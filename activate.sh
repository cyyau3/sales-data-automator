#!/bin/zsh
# Get the directory where the script is located
SCRIPT_DIR=$(dirname "$0")
# Change to the project directory
cd "$SCRIPT_DIR"
# Activate the virtual environment
source venv/bin/activate

# Print confirmation message
echo "Virtual environment activated!"
echo "Current directory: $(pwd)"