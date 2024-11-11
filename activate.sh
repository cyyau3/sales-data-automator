#!/bin/zsh

# Get the directory where the script is located
SCRIPT_DIR=${0:a:h}

# Change to the project directory
cd "$SCRIPT_DIR"

# Check if virtualenv is installed
if ! command -v virtualenv &> /dev/null; then
    echo "Error: virtualenv not found. Installing..."
    pip3 install virtualenv
fi

# Verify venv activation
if [ -z "${VIRTUAL_ENV:-}" ]; then
    echo "Error: Virtual environment activation failed"
    exit 1
fi

# Check if venv exists, if not create it with Homebrew Python
if [ ! -d "venv" ]; then
    echo "Creating new virtual environment..."
    /opt/homebrew/bin/python3 -m venv venv
fi

# Activate the virtual environment
source venv/bin/activate

# Verify Python version and location
echo "Virtual environment activated!"
echo "Current directory: $(pwd)"
echo "Using Python: $(which python3)"
echo "Python version: $(python3 -V)"