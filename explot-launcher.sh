#!/bin/bash
# ExPlot Launcher Script
# This script activates the Python virtual environment and launches ExPlot

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Change to the ExPlot directory
cd "$SCRIPT_DIR"

# Activate the virtual environment
source .venv/bin/activate

# Launch ExPlot
python launch.py

# Deactivate virtual environment when ExPlot closes
deactivate
