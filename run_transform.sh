#!/bin/bash

# Existing Comps Transformation Runner
# This script activates the virtual environment and runs the transformation

echo "Activating virtual environment..."
source .venv/bin/activate

echo "Running transformation script..."
python transform_comps.py "$@"

echo "Done!"

