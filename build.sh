#!/bin/bash

# ExPlot build script with all required packages
# Version 0.6.4

# Create a clean build directory
mkdir -p build

# Run Nuitka to build the standalone application
python -m nuitka \
  --standalone \
  --macos-create-app-bundle \
  --macos-app-name="ExPlot" \
  --macos-app-icon=explot.icns \
  --enable-plugin=tk-inter \
  --include-package=matplotlib.backends.backend_pdf \
  --output-dir=build \
  explot.py

echo "Build complete. Application is in the build directory."
echo "Manually copy pingouin and example_data.xlsx to build/ExPlot.app/Contents/MacOS."

