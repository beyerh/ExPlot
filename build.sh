#!/bin/bash

# ExPlot build script with all required packages

# Configuration
APP_NAME="ExPlot"
VERSION="0.7.2"
BUILD_DIR="build"

# Create a clean build directory
mkdir -p build

# Get Python version for reference
PYTHON_VERSION=$(python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo -e "Building ${APP_NAME} v${VERSION} with Python ${PYTHON_VERSION}..."

# Run Nuitka to build the standalone application
echo -e "Building with Nuitka..."

# If needed include:  --macos-target-arch=x86_64 \
python -m nuitka \
  --standalone \
  --macos-create-app-bundle \
  --macos-app-name="${APP_NAME}" \
  --macos-app-version="${VERSION}" \
  --macos-app-icon=explot.icns \
  --macos-signed-app-name="com.${APP_NAME}.app" \
  --enable-plugin=tk-inter \
  --include-package=matplotlib.backends.backend_pdf \
  --output-filename="${APP_NAME}" \
  --output-dir="${BUILD_DIR}" \
    launch.py

# Rename the app bundle from launch.app to ExPlot.app
mv "${BUILD_DIR}/launch.app" "${BUILD_DIR}/${APP_NAME}.app"

# Copy additional files
echo -e "Copying additional files..."
cp -r pingouin "${BUILD_DIR}/${APP_NAME}.app/Contents/MacOS/"
cp -r example_data.xlsx "${BUILD_DIR}/${APP_NAME}.app/Contents/MacOS/"

echo -e "Build completed successfully!"