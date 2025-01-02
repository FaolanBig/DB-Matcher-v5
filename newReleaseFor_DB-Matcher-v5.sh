#!/bin/bash

# Variables
REPO_URL="https://github.com/FaolanBig/DB-Matcher-v5.git"
REPO_NAME="DB-Matcher-v5"
BUILD_DIR="build_output"
LINUX_FILENAME="DB-Matcher-v5_linux_x86"
WINDOWS_FILENAME="DB-Matcher-v5_win_x86"
RELEASE_NAME="automatically-generated-binaries"

# Clone repository
echo "Cloning repository..."
git clone "$REPO_URL" "$REPO_NAME" || { echo "Error: git clone failed"; exit 1; }
cd "$REPO_NAME" || { echo "Error: repository not found"; exit 1; }

# Clean build directory
rm -rf "$BUILD_DIR"
mkdir "$BUILD_DIR"

# Build for Linux
echo "Building for Linux..."
dotnet publish -c Release -r linux-x64 --self-contained true /p:PublishSingleFile=true -o "$BUILD_DIR/linux" || { echo "Error: Linux build failed"; exit 1; }
cd "$BUILD_DIR/linux" || exit
zip "../${LINUX_FILENAME}.zip" *
tar -czf "../${LINUX_FILENAME}.tar.gz" *
cd ../../..

# Build for Windows
echo "Building for Windows..."
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true -o "$BUILD_DIR/windows" || { echo "Error: Windows build failed"; exit 1; }
cd "$BUILD_DIR/windows" || exit
zip "../${WINDOWS_FILENAME}.zip" *
tar -czf "../${WINDOWS_FILENAME}.tar.gz" *
cd ../../..

# Create GitHub Release
echo "Creating GitHub release..."
LATEST_TAG=$(git tag | tail -n 1)
gh release create "$LATEST_TAG" \
  "$BUILD_DIR/${LINUX_FILENAME}.zip" \
  "$BUILD_DIR/${LINUX_FILENAME}.tar.gz" \
  "$BUILD_DIR/${WINDOWS_FILENAME}.zip" \
  "$BUILD_DIR/${WINDOWS_FILENAME}.tar.gz" \
  --title "$LATEST_TAG" \
  --notes "$RELEASE_NAME" || { echo "Error: GitHub release creation failed"; exit 1; }

# Cleanup
echo "Cleanup..."
rm -rf "$BUILD_DIR"
rm -rf "$REPO_NAME"

echo "Done"