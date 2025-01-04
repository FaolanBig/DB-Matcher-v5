#!/bin/bash

# before running this script, make shure you are logged in with gh (github cli)
# you can see your gh status by running 'gh auth status'

# Variables
REPO_URL="https://github.com/FaolanBig/DB-Matcher-v5.git"
REPO_NAME="DB-Matcher-v5"
BUILD_DIR="build_output"
PROJECT_FILE="DB-Matcher-v5.csproj"
LINUX_FILENAME="DB-Matcher-v5_linux_x86"
WINDOWS_FILENAME="DB-Matcher-v5_win_x86"
RELEASE_NAME="automatically generated binaries with sha256 checksums"

# Clone repository
echo "cleaning..."
rm -rf $REPO_NAME
echo "cleaning successful"

echo "cloning repository..."
git clone "$REPO_URL" "$REPO_NAME" || { echo "Error: git clone failed"; exit 1; }
cd "$REPO_NAME" || { echo "Error: repository not found"; exit 1; }
echo "cloning successful"

# Clean build directory
echo "cleaning build directory"
rm -rf "../$BUILD_DIR"
echo "cleaning successful"
echo "creating build directory"
mkdir "../$BUILD_DIR"
echo "creating successful"

# Function to generate SHA256 hash files
generate_hashes() {
  local dir=$1
  echo "generating SHA256 hashes in $dir..."
  for file in "$dir"/*; do
    if [ -f "$file" ]; then
      echo "hashing file $file"
      sha256sum "$file" > "$file.sha256" || { echo "Error: Failed to generate hash for $file"; exit 1; }
      echo "hashing successful for file $file"
    fi
  done
}

# Build for Linux
echo "building for Linux..."
dotnet publish "$PROJECT_FILE" -c Release -r linux-x64 --self-contained true --nologo -v q --property WarningLevel=0 /clp:ErrorsOnly /p:PublishSingleFile=true -o "../$BUILD_DIR/linux" || { echo "Error: Linux build failed"; exit 1; }
cd "../$BUILD_DIR/linux" || exit
echo "compressing for linux as .zip"
zip "../${LINUX_FILENAME}.zip" * || exit
echo "success"
echo "compressing for linux as .tar.gz"
tar -czf "../${LINUX_FILENAME}.tar.gz" * || exit
echo "success"
cd ../../"$REPO_NAME" || exit

# Build for Windows
echo "building for Windows..."
dotnet publish "$PROJECT_FILE" -c Release -r win-x64 --self-contained true --nologo -v q --property WarningLevel=0 /clp:ErrorsOnly /p:PublishSingleFile=true -o "../$BUILD_DIR/windows" || { echo "Error: Windows build failed"; exit 1; }
cd "../$BUILD_DIR/windows" || exit
echo "compressing for windows as .zip"
zip "../${WINDOWS_FILENAME}.zip" * || exit
echo "success"
echo "compressing for windows as .tar.gz"
tar -czf "../${WINDOWS_FILENAME}.tar.gz" * || exit
echo "success"
cd ../../"$REPO_NAME" || exit

# generating hashes
echo "initializing hash generation"
generate_hashes "../$BUILD_DIR/"
echo "hashing finished successfully"

# Create GitHub Release
echo "creating GitHub release..."
LATEST_TAG=$(echo "auto_gen_$(date '+%Y-%m-%d_%H-%M-%S')")
gh release create "$LATEST_TAG" \
  "../$BUILD_DIR/${LINUX_FILENAME}.zip" \
  "../$BUILD_DIR/${LINUX_FILENAME}.tar.gz" \
  "../$BUILD_DIR/${LINUX_FILENAME}.zip.sha256" \
  "../$BUILD_DIR/${LINUX_FILENAME}.tar.gz.sha256" \
  "../$BUILD_DIR/${WINDOWS_FILENAME}.zip" \
  "../$BUILD_DIR/${WINDOWS_FILENAME}.tar.gz" \
  "../$BUILD_DIR/${WINDOWS_FILENAME}.zip.sha256" \
  "../$BUILD_DIR/${WINDOWS_FILENAME}.tar.gz.sha256" \
  --title "$LATEST_TAG" \
  --notes "$RELEASE_NAME" || { echo "Error: GitHub release creation failed"; exit 1; }
echo "generating release finished successfully"

# Cleanup
echo "cleaning..."
rm -rf "../$BUILD_DIR"
rm -rf "../$REPO_NAME"
echo "cleaning finished successfully"
echo " "
echo "Done, see you soon"

