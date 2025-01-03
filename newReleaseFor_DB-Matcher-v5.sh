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
RELEASE_NAME="automatically-generated-binaries"

# Clone repository
echo "cleaning..."
rm -rf $REPO_NAME
echo "cleaning successful"

echo "Cloning repository..."
git clone "$REPO_URL" "$REPO_NAME" || { echo "Error: git clone failed"; exit 1; }
cd "$REPO_NAME" || { echo "Error: repository not found"; exit 1; }
echo "cloning successful"

# Clean build directory
echo "cleaning build directory"
rm -rf "../$BUILD_DIR"
echo "cleaning successful"
echo "creating build directory"
mkdir "../$BUILD_DIR"
echo "creatung successful"

# Function to generate SHA256 hash files
generate_hashes() {
  local dir=$1
  echo "Generating SHA256 hashes in $dir..."
  for file in "$dir"/*; do
    if [ -f "$file" ]; then
      echo "hashing file $file"
      sha256sum "$file" > "$file.sha256" || { echo "Error: Failed to generate hash for $file"; exit 1; }
      echo "hashing successful for file $file
    fi
  done
}

# Build for Linux
echo "Building for Linux..."
dotnet publish "$PROJECT_FILE" -c Release -r linux-x64 --self-contained true --nologo -v q --property WarningLevel=0 /clp:ErrorsOnly /p:PublishSingleFile=true -o "../$BUILD_DIR/linux" || { echo "Error: Linux build failed"; exit 1; }
cd "../$BUILD_DIR/linux" || exit
zip "../${LINUX_FILENAME}.zip" * || exit
tar -czf "../${LINUX_FILENAME}.tar.gz" * || exit
cd ../../"$REPO_NAME" || exit

# Build for Windows
echo "Building for Windows..."
dotnet publish "$PROJECT_FILE" -c Release -r win-x64 --self-contained true --nologo -v q --property WarningLevel=0 /clp:ErrorsOnly /p:PublishSingleFile=true -o "../$BUILD_DIR/windows" || { echo "Error: Windows build failed"; exit 1; }
cd "../$BUILD_DIR/windows" || exit
zip "../${WINDOWS_FILENAME}.zip" * || exit
tar -czf "../${WINDOWS_FILENAME}.tar.gz" * || exit
cd ../../"$REPO_NAME" || exit

# generating hashes
echo "generating hashes"
generate_hashes "../$BUILD_DIR/"

# Create GitHub Release
echo "Creating GitHub release..."
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

# Cleanup
echo "Cleanup..."
rm -rf "../$BUILD_DIR"
rm -rf "../$REPO_NAME"

echo "Done"

