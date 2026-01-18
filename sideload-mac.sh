#!/bin/bash

# Sideload script for PowerPoint on Mac
# This script copies the manifest.xml to the wef folder

echo "🚀 Mermaid PowerPoint Plugin - Mac Sideload Script"
echo "=================================================="
echo ""

# Define the wef folder path
WEF_FOLDER="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

# Check if manifest.xml exists
if [ ! -f "manifest.xml" ]; then
    echo "❌ Error: manifest.xml not found in current directory"
    echo "   Please run this script from the project root directory"
    exit 1
fi

# Create the wef folder if it doesn't exist
echo "📁 Creating wef folder (if needed)..."
mkdir -p "$WEF_FOLDER"

if [ $? -eq 0 ]; then
    echo "✅ wef folder ready: $WEF_FOLDER"
else
    echo "❌ Error: Could not create wef folder"
    exit 1
fi

# Copy the manifest
echo ""
echo "📋 Copying manifest.xml..."
cp manifest.xml "$WEF_FOLDER/"

if [ $? -eq 0 ]; then
    echo "✅ Manifest copied successfully"
else
    echo "❌ Error: Could not copy manifest"
    exit 1
fi

# Verify the copy
echo ""
echo "🔍 Verifying..."
if [ -f "$WEF_FOLDER/manifest.xml" ]; then
    echo "✅ Manifest verified in wef folder"
    echo ""
    echo "📊 Manifest details:"
    ls -lh "$WEF_FOLDER/manifest.xml"
else
    echo "❌ Error: Manifest not found after copy"
    exit 1
fi

echo ""
echo "=================================================="
echo "✅ Sideload Complete!"
echo ""
echo "Next steps:"
echo "1. Make sure the dev server is running (npm start)"
echo "2. Quit PowerPoint completely (Cmd+Q)"
echo "3. Reopen PowerPoint"
echo "4. Go to Insert → My Add-ins"
echo "5. Look for 'Mermaid Diagrams' under Developer Add-ins"
echo ""
echo "If you don't see the add-in:"
echo "- Check that npm start is running"
echo "- Visit https://localhost:3000 in Safari and accept the certificate"
echo "- See MAC_SIDELOAD_GUIDE.md for troubleshooting"
echo ""
echo "=================================================="

# Made with Bob
