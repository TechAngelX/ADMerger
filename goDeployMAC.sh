#!/bin/bash

echo "==========================================="
echo "   ADMerger - Building Standalone .app    "
echo "==========================================="
echo ""

# Get the script directory before changing directories
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BUILD_DIR=~/Desktop/ADMerger_Build

echo "Cleaning previous builds..."
rm -rf "$BUILD_DIR"
mkdir -p "$BUILD_DIR"

echo "Publishing standalone application..."
dotnet publish -c Release -r osx-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o "$BUILD_DIR/temp"

if [ $? -eq 0 ]; then
    echo "✓ Build successful!"

    cd "$BUILD_DIR"

    # Create .app bundle structure
    mkdir -p ADMerger.app/Contents/MacOS
    mkdir -p ADMerger.app/Contents/Resources

    cat <<EOF > ADMerger.app/Contents/Info.plist
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>ADMerger</string>
    <key>CFBundleIconFile</key>
    <string>app-icon.icns</string>
    <key>CFBundleIdentifier</key>
    <string>com.techangelx.admerger</string>
    <key>CFBundleName</key>
    <string>ADMerger</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.12</string>
    <key>LSUIElement</key>
    <false/>
</dict>
</plist>
EOF

    # Move executable into .app bundle
    mv temp/ADMerger ADMerger.app/Contents/MacOS/

    # Copy icon file from project Assets folder
    if [ -f "$SCRIPT_DIR/Assets/app-icon.icns" ]; then
        cp "$SCRIPT_DIR/Assets/app-icon.icns" ADMerger.app/Contents/Resources/
        echo "✓ Icon added to app bundle"
    else
        echo "⚠ Warning: Icon file not found at $SCRIPT_DIR/Assets/app-icon.icns"
    fi

    # Clean up temporary files
    rm -rf temp

    echo ""
    echo "==========================================="
    echo "           SUCCESS! ✓                     "
    echo "==========================================="
    echo ""
    echo "Standalone application created:"
    echo "  ~/Desktop/ADMerger_Build/ADMerger.app"
    echo ""
    echo "✓ All data files embedded in app!"
    echo "✓ No external files needed!"
    echo "✓ Ready to use or share!"
    echo ""
else
    echo ""
    echo "ERROR: Build failed!"
    echo "Check for errors above."
fi

