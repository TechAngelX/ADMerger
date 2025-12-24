#!/bin/bash

dotnet publish -c Release -r osx-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o ~/Desktop/ADMerger_Build

cd ~/Desktop/ADMerger_Build

mkdir -p ADMerger.app/Contents/MacOS

cat <<EOF > ADMerger.app/Contents/Info.plist
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>ADMerger</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.admerger.app</string>
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

mv ADMerger ADMerger.app/Contents/MacOS/

mv ADMerger.app ~/Desktop/

