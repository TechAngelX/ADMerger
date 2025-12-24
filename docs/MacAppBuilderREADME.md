# 📦 Build ADMerger as a Standalone Mac App (Intel)

Follow these steps to create a clean, single-file macOS application (`.app`) that includes all necessary data and audio files.

---

### **Step 1: Publish the Single-File Binary**
Run this command from your project root folder (where `ADMerger.csproj` is located).

```

dotnet publish -c Release -r osx-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o ~/Desktop/ADMerger_Build

```
Step 2: Navigate to the Output Folder


cd ~/Desktop/ADMerger_Build
Step 3: Create the Mac App Structure


mkdir -p ADMerger.app/Contents/MacOS
Step 4: Create the Info.plist Configuration
Copy and paste this block to create the plist file:

```

cat <<EOF > ADMerger.app/Contents/Info.plist
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "[http://www.apple.com/DTDs/PropertyList-1.0.dtd](http://www.apple.com/DTDs/PropertyList-1.0.dtd)">
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
Step 5: Move the Executable
```

mv ADMerger ADMerger.app/Contents/MacOS/
Step 6: Clean Up


🎉 Done!
You will now see ADMerger.app in the folder. Double-click to run.

