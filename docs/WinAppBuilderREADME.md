# 📦 Build ADMerger as a standalone App for Windows

---

### **Step 1: Publish the Single-File Binary**
Run this command from your project root. This tells .NET to cross-compile for Windows (x64).

```
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o ~/Desktop/ADMerger_Win_Build
```

Step 2: Navigate to the Output Folder
```

cd ~/Desktop/ADMerger_Win_Build
```

You will see ADMerger.exe in the folder.

You can send this .exe file directly to any Windows user.

It is fully standalone (you do not need to install .NET).

To share it, just zip the .exe file:
Windows.zip ADMerger.exe
