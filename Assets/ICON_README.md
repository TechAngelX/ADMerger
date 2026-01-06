# App Icon Setup

## Files
- `app-icon.svg` - Source SVG icon (512x512)

## Quick Start
After converting the SVG to .ico format (see instructions below), add this line to ADMerger.csproj in the `<PropertyGroup>` section:
```xml
<ApplicationIcon>Assets/app-icon.ico</ApplicationIcon>
```

## Converting SVG to Platform-Specific Icons

### For macOS (.icns)
```bash
# Install ImageMagick if not already installed
brew install imagemagick

# Convert SVG to PNG at multiple sizes
mkdir icon.iconset
sips -z 16 16 app-icon.svg --out icon.iconset/icon_16x16.png
sips -z 32 32 app-icon.svg --out icon.iconset/icon_16x16@2x.png
sips -z 32 32 app-icon.svg --out icon.iconset/icon_32x32.png
sips -z 64 64 app-icon.svg --out icon.iconset/icon_32x32@2x.png
sips -z 128 128 app-icon.svg --out icon.iconset/icon_128x128.png
sips -z 256 256 app-icon.svg --out icon.iconset/icon_128x128@2x.png
sips -z 256 256 app-icon.svg --out icon.iconset/icon_256x256.png
sips -z 512 512 app-icon.svg --out icon.iconset/icon_256x256@2x.png
sips -z 512 512 app-icon.svg --out icon.iconset/icon_512x512.png
sips -z 1024 1024 app-icon.svg --out icon.iconset/icon_512x512@2x.png

# Convert to icns
iconutil -c icns icon.iconset -o app-icon.icns

# Clean up
rm -rf icon.iconset
```

Or use online converter:
1. Go to https://cloudconvert.com/svg-to-icns
2. Upload app-icon.svg
3. Download app-icon.icns
4. Place in Assets folder

### For Windows (.ico)
```bash
# Using ImageMagick
convert app-icon.svg -define icon:auto-resize=256,128,64,48,32,16 app-icon.ico
```

Or use online converter:
1. Go to https://convertio.co/svg-ico/
2. Upload app-icon.svg
3. Download app-icon.ico
4. Place in Assets folder

### For Linux (.png)
```bash
# Simple PNG export at 512x512
convert app-icon.svg -resize 512x512 app-icon.png
```

## Icon Design Details
- **Primary Color**: Blue gradient (#2563EB to #1E40AF)
- **Design**: "AD" lettermark with subtle document merge imagery in background
- **Style**: Modern, clean, professional
- **Text**: White "AD" with "Merger" subtitle
- **Background**: Rounded square with gradient and shine overlay
