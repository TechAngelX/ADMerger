Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "   AD Merger - Building Standalone EXE    " -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host ""

$projectPath = "C:\Users\Ricki\Documents\LOCALDEV-PC\ADMerger"
$desktopPath = "C:\Users\Ricki\Desktop"

Write-Host "Checking for running ADMerger processes..." -ForegroundColor Yellow
$runningProcesses = Get-Process -Name "ADMerger" -ErrorAction SilentlyContinue
if ($runningProcesses) {
    Write-Host "  Closing running ADMerger instances..." -ForegroundColor Yellow
    $runningProcesses | Stop-Process -Force
    Start-Sleep -Seconds 1
}

$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "  WARNING: Excel is running. Close any open files!" -ForegroundColor Red
    Write-Host "  Press any key after closing Excel, or Ctrl+C to cancel..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
Set-Location $projectPath
dotnet clean

Write-Host "Publishing standalone executable with embedded data..." -ForegroundColor Yellow
dotnet publish -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true

if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Build successful!" -ForegroundColor Green
    
    # FIXED: Changed net9.0-windows to net10.0
    $publishFolder = "$projectPath\bin\Release\net10.0\win-x64\publish"
    $publishedExe = Get-ChildItem -Path "$publishFolder\ADMerger.exe" -ErrorAction SilentlyContinue

    if ($publishedExe) {
        $desktopExe = "$desktopPath\ADMerger.exe"

        if (Test-Path $desktopExe) {
            Remove-Item $desktopExe -Force
        }

        Copy-Item $publishedExe.FullName -Destination $desktopExe -Force
        $fileSize = [math]::Round($publishedExe.Length / 1MB, 2)

        Write-Host ""
        Write-Host "===========================================" -ForegroundColor Green
        Write-Host "           SUCCESS! ✓                      " -ForegroundColor Green
        Write-Host "===========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Standalone executable created:" -ForegroundColor White
        Write-Host "  $desktopExe" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "File size: $fileSize MB" -ForegroundColor White
        Write-Host ""
        Write-Host "✓ ALL data files embedded in EXE:" -ForegroundColor Green
        Write-Host "  - THE Ranking 2026.xlsx" -ForegroundColor Gray
        Write-Host "  - ucl_degree_equivalencies_FINAL.csv" -ForegroundColor Gray
        Write-Host "  - institution_mappings.csv" -ForegroundColor Gray
        Write-Host ""
        Write-Host "✓ No external files needed!" -ForegroundColor Green
        Write-Host "✓ Single file - ready to share!" -ForegroundColor Green
        Write-Host ""
        Write-Host "🚀 Send ADMerger.exe to friends - it's 100% standalone!" -ForegroundColor Yellow
    } else {
        Write-Host "ERROR: Could not find published executable!" -ForegroundColor Red
        Write-Host "Expected at: $publishFolder\ADMerger.exe" -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "ERROR: Build failed!" -ForegroundColor Red
    Write-Host "Tip: Make sure Excel and ADMerger are closed!" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")