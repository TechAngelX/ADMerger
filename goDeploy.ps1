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
    Write-Host "  WARNING: Excel is running. Please close any CSV files!" -ForegroundColor Red
    Write-Host "  Press any key after closing Excel, or Ctrl+C to cancel..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
dotnet clean

Write-Host "Publishing self-contained executable..." -ForegroundColor Yellow
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true

if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Build successful!" -ForegroundColor Green
    
    $publishedExe = Get-ChildItem -Path "$projectPath\bin\Release\net9.0-windows\win-x64\publish\ADMerger.exe" -ErrorAction SilentlyContinue
    
    if ($publishedExe) {
        $desktopExe = "$desktopPath\ADMerger.exe"
        if (Test-Path $desktopExe) {
            Remove-Item $desktopExe -Force
        }
        
        Copy-Item $publishedExe.FullName -Destination $desktopExe -Force
        
        Write-Host ""
        Write-Host "===========================================" -ForegroundColor Green
        Write-Host "   SUCCESS! ✓                             " -ForegroundColor Green
        Write-Host "===========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Standalone executable created:" -ForegroundColor White
        Write-Host "  $desktopExe" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "File size: $([math]::Round($publishedExe.Length / 1MB, 2)) MB" -ForegroundColor White
        Write-Host ""
        Write-Host "✓ CSV data embedded in EXE - no external files needed!" -ForegroundColor Green
        Write-Host ""
        Write-Host "🚀 Double-click ADMerger.exe to run!" -ForegroundColor Yellow
    }
    else {
        Write-Host "ERROR: Could not find published executable!" -ForegroundColor Red
    }
}
else {
    Write-Host ""
    Write-Host "ERROR: Build failed!" -ForegroundColor Red
    Write-Host "Tip: Make sure Excel and ADMerger are closed!" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
