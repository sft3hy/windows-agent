# Test Excel Integration
$ErrorActionPreference = "Stop"

# Mock Write-ToolLine and Write-StatusLine if not loaded
function Write-ToolLine ($Tool, $Action, $Details) { Write-Host "[$Tool] ${Action}: $Details" -ForegroundColor Cyan }
function Write-StatusLine ($Status, $Message) { Write-Host "[$Status] $Message" -ForegroundColor Green }

# Import Excel tool
. (Join-Path $PSScriptRoot "../Tools/Excel.ps1")
. (Join-Path $PSScriptRoot "../Tools/FileSystem.ps1")

$testFile = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "AIRWAV_ExcelTest.xlsx"
if (Test-Path $testFile) { Remove-Item $testFile -Force }

Write-Host "--- Testing Invoke-ExcelCreateFile ---"
$data = "Name`tAge`tCity`nAlice`t30`tNew York`nBob`t25`tLondon"
$res = Invoke-ExcelCreateFile -FilePath $testFile -Data $data
Write-Host "Result: $res"

Write-Host "`n--- Testing Invoke-ExcelReadFile ---"
$content = Invoke-ExcelReadFile -FilePath $testFile
Write-Host "Content:`n$content"

Write-Host "`n--- Testing Invoke-ExcelGetMetadata ---"
$meta = Invoke-ExcelGetMetadata -FilePath $testFile
Write-Host "Metadata:`n$meta"

Write-Host "`n--- Testing Invoke-ExcelBeautify ---"
$res = Invoke-ExcelBeautify -FilePath $testFile
Write-Host "Result: $res"

Write-Host "`n--- Testing Invoke-ExcelWriteCell ---"
$res = Invoke-ExcelWriteCell -FilePath $testFile -Cell "A4" -Value "Charlie"
Write-Host "Result: $res"

Write-Host "`n--- Testing FileSystem.ps1 Integration ---"
$fsRes = Invoke-ReadFile -FilePath $testFile
Write-Host "FileSystem Read Result:`n$fsRes"

# Clean up
# if (Test-Path $testFile) { Remove-Item $testFile -Force }
Write-Host "`nTest Complete. File left at $testFile for manual inspection if needed."
