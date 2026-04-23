# Load configuration and tools
. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\src\Tools\FileSystem.ps1"
. "$PSScriptRoot\src\Tools\Word.ps1"

Write-Host "Searching for HELM_GUIDE.pdf..." -ForegroundColor Cyan
$pdfPath = Resolve-FuzzyFilePath -Name "HELM_GUIDE"

if (-not $pdfPath) {
    Write-Error "Could not find HELM_GUIDE.pdf in standard folders (Desktop, Documents, Downloads)."
    exit
}

Write-Host "Found PDF at: $pdfPath" -ForegroundColor Green
Write-Host "Reading PDF content..." -ForegroundColor Cyan

$content = Invoke-ReadFile -FilePath $pdfPath

if ($content -match "^ERROR:") {
    Write-Error "Failed to read PDF: $content"
    exit
}

$targetDoc = [System.IO.Path]::ChangeExtension($pdfPath, ".docx")
Write-Host "Creating Word document at: $targetDoc" -ForegroundColor Cyan

$result = Invoke-WordCreateDocument -FilePath $targetDoc -Title "HELM Guide" -Body $content -Display

Write-Host $result -ForegroundColor Green
