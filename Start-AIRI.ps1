#Requires -Version 5.1
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$ProjectRoot = $PSScriptRoot

# 1. Load Configuration
. (Join-Path $ProjectRoot "Config.ps1")

# 2. Load Modules
$folders = @("UI", "Tools", "Core")
foreach ($folder in $folders) {
    $files = Get-ChildItem (Join-Path $ProjectRoot "src\$folder") -Filter "*.ps1"
    foreach ($file in $files) {
        . $file.FullName
    }
}

# 3. Start Agent
Start-AgentSession