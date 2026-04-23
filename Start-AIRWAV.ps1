#Requires -Version 5.1
param([switch]$LoadOnly)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
# Enforce TLS 1.2 and 1.3 for enterprise API connectivity
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

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
if (-not $LoadOnly) {
    Start-AgentSession
}