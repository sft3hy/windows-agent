$script:CONFIG = @{
    ApiKey      = $env:GEMINI_API_KEY  # or hardcode: "your-key-here"
    ApiEndpoint = "https://api.genai.mil/v1/chat/completions"
    Model       = "gemini-2.5-flash"
    MaxTokens   = 4096
    Temperature = 0.3
    AgentName   = "AIRWAV"
    Version     = "1.0.0"
}

# Resolve shell folders to real paths at startup (handles redirected folders/UNC paths)
function Get-AIRWAVShellFolder {
    param([string]$ShellName, [string]$RegistryName, [string]$DefaultName)
    try {
        $comPath = (New-Object -ComObject Shell.Application).NameSpace("shell:$ShellName").Self.Path
        if ($comPath -and (Test-Path $comPath)) { return $comPath }
    } catch {}
    
    try {
        $regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        $val = (Get-ItemProperty $regKey -ErrorAction SilentlyContinue).$RegistryName
        if ($val) { 
            $expanded = [System.Environment]::ExpandEnvironmentVariables($val)
            if (Test-Path $expanded) { return $expanded }
        }
    } catch {}

    $netPath = [Environment]::GetFolderPath($DefaultName)
    if ($netPath -and (Test-Path $netPath)) { return $netPath }
    
    # User-specific fallback: RedirectedData pattern (Enterprise UNC)
    $profileRoots = @($env:USERPROFILE, $HOME, "\\nro.mil\home\s70\home01\townsesa")
    foreach ($root in ($profileRoots | Select-Object -Unique)) {
        if (-not $root) { continue }
        $redirectionPath = Join-Path $root "RedirectedData\$ShellName"
        if (Test-Path $redirectionPath) { return $redirectionPath }
        
        $directPath = Join-Path $root $ShellName
        if (Test-Path $directPath) { return $directPath }
    }
    
    $rootFallback = if ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }
    return Join-Path $rootFallback $ShellName
}

$script:ENV_PATHS = @{
    Downloads = Get-AIRWAVShellFolder -ShellName "Downloads" -RegistryName "{7d83af9a-51b5-40a0-a1ad-67acc47d6364}" -DefaultName "UserProfile"
    Desktop   = Get-AIRWAVShellFolder -ShellName "Desktop"   -RegistryName "Desktop"   -DefaultName "Desktop"
    Documents = Get-AIRWAVShellFolder -ShellName "Personal"  -RegistryName "Personal"  -DefaultName "MyDocuments"
}