# run-tests.ps1
# This script runs all Pester tests in the tests directory.

if (-not (Get-Module -ListAvailable Pester)) {
    Write-Error "Pester module not found. Please install it with: Install-Module Pester -Force"
    exit 1
}

Import-Module Pester

# Support for both Pester v5+ and older versions (v3/v4)
$pesterPath = "$PSScriptRoot/tests/Tools"

if (Get-Command New-PesterConfiguration -ErrorAction SilentlyContinue) {
    # Pester v5+
    $Config = New-PesterConfiguration
    $Config.Run.Path = $pesterPath
    $Config.Output.Verbosity = "Detailed"
    Invoke-Pester -Configuration $Config
} else {
    # Pester v3/v4 (Windows 10 default)
    Invoke-Pester -Script $pesterPath -EnableExit
}
