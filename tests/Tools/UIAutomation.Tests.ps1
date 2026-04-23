# UIAutomation.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/UIAutomation.ps1"
. $ToolPath

Describe "UIAutomation Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-UIAutomationLaunch" {
        It "Should allow launching an approved app like notepad.exe" {
            Mock Start-Process { }
            
            $result = Invoke-UIAutomationLaunch -AppPath "C:\Windows\System32\notepad.exe"
            
            $result | Should Be "Launched application successfully."
            Assert-MockCalled Start-Process -Times 1
        }

        It "Should block an unapproved app" {
            $result = Invoke-UIAutomationLaunch -AppPath "C:\Evil\virus.exe"
            
            $result | Should Match "Launch unauthorized"
        }
    }
}
