# Jabber.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/Jabber.ps1"
. $ToolPath

Describe "Jabber Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-JabberSendMessage" {
        It "Should fail gracefully if Jabber is not running" {
            Mock Get-Process { return $null }
            Mock Find-JabberPath { return $null }

            $result = Invoke-JabberSendMessage -Recipient "testuser" -Message "Hello"

            $result | Should Match "not running"
        }
    }

    Context "Invoke-JabberOpenChat" {
        It "Should try multiple URI schemes" {
            Mock Get-Process { return (New-Object -TypeName PSObject -Property @{ MainWindowHandle = [IntPtr]::Zero }) }
            Mock Start-Process { }
            Mock Set-JabberForeground { return $true }

            $result = Invoke-JabberOpenChat -Recipient "bob"

            $result | Should Match "Jabber chat opened"
            Assert-MockCalled Start-Process -Times 1
        }
    }
}
