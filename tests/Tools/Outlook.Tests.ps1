# Outlook.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/Outlook.ps1"
. $ToolPath

Describe "Outlook Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Name Variant Logic" {
        It "Should return variants for a common name like 'Sam'" {
            $variants = Get-NameVariants -First "Sam"
            # Pester 3 'Contain' checks files, not arrays. Use -contains operator instead.
            ($variants -contains "Samuel") | Should Be $true
            ($variants -contains "Samantha") | Should Be $true
            ($variants -contains "Sam") | Should Be $true
        }

        It "Should return the original name if no variants exist" {
            $variants = Get-NameVariants -First "Zorgon"
            $variants.Count | Should Be 1
            ($variants -contains "Zorgon") | Should Be $true
        }

        It "Should generate full name variations correctly" {
            $variations = Get-AllNameVariations -Name "Sam Price"
            ($variations -contains "Price, Samuel") | Should Be $true
            ($variations -contains "Samuel Price") | Should Be $true
            ($variations -contains "Price") | Should Be $true
        }
    }

    Context "Invoke-OutlookReadEmails" {
        It "Should return emails as JSON when Outlook is available" {
            # Mock the COM object and its hierarchy
            $mockItem = New-Object -TypeName PSObject
            $mockItem | Add-Member -MemberType NoteProperty -Name Subject -Value "Test Subject"
            $mockItem | Add-Member -MemberType NoteProperty -Name SenderName -Value "Test Sender"
            $mockItem | Add-Member -MemberType NoteProperty -Name SenderEmailAddress -Value "test@example.com"
            $mockItem | Add-Member -MemberType NoteProperty -Name ReceivedTime -Value (Get-Date)
            $mockItem | Add-Member -MemberType NoteProperty -Name Body -Value "Test Body content"

            $mockFolder = New-Object -TypeName PSObject
            $mockFolder | Add-Member -MemberType NoteProperty -Name Items -Value @($mockItem)

            $mockNS = New-Object -TypeName PSObject
            $mockNS | Add-Member -MemberType ScriptMethod -Name GetDefaultFolder -Value { return $mockFolder }

            $mockOutlook = New-Object -TypeName PSObject
            $mockOutlook | Add-Member -MemberType ScriptMethod -Name GetNamespace -Value { return $mockNS }

            Mock New-Object { return $mockOutlook } -ParameterFilter { $ComObject -eq "Outlook.Application" }

            $result = Invoke-OutlookReadEmails -Count 1
            $parsed = $result | ConvertFrom-Json
            
            $parsed.Subject | Should Be "Test Subject"
            $parsed.Sender | Should Be "Test Sender"
        }
    }

    Context "Test-ResolvedNameMatch" {
        It "Should return true if last name is found in resolved display name" {
            $match = Test-ResolvedNameMatch -ResolvedDisplayName "Price, Samuel" -RequestedLastName "Price"
            $match | Should Be $true
        }

        It "Should return false if last name is missing from resolved display name" {
            $match = Test-ResolvedNameMatch -ResolvedDisplayName "Smith, John" -RequestedLastName "Price"
            $match | Should Be $false
        }
    }
}
