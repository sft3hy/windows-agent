# Calendar.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/Calendar.ps1"
. $ToolPath

Describe "Calendar Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-OutlookReadCalendar" {
        It "Should return appointments as JSON" {
            $mockAppt = New-Object -TypeName PSObject
            $mockAppt | Add-Member -MemberType NoteProperty -Name Subject -Value "Test Appt"
            $mockAppt | Add-Member -MemberType NoteProperty -Name Start -Value (Get-Date)
            $mockAppt | Add-Member -MemberType NoteProperty -Name End -Value (Get-Date).AddHours(1)
            $mockAppt | Add-Member -MemberType NoteProperty -Name Location -Value "Test Room"
            $mockAppt | Add-Member -MemberType NoteProperty -Name Body -Value "Test Body"
            $mockAppt | Add-Member -MemberType NoteProperty -Name Organizer -Value "Test Org"

            $mockItems = New-Object -TypeName PSObject
            $mockItems | Add-Member -MemberType NoteProperty -Name IncludeRecurrences -Value $false
            $mockItems | Add-Member -MemberType ScriptMethod -Name Sort -Value { }
            $mockItems | Add-Member -MemberType ScriptMethod -Name Restrict -Value { return @($mockAppt) }

            $mockFolder = New-Object -TypeName PSObject
            $mockFolder | Add-Member -MemberType NoteProperty -Name Items -Value $mockItems

            $mockNS = New-Object -TypeName PSObject
            $mockNS | Add-Member -MemberType ScriptMethod -Name GetDefaultFolder -Value { return $mockFolder }

            $mockOutlook = New-Object -TypeName PSObject
            $mockOutlook | Add-Member -MemberType ScriptMethod -Name GetNamespace -Value { return $mockNS }

            Mock New-Object { return $mockOutlook } -ParameterFilter { $ComObject -eq "Outlook.Application" }

            $result = Invoke-OutlookReadCalendar
            $parsed = $result | ConvertFrom-Json
            
            $parsed.Subject | Should Be "Test Appt"
        }
    }

    Context "Invoke-OutlookCreateAppointment" {
        It "Should create a new appointment item" {
            $mockAppt = New-Object -TypeName PSObject
            $mockAppt | Add-Member -MemberType NoteProperty -Name Subject -Value ""
            $mockAppt | Add-Member -MemberType NoteProperty -Name Start -Value ""
            $mockAppt | Add-Member -MemberType NoteProperty -Name End -Value ""
            $mockAppt | Add-Member -MemberType NoteProperty -Name Location -Value ""
            $mockAppt | Add-Member -MemberType NoteProperty -Name Body -Value ""
            $mockAppt | Add-Member -MemberType NoteProperty -Name ReminderSet -Value $false
            $mockAppt | Add-Member -MemberType NoteProperty -Name ReminderMinutesBeforeStart -Value 0
            $mockAppt | Add-Member -MemberType ScriptMethod -Name Save -Value { }

            $mockOutlook = New-Object -TypeName PSObject
            $mockOutlook | Add-Member -MemberType ScriptMethod -Name CreateItem -Value { return $mockAppt }

            Mock New-Object { return $mockOutlook } -ParameterFilter { $ComObject -eq "Outlook.Application" }

            $result = Invoke-OutlookCreateAppointment -Subject "Lunch" -Start "2026-04-23 12:00"
            $result | Should Match "Calendar appointment created: 'Lunch'"
        }
    }
}
