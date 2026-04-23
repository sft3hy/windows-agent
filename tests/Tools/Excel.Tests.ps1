# Excel.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/Excel.ps1"
. $ToolPath

Describe "Excel Tool" {
    # Mocking console output
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-ExcelReadFile" {
        It "Should read data from a mocked Excel sheet" {
            $mockCell = New-Object -TypeName PSObject
            $mockCell | Add-Member -MemberType NoteProperty -Name Text -Value "CellData"

            $mockCells = New-Object -TypeName PSObject
            $mockCells | Add-Member -MemberType ScriptMethod -Name Item -Value { return $mockCell }

            $mockUsedRange = New-Object -TypeName PSObject
            $mockUsedRange | Add-Member -MemberType NoteProperty -Name Rows -Value (New-Object -TypeName PSObject -Property @{ Count = 1 })
            $mockUsedRange | Add-Member -MemberType NoteProperty -Name Columns -Value (New-Object -TypeName PSObject -Property @{ Count = 1 })

            $mockWS = New-Object -TypeName PSObject
            $mockWS | Add-Member -MemberType NoteProperty -Name UsedRange -Value $mockUsedRange
            $mockWS | Add-Member -MemberType NoteProperty -Name Cells -Value $mockCells

            $mockWB = New-Object -TypeName PSObject
            $mockWB | Add-Member -MemberType NoteProperty -Name ActiveSheet -Value $mockWS
            $mockWB | Add-Member -MemberType ScriptMethod -Name Close -Value { }

            $mockWorkbooks = New-Object -TypeName PSObject
            $mockWorkbooks | Add-Member -MemberType ScriptMethod -Name Open -Value { return $mockWB }

            $mockExcel = New-Object -TypeName PSObject
            $mockExcel | Add-Member -MemberType NoteProperty -Name Visible -Value $false
            $mockExcel | Add-Member -MemberType NoteProperty -Name Workbooks -Value $mockWorkbooks
            $mockExcel | Add-Member -MemberType ScriptMethod -Name Quit -Value { }

            Mock New-Object { return $mockExcel } -ParameterFilter { $ComObject -eq "Excel.Application" }
            Mock Test-Path { return $true }
            Mock Dispose-ComObject { }

            $result = Invoke-ExcelReadFile -FilePath "C:\fake.xlsx"
            $result.Trim() | Should Be "CellData"
        }
    }

    Context "Invoke-ExcelCreateFile" {
        It "Should split TSV data correctly and write to cells" {
            $mockCell = New-Object -TypeName PSObject
            $mockCell | Add-Member -MemberType NoteProperty -Name Value2 -Value ""

            $mockCells = New-Object -TypeName PSObject
            $mockCells | Add-Member -MemberType ScriptMethod -Name Item -Value { return $mockCell }

            $mockWS = New-Object -TypeName PSObject
            $mockWS | Add-Member -MemberType NoteProperty -Name Cells -Value $mockCells
            $mockWS | Add-Member -MemberType NoteProperty -Name Name -Value ""

            $mockWB = New-Object -TypeName PSObject
            $mockWB | Add-Member -MemberType NoteProperty -Name ActiveSheet -Value $mockWS
            $mockWB | Add-Member -MemberType ScriptMethod -Name SaveAs -Value { }
            $mockWB | Add-Member -MemberType ScriptMethod -Name Close -Value { }

            $mockWorkbooks = New-Object -TypeName PSObject
            $mockWorkbooks | Add-Member -MemberType ScriptMethod -Name Add -Value { return $mockWB }

            $mockExcel = New-Object -TypeName PSObject
            $mockExcel | Add-Member -MemberType NoteProperty -Name Visible -Value $false
            $mockExcel | Add-Member -MemberType NoteProperty -Name DisplayAlerts -Value $false
            $mockExcel | Add-Member -MemberType NoteProperty -Name Workbooks -Value $mockWorkbooks
            $mockExcel | Add-Member -MemberType ScriptMethod -Name Quit -Value { }

            Mock New-Object { return $mockExcel } -ParameterFilter { $ComObject -eq "Excel.Application" }
            Mock Test-Path { return $true }
            Mock Dispose-ComObject { }

            $result = Invoke-ExcelCreateFile -FilePath "C:\new.xlsx" -Data "Header1`tHeader2`nValue1`tValue2"
            $result | Should Match "Successfully created Excel file"
        }
    }
}
