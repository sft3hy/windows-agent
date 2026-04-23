# Word.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/Word.ps1"
. $ToolPath

Describe "Word Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-WordCreateDocument" {
        It "Should create a new document with title and body" {
            $mockStyle = New-Object -TypeName PSObject
            $mockStyles = New-Object -TypeName PSObject
            $mockStyles | Add-Member -MemberType ScriptMethod -Name Item -Value { return $mockStyle }

            $mockDoc = New-Object -TypeName PSObject
            $mockDoc | Add-Member -MemberType NoteProperty -Name Styles -Value $mockStyles
            $mockDoc | Add-Member -MemberType ScriptMethod -Name SaveAs2 -Value { }
            $mockDoc | Add-Member -MemberType ScriptMethod -Name Close -Value { }

            $mockDocuments = New-Object -TypeName PSObject
            $mockDocuments | Add-Member -MemberType ScriptMethod -Name Add -Value { return $mockDoc }

            $mockSelection = New-Object -TypeName PSObject
            $mockSelection | Add-Member -MemberType NoteProperty -Name Style -Value $null
            $mockSelection | Add-Member -MemberType ScriptMethod -Name TypeText -Value { }
            $mockSelection | Add-Member -MemberType ScriptMethod -Name TypeParagraph -Value { }

            $mockWord = New-Object -TypeName PSObject
            $mockWord | Add-Member -MemberType NoteProperty -Name Visible -Value $false
            $mockWord | Add-Member -MemberType NoteProperty -Name DisplayAlerts -Value 0
            $mockWord | Add-Member -MemberType NoteProperty -Name Documents -Value $mockDocuments
            $mockWord | Add-Member -MemberType NoteProperty -Name Selection -Value $mockSelection
            $mockWord | Add-Member -MemberType ScriptMethod -Name Quit -Value { }

            Mock New-Object { return $mockWord } -ParameterFilter { $ComObject -eq "Word.Application" }
            Mock Test-Path { return $true }

            # Use -Display so the function shows the doc instead of trying to
            # Close/Quit/ReleaseComObject (which fails on non-COM PSObjects)
            $result = Invoke-WordCreateDocument -FilePath "C:\test.docx" -Title "Title" -Body "Body" -Display
            $result | Should Be "Word document created at C:\test.docx"
        }
    }
}
