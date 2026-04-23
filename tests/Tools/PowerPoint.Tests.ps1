# PowerPoint.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/PowerPoint.ps1"
. $ToolPath

Describe "PowerPoint Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-PowerPointCreate" {
        It "Should create a presentation and handle local copy logic" {
            $mockSlide = New-Object -TypeName PSObject
            $mockSlide | Add-Member -MemberType NoteProperty -Name Shapes -Value @(
                (New-Object -TypeName PSObject -Property @{ TextFrame = (New-Object -TypeName PSObject -Property @{ TextRange = (New-Object -TypeName PSObject -Property @{ Text = "" }) }) }),
                (New-Object -TypeName PSObject -Property @{ TextFrame = (New-Object -TypeName PSObject -Property @{ TextRange = (New-Object -TypeName PSObject -Property @{ Text = "" }) }) })
            )

            $mockSlides = New-Object -TypeName PSObject
            $mockSlides | Add-Member -MemberType NoteProperty -Name Count -Value 0
            $mockSlides | Add-Member -MemberType ScriptMethod -Name Add -Value { return $mockSlide }

            $mockPres = New-Object -TypeName PSObject
            $mockPres | Add-Member -MemberType NoteProperty -Name Slides -Value $mockSlides
            $mockPres | Add-Member -MemberType ScriptMethod -Name SaveAs -Value { }

            $mockPresentations = New-Object -TypeName PSObject
            $mockPresentations | Add-Member -MemberType ScriptMethod -Name Add -Value { return $mockPres }

            $mockPPT = New-Object -TypeName PSObject
            $mockPPT | Add-Member -MemberType NoteProperty -Name Visible -Value 0
            $mockPPT | Add-Member -MemberType NoteProperty -Name Presentations -Value $mockPresentations

            Mock New-Object { return $mockPPT } -ParameterFilter { $ComObject -eq "PowerPoint.Application" }
            Mock Test-Path { return $true }
            Mock Copy-Item { }
            Mock Remove-Item { }

            $json = '[{"title":"Test Slide","content":"Bullet points"}]'
            $result = Invoke-PowerPointCreate -FilePath "C:\final.pptx" -SlidesJson $json
            
            $result | Should Be "C:\final.pptx"
            Assert-MockCalled Copy-Item -Times 1
        }
    }
}
