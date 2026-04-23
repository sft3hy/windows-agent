# Web.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/Web.ps1"
. $ToolPath

Describe "Web Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-WebFetch" {
        It "Should strip HTML tags and return plain text" {
            $mockResponse = New-Object -TypeName PSObject
            $mockResponse | Add-Member -MemberType NoteProperty -Name Content -Value "<html><body><h1>Hello</h1><script>evil()</script><p>World</p></body></html>"
            
            Mock Invoke-WebRequest { return $mockResponse }
            
            $result = Invoke-WebFetch -Url "https://example.com"
            
            $result | Should Match "Hello\s+World"
        }

        It "Should handle request failures gracefully" {
            Mock Invoke-WebRequest { throw "Connection failed" }
            
            $result = Invoke-WebFetch -Url "https://broken.com"
            
            $result | Should Match "ERROR: Connection failed"
        }
    }

    Context "Invoke-WebSearch" {
        It "Should extract URLs and call Invoke-WebFetch" {
            $mockSearchResponse = New-Object -TypeName PSObject
            # Simplified DuckDuckGo HTML result pattern
            $html = '... class="result__url" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fexample.com%2Fpage1" ...'
            $mockSearchResponse | Add-Member -MemberType NoteProperty -Name Content -Value $html
            
            Mock Invoke-WebRequest { return $mockSearchResponse }
            Mock Invoke-WebFetch { return "Page content" }
            
            $result = Invoke-WebSearch -Query "test query" -TopN 1
            
            $result | Should Match "SOURCE 1: https://example.com/page1"
            $result | Should Match "Page content"
            Assert-MockCalled Invoke-WebFetch -Times 1
        }
    }
}
