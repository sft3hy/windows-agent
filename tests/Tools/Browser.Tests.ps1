# Browser.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/Browser.ps1"
. $ToolPath

Describe "Browser Tool" {
    # Mocking Write-ToolLine and Write-StatusLine to avoid console clutter
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }

    Context "Invoke-BrowserOpen" {
        It "Should attempt to open a URL using Chrome if specified" {
            Mock Start-Process { }
            Mock Find-BrowserPath { return "C:\Program Files\Google\Chrome\Application\chrome.exe" }
            
            $result = Invoke-BrowserOpen -Url "https://google.com" -Browser "chrome"
            
            $result | Should Be "Chrome opened to: https://google.com"
            Assert-MockCalled Start-Process -Times 1 -ParameterFilter { 
                $ArgumentList -contains "https://google.com" 
            }
        }

        It "Should fallback to system default if no browsers are found" {
            Mock Start-Process { }
            Mock Find-BrowserPath { return $null }
            
            $result = Invoke-BrowserOpen -Url "https://bing.com"
            
            $result | Should Be "Browser opened to: https://bing.com"
            Assert-MockCalled Start-Process -Times 1 -ParameterFilter { 
                $FilePath -eq "https://bing.com" 
            }
        }
    }

    Context "Invoke-BrowserSearch" {
        It "Should construct a Google search URL by default" {
            Mock Invoke-BrowserOpen { param($Url, $Browser) return "Browser opened to: $Url" }
            
            $result = Invoke-BrowserSearch -Query "pester testing"
            
            $result | Should Be "Browser opened to: https://www.google.com/search?q=pester%20testing"
        }

        It "Should construct a DuckDuckGo search URL if specified" {
            Mock Invoke-BrowserOpen { param($Url, $Browser) return "Browser opened to: $Url" }
            
            $result = Invoke-BrowserSearch -Query "privacy" -Engine "duckduckgo"
            
            $result | Should Be "Browser opened to: https://duckduckgo.com/?q=privacy"
        }
    }
}
