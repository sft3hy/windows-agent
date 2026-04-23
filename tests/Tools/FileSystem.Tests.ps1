# FileSystem.Tests.ps1

$ToolPath = Resolve-Path "$PSScriptRoot/../../src/Tools/FileSystem.ps1"
. $ToolPath

Describe "FileSystem Tool" {
    function Write-ToolLine { param($Tool, $Action, $Details) }
    function Write-StatusLine { param($Status, $Message) }
    
    # Initialize ENV_PATHS for testing
    $script:ENV_PATHS = @{
        Desktop   = "$env:TEMP/Desktop"
        Downloads = "$env:TEMP/Downloads"
        Documents = "$env:TEMP/Documents"
    }
    
    # Create dummy folders
    New-Item -ItemType Directory -Path $script:ENV_PATHS.Desktop -Force | Out-Null

    Context "Invoke-ReadFile - Security Sandbox" {
        It "Should allow reading a file inside the user profile" {
            $testFile = "$env:USERPROFILE/test.txt"
            Mock Test-Path { return $true }
            Mock Get-Content { return "test content" }
            
            $result = Invoke-ReadFile -FilePath $testFile
            
            $result | Should Be "test content"
        }

        It "Should block reading a file outside allowed roots (e.g., C:\Windows)" {
            $evilFile = "C:\Windows\System32\config\SAM"
            Mock Test-Path { return $true }
            
            $result = Invoke-ReadFile -FilePath $evilFile
            
            $result | Should Match "Access Denied"
        }
    }

    Context "Resolve-FuzzyFilePath" {
        It "Should find an exact match if it exists" {
            $fileName = "ExistingFile.txt"
            $fullPath = Join-Path $script:ENV_PATHS.Desktop $fileName
            New-Item -ItemType File -Path $fullPath -Force | Out-Null
            
            $result = Resolve-FuzzyFilePath -Name "ExistingFile.txt"
            
            $result | Should Be $fullPath
        }

        It "Should find a file by appending common extensions" {
            $fileName = "ProjectReport.pdf"
            $fullPath = Join-Path $script:ENV_PATHS.Desktop $fileName
            New-Item -ItemType File -Path $fullPath -Force | Out-Null
            
            $result = Resolve-FuzzyFilePath -Name "ProjectReport"
            
            $result | Should Be $fullPath
        }

        It "Should return null if no file is found" {
            $result = Resolve-FuzzyFilePath -Name "NonExistentFile"
            
            $result | Should Be $null
        }
    }
}
