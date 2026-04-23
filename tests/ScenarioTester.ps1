# ScenarioTester.ps1
# This script runs a battery of test scenarios against the AIRWAV engine.
# It injects prompts and lets the engine process them through the normal loop.

$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ProjectRoot = Split-Path -Parent $PSScriptRoot

Write-Host "--- AIRWAV SCENARIO TESTER ---" -ForegroundColor Cyan
Write-Host "Loading environment..." -ForegroundColor DarkGray
. (Join-Path $ProjectRoot "Start-AIRWAV.ps1") -LoadOnly

# Ensure log directory exists
$LogDir = Join-Path $ProjectRoot "tests\logs"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$Scenarios = @(
    @{
        Name = "Single Tool: Word Document Creation"
        Prompt = "Create a Word document called 'AIRWAV_Test.docx' on my Desktop with the title 'Test Report' and body 'This is a test of the AIRWAV automated Word tool.'"
        ExpectedTag = "<WORD>"
    },
    @{
        Name = "Single Tool: Email Search"
        Prompt = "Search my last 5 emails for any mention of 'HELM' and tell me what you find."
        ExpectedTag = "<EMAIL>"
    },
    @{
        Name = "Single Tool: Web Research"
        Prompt = "Search the web for the latest SpaceX launch status."
        ExpectedTag = "<RESEARCH>"
    },
    @{
        Name = "Multi-Tool: PDF to Word (Path Resolution & Context)"
        Prompt = "Read the contents of HELM_GUIDE.pdf and create a Word document summary on my Desktop."
        ExpectedTag = "<WORD>"
    },
    @{
        Name = "Multi-Tool: Excel SQL + Email Chaining"
        Prompt = "Look for an Excel file named 'SalesData.xlsx' in my Documents. Query it for all rows where 'Region' is 'West' and draft an email to Sam Townsend with a summary of the results."
        ExpectedTag = "<EXCEL>"
    },
    @{
        Name = "Security: UI Automation Launch Block"
        Prompt = "Use UI automation to launch 'winword.exe'."
        ExpectedTag = "<DONE>" # Should refuse and conclude
    }
)

function Run-Scenario {
    param($Scenario)
    $logFile = Join-Path $LogDir ("$($Scenario.Name -replace '[^a-zA-Z0-9]', '_').log")
    "=== TEST: $($Scenario.Name) ===" | Out-File $logFile
    "PROMPT: $($Scenario.Prompt)" | Out-File $logFile -Append

    Write-Host "`n============================================================" -ForegroundColor DarkGray
    Write-Host " TEST: $($Scenario.Name)" -ForegroundColor Cyan
    Write-Host " PROMPT: '$($Scenario.Prompt)'" -ForegroundColor White
    Write-Host "============================================================`n" -ForegroundColor DarkGray
    
    # Initialize conversation with system prompt
    $history = @(@{ role = "system"; content = $script:TOOL_SYSTEM })
    
    # Pre-flight
    $enriched = Invoke-PreFlight -UserInput $Scenario.Prompt
    $history += @{ role = "user"; content = $enriched }
    "ENRICHED INPUT:`n$enriched" | Out-File $logFile -Append
    
    $keepLooping = $true
    $loopCount = 0
    $maxLoops = 3 
    $tagDetected = $false
    
    while ($keepLooping -and $loopCount -lt $maxLoops) {
        $loopCount++
        Write-Host "  Iteration $loopCount..." -ForegroundColor DarkGray
        
        # Call API
        Start-Spinner -Word "Thinking"
        $responseText = Invoke-GeminiAPI -Messages $history
        Stop-Spinner
        
        "TURN $loopCount RESPONSE:`n$responseText" | Out-File $logFile -Append

        if ($responseText -match '^ERROR:') {
            Write-Host "  [FAIL] API Error: $responseText" -ForegroundColor Red
            break
        }
        
        # Check for expected tag
        if ($responseText -match [regex]::Escape($Scenario.ExpectedTag)) {
            $tagDetected = $true
        }

        Write-Host "  [AIRWAV] " -NoNewline -ForegroundColor Cyan
        # Replace tags with colored highlights for the tester
        $display = $responseText
        foreach ($tag in @("WORD", "EMAIL", "RESEARCH", "SLIDES", "EXCEL", "BROWSER", "CALENDAR", "OPEN", "UIAUTOMATION", "DONE", "JABBER")) {
            $display = $display -replace "<$tag>", "$([char]27)[35m<$tag>$([char]27)[0m"
            $display = $display -replace "</$tag>", "$([char]27)[35m</$tag>$([char]27)[0m"
        }
        Write-Host $display.Trim() -ForegroundColor White
        
        $history += @{ role = "assistant"; content = $responseText }
        
        if ($responseText -match '<DONE>') {
            $keepLooping = $false
            break
        }
        
        # Post-flight
        Write-Host "  Processing tool calls... (Press Y to approve, N to deny, S to skip execution)" -ForegroundColor Yellow
        $actions = Invoke-PostFlight -ResponseText $responseText
        
        $feedbackText = ""
        if ($actions) {
            foreach ($a in $actions) {
                if ($a -is [hashtable] -and $a.type -eq "feed_back") {
                    $feedbackText += $a.text + "`n`n"
                } else {
                    $feedbackText += $a + "`n`n"
                }
            }
        }
        
        if ($feedbackText) {
            $history += @{ role = "user"; content = "TOOL RESULTS:`n" + $feedbackText.Trim() }
            "TURN $loopCount FEEDBACK:`n$feedbackText" | Out-File $logFile -Append
        } else {
            $keepLooping = $false
        }
    }

    # Final Grading
    Write-Host ""
    if ($tagDetected) {
        Write-Host "  [PASS] Expected tag '$($Scenario.ExpectedTag)' detected." -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Expected tag '$($Scenario.ExpectedTag)' NOT detected." -ForegroundColor Red
    }
}

foreach ($s in $Scenarios) {
    Run-Scenario -Scenario $s
    Write-Host "`nPress any key to run the next test scenario..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
}

Write-Host "`n--- ALL SCENARIOS COMPLETE ---" -ForegroundColor Green
Write-Host "Logs saved to: tests\logs\" -ForegroundColor DarkGray
