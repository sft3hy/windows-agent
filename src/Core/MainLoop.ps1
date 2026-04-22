function Start-AgentSession {
    Start-AnimatedBanner
    if (-not $script:CONFIG.ApiKey) {
        Write-StatusLine "ERR" "No API key found. Edit Config.ps1"
        return
    }

    $conversationHistory = @(@{ role = "system"; content = $script:TOOL_SYSTEM })

    while ($true) {
        $userInput = Read-UserInput
        if ([string]::IsNullOrWhiteSpace($userInput)) { continue }
        if ($userInput -match '^(exit|quit)$') { Write-AgentLine "Goodbye."; break }
        if ($userInput -match '(?i)^help\s*(.*)$') { Show-Help -Topic $Matches[1]; continue }
        if ($userInput -eq 'clear') { Write-Banner; $conversationHistory = @(@{ role = "system"; content = $script:TOOL_SYSTEM }); continue }

        $enrichedInput = Invoke-PreFlight -UserInput $userInput
        $conversationHistory += @{ role = "user"; content = $enrichedInput }

        $keepLooping = $true
        while ($keepLooping) {
            Write-Host ""
            Start-Spinner -Word (Get-ThinkingWord)
            $responseText = Invoke-GeminiAPI -Messages $conversationHistory
            Stop-Spinner

            $displayText = $responseText
            $displayText = $displayText -replace '<SLIDES>[\s\S]*?</SLIDES>',     '[PowerPoint queued]'
            $displayText = $displayText -replace '<JABBER>[\s\S]*?</JABBER>',     '[Jabber message queued]'
            $displayText = $displayText -replace '<EMAIL>[\s\S]*?</EMAIL>',       '[Email queued]'
            $displayText = $displayText -replace '<WORD>[\s\S]*?</WORD>',         '[Word document queued]'
            $displayText = $displayText -replace '<EXCEL>[\s\S]*?</EXCEL>',        '[Excel action queued]'
            $displayText = $displayText -replace '<BROWSER>[\s\S]*?</BROWSER>',   '[Browser action queued]'
            $displayText = $displayText -replace '<RESEARCH>[\s\S]*?</RESEARCH>', '[Web research queued]'
            $displayText = $displayText -replace '<CALENDAR>[\s\S]*?</CALENDAR>', '[Calendar action queued]'
            $displayText = $displayText -replace '<OPEN>[\s\S]*?</OPEN>',         '[File open queued]'
            $displayText = $displayText -replace '<UIAUTOMATION>[\s\S]*?</UIAUTOMATION>', '[UI Automation action queued]'
            $displayText = $displayText -replace '<DONE>', ''
            
            if ($displayText.Trim()) {
                Write-Host "`n  [AIRWAV] " -NoNewline -ForegroundColor Cyan
                Write-Host $displayText.Trim() -ForegroundColor White
            }

            $conversationHistory += @{ role = "assistant"; content = $responseText }

            if ($responseText -match '<DONE>') {
                $keepLooping = $false
                break
            }

            $actions = Invoke-PostFlight -ResponseText $responseText
            $feedbackText = ""
            
            if ($actions -and $actions.Count -gt 0) {
                Write-Divider "ACTIONS COMPLETED"
                foreach ($a in $actions) {
                    if ($a -is [hashtable] -and $a.type -eq "feed_back") {
                        Write-StatusLine "OK" "Data scraped; returning to AI for analysis."
                        $feedbackText += $a.text + "`n`n"
                    } else {
                        Write-StatusLine "OK" $a
                        $feedbackText += $a + "`n`n"
                    }
                }
            }

            if ($feedbackText) {
                $conversationHistory += @{ role = "user"; content = "TOOL RESULTS:`n" + $feedbackText.Trim() }
            } else {
                # Infinite loop prevention
                $conversationHistory += @{ role = "user"; content = "SYSTEM: You did not output any tool tags, and you did not output <DONE>. If the user's request is complete, you MUST output <DONE>. Otherwise, use a tool." }
            }
        }
    }
}