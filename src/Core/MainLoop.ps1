function Start-AgentSession {
    Write-Banner
    if (-not $script:CONFIG.ApiKey) {
        Write-StatusLine "ERR" "No API key found. Edit Config.ps1"
        return
    }

    $conversationHistory = @(@{ role = "system"; content = $script:TOOL_SYSTEM })

    while ($true) {
        $userInput = Read-UserInput
        if ([string]::IsNullOrWhiteSpace($userInput)) { continue }
        if ($userInput -match '^(exit|quit)$') { Write-AgentLine "Goodbye."; break }
        if ($userInput -eq 'help') { Show-Help; continue }
        if ($userInput -eq 'clear') { Write-Banner; $conversationHistory = @(@{ role = "system"; content = $script:TOOL_SYSTEM }); continue }

        $enrichedInput = Invoke-PreFlight -UserInput $userInput
        $conversationHistory += @{ role = "user"; content = $enrichedInput }

        Write-Host ""
        Start-Spinner -Word (Get-ThinkingWord)
        $responseText = Invoke-GeminiAPI -Messages $conversationHistory
        Stop-Spinner

        $displayText = $responseText
        $displayText = $displayText -replace '<SLIDES>[\s\S]*?</SLIDES>',     '[PowerPoint queued]'
        $displayText = $displayText -replace '<JABBER>[\s\S]*?</JABBER>',     '[Jabber message queued]'
        $displayText = $displayText -replace '<EMAIL>[\s\S]*?</EMAIL>',       '[Email queued]'
        $displayText = $displayText -replace '<WORD>[\s\S]*?</WORD>',         '[Word document queued]'
        $displayText = $displayText -replace '<BROWSER>[\s\S]*?</BROWSER>',   '[Browser action queued]'
        $displayText = $displayText -replace '<CALENDAR>[\s\S]*?</CALENDAR>', '[Calendar action queued]'
        $displayText = $displayText -replace '<OPEN>[\s\S]*?</OPEN>',         '[File open queued]'
        
        Write-Host "`n  [AIRI] " -NoNewline -ForegroundColor Cyan
        Write-Host $displayText.Trim() -ForegroundColor White

        $conversationHistory += @{ role = "assistant"; content = $responseText }

        $actions = Invoke-PostFlight -ResponseText $responseText
        if ($actions) {
            Write-Divider "ACTIONS COMPLETED"
            $actions | ForEach-Object { Write-StatusLine "OK" $_ }
        }
    }
}