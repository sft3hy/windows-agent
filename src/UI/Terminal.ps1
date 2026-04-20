function Write-Banner {
    Clear-Host
    $banner = @"
╔══════════════════════════════════════════════════════════════════════════════╗
║                                                                              ║
║      █████╗ ██╗██████╗ ██╗                                                   ║
║     ██╔══██╗██║██╔══██╗██║                                                   ║
║     ███████║██║██████╔╝██║                                                   ║
║     ██╔══██║██║██╔══██╗██║                                                   ║
║     ██║  ██║██║██║  ██║██║                                                   ║
║     ╚═╝  ╚═╝╚═╝╚═╝  ╚═╝╚═╝                                                   ║
║                                                                              ║
║        AIRI — Automated Intelligence Remote Interface                        ║
║        Version: $($script:CONFIG.Version)                                                        ║
║        Powered by GenAI.mil $($script:CONFIG.Model)                                 ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
"@
    Write-Host $banner -ForegroundColor Cyan
    Write-Host "  Tools: " -NoNewline -ForegroundColor DarkGray
    Write-Host "Outlook | Word | Excel | PowerPoint | Jabber | Chrome | Edge | Firefox | Calendar" -ForegroundColor White
    Write-Host "`n  Type your request in plain English. Type 'exit' to quit, 'help' for examples.`n" -ForegroundColor DarkGray
}

function Write-AgentLine { param([string]$Text, [string]$Color = "Cyan", [switch]$NoNewline)
    Write-Host "  [AIRI] " -NoNewline -ForegroundColor Cyan
    if ($NoNewline) { Write-Host $Text -NoNewline -ForegroundColor $Color }
    else            { Write-Host $Text -ForegroundColor $Color }
}

function Write-ToolLine { param([string]$Tool, [string]$Action, [string]$Detail = "")
    Write-Host "  [TOOL] " -NoNewline -ForegroundColor Magenta
    Write-Host "[$Tool] " -NoNewline -ForegroundColor Yellow
    Write-Host $Action -NoNewline -ForegroundColor White
    if ($Detail) { Write-Host " → $Detail" -ForegroundColor DarkGray } else { Write-Host "" }
}

function Write-StatusLine { param([string]$Status, [string]$Msg)
    $color = switch ($Status) { "OK" {"Green"} "ERR" {"Red"} "WARN" {"Yellow"} "INFO" {"Cyan"} default {"White"} }
    Write-Host "  [$Status] " -NoNewline -ForegroundColor $color
    Write-Host $Msg -ForegroundColor White
}

function Write-Divider { param([string]$Label = "")
    if ($Label) { Write-Host "  ── $Label $("─" * (60 - $Label.Length))" -ForegroundColor DarkGray } 
    else { Write-Host ("  " + "─" * 66) -ForegroundColor DarkGray }
}

function Read-UserInput {
    Write-Host "`n  ┌─ You ─────────────────────────────────────────────────────" -ForegroundColor DarkGreen
    Write-Host "  └▶ " -NoNewline -ForegroundColor DarkGreen
    return (Read-Host)
}

function Show-Help {
    Write-Divider "EXAMPLE REQUESTS"
    @(
      "Read the PDF at C:\Reports\Q1.pdf and create a PowerPoint summary",
      "Draft an email to the team about the project status",
      "Create a Word document summarising today's key wins",
      "Schedule a meeting with Sam Townsend tomorrow at 10am for 1 hour",
      "What's on my calendar for the next 7 days?",
      "Open https://news.google.com in Edge",
      "Search DuckDuckGo for the latest AI news",
      "Lookup the email address for 'John Smith'",
      "Send a Jabber message to jane.doe saying 'See you at the meeting!'"
    ) | ForEach-Object { Write-Host "  • " -NoNewline -ForegroundColor Yellow; Write-Host $_ -ForegroundColor White }
    Write-Divider
}