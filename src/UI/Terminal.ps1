$script:AIRWAV_ASCII = @(
    "           █████╗  ██╗ ██████╗  ██╗    ██╗  █████╗  ██╗   ██╗"
    "          ██╔══██╗ ██║ ██╔══██╗ ██║    ██║ ██╔══██╗ ██║   ██║"
    "          ███████║ ██║ ██████╔╝ ██║ █╗ ██║ ███████║ ██║   ██║"
    "          ██╔══██║ ██║ ██╔══██╗ ██║███╗██║ ██╔══██║ ╚██╗ ██╔╝"
    "          ██║  ██║ ██║ ██║  ██║ ╚███╔███╔╝ ██║  ██║  ╚████╔╝"
    "          ╚═╝  ╚═╝ ╚═╝ ╚═╝  ╚═╝  ╚══╝╚══╝  ╚═╝  ╚═╝   ╚═══╝"
)

function Format-ToolStr {
    param([string]$Tool)
    $esc = "$([char]27)"
    $reset = "$esc[0m"
    switch ($Tool.ToLower()) {
        "outlook"    { return "$esc[38;2;0;114;198m$Tool$reset" }
        "word"       { return "$esc[38;2;43;87;154m$Tool$reset" }
        "excel"      { return "$esc[38;2;29;111;71m$Tool$reset" }
        "powerpoint" { return "$esc[38;2;208;68;35m$Tool$reset" }
        "jabber"     { return "$esc[38;2;0;188;235m$Tool$reset" }
        "chrome"     { 
            $cr = "$esc[38;2;234;67;53m"
            $cy = "$esc[38;2;251;188;4m"
            $cb = "$esc[38;2;71;133;234m"
            $cg = "$esc[38;2;52;168;83m"
            if ($Tool -eq "Chrome") { return "${cr}C${cb}h${cg}r${cy}o${cr}m${cb}e$reset" }
            else { return "${cr}c${cy}hro${cb}me$reset" }
        }
        "edge" { return "$esc[38;2;43;195;210m$Tool$reset" }
        "firefox"    { return "$esc[38;2;255;102;17m$Tool$reset" }
        "calendar"   { return "$esc[38;2;0;120;212m$Tool$reset" }
        default      { return "$esc[38;2;255;255;255m$Tool$reset" }
    }
}

function Get-ToolsLine {
    $tools = @("Outlook", "Word", "Excel", "PowerPoint", "Jabber", "Chrome", "Edge", "Firefox", "Calendar")
    $esc = "$([char]27)"
    $colored = $tools | ForEach-Object { Format-ToolStr $_ }
    return ($colored -join "$esc[38;2;128;128;128m | $esc[0m")
}

function Write-Banner {
    param([switch]$NoClear)
    if (-not $NoClear) { Clear-Host }
    else { try { [Console]::SetCursorPosition(0, 0) } catch {} }
    
    $topLine    = "╔═══════════════════════════════════════════════════════════════════════╗"
    $emptyLine  = "║                                                                       ║"
    $bottomLine = "╚═══════════════════════════════════════════════════════════════════════╝"
    $version    = "               Version: $($script:CONFIG.Version)".PadRight(71)
    $powered    = "               Powered by GenAI.mil $($script:CONFIG.Model)".PadRight(71)
    $airwavLine = "               AIRWAV — AI Runtime for Windows Automation & Vision".PadRight(71)
    
    Write-Host $topLine -ForegroundColor Cyan
    Write-Host $emptyLine -ForegroundColor Cyan
    Write-Host $emptyLine -ForegroundColor Cyan
    
    foreach ($line in $script:AIRWAV_ASCII) {
        Write-Host "║" -NoNewline -ForegroundColor Cyan
        Write-Host $line.PadRight(71) -NoNewline -ForegroundColor Cyan
        Write-Host "║" -ForegroundColor Cyan
    }

    Write-Host $emptyLine -ForegroundColor Cyan
    Write-Host $emptyLine -ForegroundColor Cyan
    Write-Host "║$airwavLine║" -ForegroundColor Cyan
    Write-Host "║$version║" -ForegroundColor Cyan
    Write-Host "║$powered║" -ForegroundColor Cyan
    Write-Host $emptyLine -ForegroundColor Cyan
    Write-Host $bottomLine -ForegroundColor Cyan

    Write-Host "  Tools: " -NoNewline -ForegroundColor DarkGray
    Write-Host $(Get-ToolsLine)
    Write-Host "  Type your request in plain English. Type 'exit' to quit, 'help' for examples.`n" -ForegroundColor DarkGray
}

function Start-AnimatedBanner {
    if (-not $script:CONFIG.ApiKey) { Write-Banner; return }
    Clear-Host
    
    $script:BannerShared = [hashtable]::Synchronized(@{
        Running = $true
        ASCII   = $script:AIRWAV_ASCII
        Version = $script:CONFIG.Version
        Model   = $script:CONFIG.Model
        ToolsStr= (Get-ToolsLine)
    })

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"; $rs.ThreadOptions = "ReuseThread"; $rs.Open()
    $rs.SessionStateProxy.SetVariable("Shared", $script:BannerShared)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        $topLine    = "╔═══════════════════════════════════════════════════════════════════════╗"
        $emptyLine  = "║                                                                       ║"
        $bottomLine = "╚═══════════════════════════════════════════════════════════════════════╝"
        $version    = "          Version: $($Shared.Version)".PadRight(71)
        $powered    = "          Powered by GenAI.mil $($Shared.Model)".PadRight(71)
        $airwavLine = "          AIRWAV — AI Runtime for Windows Automation & Vision".PadRight(71)
        
        $amplitude = 0.7; $frequency = 0.3; $speed = 0.15; $frames = 150
        $maxLen = 0; foreach ($line in $Shared.ASCII) { if ($line.Length -gt $maxLen) { $maxLen = $line.Length } }

        # Detect letter groups: contiguous non-space column runs
        $letterGroups = [System.Collections.Generic.List[hashtable]]::new()
        $inLetter = $false; $groupStart = 0
        for ($x = 0; $x -lt $maxLen; $x++) {
            $hasContent = $false
            foreach ($line in $Shared.ASCII) {
                if ($x -lt $line.Length -and $line[$x] -ne ' ') { $hasContent = $true; break }
            }
            if ($hasContent -and -not $inLetter) { $groupStart = $x; $inLetter = $true }
            elseif (-not $hasContent -and $inLetter) { $letterGroups.Add(@{ S = $groupStart; E = $x - 1 }); $inLetter = $false }
        }
        if ($inLetter) { $letterGroups.Add(@{ S = $groupStart; E = $maxLen - 1 }) }

        # Build a lookup: column -> letter group index (-1 = space)
        $colGroup = [int[]]::new($maxLen)
        for ($x = 0; $x -lt $maxLen; $x++) { $colGroup[$x] = -1 }
        for ($g = 0; $g -lt $letterGroups.Count; $g++) {
            for ($x = $letterGroups[$g].S; $x -le $letterGroups[$g].E; $x++) { $colGroup[$x] = $g }
        }
        $numLetters = $letterGroups.Count

        for ($f = 0; $f -lt $frames -and $Shared.Running; $f++) {
            $t = $f / $frames
            $phase = $f * $speed
            # Settle front sweeps over letter indices: starts at 0.2, ends at 1.0
            $letterFront = if ($t -lt 0.2) { -1.0 } else { ($t - 0.2) / 0.8 * $numLetters }
            
            $animBuffer = @()
            for ($i=0; $i -lt 8; $i++) { $animBuffer += [System.Text.StringBuilder]::new((" " * 71)) }

            for ($x = 0; $x -lt $maxLen; $x++) {
                $g = $colGroup[$x]
                # Space column — skip
                if ($g -lt 0) { continue }
                # Per-letter damping: smooth cosine over a range of 1 letter-unit
                $distFromFront = $g - $letterFront
                $damp = if ($distFromFront -le -1.0) { 0.0 } `
                        elseif ($distFromFront -ge 0) { 1.0 } `
                        else { ([Math]::Cos($distFromFront * [Math]::PI) + 1.0) / 2.0 }
                $yOff = [Math]::Round($amplitude * $damp * [Math]::Sin($phase - ($x * $frequency)))
                for ($y = 0; $y -lt $Shared.ASCII.Count; $y++) {
                    if ($x -lt $Shared.ASCII[$y].Length) {
                        $c = $Shared.ASCII[$y][$x]
                        if ($c -ne ' ') {
                            $dy = $y + 1 + $yOff
                            if ($dy -ge 0 -and $dy -lt 8) { $animBuffer[$dy][$x] = $c }
                        }
                    }
                }
            }

            # Save and Restore Cursor for non-blocking feel
            $origL = [Console]::CursorLeft; $origT = [Console]::CursorTop
            [Console]::SetCursorPosition(0, 0); [Console]::ForegroundColor = "Cyan"
            
            [Console]::WriteLine($topLine)
            [Console]::WriteLine($emptyLine); [Console]::WriteLine($emptyLine)
            for ($i=0; $i -lt 8; $i++) { [Console]::WriteLine("║" + $animBuffer[$i].ToString() + "║") }
            [Console]::WriteLine($emptyLine)
            [Console]::WriteLine("║$airwavLine║"); [Console]::WriteLine("║$version║"); [Console]::WriteLine("║$powered║")
            [Console]::WriteLine($emptyLine); [Console]::WriteLine($bottomLine)
            [Console]::WriteLine("  Tools: " + $Shared.ToolsStr)
            [Console]::ForegroundColor = "DarkGray"
            [Console]::WriteLine("  Type your request in plain English. Type 'exit' to quit, 'help' for examples.")
            [Console]::WriteLine("")
            [Console]::ForegroundColor = "DarkGreen"
            [Console]::Write("  └▶ ")
            [Console]::ResetColor()

            # Park cursor at the input position (after └▶ ) so keypresses land there
            try { [Console]::SetCursorPosition(5, 20) } catch {}
            [Threading.Thread]::Sleep(25)
        }

        # Final frame: draw all letters at yOff=0 to guarantee a clean horizontal baseline
        $finalBuffer = @()
        for ($i=0; $i -lt 8; $i++) { $finalBuffer += [System.Text.StringBuilder]::new((" " * 71)) }
        for ($x = 0; $x -lt $maxLen; $x++) {
            for ($y = 0; $y -lt $Shared.ASCII.Count; $y++) {
                if ($x -lt $Shared.ASCII[$y].Length) {
                    $c = $Shared.ASCII[$y][$x]
                    if ($c -ne ' ') {
                        $dy = $y + 1  # zero offset
                        if ($dy -ge 0 -and $dy -lt 8) { $finalBuffer[$dy][$x] = $c }
                    }
                }
            }
        }
        $origL = [Console]::CursorLeft; $origT = [Console]::CursorTop
        [Console]::SetCursorPosition(0, 0); [Console]::ForegroundColor = "Cyan"
        [Console]::WriteLine($topLine)
        [Console]::WriteLine($emptyLine); [Console]::WriteLine($emptyLine)
        for ($i=0; $i -lt 8; $i++) { [Console]::WriteLine("║" + $finalBuffer[$i].ToString() + "║") }
        [Console]::WriteLine($emptyLine)
        [Console]::WriteLine("║$airwavLine║"); [Console]::WriteLine("║$version║"); [Console]::WriteLine("║$powered║")
        [Console]::WriteLine($emptyLine); [Console]::WriteLine($bottomLine)
        [Console]::WriteLine("  Tools: " + $Shared.ToolsStr)
        [Console]::ForegroundColor = "DarkGray"
        [Console]::WriteLine("  Type your request in plain English. Type 'exit' to quit, 'help' for examples.")
        [Console]::WriteLine("")
        [Console]::ForegroundColor = "DarkGreen"
        [Console]::Write("  └▶ ")
        [Console]::ResetColor()
        try { [Console]::SetCursorPosition(5, 20) } catch {}

        $Shared.Running = $false
    })
    $script:BannerPS = $ps; $script:BannerRS = $rs; $script:BannerHandle = $ps.BeginInvoke()
}

function Stop-AnimatedBanner {
    if ($script:BannerShared) {
        $script:BannerShared.Running = $false
        Start-Sleep -Milliseconds 100
        try { 
            $script:BannerPS.EndInvoke($script:BannerHandle); $script:BannerPS.Dispose(); $script:BannerRS.Close() 
            $script:BannerShared = $null
        } catch {}
        # Final static draw to ensure perfect settlement
        Write-Banner -NoClear
    }
}

function Write-AgentLine { param([string]$Text, [string]$Color = "Cyan", [switch]$NoNewline)
    Write-Host "  [AIRWAV] " -NoNewline -ForegroundColor Cyan
    if ($NoNewline) { Write-Host $Text -NoNewline -ForegroundColor $Color }
    else            { Write-Host $Text -ForegroundColor $Color }
}

function Write-ToolLine { param([string]$Tool, [string]$Action, [string]$Detail = "")
    Write-Host "  [TOOL] " -NoNewline -ForegroundColor Magenta
    Write-Host "[$([char]27)[0m$(Format-ToolStr $Tool)$([char]27)[33m] " -NoNewline -ForegroundColor Yellow
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
    else { Write-Host ("  " + "─" * 71) -ForegroundColor DarkGray }
}

function Read-UserInput {
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