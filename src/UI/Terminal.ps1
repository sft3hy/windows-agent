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
        
        $amplitude = 0.7; $frequency = 0.3; $speed = 0.15; $frames = 30
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
    param([string]$Topic = "")
    $Topic = $Topic.Trim()
    
    if (-not $Topic) {
        Write-Divider "GENERAL HELP"
        Write-Host "  Type 'help <topic>' for more details. Available topics:" -ForegroundColor White
        Write-Host "  • Outlook / Email" -ForegroundColor Yellow
        Write-Host "  • Excel" -ForegroundColor Yellow
        Write-Host "  • Word" -ForegroundColor Yellow
        Write-Host "  • PowerPoint" -ForegroundColor Yellow
        Write-Host "  • FileSystem / PDF" -ForegroundColor Yellow
        Write-Host "  • Web / Browser" -ForegroundColor Yellow
        Write-Host "  • Calendar" -ForegroundColor Yellow
        Write-Host "  • Jabber" -ForegroundColor Yellow
        Write-Divider "QUICK EXAMPLES"
        @(
          "Read the PDF at C:\Reports\Q1.pdf and create a PowerPoint summary",
          "Draft an email to the team about the project status",
          "What's on my calendar for the next 7 days?"
        ) | ForEach-Object { Write-Host "  • " -NoNewline -ForegroundColor Yellow; Write-Host $_ -ForegroundColor White }
        Write-Divider
        return
    }

    Write-Divider "HELP: $Topic"
    switch -Regex ($Topic) {
        "(?i)excel" {
            Write-Host "  Excel Automation:" -ForegroundColor Cyan
            Write-Host "  • Read/Write cells and ranges" -ForegroundColor White
            Write-Host "  • Convert raw data or CSVs into new Workbooks" -ForegroundColor White
            Write-Host "  • Apply heatmaps and beautify tables" -ForegroundColor White
            Write-Host "  • Add Charts (Bar, Line, Pie)" -ForegroundColor White
            Write-Host "  • Query data using SQL directly on sheets" -ForegroundColor White
            Write-Host "  • Remove duplicates" -ForegroundColor White
            Write-Host "  • Export whole workbooks to PDF" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > Please beautify the sales report at C:\Data.xlsx" -ForegroundColor Yellow
            Write-Host "  > Query C:\Data.xlsx and tell me who has the most sales, then create a bar chart." -ForegroundColor Yellow
            Write-Host "  > Export my excel file to a PDF." -ForegroundColor Yellow
        }
        "(?i)pdf|file" {
            Write-Host "  FileSystem & PDF:" -ForegroundColor Cyan
            Write-Host "  • Read text from local files (TXT, CSV, MD, JSON, LOG)" -ForegroundColor White
            Write-Host "  • Extract text securely from PDFs and Word Documents" -ForegroundColor White
            Write-Host "  • Find files on your Desktop or Documents without exact paths" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > Read the Q3 Report pdf from my documents and summarize it." -ForegroundColor Yellow
            Write-Host "  > Can you read the system.log file on my desktop?" -ForegroundColor Yellow
        }
        "(?i)powerpoint|ppt" {
            Write-Host "  PowerPoint Automation:" -ForegroundColor Cyan
            Write-Host "  • Generate new slide decks from scratch" -ForegroundColor White
            Write-Host "  • Open existing presentations" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > Read my Word doc and convert it into a 5-slide PowerPoint presentation." -ForegroundColor Yellow
            Write-Host "  > Create a presentation about AI with 3 slides." -ForegroundColor Yellow
        }
        "(?i)outlook|email" {
            Write-Host "  Outlook / Email:" -ForegroundColor Cyan
            Write-Host "  • Read your Inbox or other folders" -ForegroundColor White
            Write-Host "  • Send emails (with attachments and auto-resolving contacts)" -ForegroundColor White
            Write-Host "  • Draft emails for review" -ForegroundColor White
            Write-Host "  • Reply, Forward, Delete, or Flag specific emails" -ForegroundColor White
            Write-Host "  • Lookup Active Directory contacts" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > Draft an email to Sam Townsend with the Q1 report attached." -ForegroundColor Yellow
            Write-Host "  > Read my last 5 emails and reply to the one from the boss." -ForegroundColor Yellow
            Write-Host "  > What is John Doe's email address?" -ForegroundColor Yellow
        }
        "(?i)word" {
            Write-Host "  Word Document Automation:" -ForegroundColor Cyan
            Write-Host "  • Create new documents with formatted titles and text" -ForegroundColor White
            Write-Host "  • Append new text to existing documents" -ForegroundColor White
            Write-Host "  • Perform global Find and Replace" -ForegroundColor White
            Write-Host "  • Open documents for viewing" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > Draft a meeting agenda in Word and save it to my Desktop." -ForegroundColor Yellow
            Write-Host "  > In my template.docx, replace [NAME] with Alice." -ForegroundColor Yellow
        }
        "(?i)web|browser" {
            Write-Host "  Web & Browser Automation:" -ForegroundColor Cyan
            Write-Host "  • Perform web research via search engines" -ForegroundColor White
            Write-Host "  • Fetch and read text content directly from websites" -ForegroundColor White
            Write-Host "  • Open URLs in Chrome or Edge" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > Research the latest news on SpaceX." -ForegroundColor Yellow
            Write-Host "  > Go to https://docs.microsoft.com and summarize the page." -ForegroundColor Yellow
            Write-Host "  > Open google.com in Chrome." -ForegroundColor Yellow
        }
        "(?i)calendar" {
            Write-Host "  Calendar Automation:" -ForegroundColor Cyan
            Write-Host "  • Read upcoming appointments and meetings" -ForegroundColor White
            Write-Host "  • Schedule new calendar events" -ForegroundColor White
            Write-Host "  • Delete existing meetings" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > What is on my calendar for today?" -ForegroundColor Yellow
            Write-Host "  > Schedule a 1 hour sync with the engineering team for tomorrow at 2 PM." -ForegroundColor Yellow
        }
        "(?i)jabber" {
            Write-Host "  Jabber Automation:" -ForegroundColor Cyan
            Write-Host "  • Send instant messages" -ForegroundColor White
            Write-Host "  • Open chat windows" -ForegroundColor White
            Write-Host "`n  Examples:" -ForegroundColor DarkGray
            Write-Host "  > Send a Jabber message to Sam saying 'Hello!'" -ForegroundColor Yellow
        }
        default {
            Write-Host "  Unknown topic: $Topic" -ForegroundColor Red
            Write-Host "  Try 'help' for a list of available topics." -ForegroundColor White
        }
    }
    Write-Divider
}