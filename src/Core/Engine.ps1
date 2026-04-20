$script:TOOL_SYSTEM = @"
IDENTITY:
You are AIRWAV — Automated Intelligence Remote Interface.
You are powered by Gemini 2.5 Flash, running LOCALLY on the user's Windows workstation.
You have direct access to: Outlook, Word, Excel, PowerPoint, Jabber, Chrome, Edge, Firefox, Local filesystem, Outlook Calendar.

CRITICAL RULES:
1. USE EXISTING FILES in priority. If a file is read into context (see FILE CONTENTS), use its content.
2. DO NOT output <SLIDES> to recreate a file that already exists unless asked to MODIFY it.
3. OUTPUT FORMAT (use ONLY ONE tag per action type):
   - PowerPoint (new/update): <SLIDES>[{"title": "t", "content": "c"}]</SLIDES>
   - Jabber:    <JABBER>{"recipient": "user", "message": "text"}</JABBER>
   - Email:     <EMAIL>{"action": "draft|reply|forward|delete|flag", "to": "name", "subject": "subj", "body": "txt", "query": "email to reply/fwd/del", "attachments": []}</EMAIL>
   - Word doc:  <WORD>{"action": "create|open|append|replace", "path": "file.docx", "title": "t", "body": "content", "search": "old", "replace": "new"}</WORD>
   - Excel:     <EXCEL>{"action": "read|create|write", "path": "file.xlsx", "cell": "B2", "value": "100"}</EXCEL>
   - Browser:   <BROWSER>{"url": "https://...", "browser": "chrome|edge|firefox"}</BROWSER>
   - Search:    <BROWSER>{"search": "query text", "engine": "google|bing", "browser": ""}</BROWSER>
   - Calendar:  <CALENDAR>{"action": "create|read|delete", "subject": "matching subj", "start": "2026-04-17 14:00", "end": "2026-04-17 15:00", "location": "", "body": "", "attendees": []}</CALENDAR>
   - Open file: <OPEN>{"path": "Documents/file.ext"}</OPEN>
   - UI Task:   <UIAUTOMATION>{"action": "launch|inspect|click|type", "path": "notepad.exe", "window": "Notepad", "element_name": "File", "id": "12", "text": "hello"}</UIAUTOMATION>
4. For email body text, keep it on ONE line using \n for line breaks. Do NOT use smart quotes or markdown in JSON values.
   IMPORTANT: In ALL JSON path values, use FORWARD SLASHES only (e.g. "Documents/Report.docx"). Do NOT use backslashes. Use relative paths like "Desktop/file.docx" or "Documents/file.docx" — the engine resolves them automatically.
5. Multiple email recipients: put them ALL in the "to" field separated by commas, e.g. "to": "Sam Price, Demetra Drizis". Never split them into separate EMAIL tags.
6. When you create a PowerPoint and then email it, set "attachments": ["AIRWAV_Presentation.pptx"] — the engine will resolve the real path automatically.
7. EMAIL action is ALWAYS "draft" — never "send". The user reviews and sends from Outlook themselves. Do not say you sent an email.

User paths: Desktop=$($script:ENV_PATHS.Desktop) Downloads=$($script:ENV_PATHS.Downloads) Documents=$($script:ENV_PATHS.Documents)
"@

function Find-FilePaths {
    param([string]$Text)
    $paths = @()
    # 1. Full UNC or Local paths
    foreach ($m in [regex]::Matches($Text, '\\\\[^\s"'']+')) { $paths += $m.Value }
    foreach ($m in [regex]::Matches($Text, '[A-Za-z]:\\[^\s"'']+')) { $paths += $m.Value }

    # 2. Relative "Desktop\file" or "Downloads\file" style paths
    foreach ($m in [regex]::Matches($Text, '(?i)(Downloads|Desktop|Documents)[/\\][^\s"'']+')) {
        $rel = $m.Value -replace '(?i)^Downloads', $script:ENV_PATHS.Downloads `
                        -replace '(?i)^Desktop',   $script:ENV_PATHS.Desktop `
                        -replace '(?i)^Documents', $script:ENV_PATHS.Documents
        $paths += $rel
    }

    # 3. Bare filenames — try with full name first (preserving ext), then without ext
    #    Also handles filenames with spaces via quoted-string patterns
    $barePattern = '(?i)\b([a-z0-9_\-\. ]+\.(?:pptx|pdf|docx|txt|xlsx|csv|md))\b'
    foreach ($m in [regex]::Matches($Text, $barePattern)) {
        $filename = $m.Value.Trim()
        if ($filename -notmatch '[:\\]') {
            # Try full filename first (Resolve-FuzzyFilePath with ext already included)
            $resolved = Resolve-FuzzyFilePath -Name $filename
            if (-not $resolved) {
                # Strip extension and try again
                $resolved = Resolve-FuzzyFilePath -Name ($filename -replace '\.[^.]+$', '')
            }
            if ($resolved) { $paths += $resolved }
        }
    }

    return $paths | Sort-Object -Unique -Property { $_.ToLower() }
}

function Invoke-PreFlight {
    param([string]$UserInput)
    $filePaths = Find-FilePaths -Text $UserInput
    $fileContents = @()
    foreach ($fp in $filePaths) {
        Write-ToolLine "File" "Auto-reading" $fp
        $result = Invoke-ReadFile -FilePath $fp
        if ($result -notmatch '^ERROR:') { $fileContents += "`n--- FILE: $fp ---`n$result`n--- END ---`n" }
    }
    
    $enriched = $UserInput
    if ($fileContents) { $enriched += "`n`nFILE CONTENTS:`n" + ($fileContents -join "`n") }
    
    # Only remind if file content was NOT found for that type
    if ($UserInput -match '(?i)(powerpoint|pptx)' -and $enriched -notmatch '\.pptx ---') { 
        $enriched += "`n`nTASK REMINDER: Output <SLIDES> JSON array if creating a NEW presentation." 
    }
    if ($UserInput -match '(?i)(jabber|im\b|chat)') { 
        $enriched += "`n`nTASK REMINDER: Output <JABBER> JSON block for messaging." 
    }
    if ($UserInput -match '(?i)(email|outlook|draft)') { 
        $enriched += "`n`nTASK REMINDER: Output <EMAIL> JSON block for Outlook actions." 
    }
    
    return $enriched
}

# Normalize JSON emitted by the LLM: fix bad escapes from Windows paths.
# A regex cannot distinguish \\n (escaped-backslash then 'n') from \n (newline escape)
# from \n (start of path like \nro.mil).  We use a sequential scanner instead.
function ConvertFrom-LlmJson {
    param([string]$Raw)
    # 1. Strip hallucinated markdown escapes: \_ \* \# etc.
    $s = $Raw -replace '\\([_*#%])', '$1'
    # 2. Walk the string left-to-right, fixing invalid JSON escape sequences.
    #    Valid JSON escapes after \: " \ / b f n r t u
    #    Anything else (like \h \S \R from Windows paths) gets doubled: \ → \\
    $sb = New-Object System.Text.StringBuilder ($s.Length + 64)
    $i = 0
    while ($i -lt $s.Length) {
        $ch = $s[$i]
        if ($ch -eq '\' -and ($i + 1) -lt $s.Length) {
            $nxt = $s[$i + 1]
            if ($nxt -in '"', '\', '/', 'b', 'f', 'n', 'r', 't', 'u') {
                # Valid JSON escape pair — keep both chars
                [void]$sb.Append($ch)
                [void]$sb.Append($nxt)
                $i += 2
            } else {
                # Invalid escape like \h \S \R — double the backslash to make \\X
                [void]$sb.Append('\')
                [void]$sb.Append('\')
                # Do NOT consume $nxt — let it be processed as a normal char next pass
                $i += 1
            }
        } else {
            [void]$sb.Append($ch)
            $i += 1
        }
    }
    return $sb.ToString() | ConvertFrom-Json
}

# Resolve an array of attachment values (may be bare filenames) to full paths
function Resolve-Attachments {
    param([array]$Attachments)
    $resolved = @()
    foreach ($a in $Attachments) {
        if (-not $a) { continue }
        $a = [string]$a
        if (Test-Path $a) { $resolved += $a; continue }
        # Try fuzzy resolve: strip extension, search common folders
        $noExt  = $a -replace '\.[^.]+$', ''
        $hit    = Resolve-FuzzyFilePath -Name $a   # try full name first
        if (-not $hit) { $hit = Resolve-FuzzyFilePath -Name $noExt }
        # Final fallback: if the attachment is a .pptx and we just created one this session, use it
        if (-not $hit -and $a -match '\.pptx$' -and $script:LastCreatedPptx -and (Test-Path $script:LastCreatedPptx)) {
            Write-StatusLine "INFO" "Using last created PPTX for '$a': $script:LastCreatedPptx"
            $hit = $script:LastCreatedPptx
        }
        # Same fallback for .docx
        if (-not $hit -and $a -match '\.docx$' -and $script:LastCreatedDocx -and (Test-Path $script:LastCreatedDocx)) {
            Write-StatusLine "INFO" "Using last created DOCX for '$a': $script:LastCreatedDocx"
            $hit = $script:LastCreatedDocx
        }
        if ($hit) { Write-StatusLine "INFO" "Resolved attachment '$a' → $hit"; $resolved += $hit }
        else      { Write-StatusLine "WARN" "Attachment not found: $a" }
    }
    return $resolved
}

function Invoke-PostFlight {
    param([string]$ResponseText)
    $actionsRun = @()

    # ═══════════════════════════════════════════════════════════════════════════
    # EXECUTION ORDER matters!  Dependencies flow left-to-right:
    #   1. Browser/Search  (gather info)
    #   2. PPTX / Word     (create files — may be attached later)
    #   3. Email           (draft with attachments that now exist on disk)
    #   4. Jabber / Calendar / Open  (independent actions)
    # ═══════════════════════════════════════════════════════════════════════════

    # ── 1. Browser ───────────────────────────────────────────────────────────
    if ($ResponseText -match '<BROWSER>([\s\S]*?)</BROWSER>') {
        $browserRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $browserRaw
            $browser = if ($obj.browser) { $obj.browser } else { "" }
            if ($obj.search) {
                $engine = if ($obj.engine) { $obj.engine } else { "google" }
                Write-Host "  Search '$($obj.search)' in $browser? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-BrowserSearch -Query $obj.search -Engine $engine -Browser $browser }
            } elseif ($obj.url) {
                Write-Host "  Open '$($obj.url)' in $browser? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-BrowserOpen -Url $obj.url -Browser $browser }
            }
        } catch { Write-StatusLine "ERR" "Browser JSON parse failed: $_" }
    }

    # ── 2a. PowerPoint ───────────────────────────────────────────────────────
    if ($ResponseText -match '<SLIDES>([\s\S]*?)</SLIDES>') {
        $slidesRaw = $Matches[1].Trim()
        $savePath = Join-Path $script:ENV_PATHS.Desktop "AIRWAV_Presentation.pptx"
        Write-Host "  Generate PPTX? [Y/N]: " -NoNewline -ForegroundColor Yellow
        if ((Read-Host) -match '^[Yy]') {
            $pptResult = Invoke-PowerPointCreate -FilePath $savePath -SlidesJson $slidesRaw
            if ($pptResult -and (Test-Path $pptResult)) {
                $script:LastCreatedPptx = $pptResult
                $actionsRun += "PowerPoint created: $pptResult"
            } else {
                $actionsRun += "PowerPoint creation failed"
            }
        }
    }

    # ── 2b. Word Document ────────────────────────────────────────────────────
    if ($ResponseText -match '<WORD>([\s\S]*?)</WORD>') {
        $wordRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $wordRaw
            # Resolve relative / forward-slash paths from the model
            $docPath = if ($obj.path) { $obj.path } else { "Desktop/AIRWAV_Document.docx" }
            $docPath = $docPath -replace '/', '\'
            if ($docPath -match '^(?i)(Desktop|Documents|Downloads)[\\\/]') {
                $docPath = $docPath -replace '(?i)^Desktop',   $script:ENV_PATHS.Desktop `
                                   -replace '(?i)^Documents', $script:ENV_PATHS.Documents `
                                   -replace '(?i)^Downloads', $script:ENV_PATHS.Downloads
            } elseif ($docPath -notmatch '^[A-Za-z]:\\' -and $docPath -notmatch '^\\\\') {
                $docPath = Join-Path $script:ENV_PATHS.Desktop $docPath
            }
            $action = if ($obj.action) { $obj.action } else { "create" }
            Write-Host "  $action Word doc at $docPath? [Y/N]: " -NoNewline -ForegroundColor Yellow
            if ((Read-Host) -match '^[Yy]') {
                if ($action -eq "open") {
                    $actionsRun += Invoke-WordOpenDocument -FilePath $docPath
                } elseif ($action -eq "append") {
                    $actionsRun += Invoke-WordAppendText -FilePath $docPath -Text $obj.body
                } elseif ($action -eq "replace") {
                    $actionsRun += Invoke-WordReplaceText -FilePath $docPath -SearchText $obj.search -ReplaceText $obj.replace
                } else {
                    $result = Invoke-WordCreateDocument -FilePath $docPath -Title $obj.title -Body $obj.body -Display
                    $actionsRun += $result
                    if ($result -and $result -notmatch '^ERROR' -and (Test-Path $docPath)) {
                        $script:LastCreatedDocx = $docPath
                        Write-StatusLine "OK" "DOCX ready for attachment: $docPath"
                    }
                }
            }
        } catch { Write-StatusLine "ERR" "Word JSON parse failed: $_" }
    }
    
    # ── 2c. Excel ────────────────────────────────────────────────────────────
    if ($ResponseText -match '<EXCEL>([\s\S]*?)</EXCEL>') {
        $excelRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $excelRaw
            $exPath = $obj.path -replace '/', '\'
            if ($exPath -match '^(?i)(Desktop|Documents|Downloads)[\\\/]') {
                $exPath = $exPath -replace '(?i)^Desktop', $script:ENV_PATHS.Desktop -replace '(?i)^Documents', $script:ENV_PATHS.Documents -replace '(?i)^Downloads', $script:ENV_PATHS.Downloads
            } elseif ($exPath -notmatch '^[A-Za-z]:\\' -and $exPath -notmatch '^\\\\') { $exPath = Join-Path $script:ENV_PATHS.Desktop $exPath }
            
            $action = if ($obj.action) { $obj.action } else { "read" }
            Write-Host "  $action Excel file at $exPath? [Y/N]: " -NoNewline -ForegroundColor Yellow
            if ((Read-Host) -match '^[Yy]') {
                if ($action -eq "write") {
                    $actionsRun += Invoke-ExcelWriteCell -FilePath $exPath -Cell $obj.cell -Value $obj.value
                }
            }
        } catch { Write-StatusLine "ERR" "Excel JSON parse failed: $_" }
    }

    # ── 3. Email (LAST of the file-dependent actions) ────────────────────────
    if ($ResponseText -match '<EMAIL>([\s\S]*?)</EMAIL>') {
        $emailRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $emailRaw
            $emailAction = if ($obj.action) { $obj.action } else { "draft" }
            if ($emailAction -eq "reply") {
                Write-Host "  Reply to '$($obj.query)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-OutlookReplyEmail -Query $obj.query -Body $obj.body }
            } elseif ($emailAction -eq "forward") {
                Write-Host "  Forward '$($obj.query)' to '$($obj.to)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-OutlookForwardEmail -Query $obj.query -To $obj.to -Body $obj.body }
            } elseif ($emailAction -eq "delete") {
               Write-Host "  Delete email '$($obj.query)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
               if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-OutlookDeleteEmail -Query $obj.query }
            } elseif ($emailAction -eq "flag") {
               Write-Host "  Flag email '$($obj.query)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
               if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-OutlookFlagEmail -Query $obj.query }
            } else {
                if ($obj.to -or $obj.subject) {
                    $attachList = @()
                    if ($obj.attachments) { $attachList = Resolve-Attachments -Attachments @($obj.attachments) }
                    Write-Host "  Draft email to $($obj.to)? [Y/N]: " -NoNewline -ForegroundColor Yellow
                    if ((Read-Host) -match '^[Yy]') {
                        $actionsRun += Invoke-OutlookDraftEmail -To $obj.to -Subject $obj.subject -Body $obj.body -Attachments $attachList
                    }
                }
            }
        } catch { 
            $errS = if ($emailRaw) { $emailRaw.Substring(0,[Math]::Min(200,$emailRaw.Length)) } else { "EMPTY" }
            Write-StatusLine "ERR" "Email JSON parse failed: $_ | Raw: $errS" 
        }
    }

    # ── 4. Jabber ────────────────────────────────────────────────────────────
    if ($ResponseText -match '<JABBER>([\s\S]*?)</JABBER>') {
        $jabberRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $jabberRaw
            if ($obj.recipient) {
                Write-Host "  Send Jabber to $($obj.recipient)? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-JabberSendMessage -Recipient $obj.recipient -Message $obj.message }
            }
        } catch { Write-StatusLine "ERR" "Jabber JSON parse failed: $_" }
    }

    # ── 5. Calendar ──────────────────────────────────────────────────────────
    if ($ResponseText -match '<CALENDAR>([\s\S]*?)</CALENDAR>') {
        $calRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $calRaw
            $calAction = if ($obj.action) { $obj.action } else { "create" }
            if ($calAction -eq "read") {
                $days = if ($obj.days) { [int]$obj.days } else { 7 }
                $actionsRun += Invoke-OutlookReadCalendar -DaysAhead $days
            } elseif ($calAction -eq "delete") {
                Write-Host "  Delete appointment '$($obj.subject)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-OutlookDeleteAppointment -Subject $obj.subject }
            } else {
                $attendees = if ($obj.attendees) { @($obj.attendees) } else { @() }
                if ($attendees.Count -gt 0) {
                    Write-Host "  Create MEETING '$($obj.subject)' ($($attendees.Count) attendees)? [Y/N]: " -NoNewline -ForegroundColor Yellow
                    if ((Read-Host) -match '^[Yy]') {
                        $actionsRun += Invoke-OutlookCreateMeeting -Subject $obj.subject -Start $obj.start -End $obj.end `
                                           -Location $obj.location -Body $obj.body -Attendees $attendees -Display
                    }
                } else {
                    Write-Host "  Create appointment '$($obj.subject)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                    if ((Read-Host) -match '^[Yy]') {
                        $actionsRun += Invoke-OutlookCreateAppointment -Subject $obj.subject -Start $obj.start -End $obj.end `
                                           -Location $obj.location -Body $obj.body -Display
                    }
                }
            }
        } catch { Write-StatusLine "ERR" "Calendar JSON parse failed: $_" }
    }

    # ── 6. Open file ─────────────────────────────────────────────────────────
    if ($ResponseText -match '<OPEN>([\s\S]*?)</OPEN>') {
        $openRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $openRaw
            if ($obj.path -and (Test-Path $obj.path)) {
                $ext = [System.IO.Path]::GetExtension($obj.path).ToLower()
                $safeExts = @(".pdf", ".docx", ".xlsx", ".pptx", ".txt", ".csv", ".md", ".jpg", ".png")
                
                if ($ext -notin $safeExts) {
                    Write-StatusLine "ERR" "SECURITY VIOLATION: Cannot autonomously open potentially dangerous file extension '$ext'."
                } else {
                    Write-Host "  Open '$($obj.path)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                    if ((Read-Host) -match '^[Yy]') {
                        Start-Process $obj.path
                        $actionsRun += "Opened: $($obj.path)"
                    }
                }
            }
        } catch { Write-StatusLine "ERR" "Open JSON parse failed: $_" }
    }

    # ── 7. UI Automation ─────────────────────────────────────────────────────
    if ($ResponseText -match '<UIAUTOMATION>([\s\S]*?)</UIAUTOMATION>') {
        $uiaRaw = $Matches[1]
        try {
            $obj = ConvertFrom-LlmJson -Raw $uiaRaw
            $uiaAction = if ($obj.action) { $obj.action } else { "inspect" }
            if ($uiaAction -eq "launch") {
                Write-Host "  Launch '$($obj.path)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-UIAutomationLaunch -AppPath $obj.path }
            } elseif ($uiaAction -eq "inspect") {
                Write-Host "  Inspect UI of '$($obj.window)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-UIAutomationInspect -WindowTitle $obj.window }
            } elseif ($uiaAction -eq "click") {
                Write-Host "  Click element '$($obj.element_name)' in '$($obj.window)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-UIAutomationClick -WindowTitle $obj.window -ElementName $obj.element_name -AutomationId $obj.id }
            } elseif ($uiaAction -eq "type") {
                Write-Host "  Type text into '$($obj.element_name)' in '$($obj.window)'? [Y/N]: " -NoNewline -ForegroundColor Yellow
                if ((Read-Host) -match '^[Yy]') { $actionsRun += Invoke-UIAutomationType -WindowTitle $obj.window -ElementName $obj.element_name -AutomationId $obj.id -Text $obj.text }
            }
        } catch { Write-StatusLine "ERR" "UIA JSON parse failed: $_" }
    }

    return $actionsRun
}