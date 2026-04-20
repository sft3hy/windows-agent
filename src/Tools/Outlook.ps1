# ============================================================
# OUTLOOK INTEGRATION
# ============================================================

function Invoke-OutlookReadEmails {
    param([int]$Count = 10, [string]$Folder = "Inbox", [string]$Filter = "")
    Write-ToolLine "Outlook" "Reading emails" "$Folder (last $Count)"
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns      = $outlook.GetNamespace("MAPI")
        $targetFolder = $ns.GetDefaultFolder(6)  # 6 = olFolderInbox

        if ($Folder -ne "Inbox") {
            # Walk the first store's folder tree to find a matching sub-folder
            $storeFolder = $ns.Folders.Item(1)  # first mail account store
            $found = $storeFolder.Folders | Where-Object { $_.Name -like "*$Folder*" } | Select-Object -First 1
            if (-not $found) {
                # Recursively check sub-folders of Inbox
                $found = $ns.GetDefaultFolder(6).Folders | Where-Object { $_.Name -like "*$Folder*" } | Select-Object -First 1
            }
            if ($found) { $targetFolder = $found }
            else         { Write-StatusLine "WARN" "Folder '$Folder' not found, defaulting to Inbox" }
        }

        $items = $targetFolder.Items
        try { $items.Sort("[ReceivedTime]", $true) } catch { <# ignore for non-mail folders #> }
        $results = @()
        $i = 0
        foreach ($item in $items) {
            if ($i -ge $Count) { break }
            if ($Filter -and $item.Subject -notlike "*$Filter*" -and $item.SenderName -notlike "*$Filter*") { continue };
            $results += @{
                Index       = $i + 1
                Subject     = $item.Subject
                Sender      = $item.SenderName
                SenderEmail = $item.SenderEmailAddress
                Received    = try { $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm") } catch { "" };
                Body        = $item.Body.Substring(0, [Math]::Min(500, $item.Body.Length))
            }
            $i++
        }
        Write-StatusLine "OK" "Retrieved $($results.Count) emails"
        return $results | ConvertTo-Json -Depth 3
    } catch {
        Write-StatusLine "ERR" "Outlook error: $_"
        return "ERROR: $_"
    }
}

# ============================================================
# SHARED NICKNAME / DIMINUTIVE TABLE
# ============================================================
# Maps common short names → formal equivalents.  Used by both
# contact-lookup and the Recipients.Add resolution path.
$script:NicknameMap = @{
    'sam'     = @('samuel','samantha');  'mike'    = @('michael')
    'matt'    = @('matthew');            'dan'     = @('daniel')
    'dave'    = @('david');              'chris'   = @('christopher','christina','christine')
    'greg'    = @('gregory');            'steve'   = @('steven','stephen')
    'tony'    = @('anthony');            'tom'     = @('thomas')
    'bill'    = @('william');            'bob'     = @('robert')
    'dick'    = @('richard');            'rick'    = @('richard')
    'rich'    = @('richard');            'rob'     = @('robert')
    'don'     = @('donald');             'ron'     = @('ronald')
    'joe'     = @('joseph');             'jim'     = @('james')
    'jimmy'   = @('james');             'jen'     = @('jennifer')
    'jenn'    = @('jennifer');           'ed'      = @('edward','edwin','edmund')
    'ted'     = @('theodore','edward');  'al'      = @('albert','alan','alexander')
    'alex'    = @('alexander','alexandra','alexis')
    'andy'    = @('andrew');             'drew'    = @('andrew')
    'ben'     = @('benjamin');           'beth'    = @('elizabeth','bethany')
    'liz'     = @('elizabeth');          'kate'    = @('katherine','kathleen','kathryn')
    'kathy'   = @('katherine','kathleen','kathryn')
    'nick'    = @('nicholas');           'pat'     = @('patrick','patricia')
    'will'    = @('william');            'charlie' = @('charles')
    'chuck'   = @('charles');            'jack'    = @('john','jackson')
    'johnny'  = @('john');               'jon'     = @('jonathan','john')
    'larry'   = @('lawrence');           'ray'     = @('raymond')
    'tim'     = @('timothy');            'demi'    = @('demetra','demetria')
    'sue'     = @('susan','suzanne');    'doug'    = @('douglas')
    'frank'   = @('francis','franklin'); 'fred'    = @('frederick')
    'hank'    = @('henry');              'hal'     = @('harold','henry')
    'harry'   = @('harold','henry');     'lori'    = @('lorraine')
    'marge'   = @('margaret');           'margie'  = @('margaret')
    'peg'     = @('margaret');           'peggy'   = @('margaret')
    'meg'     = @('megan','margaret');   'phil'    = @('philip','phillip')
    'barb'    = @('barbara');            'sandy'   = @('sandra','alexander')
    'terry'   = @('terrence','theresa'); 'vince'   = @('vincent')
    'wally'   = @('walter','wallace')
}

# Build reverse map once so formal→short also works ("samuel" → "sam")
$script:NicknameReverseMap = @{}
foreach ($short in $script:NicknameMap.Keys) {
    foreach ($formal in $script:NicknameMap[$short]) {
        if (-not $script:NicknameReverseMap.ContainsKey($formal)) {
            $script:NicknameReverseMap[$formal] = @()
        }
        $script:NicknameReverseMap[$formal] += $short
    }
}

# Return all plausible alternate first-name forms for a given name.
# e.g. "Sam" → @("Sam", "Samuel", "Samantha")
function Get-NameVariants {
    param([string]$First)
    $lc = $First.ToLower()
    $variants = @($First)
    if ($script:NicknameMap.ContainsKey($lc)) {
        foreach ($f in $script:NicknameMap[$lc]) {
            $variants += (Get-Culture).TextInfo.ToTitleCase($f)
        }
    }
    if ($script:NicknameReverseMap.ContainsKey($lc)) {
        foreach ($s in $script:NicknameReverseMap[$lc]) {
            $variants += (Get-Culture).TextInfo.ToTitleCase($s)
        }
    }
    return $variants | Sort-Object -Unique
}

# Build all name orderings for a multi-part name + nickname expansions.
# Returns: @("Price, Samuel", "Samuel Price", "Price Samuel", "Price, Sam", ...)
function Get-AllNameVariations {
    param([string]$Name)
    $parts = $Name.Trim() -split '\s+'
    $variations = @()
    if ($parts.Count -ge 2) {
        $first = $parts[0]; $last = $parts[-1]
        $firstVariants = Get-NameVariants -First $first
        foreach ($fn in $firstVariants) {
            $variations += "$last, $fn"    # GAL: Price, Samuel
            $variations += "$fn $last"     # Natural: Samuel Price
            $variations += "$last $fn"     # Reversed: Price Samuel
        }
        if ($last.Length -gt 3) { $variations += $last }   # Last name only
        $variations = $variations | Sort-Object -Unique
    } else {
        $variations += $Name
    }
    return $variations
}

# ============================================================
# CONTACT LOOKUP (informational — extracts email if possible)
# ============================================================
function Invoke-OutlookContactLookup {
    param([string]$Name)
    Write-ToolLine "Outlook" "Looking up contact" $Name
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace("MAPI")

        $variations = Get-AllNameVariations -Name $Name
        $parts = $Name.Trim() -split '\s+'

        Write-StatusLine "INFO" "Trying $($variations.Count) name variations for '$Name'"

        # ── Pass 1: CreateRecipient + Resolve() ──────────────────────────
        foreach ($v in $variations) {
            try {
                $recip = $ns.CreateRecipient($v)
                if ($recip.Resolve()) {
                    $email = $null
                    # Method 1: ExchangeUser SMTP
                    try { $email = $recip.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {}
                    # Method 2: PropertyAccessor for SMTP address
                    if (-not $email -or $email -notmatch '@') {
                        try {
                            $PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                            $email = $recip.AddressEntry.PropertyAccessor.GetProperty($PR_SMTP)
                        } catch {}
                    }
                    # Method 3: Raw .Address (may be Exchange DN)
                    if (-not $email -or $email -notmatch '@') {
                        try { $email = $recip.AddressEntry.Address } catch {}
                    }
                    if ($email -and $email -match '@') {
                        Write-StatusLine "OK" "GAL resolved '$v' → $email"
                        return $email
                    }
                    # If Resolve() succeeded but we can't extract SMTP, return
                    # the variation that resolved — the caller can use it with
                    # Recipients.Add which will work since Outlook knows it.
                    $resolvedName = try { $recip.AddressEntry.Name } catch { $v }
                    Write-StatusLine "OK" "GAL resolved '$v' (Exchange DN — returning display name '$resolvedName')"
                    return "RESOLVED:$resolvedName"
                }
            } catch {}
        }

        # ── Pass 2: Local Contacts folder ────────────────────────────────
        try {
            $contactsFolder = $ns.GetDefaultFolder(10)  # 10 = olFolderContacts
            $contacts = $contactsFolder.Items
            foreach ($v in $variations) {
                $found = try { $contacts.Find("[FullName] = '$v'") } catch { $null }
                if (-not $found -and $parts.Count -ge 2) {
                    $found = try { $contacts.Find("[LastName] = '$($parts[-1])'") } catch { $null }
                }
                if (-not $found) {
                    $escV = [regex]::Escape($v)
                    $found = $contacts | Where-Object {
                        ($_.FullName -match $escV) -or ($_.FileAs -match $escV)
                    } | Select-Object -First 1
                }
                if ($found -and $found.Email1Address) {
                    Write-StatusLine "OK" "Contacts folder: $($found.FullName) → $($found.Email1Address)"
                    return $found.Email1Address
                }
            }
        } catch {}

        # ── Pass 3: Recent Emails (Inbox & Sent) ─────────────────────────
        try {
            Write-StatusLine "INFO" "Checking recent emails for '$Name'"
            $inbox = $ns.GetDefaultFolder(6).Items
            $sent  = $ns.GetDefaultFolder(5).Items
            try { $inbox.Sort("[ReceivedTime]", $true); $sent.Sort("[SentOn]", $true) } catch {}
            
            $lastName = if ($parts.Count -ge 2) { $parts[-1] } else { $Name }
            $escLastName = '(?i)' + [regex]::Escape($lastName)
            
            for ($i = 1; $i -le 50; $i++) {
                # Check Inbox (Senders)
                if ($i -le $inbox.Count) {
                    $item = try { $inbox.Item($i) } catch { $null }
                    if ($item -and $item.SenderName -match $escLastName) {
                        $email = try { $item.Sender.GetExchangeUser().PrimarySmtpAddress } catch { $null }
                        if (-not $email -or $email -notmatch '@') { $email = try { $item.SenderEmailAddress } catch { "" } }
                        if ($email -and $email -match '@') {
                            Write-StatusLine "OK" "Recent emails (Inbox): $($item.SenderName) → $email"
                            return $email
                        }
                    }
                }
                # Check Sent Items (Recipients)
                if ($i -le $sent.Count) {
                    $item = try { $sent.Item($i) } catch { $null }
                    if ($item) {
                        $recips = try { $item.Recipients } catch { $null }
                        if ($recips) {
                            foreach ($r in $recips) {
                                if ($r.Name -match $escLastName) {
                                    $email = try { $r.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch { $null }
                                    if (-not $email -or $email -notmatch '@') { $email = try { $r.Address } catch { "" } }
                                    if ($email -and $email -match '@') {
                                        Write-StatusLine "OK" "Recent emails (Sent): $($r.Name) → $email"
                                        return $email
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch { Write-StatusLine "WARN" "Recent emails search failed: $_" }

        Write-StatusLine "WARN" "No contact found for: $Name"
        return $null
    } catch {
        Write-StatusLine "ERR" "Contact lookup failed: $_"
        return $null
    }
}

# ============================================================
# RECIPIENTS API — the correct way to add + resolve names
# ============================================================
# Uses $mail.Recipients.Add() + .Resolve() directly on the mail
# item.  This is identical to typing a name in the To: box and
# pressing Tab — Outlook's own GAL engine does the resolution.
# Falls back through nickname variations until one resolves.
function Add-MailRecipients {
    param(
        [object]$Mail,
        [string]$To,
        [int]$RecipientType = 1   # 1 = olTo, 2 = olCC, 3 = olBCC
    )

    # Tokenise: split on ", " | " and " | ";"
    $tokens = $To -split '\s*(?:,\s*|\s+and\s+|;\s*)\s*' | Where-Object { $_.Trim() -ne '' }

    foreach ($token in $tokens) {
        $token = $token.Trim()
        if (-not $token) { continue }

        try {
            # If it's already an email, just add it directly
            if ($token -match '@') {
                $r = $Mail.Recipients.Add($token)
                $r.Type = $RecipientType
                $r.Resolve() | Out-Null
                Write-StatusLine "OK" "Added recipient: $token"
                continue
            }

            Write-StatusLine "INFO" "Resolving recipient: '$token'"

            # First try: our contact lookup (which now returns RESOLVED:name for
            # Exchange DN matches)
            $lookupResult = Invoke-OutlookContactLookup -Name $token

            if ($lookupResult -and $lookupResult -match '@') {
                # Got a real SMTP email
                $r = $Mail.Recipients.Add($lookupResult)
                $r.Type = $RecipientType
                $r.Resolve() | Out-Null
                Write-StatusLine "OK" "Added resolved email: $lookupResult"
                continue
            }

            if ($lookupResult -and $lookupResult -match '^RESOLVED:(.+)$') {
                $resolvedName = $Matches[1] # Safe: match was on the line above
                $r = $Mail.Recipients.Add($resolvedName)
                $r.Type = $RecipientType
                $r.Resolve() | Out-Null
                Write-StatusLine "OK" "Added GAL-resolved name: $resolvedName"
                continue
            }

            # Lookup completely failed — try Recipients.Add with each variation
            # and see if Outlook can resolve any of them natively.
            $variations = Get-AllNameVariations -Name $token
            $added = $false

            foreach ($v in $variations) {
                # Outlook's Recipients.Add can sometimes throw if input is very weird
                $r = try { $Mail.Recipients.Add($v) } catch { $null }
                if ($r) {
                    $r.Type = $RecipientType
                    if ($r.Resolve()) {
                        Write-StatusLine "OK" "Recipients.Add resolved '$v'"
                        $added = $true
                        break
                    } else {
                        $r.Delete()      # Remove failed attempt
                    }
                }
            }

            if (-not $added) {
                # Last resort: add the original name as-is
                $r = $Mail.Recipients.Add($token)
                $r.Type = $RecipientType
                Write-StatusLine "WARN" "Could not auto-resolve '$token' — added for manual review"
            }
        } catch {
            Write-StatusLine "WARN" "Error adding recipient '$token': $_"
        }
    }

    # Final attempt to resolve any remaining unresolved recipients
    try { $Mail.Recipients.ResolveAll() | Out-Null } catch {}
}

# Legacy Resolve-EmailAddress kept for any external callers.
# New code should use Add-MailRecipients instead.
function Resolve-EmailAddress {
    param([string]$To)
    $raw = $To -split '\s*(?:,\s*|\s+and\s+|;\s*)\s*' | Where-Object { $_.Trim() -ne '' }
    $resolved = @()
    foreach ($token in $raw) {
        $token = $token.Trim()
        if (-not $token) { continue }
        if ($token -match '@') { $resolved += $token; continue }
        Write-StatusLine "INFO" "Resolving contact: '$token'"
        $email = Invoke-OutlookContactLookup -Name $token
        if ($email) {
            if ($email -match '@') { $resolved += $email }
            elseif ($email -match '^RESOLVED:(.+)$') { $resolved += $Matches[1] }
            else { $resolved += $email }
        } else {
            Write-StatusLine "WARN" "Could not resolve '$token' - using as-is"; $resolved += $token
        }
    }
    return ($resolved -join "; ")
}

# Get the user's Outlook signature or build a fallback
function Get-OutlookSignature {
    try {
        # Outlook stores signatures in the AppData folder
        $sigPath = Join-Path $env:APPDATA "Microsoft\Signatures"
        if (Test-Path $sigPath) {
            # Look for .txt signature files (plain text versions)
            $sigFiles = Get-ChildItem $sigPath -Filter "*.txt" | Sort-Object LastWriteTime -Descending
            if ($sigFiles.Count -gt 0) {
                $sigContent = Get-Content $sigFiles[0].FullName -Raw
                if ($sigContent.Trim().Length -gt 5) {
                    Write-StatusLine "OK" "Using Outlook signature: $($sigFiles[0].BaseName)"
                    return "`n`n$sigContent"
                }
            }
        }
    } catch {}
    
    # Fallback: build from current user info
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace("MAPI")
        $user = $ns.CurrentUser
        $email = try { $ns.Accounts.Item(1).SmtpAddress } catch { "" }
        $fullName = $user.Name
        $parts = $fullName -split ',\s*|\s+'
        # Handle "LastName, FirstName" or "FirstName LastName"
        if ($fullName -match ',') {
            $firstName = $parts[1]
            $lastName  = $parts[0]
        } else {
            $firstName = $parts[0]
            if ($parts.Count -gt 1) { $lastName = $parts[-1] } else { $lastName = "" }
        }
        Write-StatusLine "OK" "Built fallback signature for $firstName $lastName"
        return "`n`nv/r,`n$firstName $lastName`n$email"
    } catch {
        return "`n`nv/r,"
    }
}

function Invoke-OutlookSendEmail {
    param([string]$To, [string]$Subject, [string]$Body, [string]$CC = "", [array]$Attachments = @())
    Write-ToolLine "Outlook" "Sending email" "To: $To | Subject: $Subject"
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail    = $outlook.CreateItem(0)  # 0 = olMailItem
        $mail.Subject = $Subject
        
        # Append signature to body
        $sig = Get-OutlookSignature
        $mail.Body = $Body + $sig

        # Use Recipients API for proper GAL resolution
        Add-MailRecipients -Mail $mail -To $To -RecipientType 1
        if ($CC) { Add-MailRecipients -Mail $mail -To $CC -RecipientType 2 }
        
        foreach ($file in $Attachments) {
            if ($file -and (Test-Path $file)) {
                $mail.Attachments.Add($file) | Out-Null
                Write-StatusLine "OK" "Attached: $file"
            }
        }

        $mail.Send()
        Write-StatusLine "OK" "Email sent"
        return "Email sent successfully"
    } catch {
        Write-StatusLine "ERR" "Send email failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-OutlookDraftEmail {
    param([string]$To, [string]$Subject, [string]$Body, [array]$Attachments = @())
    Write-ToolLine "Outlook" "Opening draft email" "To: $To"
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail    = $outlook.CreateItem(0)
        $mail.Subject = $Subject
        
        # Append signature to body
        $sig = Get-OutlookSignature
        $mail.Body = $Body + $sig

        # Use Recipients API for proper GAL resolution — this is identical
        # to typing a name in the To: field and pressing Tab.
        Add-MailRecipients -Mail $mail -To $To -RecipientType 1
        
        foreach ($file in $Attachments) {
            if ($file -and (Test-Path $file)) {
                $mail.Attachments.Add($file) | Out-Null
                Write-StatusLine "OK" "Attached: $file"
            }
        }

        $mail.Display()
        Write-StatusLine "OK" "Draft opened in Outlook"
        return "Draft email opened in Outlook for review"
    } catch {
        Write-StatusLine "ERR" "Draft failed: $_"
        return "ERROR: $_"
    }
}

# ============================================================
# ADVANCED OUTLOOK COM (Reply, Forward, Delete, Flag)
# ============================================================

# Helper to find an email
function Get-OutlookEmail {
    param([string]$Query)
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6).Items
    try { $inbox.Sort("[ReceivedTime]", $true) } catch {}
    
    if ($Query -match '^\d+$') {
        $idx = [int]$Query
        if ($idx -le $inbox.Count -and $idx -ge 1) { return $inbox.Item($idx) }
    }
    
    $QueryEscaped = $Query -replace "'", "''"
    $item = try { $inbox.Find("@SQL=""urn:schemas:httpmail:subject"" LIKE '%$QueryEscaped%'") } catch { $null }
    if ($item) { return $item }
    
    for ($i = 1; $i -le 50; $i++) {
        if ($i -le $inbox.Count) {
            $item = $inbox.Item($i)
            if ($item.Subject -match [regex]::Escape($Query) -or $item.SenderName -match [regex]::Escape($Query)) {
                return $item
            }
        }
    }
    return $null
}

function Invoke-OutlookReplyEmail {
    param([string]$Query, [string]$Body)
    Write-ToolLine "Outlook" "Replying to email" $Query
    try {
        $item = Get-OutlookEmail -Query $Query
        if (-not $item) { return "ERROR: Email not found matching '$Query'" }
        $reply = $item.ReplyAll()
        $sig = Get-OutlookSignature
        $htmlBody = "<p>" + ($Body -replace "`n", "<br>") + "</p><br>" + ($sig -replace "`n", "<br>")
        $reply.HTMLBody = $htmlBody + $reply.HTMLBody
        $reply.Display()
        Write-StatusLine "OK" "Draft reply opened"
        return "Draft reply opened in Outlook."
    } catch { return "ERROR: $_" }
}

function Invoke-OutlookForwardEmail {
    param([string]$Query, [string]$To, [string]$Body)
    Write-ToolLine "Outlook" "Forwarding email" $Query
    try {
        $item = Get-OutlookEmail -Query $Query
        if (-not $item) { return "ERROR: Email not found matching '$Query'" }
        $fwd = $item.Forward()
        Add-MailRecipients -Mail $fwd -To $To -RecipientType 1
        $sig = Get-OutlookSignature
        $htmlBody = "<p>" + ($Body -replace "`n", "<br>") + "</p><br>" + ($sig -replace "`n", "<br>")
        $fwd.HTMLBody = $htmlBody + $fwd.HTMLBody
        $fwd.Display()
        Write-StatusLine "OK" "Draft forward opened"
        return "Draft forward opened in Outlook."
    } catch { return "ERROR: $_" }
}

function Invoke-OutlookDeleteEmail {
    param([string]$Query)
    Write-ToolLine "Outlook" "Deleting email" $Query
    try {
        $item = Get-OutlookEmail -Query $Query
        if (-not $item) { return "ERROR: Email not found matching '$Query'" }
        $subj = $item.Subject
        Write-Host "  Delete email '$subj'? [Y/N]: " -NoNewline -ForegroundColor Yellow
        if ((Read-Host) -match '^[Yy]') {
            $item.Delete()
            Write-StatusLine "OK" "Email softly deleted"
            return "Email successfully deleted."
        }
        return "Delete cancelled."
    } catch { return "ERROR: $_" }
}

function Invoke-OutlookFlagEmail {
    param([string]$Query)
    Write-ToolLine "Outlook" "Flagging email" $Query
    try {
        $item = Get-OutlookEmail -Query $Query
        if (-not $item) { return "ERROR: Email not found matching '$Query'" }
        $item.MarkAsTask(4)  # olMarkThisWeek
        $item.Save()
        Write-StatusLine "OK" "Email flagged"
        return "Email flagged successfully."
    } catch { return "ERROR: $_" }
}