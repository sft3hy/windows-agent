# ============================================================
# OUTLOOK INTEGRATION
# ============================================================

function Invoke-OutlookReadEmails {
    <#
    .SYNOPSIS
    Reads recent emails from an Outlook folder.

    .DESCRIPTION
    Connects to the specified Outlook folder (Inbox by default) and retrieves the most recent emails, returning them as a JSON string. Can optionally filter by subject or sender.

    .PARAMETER Count
    The maximum number of emails to retrieve. Default is 10.

    .PARAMETER Folder
    The name of the folder to read from. Defaults to 'Inbox'.

    .PARAMETER Filter
    An optional search string to filter emails by Subject or SenderName.

    .EXAMPLE
    Invoke-OutlookReadEmails -Count 5 -Folder "Important"
    Reads the 5 most recent emails from the 'Important' folder.
    #>
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
# FIX: Full-name variations come FIRST; last-name-only is a LAST RESORT.
# Previously, last-name-only was first in the list, causing the wrong person
# to be matched whenever multiple contacts share a surname.
function Get-AllNameVariations {
    param([string]$Name)
    $parts = $Name.Trim() -split '\s+'
    $variations = @()
    if ($parts.Count -ge 2) {
        $first = $parts[0]; $last = $parts[-1]
        $firstVariants = Get-NameVariants -First $first
        foreach ($fn in $firstVariants) {
            $variations += "$last, $fn"    # GAL: Price, Samuel  (highest hit rate, full name)
            $variations += "$fn $last"     # Natural: Samuel Price
            $variations += "$last $fn"     # Reversed: Price Samuel
        }
        # Last-name-only goes at the END as a last resort — not first.
        # Only attempt if the surname is long enough to be distinctive.
        if ($last.Length -gt 4) { $variations += $last }
        $variations = $variations | Select-Object -Unique
    } else {
        $variations += $Name
    }
    return $variations
}

# FIX: Validates that a resolved GAL entry actually belongs to the person
# we searched for (by checking the last name appears in the resolved display name).
# This prevents returning the wrong "Price" when the GAL resolves to a different one.
function Test-ResolvedNameMatch {
    param([string]$ResolvedDisplayName, [string]$RequestedLastName)
    if (-not $RequestedLastName -or $RequestedLastName.Length -le 2) { return $true }
    return $ResolvedDisplayName -match [regex]::Escape($RequestedLastName)
}

# ============================================================
# CONTACT LOOKUP (informational — extracts email if possible)
# ============================================================
function Invoke-OutlookContactLookup {
    <#
    .SYNOPSIS
    Look up a contact's email address using Active Directory.

    .DESCRIPTION
    Bypasses Outlook COM to perform a highly efficient LDAP query against Active Directory to resolve a person's name to an email address. Useful for finding colleagues in the Global Address List.

    .PARAMETER Name
    The partial or full name to search for (e.g., 'Sam Price' or 'Demetra').

    .EXAMPLE
    Invoke-OutlookContactLookup -Name "John Doe"
    Returns the email address for John Doe if found in AD.
    #>
    param([string]$Name)
    Write-ToolLine "Outlook" "Looking up contact" $Name
    try {
        Write-StatusLine "INFO" "Attempting Active Directory (LDAP) search for '$Name'"
        
        # Bypass Outlook COM entirely and query AD
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        if ($null -ne $domain) {
            $searcher = [adsisearcher]""
            $parts    = $Name.Trim() -split '\s+'
            
            $queries = @()
            if ($parts.Count -ge 2) {
                $f = $parts[0]
                $l = $parts[-1]
                # Priority 1: Last First (user requested this order)
                # Priority 2: First Last
                # Priority 3: DisplayName contains Last and First
                # Priority 4: DisplayName contains First and Last
                # Priority 5: Fallback to just Last Name
                $queries += "(&(objectCategory=person)(objectClass=user)(anr=$l $f))"
                $queries += "(&(objectCategory=person)(objectClass=user)(anr=$Name))"
                $queries += "(&(objectCategory=person)(objectClass=user)(displayName=*$l*$f*))"
                $queries += "(&(objectCategory=person)(objectClass=user)(displayName=*$f*$l*))"
                $queries += "(&(objectCategory=person)(objectClass=user)(sn=$l))"
            } else {
                $queries += "(&(objectCategory=person)(objectClass=user)(|(anr=$Name)(givenName=$Name*)(sn=$Name*)(displayName=*$Name*)))"
            }
            
            $searcher.PropertiesToLoad.Add("displayname") | Out-Null
            $searcher.PropertiesToLoad.Add("mail") | Out-Null
            
            $lastName = $parts[-1]
            $escLastName = '(?i)' + [regex]::Escape($lastName)
            $escFirstName = if ($parts.Count -ge 2) { '(?i)' + [regex]::Escape($parts[0]) } else { "" }
            
            $foundMatch = $false
            foreach ($query in $queries) {
                $searcher.Filter = $query
                $results = $searcher.FindAll()
                
                foreach ($res in $results) {
                    $prop = $res.Properties
                    $email = if ($prop["mail"]) { $prop["mail"][0] } else { "" }
                    $dispName = if ($prop["displayname"]) { $prop["displayname"][0] } else { "" }
                    
                    if ($email -and $email -match '@') {
                        # Strict validation for broad queries (like sn=$l)
                        if ($query -match "sn=" -and $parts.Count -ge 2) {
                            if (-not ($dispName -match $escFirstName)) {
                                continue # Skip if first name isn't in display name for fallback query
                            }
                        }
                        
                        if ($parts.Count -eq 1 -or $dispName -match $escLastName) {
                            Write-StatusLine "OK" "AD search resolved '$Name' → $email"
                            return $email
                        }
                    }
                }
            }
            
            Write-StatusLine "WARN" "AD found no exact match for '$Name'."
        } else {
            Write-StatusLine "WARN" "Not connected to an Active Directory domain."
        }

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
        [int]$RecipientType = 1
    )

    if ($null -eq $Mail) {
        Write-StatusLine "WARN" "Add-MailRecipients: Invalid Mail object."
        return
    }

    $tokens = $To -split '\s*(?:,\s*|\s+and\s+|;\s*)\s*' | Where-Object { $_.Trim() -ne '' }
    $resolvedList = @()

    foreach ($token in $tokens) {
        $token = $token.Trim()
        if (-not $token) { continue }

        try {
            if ($token -match '@') {
                $r = try { $Mail.Recipients.Add($token) } catch { $null }
                if ($r) { 
                    $r.Type = $RecipientType; try { $r.Resolve() | Out-Null } catch {} 
                    $resolvedList += $token
                }
                Write-StatusLine "OK" "Added recipient: $token"
                continue
            }

            Write-StatusLine "INFO" "Resolving recipient: '$token'"
            $parts    = $token.Trim() -split '\s+'
            $lastName = $parts[-1]

            $lookupResult = Invoke-OutlookContactLookup -Name $token

            if ($lookupResult -and $lookupResult -match '@') {
                $r = try { $Mail.Recipients.Add($lookupResult) } catch { $null }
                if ($r) { 
                    $r.Type = $RecipientType; try { $r.Resolve() | Out-Null } catch {} 
                    $resolvedList += $lookupResult
                }
                Write-StatusLine "OK" "Added resolved email: $lookupResult"
                continue
            }

            if ($lookupResult -and $lookupResult -match '^RESOLVED:(.+)$') {
                $resolvedName = $Matches[1]
                if (-not (Test-ResolvedNameMatch -ResolvedDisplayName $resolvedName -RequestedLastName $lastName)) {
                    Write-StatusLine "WARN" "RESOLVED name '$resolvedName' does not match requested '$lastName' — falling through to variation search"
                } else {
                    $r = try { $Mail.Recipients.Add($resolvedName) } catch { $null }
                    if ($r) { 
                        $r.Type = $RecipientType; try { $r.Resolve() | Out-Null } catch {} 
                        $resolvedList += $resolvedName
                    }
                    Write-StatusLine "OK" "Added GAL-resolved name: $resolvedName"
                    continue
                }
            }

            # Lookup failed or mismatch — try Recipients.Add with each variation directly.
            $variations = Get-AllNameVariations -Name $token
            $added = $false

            foreach ($v in $variations) {
                if (-not $v) { continue }
                $r = try { $Mail.Recipients.Add($v) } catch { $null }
                if ($r) {
                    $r.Type = $RecipientType
                    $resolved = $false
                    try { $resolved = $r.Resolve() } catch { $resolved = $false }

                    if ($resolved) {
                        $resolvedDisplayName = try { $r.AddressEntry.Name } catch { $v }
                        if (Test-ResolvedNameMatch -ResolvedDisplayName $resolvedDisplayName -RequestedLastName $lastName) {
                            Write-StatusLine "OK" "Recipients.Add resolved '$v' → '$resolvedDisplayName'"
                            $resolvedList += $resolvedDisplayName
                            $added = $true
                            break
                        } else {
                            Write-StatusLine "WARN" "Recipients.Add resolved '$v' → '$resolvedDisplayName' but last name mismatch — skipping"
                            try { $r.Delete() } catch {}
                        }
                    } else {
                        try { $r.Delete() } catch {}
                    }
                }
            }

            if (-not $added) {
                $r = try { $Mail.Recipients.Add($token) } catch { $null }
                if ($r) { 
                    $r.Type = $RecipientType 
                    $resolvedList += $token
                }
                Write-StatusLine "WARN" "Could not auto-resolve '$token' — added for manual review"
            }
        } catch {
            Write-StatusLine "WARN" "Error adding recipient '$token': $_"
        }
    }

    try { $Mail.Recipients.ResolveAll() | Out-Null } catch {}

    # FIX: Explicitly set To/CC properties to ensure they appear in the UI pre-filled.
    # This resolves issues where Recipients.Add() doesn't immediately reflect in the To: text box.
    # We do this OUTSIDE the ResolveAll try/catch so it doesn't get skipped if resolution fails.
    if ($resolvedList.Count -gt 0) {
        $combined = $resolvedList -join "; "
        if ($RecipientType -eq 1) { try { $Mail.To = $combined } catch {} }
        elseif ($RecipientType -eq 2) { try { $Mail.CC = $combined } catch {} }
    }
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
    <#
    .SYNOPSIS
    Sends an email automatically via Outlook.

    .DESCRIPTION
    Creates and immediately sends an email. Automatically resolves recipient names using the GAL, appends the user's signature, and attaches any specified files.

    .PARAMETER To
    A semicolon-separated list of recipient names or email addresses.

    .PARAMETER Subject
    The subject line of the email.

    .PARAMETER Body
    The body text of the email.

    .PARAMETER CC
    Optional. A semicolon-separated list of CC recipients.

    .PARAMETER Attachments
    Optional. An array of absolute file paths to attach.

    .EXAMPLE
    Invoke-OutlookSendEmail -To "asmith@example.com" -Subject "Report" -Body "Here is the report."
    Sends an email to asmith.
    #>
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
    <#
    .SYNOPSIS
    Drafts an email and displays it in Outlook for user review.

    .DESCRIPTION
    Creates a new email item, pre-fills the To, Subject, Body, Signature, and Attachments, and then displays the Outlook Inspector window so the user can review and send it manually.

    .PARAMETER To
    A semicolon-separated list of recipient names or email addresses.

    .PARAMETER Subject
    The subject line of the email.

    .PARAMETER Body
    The body text of the email.

    .PARAMETER Attachments
    Optional. An array of absolute file paths to attach.

    .EXAMPLE
    Invoke-OutlookDraftEmail -To "Jane Doe" -Subject "Draft" -Body "Please review."
    Opens an email draft addressed to Jane Doe.
    #>
    param([string]$To, [string]$Subject, [string]$Body, [array]$Attachments = @())
    Write-ToolLine "Outlook" "Opening draft email" "To: $To"
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail    = $outlook.CreateItem(0)
        $mail.Subject = $Subject

        # Append signature to body
        $sig = Get-OutlookSignature
        $mail.Body = $Body + $sig

        # Resolve each recipient token to a raw SMTP email address and set
        # $mail.To directly as a plain string — this is the most reliable way
        # to pre-fill the To: field without depending on the Recipients COM API.
        $tokens = $To -split '\s*(?:,\s*|\s+and\s+|;\s*)\s*' | Where-Object { $_.Trim() -ne '' }
        $rawEmails = @()
        foreach ($token in $tokens) {
            $token = $token.Trim()
            if (-not $token) { continue }
            if ($token -match '@') {
                $rawEmails += $token
                Write-StatusLine "OK" "Using raw email: $token"
            } else {
                $resolved = Invoke-OutlookContactLookup -Name $token
                if ($resolved -and $resolved -match '@') {
                    $rawEmails += $resolved
                    Write-StatusLine "OK" "Resolved '$token' → $resolved"
                } else {
                    # Fallback: use name as-is and let Outlook resolve on open
                    $rawEmails += $token
                    Write-StatusLine "WARN" "Could not resolve '$token' — using name as-is"
                }
            }
        }

        if ($rawEmails.Count -gt 0) {
            $mail.To = $rawEmails -join "; "
        }

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
    <#
    .SYNOPSIS
    Draft a reply to a specific email.

    .DESCRIPTION
    Finds an email matching the given query (index number or subject search) in the Inbox and generates a Reply-All draft containing the specified body text.

    .PARAMETER Query
    The subject text or the index number of the email in the Inbox to reply to.

    .PARAMETER Body
    The message text to insert above the original email thread.

    .EXAMPLE
    Invoke-OutlookReplyEmail -Query "Project Update" -Body "Thanks for the update."
    Opens a draft reply to the 'Project Update' email.
    #>
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
    <#
    .SYNOPSIS
    Draft a forward of a specific email.

    .DESCRIPTION
    Finds an email matching the given query and generates a Forward draft to the specified recipient(s).

    .PARAMETER Query
    The subject text or the index number of the email to forward.

    .PARAMETER To
    The recipient(s) to forward the email to.

    .PARAMETER Body
    The message text to insert above the forwarded thread.

    .EXAMPLE
    Invoke-OutlookForwardEmail -Query 1 -To "Boss" -Body "FYI on this thread."
    Forwards the most recent email in the Inbox to 'Boss'.
    #>
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
    <#
    .SYNOPSIS
    Delete a specific email.

    .DESCRIPTION
    Finds an email matching the given query and prompts the user for confirmation before moving it to the Deleted Items folder.

    .PARAMETER Query
    The subject text or index number of the email to delete.

    .EXAMPLE
    Invoke-OutlookDeleteEmail -Query "Spam Offer"
    Prompts to delete the matching email.
    #>
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
    <#
    .SYNOPSIS
    Flag a specific email for follow-up.

    .DESCRIPTION
    Finds an email matching the given query and marks it as a task (flagged) in Outlook.

    .PARAMETER Query
    The subject text or index number of the email to flag.

    .EXAMPLE
    Invoke-OutlookFlagEmail -Query 2
    Flags the second most recent email in the Inbox.
    #>
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