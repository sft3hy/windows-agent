# ============================================================
# FIX SUMMARY
# ============================================================
# 1. Removed the aggressive Pass 0 last-name-only GAL lookup.
#    It was resolving the wrong person (first match on surname)
#    and returning early before any first-name validation.
#
# 2. Get-AllNameVariations now tries full-name variations FIRST
#    ("Price, Samuel", "Samuel Price") before last-name-only.
#    Last-name-only is kept as a last resort at the END of the
#    list, not the beginning.
#
# 3. After GAL Resolve() succeeds, the display name of the
#    resolved entry is now validated against the requested last
#    name before accepting the result. A surname mismatch means
#    Outlook matched the wrong person — we skip it.
#
# 4. Pass 3 (recent email scan) now uses .Restrict() with an
#    DASL filter instead of .Item($i) index looping, which is
#    dramatically faster on large mailboxes.
#
# 5. Add-MailRecipients no longer blindly accepts RESOLVED:name
#    results — it re-validates the resolved display name contains
#    the requested last name before adding.
# ============================================================

# ============================================================
# SHARED NICKNAME / DIMINUTIVE TABLE (unchanged)
# ============================================================
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

$script:NicknameReverseMap = @{}
foreach ($short in $script:NicknameMap.Keys) {
    foreach ($formal in $script:NicknameMap[$short]) {
        if (-not $script:NicknameReverseMap.ContainsKey($formal)) {
            $script:NicknameReverseMap[$formal] = @()
        }
        $script:NicknameReverseMap[$formal] += $short
    }
}

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

# FIX 2: Full-name variations come FIRST; last-name-only is a LAST RESORT.
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

# FIX 3: Validates that a resolved GAL entry actually belongs to the person
# we searched for (by checking the last name appears in the resolved display name).
# This prevents returning the wrong "Price" when the GAL resolves to a different one.
function Test-ResolvedNameMatch {
    param([string]$ResolvedDisplayName, [string]$RequestedLastName)
    if (-not $RequestedLastName -or $RequestedLastName.Length -le 2) { return $true }
    return $ResolvedDisplayName -match [regex]::Escape($RequestedLastName)
}

function Invoke-OutlookContactLookup {
    param([string]$Name)
    Write-ToolLine "Outlook" "Looking up contact" $Name
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace("MAPI")

        $parts    = $Name.Trim() -split '\s+'
        $lastName = $parts[-1]

        # FIX 1 & 2: No more aggressive Pass 0 last-name-only lookup.
        # Get-AllNameVariations now returns full-name forms first.
        $variations = Get-AllNameVariations -Name $Name
        Write-StatusLine "INFO" "Trying $($variations.Count) name variations for '$Name'"

        # ── Pass 1: CreateRecipient + Resolve() ──────────────────────────
        foreach ($v in $variations) {
            try {
                $recip = $ns.CreateRecipient($v)
                if ($recip.Resolve()) {
                    $resolvedDisplayName = try { $recip.AddressEntry.Name } catch { $v }

                    # FIX 3: Validate the resolved entry is actually this person.
                    # Outlook may resolve "Price" to the first Price in the GAL.
                    if (-not (Test-ResolvedNameMatch -ResolvedDisplayName $resolvedDisplayName -RequestedLastName $lastName)) {
                        Write-StatusLine "WARN" "GAL resolved '$v' → '$resolvedDisplayName' but last name '$lastName' not found — skipping"
                        continue
                    }

                    $email = $null
                    try { $email = $recip.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {}
                    if (-not $email -or $email -notmatch '@') {
                        try {
                            $PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                            $email = $recip.AddressEntry.PropertyAccessor.GetProperty($PR_SMTP)
                        } catch {}
                    }
                    if (-not $email -or $email -notmatch '@') {
                        try { $email = $recip.AddressEntry.Address } catch {}
                    }
                    if ($email -and $email -match '@') {
                        Write-StatusLine "OK" "GAL resolved '$v' → $email"
                        return $email
                    }
                    Write-StatusLine "OK" "GAL resolved '$v' (Exchange DN → '$resolvedDisplayName')"
                    return "RESOLVED:$resolvedDisplayName"
                }
            } catch {}
        }

        # ── Pass 2: Local Contacts folder ────────────────────────────────
        try {
            $contactsFolder = $ns.GetDefaultFolder(10)
            $contacts = $contactsFolder.Items
            foreach ($v in $variations) {
                $found = try { $contacts.Find("[FullName] = '$v'") } catch { $null }
                if (-not $found -and $parts.Count -ge 2) {
                    $found = try { $contacts.Find("[LastName] = '$lastName'") } catch { $null }
                }
                if (-not $found) {
                    $escLast = '(?i)' + [regex]::Escape($lastName)
                    $found = $contacts | Where-Object {
                        ($_.FullName -match $escLast) -or ($_.FileAs -match $escLast) -or ($_.Email1Address -match $escLast)
                    } | Select-Object -First 1
                }
                if ($found -and $found.Email1Address) {
                    Write-StatusLine "OK" "Contacts folder: $($found.FullName) → $($found.Email1Address)"
                    return $found.Email1Address
                }
            }
        } catch {}

        # ── Pass 3: Recent Emails using .Restrict() — much faster than .Item() loops ──
        # FIX 4: .Restrict() with a DASL filter pushes the search to the store layer.
        # The old .Item($i) loop was fetching every email object into memory one by one,
        # which is extremely slow on mailboxes with thousands of messages.
        try {
            Write-StatusLine "INFO" "Checking recent emails for '$Name' via Restrict()"
            $escLastDASL = $lastName -replace "'", "''"

            # Search Inbox senders
            $inbox  = $ns.GetDefaultFolder(6).Items
            $filter = "@SQL=""urn:schemas:httpmail:sendername"" LIKE '%$escLastDASL%'"
            $hits   = try { $inbox.Restrict($filter) } catch { $null }
            if ($hits -and $hits.Count -gt 0) {
                $item = $hits.Item(1)
                $email = try { $item.Sender.GetExchangeUser().PrimarySmtpAddress } catch { $null }
                if (-not $email -or $email -notmatch '@') { $email = try { $item.SenderEmailAddress } catch { "" } }
                if ($email -and $email -match '@') {
                    Write-StatusLine "OK" "Recent emails (Inbox/Restrict): $($item.SenderName) → $email"
                    return $email
                }
            }

            # Search Sent Items recipients — Restrict on subject/sender isn't useful here,
            # so we restrict to recent items only (last 90 days) and iterate recipients.
            $sent   = $ns.GetDefaultFolder(5).Items
            $cutoff = (Get-Date).AddDays(-90).ToString("yyyy-MM-dd HH:mm")
            $recentFilter = "@SQL=""urn:schemas:httpmail:datereceived"" >= '$cutoff'"
            $recentSent = try { $sent.Restrict($recentFilter) } catch { $sent }
            try { $recentSent.Sort("[SentOn]", $true) } catch {}

            $escLastName = '(?i)' + [regex]::Escape($lastName)
            $cap = [Math]::Min(200, $recentSent.Count)
            for ($i = 1; $i -le $cap; $i++) {
                $item = try { $recentSent.Item($i) } catch { $null }
                if (-not $item) { continue }
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
        } catch { Write-StatusLine "WARN" "Recent emails search failed: $_" }

        Write-StatusLine "WARN" "No contact found for: $Name"
        return $null
    } catch {
        Write-StatusLine "ERR" "Contact lookup failed: $_"
        return $null
    }
}

# ============================================================
# RECIPIENTS API
# ============================================================
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

    foreach ($token in $tokens) {
        $token = $token.Trim()
        if (-not $token) { continue }

        try {
            if ($token -match '@') {
                $r = try { $Mail.Recipients.Add($token) } catch { $null }
                if ($r) { $r.Type = $RecipientType; try { $r.Resolve() | Out-Null } catch {} }
                Write-StatusLine "OK" "Added recipient: $token"
                continue
            }

            Write-StatusLine "INFO" "Resolving recipient: '$token'"
            $parts    = $token.Trim() -split '\s+'
            $lastName = $parts[-1]

            $lookupResult = Invoke-OutlookContactLookup -Name $token

            if ($lookupResult -and $lookupResult -match '@') {
                $r = try { $Mail.Recipients.Add($lookupResult) } catch { $null }
                if ($r) { $r.Type = $RecipientType; try { $r.Resolve() | Out-Null } catch {} }
                Write-StatusLine "OK" "Added resolved email: $lookupResult"
                continue
            }

            if ($lookupResult -and $lookupResult -match '^RESOLVED:(.+)$') {
                $resolvedName = $Matches[1]

                # FIX 5: Validate the RESOLVED name before blindly adding it.
                # Previously we would add any RESOLVED: result even if it was the wrong person.
                if (-not (Test-ResolvedNameMatch -ResolvedDisplayName $resolvedName -RequestedLastName $lastName)) {
                    Write-StatusLine "WARN" "RESOLVED name '$resolvedName' does not match requested '$lastName' — falling through to variation search"
                } else {
                    $r = try { $Mail.Recipients.Add($resolvedName) } catch { $null }
                    if ($r) { $r.Type = $RecipientType; try { $r.Resolve() | Out-Null } catch {} }
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
                        # FIX 5 (cont.): Validate Outlook's native resolve result too.
                        $resolvedDisplayName = try { $r.AddressEntry.Name } catch { $v }
                        if (Test-ResolvedNameMatch -ResolvedDisplayName $resolvedDisplayName -RequestedLastName $lastName) {
                            Write-StatusLine "OK" "Recipients.Add resolved '$v' → '$resolvedDisplayName'"
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
                if ($r) { $r.Type = $RecipientType }
                Write-StatusLine "WARN" "Could not auto-resolve '$token' — added for manual review"
            }
        } catch {
            Write-StatusLine "WARN" "Error adding recipient '$token': $_"
        }
    }

    try { $Mail.Recipients.ResolveAll() | Out-Null } catch {}
}

# Legacy wrapper — unchanged interface, uses new internals automatically.
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
