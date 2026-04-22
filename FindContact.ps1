#Requires -Version 5.1
<#
.SYNOPSIS
    Fuzzy-searches the Outlook Global Address List (GAL) for a contact by name,
    displays their info, and opens a new draft email to the best match.


.DESCRIPTION
    Uses Outlook's native Ambiguous Name Resolution (ANR) -- the same mechanism
    the Outlook "Address Book" search dialog uses -- to query the Exchange GAL.
    Presents a ranked pick-list when multiple candidates are found, then opens
    a pre-addressed draft email in Outlook.


.PARAMETER Name
    The name (or partial name) to search for. Supports partial first/last name,
    e.g. "Sam Town" will find "Sam Townsend".


.PARAMETER TopN
    How many ranked candidates to show. Default: 10.


.EXAMPLE
    .\Search-GALContact.ps1 -Name "Sam Townsend"


.EXAMPLE
    .\Search-GALContact.ps1 -Name "sarah j" -TopN 20
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0, HelpMessage = "Name to search for in the GAL")]
    [ValidateNotNullOrEmpty()]
    [string]$Name,


    [Parameter()]
    [ValidateRange(1, 50)]
    [int]$TopN = 10
)


# Outlook COM objects are late-bound; strict mode causes false errors on valid properties.
$ErrorActionPreference = 'Stop'


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────


function Get-LevenshteinDistance {
    param([string]$a, [string]$b)
    $a = $a.ToLower(); $b = $b.ToLower()
    $la = $a.Length;   $lb = $b.Length
    if ($la -eq 0) { return $lb }
    if ($lb -eq 0) { return $la }
    $prev = 0..$lb
    for ($i = 1; $i -le $la; $i++) {
        $curr = @(0) * ($lb + 1)
        $curr[0] = $i
        for ($j = 1; $j -le $lb; $j++) {
            $cost = if ($a[$i-1] -eq $b[$j-1]) { 0 } else { 1 }
            $curr[$j] = [Math]::Min([Math]::Min($prev[$j]+1, $curr[$j-1]+1), $prev[$j-1]+$cost)
        }
        $prev = $curr
    }
    return $prev[$lb]
}


function Get-FuzzyScore {
    param([string]$query, [string]$candidate)
    $q = $query.Trim().ToLower()
    $c = $candidate.Trim().ToLower()
    if ($c -eq $q)       { return 100 }
    if ($c.Contains($q)) { return [int](90 * ($q.Length / $c.Length) + 10) }


    # Token match: all query words found somewhere in candidate words
    $qToks = $q -split '\s+' | Where-Object { $_ }
    $cToks = $c -split '\s+' | Where-Object { $_ }
    $hits  = ($qToks | Where-Object { $tok = $_; $cToks | Where-Object { $_.Contains($tok) } }).Count
    if ($qToks.Count -gt 0 -and $hits -eq $qToks.Count) {
        return [int](80 * ($hits / [Math]::Max($qToks.Count, $cToks.Count)) + 10)
    }


    # Levenshtein fallback
    $dist = Get-LevenshteinDistance $q $c
    $maxL = [Math]::Max($q.Length, $c.Length)
    return [int]((1 - $dist / $maxL) * 70)
}


function Format-ContactInfo {
    param($addressEntry)
    if ($addressEntry -is [PSCustomObject]) { return $addressEntry }
    try {
        $ex = $addressEntry.GetExchangeUser()
        if ($null -ne $ex) {
            return [PSCustomObject]@{
                DisplayName = $ex.Name
                Email       = $ex.PrimarySmtpAddress
                Title       = $ex.JobTitle
                Department  = $ex.Department
                Phone       = $ex.BusinessTelephoneNumber
                Office      = $ex.OfficeLocation
                Company     = $ex.CompanyName
            }
        }
    } catch { }
    return [PSCustomObject]@{
        DisplayName = $addressEntry.Name
        Email       = $addressEntry.Address
        Title = ''; Department = ''; Phone = ''; Office = ''; Company = ''
    }
}


# ─────────────────────────────────────────────────────────────────────────────
# CONNECT TO OUTLOOK
# ─────────────────────────────────────────────────────────────────────────────


Write-Host "`n  Connecting to Outlook..." -ForegroundColor Cyan


try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Error "Could not create Outlook COM object. Is Outlook installed?`n$_"
    exit 1
}


$ns = $outlook.GetNamespace("MAPI")
if ($null -eq $ns) {
    Write-Error "Could not access Outlook MAPI namespace. Please ensure Outlook is open and signed in."
    exit 1
}


# ─────────────────────────────────────────────────────────────────────────────
# SEARCH STRATEGY
#
# Because Office LTSC 2024 has aggressive Object Model (OM) Security policies
# that block direct COM resolution of contacts unless running as Admin, we
# bypass Outlook COM completely and query Active Directory via LDAP.
# ─────────────────────────────────────────────────────────────────────────────


Write-Host "  Querying Directory for '$Name'..." -ForegroundColor Cyan


$candidates = [System.Collections.Generic.List[object]]::new()


# ── Active Directory LDAP Search (Bypasses COM Security) ─────────────────
Write-Host "  Attempting Active Directory (LDAP) search..." -ForegroundColor DarkGray
try {
    # Check if we are on an AD domain
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    if ($null -ne $domain) {
        $searcher = [adsisearcher]""
        $parts = $Name.Trim() -split '\s+'
        
        $queries = @()
        if ($parts.Count -ge 2) {
            $f = $parts[0]
            $l = $parts[-1]
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
        $searcher.PropertiesToLoad.Add("title") | Out-Null
        $searcher.PropertiesToLoad.Add("department") | Out-Null
        $searcher.PropertiesToLoad.Add("telephonenumber") | Out-Null
        $searcher.PropertiesToLoad.Add("company") | Out-Null
        
        $lastName = $parts[-1]
        $escLastName = '(?i)' + [regex]::Escape($lastName)
        $escFirstName = if ($parts.Count -ge 2) { '(?i)' + [regex]::Escape($parts[0]) } else { "" }
        
        $foundEmails = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        
        foreach ($query in $queries) {
            $searcher.Filter = $query
            $results = $searcher.FindAll()
            
            foreach ($res in $results) {
                $prop = $res.Properties
                $email = if ($prop["mail"]) { $prop["mail"][0] } else { "" }
                $dispName = if ($prop["displayname"]) { $prop["displayname"][0] } else { "" }
                
                if ($dispName -and $email) {
                    # Skip if we already added this email
                    if (-not $foundEmails.Add($email)) { continue }
                    
                    $candidates.Add([PSCustomObject]@{
                        Name        = $dispName
                        Address     = $email
                        DisplayName = $dispName
                        Email       = $email
                        Title       = if ($prop["title"]) { $prop["title"][0] } else { "" }
                        Department  = if ($prop["department"]) { $prop["department"][0] } else { "" }
                        Phone       = if ($prop["telephonenumber"]) { $prop["telephonenumber"][0] } else { "" }
                        Office      = ""
                        Company     = if ($prop["company"]) { $prop["company"][0] } else { "" }
                        IsADHit     = $true
                    })
                }
            }
        }
        Write-Host "  AD search found $($candidates.Count) match(es)" -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  AD search unavailable or failed: $_" -ForegroundColor DarkYellow
}




# ─────────────────────────────────────────────────────────────────────────────
# NO RESULTS
# ─────────────────────────────────────────────────────────────────────────────


if ($candidates.Count -eq 0) {
    Write-Host "`n  No contacts found for '$Name' using any programmatic strategy." -ForegroundColor Yellow
    Write-Host "  Falling back to Outlook's native GUI resolution..." -ForegroundColor Cyan
    
    try {
        $mail = $outlook.CreateItem(0)
        $mail.To = $Name
        $mail.Display($false)
        Write-Host "  Draft opened in Outlook. Use Ctrl+K (Check Names) to resolve the contact." -ForegroundColor Green
        exit 0
    } catch {
        Write-Host "  Failed to open draft: $_" -ForegroundColor Red
        exit 1
    }
}


# ─────────────────────────────────────────────────────────────────────────────
# DEDUPLICATE + FUZZY RANK
# ─────────────────────────────────────────────────────────────────────────────


$seenNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$scored    = [System.Collections.Generic.List[PSCustomObject]]::new()


foreach ($ae in $candidates) {
    try {
        if (-not $seenNames.Add($ae.Name)) { continue }
        $score = Get-FuzzyScore -query $Name -candidate $ae.Name
        $scored.Add([PSCustomObject]@{ Score = $score; Entry = $ae })
    } catch { }
}


$ranked = $scored | Sort-Object -Property Score -Descending | Select-Object -First $TopN


# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY RESULTS
# ─────────────────────────────────────────────────────────────────────────────


Write-Host "`n  Found $($ranked.Count) result(s):`n" -ForegroundColor Green


$contacts = [System.Collections.Generic.List[PSCustomObject]]::new()
$i = 1


foreach ($r in $ranked) {
    $info = Format-ContactInfo $r.Entry
    $contacts.Add($info)
    $bar  = ('=' * [int]($r.Score / 10)).PadRight(10)


    Write-Host ("  [{0,2}] {1}" -f $i, $info.DisplayName) -ForegroundColor White
    Write-Host ("        Match: [{0}] {1}%" -f $bar, $r.Score) -ForegroundColor DarkGreen
    if ($info.Email)      { Write-Host ("        Email:  $($info.Email)")      -ForegroundColor Cyan    }
    if ($info.Title)      { Write-Host ("        Title:  $($info.Title)")      -ForegroundColor DarkGray }
    if ($info.Department) { Write-Host ("        Dept:   $($info.Department)") -ForegroundColor DarkGray }
    if ($info.Phone)      { Write-Host ("        Phone:  $($info.Phone)")      -ForegroundColor DarkGray }
    if ($info.Office)     { Write-Host ("        Office: $($info.Office)")     -ForegroundColor DarkGray }
    Write-Host ""
    $i++
}


# ─────────────────────────────────────────────────────────────────────────────
# SELECT & OPEN DRAFT
# ─────────────────────────────────────────────────────────────────────────────


if ($contacts.Count -eq 1) {
    $choice = 1
    Write-Host "  Single match -- selecting automatically.`n" -ForegroundColor DarkGray
} else {
    do {
        $raw    = Read-Host "  Select a contact [1-$($contacts.Count)], or Q to quit"
        if ($raw -match '^[Qq]') { Write-Host "  Cancelled."; exit 0 }
        $choice = $raw -as [int]
    } while ($null -eq $choice -or $choice -lt 1 -or $choice -gt $contacts.Count)
}


$selected = $contacts[$choice - 1]


if ([string]::IsNullOrWhiteSpace($selected.Email)) {
    Write-Warning "No SMTP address found -- Outlook will try to resolve by display name."
}


Write-Host "`n  Opening draft to: $($selected.DisplayName) <$($selected.Email)>" -ForegroundColor Cyan


try {
    $mail         = $outlook.CreateItem(0)
    $mail.To      = if ($selected.Email) { $selected.Email } else { $selected.DisplayName }
    $mail.Subject = ""
    $mail.Display($false)   # non-modal -- doesn't block the terminal
    Write-Host "  Draft opened in Outlook!`n" -ForegroundColor Green
} catch {
    Write-Error "Failed to open draft: $_"
    exit 1
}




