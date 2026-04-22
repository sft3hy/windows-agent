# ============================================================
# FILE SYSTEM & CONTENT EXTRACTION (PDF, Word, PPT)
# ============================================================

function Invoke-ReadFile {
    <#
    .SYNOPSIS
    Reads the content of various file types.

    .DESCRIPTION
    A universal file reader that securely extracts text from TXT, CSV, MD, JSON, XML, LOG, PDF, DOCX, PPTX, and XLSX files. Implements a security sandbox to prevent directory traversal attacks.

    .PARAMETER FilePath
    The absolute path to the file.

    .EXAMPLE
    Invoke-ReadFile -FilePath "C:\Users\John\Documents\Report.pdf"
    Extracts text content from the PDF file.
    #>
    param([string]$FilePath)
    Write-ToolLine "File" "Reading file" $FilePath
    try {
        if (-not (Test-Path $FilePath)) { return "ERROR: File not found: $FilePath" }
        
        # Security Sandbox: Ensure file exists inside User Profile, Temp, or known shell folders
        # This handles Enterprise Folder Redirection where Documents might be on a UNC share (e.g. \\server\share)
        $fullPath = [System.IO.Path]::GetFullPath($FilePath)
        
        $allowedRoots = @(
            [System.IO.Path]::GetFullPath($env:USERPROFILE),
            [System.IO.Path]::GetFullPath($env:TEMP),
            [System.IO.Path]::GetFullPath([Environment]::GetFolderPath("MyDocuments")),
            [System.IO.Path]::GetFullPath([Environment]::GetFolderPath("Desktop"))
        ) | Select-Object -Unique

        $isAllowed = $false
        foreach ($root in $allowedRoots) {
            if ($fullPath.StartsWith($root, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                $isAllowed = $true
                break
            }
        }

        if (-not $isAllowed) {
            Write-StatusLine "ERR" "SECURITY VIOLATION: Access denied to $fullPath. Files must be within User Profile, Temp, or standard shell folders."
            return "ERROR: Access Denied. Cannot read files outside of designated user space."
        }

        $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()

        if ($ext -eq ".pdf") {
            # Attempt Word COM with a hard timeout to prevent hanging
            Write-StatusLine "INFO" "Extracting PDF via Word COM (30s timeout)..."
            $text = Invoke-WordExtract -Path $FilePath -TimeoutSeconds 30
            if ($text -and $text -notmatch '^ERROR:') {
                Write-StatusLine "OK" "Extracted PDF text via Word ($($text.Length) chars)"
                return $text.Substring(0, [Math]::Min(8000, $text.Length))
            }
            # Fallback: raw binary text extraction (no Word dependency)
            Write-StatusLine "WARN" "Word timed out or failed, using raw text extraction"
            return Invoke-RawTextExtract -Path $FilePath

        } elseif ($ext -in ".txt", ".csv", ".md", ".json", ".xml", ".log") {
            $content = Get-Content $FilePath -Raw
            Write-StatusLine "OK" "Read text file ($($content.Length) chars)"
            return $content.Substring(0, [Math]::Min(8000, $content.Length))

        } elseif ($ext -eq ".docx") {
            Write-StatusLine "INFO" "Extracting DOCX via Word COM (30s timeout)..."
            $text = Invoke-WordExtract -Path $FilePath -TimeoutSeconds 30
            if ($text -and $text -notmatch '^ERROR:') {
                Write-StatusLine "OK" "Read DOCX ($($text.Length) chars)"
                return $text.Substring(0, [Math]::Min(8000, $text.Length))
            }
            return "ERROR: Could not read DOCX - Word COM unavailable or timed out"

        } elseif ($ext -eq ".pptx") {
            Write-StatusLine "INFO" "Extracting PPTX via PowerPoint COM (30s timeout)..."
            $text = Invoke-PowerPointExtract -Path $FilePath -TimeoutSeconds 30
            if ($text -and $text -notmatch '^ERROR:') {
                Write-StatusLine "OK" "Read PPTX ($($text.Length) chars)"
                return $text.Substring(0, [Math]::Min(8000, $text.Length))
            }
            return "ERROR: Could not read PPTX - PowerPoint COM unavailable or timed out"

        } elseif ($ext -eq ".xlsx") {
            $text = Invoke-ExcelReadFile -FilePath $FilePath -MaxRows 100
            if ($text -and $text -notmatch '^ERROR:') {
                Write-StatusLine "OK" "Read XLSX ($($text.Length) chars)"
                return $text.Substring(0, [Math]::Min(8000, $text.Length))
            }
            return "ERROR: Could not read XLSX"

        } else {
            return "ERROR: Unsupported file type: $ext"
        }

    } catch {
        Write-StatusLine "ERR" "Read file failed: $_"
        return "ERROR: $_"
    }
}

# Run Word COM in a separate Runspace with a hard timeout
function Invoke-WordExtract {
    <#
    .SYNOPSIS
    Extracts text from Word documents using COM interop.

    .DESCRIPTION
    Spawns a hidden Word process in an STA thread to securely extract text from .docx or .pdf files. Enforces a strict timeout to prevent hung COM processes.

    .PARAMETER Path
    The absolute path to the document.

    .PARAMETER TimeoutSeconds
    The maximum time to wait for extraction before killing the Word process. Defaults to 30.
    #>
    param([string]$Path, [int]$TimeoutSeconds = 30)

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.Open()
    $rs.SessionStateProxy.SetVariable("FilePath", $Path)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            # ConfirmConversions = $false prevents the PDF conversion dialog
            $doc = $word.Documents.Open($FilePath, $false, $true, $false, "", "", $true, "", "", 0, [Type]::Missing, $false)
            $text = $doc.Content.Text
            $doc.Close($false)
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            return $text
        } catch {
            try { $word.Quit() } catch {}
            return "ERROR: $_"
        }
    })

    $handle = $ps.BeginInvoke()
    $completed = $handle.AsyncWaitHandle.WaitOne($TimeoutSeconds * 1000)

    if ($completed) {
        $result = $ps.EndInvoke($handle)
        $ps.Dispose()
        $rs.Close()
        if ($result -and $result.Count -gt 0) { return [string]$result[0] }
        return "ERROR: Word returned no text"
    } else {
        # Timed out — kill any Word processes we spawned
        $ps.Stop()
        $ps.Dispose()
        $rs.Close()
        try { Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force } catch {}
        return "ERROR: Word COM timed out after ${TimeoutSeconds}s"
    }
}

# Run PowerPoint COM in a separate Runspace with a hard timeout
function Invoke-PowerPointExtract {
    <#
    .SYNOPSIS
    Extracts text from PowerPoint presentations using COM interop.

    .DESCRIPTION
    Spawns a hidden PowerPoint process in an STA thread to extract text from all slides and shapes in a .pptx file. Enforces a strict timeout.

    .PARAMETER Path
    The absolute path to the presentation.

    .PARAMETER TimeoutSeconds
    The maximum time to wait for extraction. Defaults to 30.
    #>
    param([string]$Path, [int]$TimeoutSeconds = 30)

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.Open()
    $rs.SessionStateProxy.SetVariable("FilePath", $Path)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        try {
            $ppt = New-Object -ComObject PowerPoint.Application
            # Open presentation: FileName, ReadOnly, Untitled, WithWindow
            $pres = $ppt.Presentations.Open($FilePath, -1, 0, 0)
            
            $textBuilder = New-Object System.Text.StringBuilder
            foreach ($slide in $pres.Slides) {
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                        [void]$textBuilder.AppendLine($shape.TextFrame.TextRange.Text)
                    }
                }
            }
            $pres.Close()
            # Note: Do not quit PPT as it might be used by the user, but we're invisible so maybe fine
            # Actually, we should quit if we are the only reference
            if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
            return $textBuilder.ToString()
        } catch {
            return "ERROR: $_"
        }
    })

    $handle = $ps.BeginInvoke()
    $completed = $handle.AsyncWaitHandle.WaitOne($TimeoutSeconds * 1000)

    if ($completed) {
        $result = $ps.EndInvoke($handle)
        $ps.Dispose()
        $rs.Close()
        if ($result -and $result.Count -gt 0) { return [string]$result[0] }
        return "ERROR: PowerPoint returned no text"
    } else {
        $ps.Stop()
        $ps.Dispose()
        $rs.Close()
        try { Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force } catch {}
        return "ERROR: PowerPoint COM timed out after ${TimeoutSeconds}s"
    }
}

# Raw text extraction fallback — reads PDF binary and strips non-printable chars
function Invoke-RawTextExtract {
    <#
    .SYNOPSIS
    Fallback method to extract raw text from binary PDF files.

    .DESCRIPTION
    If Word COM fails, this function reads the raw binary of a PDF and uses regular expressions to extract printable text streams.

    .PARAMETER Path
    The absolute path to the file.
    #>
    param([string]$Path)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $raw   = [System.Text.Encoding]::UTF8.GetString($bytes)
        # Extract readable text between PDF stream markers
        $textChunks = @()
        $matches = [regex]::Matches($raw, '(?<=BT\s)(.*?)(?=\sET)', 'Singleline')
        foreach ($m in $matches) {
            $chunk = $m.Value -replace '\(([^)]*)\)', '$1' -replace '[^A-Za-z0-9\s.,;:!?\-/\\''\"@#$%&*+=]', ''
            if ($chunk.Trim().Length -gt 2) { $textChunks += $chunk.Trim() }
        }
        if ($textChunks.Count -gt 0) {
            $text = $textChunks -join "`n"
            Write-StatusLine "OK" "Raw PDF extraction: $($text.Length) chars from $($textChunks.Count) text blocks"
            return $text.Substring(0, [Math]::Min(8000, $text.Length))
        }
        # Absolute last resort: just strip non-printable from the whole file
        $text = $raw -replace '[^\x20-\x7E\r\n]', ' ' -replace '\s{3,}', ' '
        Write-StatusLine "WARN" "Used brute-force text strip ($($text.Length) chars)"
        return $text.Substring(0, [Math]::Min(8000, $text.Length))
    } catch {
        return "ERROR: Raw extraction failed: $_"
    }
}

# Brute-force file discovery in common shell folders
function global:Resolve-FuzzyFilePath {
    <#
    .SYNOPSIS
    Fuzzy-matches a filename against common user directories.

    .DESCRIPTION
    If a user asks to "read the sales report", this function searches the Desktop, Documents, Downloads, and OneDrive folders for any matching files with common extensions (.pdf, .docx, .xlsx, etc.).

    .PARAMETER Name
    The partial or exact filename to search for.

    .EXAMPLE
    Resolve-FuzzyFilePath -Name "QuarterlyEarnings"
    Searches standard directories for QuarterlyEarnings.pdf, .xlsx, etc.
    #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }

    # Extensions to try (blank = exact match including ext already in $Name)
    $exts = @("", ".pptx", ".pdf", ".docx", ".txt", ".xlsx", ".csv", ".md")

    # Build the folder list — add OneDrive and USERPROFILE root
    $oneDrive = @(
        $env:OneDrive,
        $env:OneDriveConsumer,
        $env:OneDriveCommercial
    ) | Where-Object { $_ -and (Test-Path $_) }

    $folders = @(
        $script:ENV_PATHS.Desktop,
        $script:ENV_PATHS.Downloads,
        $script:ENV_PATHS.Documents,
        $env:USERPROFILE
    ) + $oneDrive | Where-Object { $_ }

    foreach ($folder in $folders) {
        foreach ($ext in $exts) {
            $check = Join-Path $folder ($Name + $ext)
            if (Test-Path $check) { return $check }
        }
    }
    return $null
}