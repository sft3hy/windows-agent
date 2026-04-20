# ============================================================
# MICROSOFT WORD INTEGRATION
# ============================================================

function Invoke-WordCreateDocument {
    param(
        [string]$FilePath,
        [string]$Title   = "",
        [string]$Body    = "",
        [switch]$Display        # $true = show Word to user after creating
    )
    Write-ToolLine "Word" "Creating document" $FilePath
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $Display.IsPresent
        $word.DisplayAlerts = 0

        $doc = $word.Documents.Add()

        # Insert title as Heading 1 if provided
        if ($Title) {
            $sel = $word.Selection
            $sel.Style = $doc.Styles.Item("Heading 1")
            $sel.TypeText($Title)
            $sel.TypeParagraph()
            $sel.Style = $doc.Styles.Item("Normal")
        }

        # Insert body text, split on \n for paragraph breaks
        if ($Body) {
            $paragraphs = $Body -split "(?:\\n|\r?\n)"
            foreach ($para in $paragraphs) {
                $word.Selection.TypeText($para.Trim())
                $word.Selection.TypeParagraph()
            }
        }

        # Ensure directory exists
        $dir = [System.IO.Path]::GetDirectoryName($FilePath)
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        # Save as DOCX (format 16 = wdFormatXMLDocument)
        $doc.SaveAs2($FilePath, 16)

        if ($Display) {
            $word.Visible = $true
        } else {
            $doc.Close($false)
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }

        Write-StatusLine "OK" "Word document created: $FilePath"
        return "Word document created at $FilePath"
    } catch {
        Write-StatusLine "ERR" "Word create failed: $_"
        try { $word.Quit() } catch {}
        return "ERROR: $_"
    }
}

function Invoke-WordOpenDocument {
    param([string]$FilePath)
    Write-ToolLine "Word" "Opening document" $FilePath
    try {
        if (-not (Test-Path $FilePath)) { return "ERROR: File not found: $FilePath" }
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $word.Documents.Open($FilePath) | Out-Null
        Write-StatusLine "OK" "Opened $FilePath in Word"
        return "Word document opened: $FilePath"
    } catch {
        Write-StatusLine "ERR" "Word open failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-WordAppendText {
    param([string]$FilePath, [string]$Text)
    Write-ToolLine "Word" "Appending text" $FilePath
    try {
        if (-not (Test-Path $FilePath)) { return "ERROR: File not found: $FilePath" }
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($FilePath)

        # Move to end of document and append
        $word.Selection.EndKey(6) | Out-Null  # 6 = wdStory
        $paragraphs = $Text -split "(?:\\n|\r?\n)"
        foreach ($para in $paragraphs) {
            $word.Selection.TypeParagraph()
            $word.Selection.TypeText($para.Trim())
        }

        $doc.Save()
        $doc.Close($false)
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        Write-StatusLine "OK" "Appended text to $FilePath"
        return "Text appended to Word document: $FilePath"
    } catch {
        Write-StatusLine "ERR" "Word append failed: $_"
        try { $word.Quit() } catch {}
        return "ERROR: $_"
    }
}

function Invoke-WordReplaceText {
    param([string]$FilePath, [string]$SearchText, [string]$ReplaceText)
    Write-ToolLine "Word" "Replacing text" "'$SearchText' -> '$ReplaceText'"
    try {
        if (-not (Test-Path $FilePath)) { return "ERROR: File not found: $FilePath" }
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($FilePath)

        $find = $word.Selection.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        
        $wdFindContinue = 1
        $wdReplaceAll = 2
        
        $result = $find.Execute($SearchText, $false, $false, $false, $false, $false, $true, $wdFindContinue, $false, $ReplaceText, $wdReplaceAll)
        
        $doc.Save()
        $doc.Close($false)
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        
        Write-StatusLine "OK" "Replaced text in $FilePath"
        return "Text replaced successfully in $FilePath"
    } catch {
        Write-StatusLine "ERR" "Word replace failed: $_"
        try { $word.Quit() } catch {}
        return "ERROR: $_"
    }
}
