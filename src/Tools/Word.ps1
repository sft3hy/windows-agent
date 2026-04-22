# ============================================================
# MICROSOFT WORD INTEGRATION
# ============================================================

function Invoke-WordCreateDocument {
    <#
    .SYNOPSIS
    Create a new Microsoft Word document.

    .DESCRIPTION
    Uses Word COM automation to generate a new .docx file. You can specify an optional Title (formatted as Heading 1) and Body text (split into paragraphs). Can optionally display the document immediately upon creation.

    .PARAMETER FilePath
    The absolute path where the .docx file will be saved.

    .PARAMETER Title
    Optional. A title string inserted at the top of the document as a 'Heading 1'.

    .PARAMETER Body
    Optional. The main text content of the document.

    .PARAMETER Display
    Optional switch. If specified, the Word window remains open and visible after creation. Otherwise, the file is saved and closed silently.

    .EXAMPLE
    Invoke-WordCreateDocument -FilePath "C:\Doc.docx" -Title "Meeting Notes" -Body "Discussed project timelines."
    Creates a new Word document with a title and single paragraph.
    #>
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
    <#
    .SYNOPSIS
    Open a Microsoft Word document in the UI.

    .DESCRIPTION
    Launches the Word application and makes the specified document visible to the user for editing.

    .PARAMETER FilePath
    The absolute path to the .docx file.

    .EXAMPLE
    Invoke-WordOpenDocument -FilePath "C:\Doc.docx"
    Opens Doc.docx in Word.
    #>
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
    <#
    .SYNOPSIS
    Append text to an existing Word document.

    .DESCRIPTION
    Opens a Word document in the background, moves the cursor to the very end of the file, inserts the specified text as new paragraphs, and saves the file.

    .PARAMETER FilePath
    The absolute path to the .docx file.

    .PARAMETER Text
    The text to append.

    .EXAMPLE
    Invoke-WordAppendText -FilePath "C:\Doc.docx" -Text "Addendum: approved."
    Adds the text to the end of the document.
    #>
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
    <#
    .SYNOPSIS
    Find and replace text in a Word document.

    .DESCRIPTION
    Performs a global find-and-replace operation across the entire Word document in the background, then saves the file.

    .PARAMETER FilePath
    The absolute path to the .docx file.

    .PARAMETER SearchText
    The string to look for.

    .PARAMETER ReplaceText
    The string to replace it with.

    .EXAMPLE
    Invoke-WordReplaceText -FilePath "C:\Template.docx" -SearchText "[NAME]" -ReplaceText "Alice"
    Replaces all instances of [NAME] with Alice.
    #>
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
