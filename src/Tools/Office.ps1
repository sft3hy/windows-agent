# ============================================================
# EXCEL & POWERPOINT INTEGRATION
# ============================================================

function Invoke-ExcelReadFile {
    param([string]$FilePath, [string]$SheetName = "", [int]$MaxRows = 100)
    Write-ToolLine "Excel" "Reading file" $FilePath
    try {
        if (-not (Test-Path $FilePath)) { return "ERROR: File not found: $FilePath" }
        $excel    = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb       = $excel.Workbooks.Open($FilePath)
        if ($SheetName) { $ws = $wb.Sheets[$SheetName] } else { $ws = $wb.ActiveSheet }
        $used     = $ws.UsedRange
        $rows     = $used.Rows.Count
        $cols     = $used.Columns.Count
        $data     = @()
        $maxR     = [Math]::Min($rows, $MaxRows)
        for ($r = 1; $r -le $maxR; $r++) {
            $row = @()
            for ($c = 1; $c -le $cols; $c++) {
                $row += $ws.Cells($r, $c).Text
            }
            $data += ,($row -join "`t")
        }
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-StatusLine "OK" "Read $maxR rows × $cols cols from $($ws.Name)"
        return ($data -join "`n")
    } catch {
        Write-StatusLine "ERR" "Excel read failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-ExcelCreateFile {
    param([string]$FilePath, [string]$Data, [string]$SheetName = "Sheet1")
    Write-ToolLine "Excel" "Creating file" $FilePath
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb    = $excel.Workbooks.Add()
        $ws    = $wb.ActiveSheet
        $ws.Name = $SheetName
        $lines = $Data -split "`n"
        $r = 1
        foreach ($line in $lines) {
            $cells = $line -split "`t"
            $c = 1
            foreach ($cell in $cells) {
                $ws.Cells($r, $c) = $cell
                $c++
            }
            $r++
        }
        # Use format 51 (xlOpenXMLWorkbook) to ensure proper .xlsx output
        $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
        $fmt = if ($ext -eq ".xlsx") { 51 } else { [Type]::Missing }
        if ($fmt -eq [Type]::Missing) { $wb.SaveAs($FilePath) } else { $wb.SaveAs($FilePath, $fmt) }
        $wb.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-StatusLine "OK" "Created $FilePath"
        return "Excel file created at $FilePath"
    } catch {
        Write-StatusLine "ERR" "Excel create failed: $_"
        try { $excel.Quit() } catch {}
        return "ERROR: $_"
    }
}

function Invoke-PowerPointCreate {
    param([string]$FilePath, [string]$SlidesJson)
    Write-ToolLine "PowerPoint" "Creating presentation" $FilePath
    try {
        $slides = $SlidesJson | ConvertFrom-Json
        $ppt    = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = 1
        $pres   = $ppt.Presentations.Add()

        foreach ($slide in $slides) {
            $sl = $pres.Slides.Add($pres.Slides.Count + 1, 1)  # 1 = ppLayoutTitle
            # Shape 1 = title placeholder (always present in layout 1)
            if ($slide.title -and $sl.Shapes.Count -ge 1) {
                try { $sl.Shapes[1].TextFrame.TextRange.Text = $slide.title } catch {}
            }
            # Shape 2 = content placeholder (present only when layout has one)
            if ($slide.content -and $sl.Shapes.Count -ge 2) {
                try { $sl.Shapes[2].TextFrame.TextRange.Text = $slide.content } catch {}
            }
        }

        # PowerPoint COM SaveAs can fail on UNC / redirected shell folders.
        # Strategy: save to a guaranteed-local %TEMP% path first, then copy.
        $localTemp = Join-Path $env:TEMP "AIRI_PPT_$(Get-Random).pptx"
        $pres.SaveAs($localTemp)

        # If the final destination is different from temp, copy it there
        $finalPath = $FilePath
        if ($localTemp -ne $FilePath) {
            try {
                # Ensure destination directory exists
                $destDir = [System.IO.Path]::GetDirectoryName($FilePath)
                if ($destDir -and -not (Test-Path $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }
                Copy-Item -Path $localTemp -Destination $FilePath -Force
                Remove-Item $localTemp -Force -ErrorAction SilentlyContinue
            } catch {
                # Copy failed — fall back to the temp path so the file is still usable
                Write-StatusLine "WARN" "Could not copy to $FilePath, using temp: $localTemp"
                $finalPath = $localTemp
            }
        }

        Write-StatusLine "OK" "Created $($slides.Count)-slide presentation: $finalPath"
        # Return the actual path so callers (Engine.ps1) can thread it to email
        return $finalPath
    } catch {
        Write-StatusLine "ERR" "PowerPoint create failed: $_"
        return $null
    }
}

function Invoke-PowerPointOpen {
    param([string]$FilePath)
    Write-ToolLine "PowerPoint" "Opening file" $FilePath
    try {
        $ppt  = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = 1
        $pres = $ppt.Presentations.Open($FilePath)
        Write-StatusLine "OK" "Opened $($pres.Slides.Count) slides"
        return "Opened presentation with $($pres.Slides.Count) slides"
    } catch {
        Write-StatusLine "ERR" "PowerPoint open failed: $_"
        return "ERROR: $_"
    }
}