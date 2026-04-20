# ============================================================
# EXCEL INTEGRATION
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

function Invoke-ExcelWriteCell {
    param([string]$FilePath, [string]$SheetName = "Sheet1", [string]$Cell, [string]$Value)
    Write-ToolLine "Excel" "Writing cell" "$Cell in $SheetName"
    try {
        if (-not (Test-Path $FilePath)) { return "ERROR: File not found: $FilePath" }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $wb = $excel.Workbooks.Open($FilePath)
        $ws = try { $wb.Sheets[$SheetName] } catch { $wb.ActiveSheet }
        
        $ws.Range($Cell).Value2 = $Value
        
        $wb.Save()
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-StatusLine "OK" "Wrote '$Value' to cell $Cell"
        return "Cell $Cell successfully updated to '$Value' in $FilePath."
    } catch {
        Write-StatusLine "ERR" "Excel write failed: $_"
        try { $excel.Quit() } catch {}
        return "ERROR: $_"
    }
}
