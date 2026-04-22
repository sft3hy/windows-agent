# ============================================================
# POWERPOINT INTEGRATION
# ============================================================

function Invoke-PowerPointCreate {
    <#
    .SYNOPSIS
    Create a new PowerPoint presentation from JSON data.

    .DESCRIPTION
    Launches PowerPoint and dynamically builds a presentation slide-by-slide based on a provided JSON array containing title and content fields. Features a robust save mechanism that bypasses COM save errors on network drives by saving locally and copying.

    .PARAMETER FilePath
    The absolute path where the new .pptx file should be saved.

    .PARAMETER SlidesJson
    A JSON string array containing slide objects (e.g., '[{"title":"Slide 1","content":"Bullet 1\nBullet 2"}]').

    .EXAMPLE
    Invoke-PowerPointCreate -FilePath "C:\Pres.pptx" -SlidesJson '[{"title":"Intro","content":"Hello"}]'
    Creates a 1-slide presentation.
    #>
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
        $localTemp = Join-Path $env:TEMP "AIRWAV_PPT_$(Get-Random).pptx"
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
    <#
    .SYNOPSIS
    Open a PowerPoint presentation.

    .DESCRIPTION
    Launches the PowerPoint application and makes the specified presentation visible to the user.

    .PARAMETER FilePath
    The absolute path to the .pptx file.

    .EXAMPLE
    Invoke-PowerPointOpen -FilePath "C:\Pres.pptx"
    Opens Pres.pptx in the PowerPoint UI.
    #>
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
