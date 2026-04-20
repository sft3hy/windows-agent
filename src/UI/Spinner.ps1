function Get-ThinkingWord {
    $wordsFile = Join-Path $PSScriptRoot "..\..\thinking-words.txt"
    if (Test-Path $wordsFile) {
        $words = Get-Content $wordsFile | Where-Object { $_.Trim() -ne "" }
        if ($words.Count -gt 0) { return $words[(Get-Random -Maximum $words.Count)] }
    }
    return "Thinking"
}

function Start-Spinner {
    param([string]$Word = "Thinking")
    $script:SpinnerShared = [hashtable]::Synchronized(@{ Running = $true })
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"; $rs.ThreadOptions = "ReuseThread"; $rs.Open()
    $rs.SessionStateProxy.SetVariable("Shared", $script:SpinnerShared)
    $rs.SessionStateProxy.SetVariable("SpinWord", $Word)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        $frames = @('|','/','-','\'); $colors = @(11, 13, 9, 3) # ConsoleColors
        $i = 0; $label = "  $SpinWord "
        while ($Shared.Running) {
            [Console]::ForegroundColor = $colors[$i % $colors.Count]
            [Console]::Write("`r" + $label + $frames[$i % $frames.Count])
            [Console]::ResetColor(); $i++; [Threading.Thread]::Sleep(100)
        }
        [Console]::Write("`r" + (' ' * ($label.Length + 4)) + "`r")
    })
    $script:SpinnerPS = $ps; $script:SpinnerRS = $rs; $script:SpinnerHandle = $ps.BeginInvoke()
}

function Stop-Spinner {
    if ($script:SpinnerShared) {
        $script:SpinnerShared.Running = $false
        Start-Sleep -Milliseconds 200
        try { $script:SpinnerPS.EndInvoke($script:SpinnerHandle); $script:SpinnerPS.Dispose(); $script:SpinnerRS.Close() } catch {}
    }
}