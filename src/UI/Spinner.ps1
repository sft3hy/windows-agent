function Get-ThinkingWord {
    $wordsFile = Join-Path $PSScriptRoot "..\..\thinking-words.txt"
    if (Test-Path $wordsFile) {
        $words = Get-Content $wordsFile | Where-Object { $_.Trim() -ne "" }
        if ($words.Count -gt 0) { return $words[(Get-Random -Maximum $words.Count)] }
    }
    return "Thinking"
}

function Start-Spinner {
    param([string]$Word = (Get-ThinkingWord))
    $script:SpinnerShared = [hashtable]::Synchronized(@{ Running = $true })
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"; $rs.ThreadOptions = "ReuseThread"; $rs.Open()
    $rs.SessionStateProxy.SetVariable("Shared", $script:SpinnerShared)
    $rs.SessionStateProxy.SetVariable("SpinWord", $Word)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        $frames = @('.', '..', '...', '....', '...', '..', '.')
        $i = 0; $label = "  $SpinWord "
        while ($Shared.Running) {
            # Note: $colors should be defined; if not, we fallback to default
            try { [Console]::ForegroundColor = [ConsoleColor]::Magenta } catch {}
            [Console]::Write("`r" + $label + $frames[$i % $frames.Count].PadRight(5))
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
        try { 
            if ($script:SpinnerPS) {
                $script:SpinnerPS.EndInvoke($script:SpinnerHandle)
                $script:SpinnerPS.Dispose()
            }
            if ($script:SpinnerRS) {
                $script:SpinnerRS.Close()
                $script:SpinnerRS.Dispose()
            }
        } catch {}
        $script:SpinnerShared = $null
        $script:SpinnerPS = $null
        $script:SpinnerRS = $null
        $script:SpinnerHandle = $null
    }
}