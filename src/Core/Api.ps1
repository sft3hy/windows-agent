function Invoke-GeminiAPI {
    param([array]$Messages)

    $modMessages = @()
    foreach ($m in $Messages) {
        $modMessages += @{ role = $m.role; content = $m.content }
    }

    # If no system message found in history, prepend it
    if (-not ($modMessages | Where-Object { $_.role -eq "system" })) {
        $modMessages = @(@{ role = "system"; content = $script:TOOL_SYSTEM }) + $modMessages
    }

    $body = @{
        model = $script:CONFIG.Model; messages = $modMessages
        max_tokens = $script:CONFIG.MaxTokens; temperature = $script:CONFIG.Temperature
    } | ConvertTo-Json -Depth 10

    try {
        $maxRetries = 3
        $retryCount = 0
        $success = $false
        $resp = $null

        while (-not $success -and $retryCount -lt $maxRetries) {
            try {
                $resp = Invoke-RestMethod -Uri $script:CONFIG.ApiEndpoint -Method POST -Body $body -Headers @{
                    "Authorization" = "Bearer $($script:CONFIG.ApiKey)"
                    "Content-Type"  = "application/json; charset=utf-8"
                } -TimeoutSec 60
                $success = $true
            } catch {
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Write-StatusLine "WARN" "API call failed (Attempt $retryCount/$maxRetries): $_. Retrying in 2s..."
                    Start-Sleep -Seconds 2
                } else {
                    throw $_
                }
            }
        }

        if ($resp.choices[0].message.content) {
            return $resp.choices[0].message.content
        } else {
            return ""
        }
    } catch {
        Write-StatusLine "ERR" "API call failed after $maxRetries attempts: $_"
        return "ERROR: API unavailable - $_"
    }
}