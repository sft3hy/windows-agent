function Invoke-GeminiAPI {
    param([array]$Messages)

    $modMessages = @()
    $systemInjected = $false
    foreach ($m in $Messages) {
        if ($m.role -eq "system") { continue }
        $clone = @{ role = $m.role; content = $m.content }
        if ($m.role -eq "user" -and -not $systemInjected) {
            $clone.content = "[SYSTEM]`n" + $script:TOOL_SYSTEM + "`n[/SYSTEM]`n`n" + $clone.content
            $systemInjected = $true
        }
        $modMessages += $clone
    }

    $body = @{
        model = $script:CONFIG.Model; messages = $modMessages
        max_tokens = $script:CONFIG.MaxTokens; temperature = $script:CONFIG.Temperature
    } | ConvertTo-Json -Depth 10

    try {
        $resp = Invoke-RestMethod -Uri $script:CONFIG.ApiEndpoint -Method POST -Body $body -Headers @{
            "Authorization" = "Bearer $($script:CONFIG.ApiKey)"
            "Content-Type"  = "application/json"
        }
        if ($resp.choices[0].message.content) {
            return $resp.choices[0].message.content
        } else {
            return ""
        }
    } catch {
        Write-StatusLine "ERR" "API call failed: $_"
        return "ERROR: API unavailable - $_"
    }
}