# ============================================================
# BROWSER INTEGRATION — Chrome, Edge, Firefox
# ============================================================

# Find the first available path from a list of candidates
function Find-BrowserPath {
    param([array]$Candidates)
    return $Candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

# -----------  Chrome  -----------
function Invoke-ChromeOpen {
    <#
    .SYNOPSIS
    Open a URL in Google Chrome.

    .DESCRIPTION
    Launches Google Chrome and navigates to the specified URL. Searches standard installation paths to locate the executable.

    .PARAMETER Url
    The web address to open (e.g., 'https://www.google.com').

    .EXAMPLE
    Invoke-ChromeOpen -Url "https://www.example.com"
    Opens example.com in Google Chrome.
    #>
    param([string]$Url)
    Write-ToolLine "Chrome" "Opening URL" $Url
    try {
        $path = Find-BrowserPath @(
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
            "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
            "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
        )
        if ($path) { Start-Process $path -ArgumentList $Url }
        else        { Start-Process $Url }
        Write-StatusLine "OK" "Chrome → $Url"
        return "Chrome opened to: $Url"
    } catch {
        Write-StatusLine "ERR" "Chrome open failed: $_"
        return "ERROR: $_"
    }
}

# -----------  Edge  -----------
function Invoke-EdgeOpen {
    <#
    .SYNOPSIS
    Open a URL in Microsoft Edge.

    .DESCRIPTION
    Launches Microsoft Edge and navigates to the specified URL. Searches standard installation paths to locate the executable.

    .PARAMETER Url
    The web address to open.

    .EXAMPLE
    Invoke-EdgeOpen -Url "https://www.example.com"
    Opens example.com in Microsoft Edge.
    #>
    param([string]$Url)
    Write-ToolLine "Edge" "Opening URL" $Url
    try {
        $path = Find-BrowserPath @(
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
            "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
            "$env:LocalAppData\Microsoft\Edge\Application\msedge.exe"
        )
        if ($path) { Start-Process $path -ArgumentList $Url }
        else        { Start-Process $Url }
        Write-StatusLine "OK" "Edge → $Url"
        return "Edge opened to: $Url"
    } catch {
        Write-StatusLine "ERR" "Edge open failed: $_"
        return "ERROR: $_"
    }
}

# -----------  Firefox  -----------
function Invoke-FirefoxOpen {
    <#
    .SYNOPSIS
    Open a URL in Mozilla Firefox.

    .DESCRIPTION
    Launches Mozilla Firefox and navigates to the specified URL. Searches standard installation paths to locate the executable.

    .PARAMETER Url
    The web address to open.

    .EXAMPLE
    Invoke-FirefoxOpen -Url "https://www.example.com"
    Opens example.com in Mozilla Firefox.
    #>
    param([string]$Url)
    Write-ToolLine "Firefox" "Opening URL" $Url
    try {
        $path = Find-BrowserPath @(
            "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
            "$env:ProgramFiles(x86)\Mozilla Firefox\firefox.exe",
            "$env:LocalAppData\Mozilla Firefox\firefox.exe"
        )
        if ($path) { Start-Process $path -ArgumentList $Url }
        else        { Start-Process $Url }
        Write-StatusLine "OK" "Firefox → $Url"
        return "Firefox opened to: $Url"
    } catch {
        Write-StatusLine "ERR" "Firefox open failed: $_"
        return "ERROR: $_"
    }
}

# -----------  Smart launcher  -----------
# Tries Chrome → Edge → Firefox → Windows default
function Invoke-BrowserOpen {
    <#
    .SYNOPSIS
    Open a URL in a specified or automatically detected web browser.

    .DESCRIPTION
    Attempts to launch the URL in the specified browser. If no browser is specified, it will intelligently fallback through Chrome, Edge, and the system default browser.

    .PARAMETER Url
    The web address to navigate to.

    .PARAMETER Browser
    Optional. The preferred browser to use. Valid options: 'chrome', 'edge', 'firefox'.

    .EXAMPLE
    Invoke-BrowserOpen -Url "https://github.com"
    Opens GitHub in the default detected browser (Chrome -> Edge -> System Default).

    .EXAMPLE
    Invoke-BrowserOpen -Url "https://github.com" -Browser "firefox"
    Explicitly forces the URL to open in Mozilla Firefox.
    #>
    param([string]$Url, [string]$Browser = "")
    Write-ToolLine "Browser" "Opening URL" $Url

    if ($Browser -eq "chrome") { return Invoke-ChromeOpen -Url $Url }
    if ($Browser -eq "edge")   { return Invoke-EdgeOpen   -Url $Url }
    if ($Browser -eq "firefox") { return Invoke-FirefoxOpen -Url $Url }

    # Auto-detect: try Chrome first, then Edge, then Firefox, then system default
    $chromePath = Find-BrowserPath @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )
    $edgePath = Find-BrowserPath @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
    )

    try {
        if     ($chromePath) { Start-Process $chromePath  -ArgumentList $Url; Write-StatusLine "OK" "Chrome → $Url" }
        elseif ($edgePath)   { Start-Process $edgePath    -ArgumentList $Url; Write-StatusLine "OK" "Edge → $Url" }
        else                 { Start-Process $Url;                             Write-StatusLine "OK" "System browser → $Url" }
        return "Browser opened to: $Url"
    } catch {
        Write-StatusLine "ERR" "Browser open failed: $_"
        return "ERROR: $_"
    }
}

# -----------  Search  -----------
# Opens a web search in the preferred browser
function Invoke-BrowserSearch {
    <#
    .SYNOPSIS
    Perform a web search in the browser using the specified engine.

    .DESCRIPTION
    Constructs a valid search URL by encoding your query and launches it in a web browser. Supports Google, Bing, and DuckDuckGo.

    .PARAMETER Query
    The search terms to look up.

    .PARAMETER Engine
    Optional. The search engine to use. Valid options: 'google', 'bing', 'duckduckgo'. Default is 'google'.

    .PARAMETER Browser
    Optional. The preferred browser to use (e.g., 'chrome', 'edge').

    .EXAMPLE
    Invoke-BrowserSearch -Query "Powershell documentation"
    Searches Google for 'Powershell documentation' using the default browser.

    .EXAMPLE
    Invoke-BrowserSearch -Query "Cybersecurity news" -Engine "duckduckgo" -Browser "firefox"
    Searches DuckDuckGo using Firefox.
    #>
    param(
        [string]$Query,
        [string]$Engine = "google",  # google | bing | duckduckgo
        [string]$Browser = ""
    )
    Write-ToolLine "Browser" "Searching" "$Engine`: $Query"
    $encoded = [Uri]::EscapeDataString($Query)
    $baseUrl = switch ($Engine.ToLower()) {
        "bing"       { "https://www.bing.com/search?q=" }
        "duckduckgo" { "https://duckduckgo.com/?q=" }
        default      { "https://www.google.com/search?q=" }
    }
    $url = $baseUrl + $encoded
    Write-StatusLine "INFO" "Search URL: $url"
    return Invoke-BrowserOpen -Url $url -Browser $Browser
}
