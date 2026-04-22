# ============================================================
# WEB — HTTP fetch + Google search scraping
# ============================================================

# Fetch a single URL and return readable plain text (strips HTML tags).
# Returns up to $MaxChars characters so it fits in the model context window.
function Invoke-WebFetch {
    <#
    .SYNOPSIS
    Fetch and extract readable text from a URL.

    .DESCRIPTION
    Performs an HTTP GET request to the specified URL using a modern User-Agent. It strips all HTML tags, scripts, and styles, returning only the human-readable text content. Designed to feed web content into an LLM context window.

    .PARAMETER Url
    The web address to fetch (e.g., 'https://en.wikipedia.org/wiki/PowerShell').

    .PARAMETER MaxChars
    Optional. The maximum number of characters to return to prevent overflowing the LLM context. Defaults to 4000.

    .EXAMPLE
    Invoke-WebFetch -Url "https://news.ycombinator.com"
    Returns the readable text from the Hacker News homepage.
    #>
    param(
        [string]$Url,
        [int]$MaxChars = 4000
    )
    Write-ToolLine "Browser" "Fetching URL" $Url
    try {
        $headers = @{
            "User-Agent"                = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
            "Accept"                    = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"
            "Accept-Language"           = "en-US,en;q=0.9"
            "Sec-Ch-Ua"                 = "`"Chromium`";v=`"124`", `"Google Chrome`";v=`"124`", `"Not-A.Brand`";v=`"99`""
            "Sec-Ch-Ua-Mobile"          = "?0"
            "Sec-Ch-Ua-Platform"        = "`"Windows`""
            "Sec-Fetch-Dest"            = "document"
            "Sec-Fetch-Mode"            = "navigate"
            "Sec-Fetch-Site"            = "none"
            "Sec-Fetch-User"            = "?1"
            "Upgrade-Insecure-Requests" = "1"
        }
        $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 15 -Headers $headers
        # Strip HTML: remove tags, decode entities, collapse whitespace
        $text = $response.Content
        $text = [regex]::Replace($text, '<script[^>]*>[\s\S]*?</script>', '', 'IgnoreCase')
        $text = [regex]::Replace($text, '<style[^>]*>[\s\S]*?</style>',  '', 'IgnoreCase')
        $text = [regex]::Replace($text, '<[^>]+>', ' ')
        $text = [System.Net.WebUtility]::HtmlDecode($text)
        $text = [regex]::Replace($text, '\s{3,}', "`n`n")
        $text = $text.Trim()
        if ($text.Length -gt $MaxChars) { $text = $text.Substring(0, $MaxChars) + "`n...[truncated]" }
        Write-StatusLine "OK" "Fetched $($text.Length) chars from $Url"
        return $text
    } catch {
        Write-StatusLine "ERR" "Web fetch failed for '$Url': $_"
        return "ERROR: $_"
    }
}

# Web search → scrape the top N result URLs → return all content.
# Uses DuckDuckGo HTML results page to avoid captchas/consent redirects.
function Invoke-WebSearch {
    <#
    .SYNOPSIS
    Perform a web search and scrape the top results.

    .DESCRIPTION
    Uses DuckDuckGo (to avoid captchas) to search for a query, extracts the top N result URLs, and then uses Invoke-WebFetch to download and read the content of those pages. Returns an aggregated text block containing all source material.

    .PARAMETER Query
    The search terms.

    .PARAMETER TopN
    Optional. The number of top search results to visit and read. Defaults to 3.

    .PARAMETER MaxChars
    Optional. The maximum characters to read per page. Defaults to 3000.

    .EXAMPLE
    Invoke-WebSearch -Query "Latest PowerShell features"
    Searches the web and returns text summaries from the top 3 pages.
    #>
    param(
        [string]$Query,
        [int]$TopN     = 3,
        [int]$MaxChars = 3000   # per page
    )
    Write-ToolLine "Browser" "Web research" "Top $TopN results for: $Query"
    try {
        $encoded    = [Uri]::EscapeDataString($Query)
        $searchUrl  = "https://html.duckduckgo.com/html/?q=$encoded"
        $searchResp = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -TimeoutSec 15 `
                          -Headers @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" }

        # Extract DuckDuckGo proxy links: href="//duckduckgo.com/l/?uddg="
        $rawLinks = [regex]::Matches($searchResp.Content, 'class="result__url" href="//duckduckgo\.com/l/\?uddg=([^"&]+)') |
                    ForEach-Object { [Uri]::UnescapeDataString($_.Groups[1].Value) } |
                    Where-Object { $_ -notmatch 'google\.com|youtube\.com' -and $_ -match '^https?://' } |
                    Select-Object -Unique |
                    Select-Object -First $TopN

        if (-not $rawLinks) {
            Write-StatusLine "WARN" "No result URLs extracted from DuckDuckGo — returning search page text"
            return Invoke-WebFetch -Url $searchUrl -MaxChars ($MaxChars * $TopN)
        }

        $results = @()
        $i = 1
        foreach ($link in $rawLinks) {
            Write-StatusLine "INFO" "Reading result $i/$($rawLinks.Count): $link"
            $content = Invoke-WebFetch -Url $link -MaxChars $MaxChars
            $results += "=== SOURCE $i`: $link ===`n$content`n"
            $i++
        }
        return $results -join "`n"
    } catch {
        Write-StatusLine "ERR" "Web search failed: $_"
        return "ERROR: $_"
    }
}