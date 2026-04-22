$r = Invoke-WebRequest -Uri 'https://html.duckduckgo.com/html/?q=iran+news'
$r.Content | Select-String -Pattern 'href="([^"]+)"' -AllMatches | ForEach-Object { $_.Matches.Groups[1].Value } | Select-Object -First 20
