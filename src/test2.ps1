$r = Invoke-WebRequest -Uri 'https://lite.duckduckgo.com/lite/' -Method Post -Body @{ q = "iran news" } -UseBasicParsing -TimeoutSec 10 -Headers @{ "User-Agent" = "Mozilla/5.0" }
$r.Content | Select-String -Pattern 'href="([^"]+)"' -AllMatches | ForEach-Object { $_.Matches.Groups[1].Value } | Select-Object -Unique | Select-Object -First 10
