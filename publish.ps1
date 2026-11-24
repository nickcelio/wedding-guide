# Publishes the latest guide to https://nickcelio.github.io/wedding-guide/
# Usage: right-click > Run with PowerShell, or run .\publish.ps1 in a terminal
$env:Path += ';C:\Program Files\Git\cmd'
Set-Location $PSScriptRoot
Copy-Item .\nicaragua-travel-guide.html .\index.html -Force
git add -A
git commit -m "Update travel guide"
git push
Write-Host "`nPublished! Live in ~1 minute at https://nickcelio.github.io/wedding-guide/" -ForegroundColor Green
