# Quick fix for "Invalid JSON primitive" error in NTP installer
# Run this script as Administrator before running install_ntp_timing_guided.ps1

Write-Host "Fixing JSON BOM issue..." -ForegroundColor Cyan

# Delete cached JSON file if it exists
$cachePath = "C:\Users\micha\Downloads\config\ntp-country-servers.json"
if (Test-Path -LiteralPath $cachePath) {
    Remove-Item -LiteralPath $cachePath -Force
    Write-Host "Deleted cached JSON file: $cachePath" -ForegroundColor Green
}

# Also check other possible cache locations
$altCache = Join-Path $env:TEMP "occultation-ntp-installer\config\ntp-country-servers.json"
if (Test-Path -LiteralPath $altCache) {
    Remove-Item -LiteralPath $altCache -Force
    Write-Host "Deleted cached JSON file: $altCache" -ForegroundColor Green
}

# Test the GitHub download with BOM stripping
Write-Host "Testing GitHub JSON download..." -ForegroundColor Cyan
try {
    $resp = Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/config/ntp-country-servers.json' -UseBasicParsing
    $text = $resp.Content
    
    # Strip UTF-8 BOM if present
    $bom = [System.Text.Encoding]::UTF8.GetPreamble()
    $bomString = [System.Text.Encoding]::UTF8.GetString($bom)
    if ($text.StartsWith($bomString)) {
        $text = $text.Substring($bomString.Length)
    }
    
    # Trim all leading whitespace
    $text = $text.TrimStart()
    
    # Try to parse JSON
    $obj = $text | ConvertFrom-Json
    Write-Host "SUCCESS: JSON parses correctly after BOM stripping" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "Now run the updated installer from:" -ForegroundColor Cyan
Write-Host "C:\Users\AstroPC\Documents\GitHub\occultation-ntp-installer\install_ntp_timing_guided.ps1" -ForegroundColor Yellow
Write-Host ""
Write-Host "If you're running from Downloads, copy the updated script from above location." -ForegroundColor Cyan