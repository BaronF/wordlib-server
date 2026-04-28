$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logoPath = 'C:\Users\Baron\Desktop\logo.jpg'
if (-not (Test-Path $logoPath)) {
    Write-Host "ERROR: logo.jpg not found"
    exit 1
}
$bytes = [System.IO.File]::ReadAllBytes($logoPath)
Write-Host "File size: $($bytes.Length) bytes"
Write-Host "First 4 bytes: $($bytes[0]) $($bytes[1]) $($bytes[2]) $($bytes[3])"
$b64 = [Convert]::ToBase64String($bytes)
Write-Host "Base64 length: $($b64.Length)"
$outFile = Join-Path $scriptDir "logo_base64.txt"
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($outFile, $b64, $utf8)
Write-Host "OK saved to $outFile"
