$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$htmlFile = Join-Path $scriptDir "index.html"
$b64File = Join-Path $scriptDir "logo_base64.txt"
$b64 = [System.IO.File]::ReadAllText($b64File).Trim()
$html = [System.IO.File]::ReadAllText($htmlFile, [System.Text.Encoding]::UTF8)
$oldTag = '<img src="data:image/jpeg;base64,'
$startIdx = $html.IndexOf($oldTag)
if ($startIdx -lt 0) {
    Write-Host "ERROR: img tag not found"
    exit 1
}
$srcStart = $html.IndexOf('base64,', $startIdx) + 7
$srcEnd = $html.IndexOf('"', $srcStart)
$oldB64 = $html.Substring($srcStart, $srcEnd - $srcStart)
Write-Host "Old base64 length: $($oldB64.Length)"
Write-Host "New base64 length: $($b64.Length)"
$newHtml = $html.Substring(0, $srcStart) + $b64 + $html.Substring($srcEnd)
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlFile, $newHtml, $utf8)
Write-Host "OK: logo embedded"
