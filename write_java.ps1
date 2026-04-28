$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $scriptDir "java_code.txt"
$dst = Join-Path $scriptDir "WordLibServer.java"
$content = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::UTF8)
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($dst, $content, $utf8)
Write-Host "OK: $dst"
