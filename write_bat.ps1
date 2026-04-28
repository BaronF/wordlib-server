$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $scriptDir "start_content.txt"
$dst = "C:\Users\Baron\Desktop\wordlib-system\start.bat"
$content = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::UTF8)
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($dst, $content, $utf8)
Write-Host "OK: $dst"
