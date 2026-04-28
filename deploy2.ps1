$dst = Join-Path $env:USERPROFILE "Desktop\wordlib-system"
$src = Join-Path (Get-Location).Path "wordlib-server"
[System.IO.File]::Copy((Join-Path $src "server.py"), (Join-Path $dst "server.py"), $true)
[System.IO.File]::Copy((Join-Path $src "index.html"), (Join-Path $dst "index.html"), $true)
Write-Host "done"
Get-ChildItem $dst | Select-Object Name,Length
