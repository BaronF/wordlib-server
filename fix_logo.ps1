$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$htmlFile = Join-Path $scriptDir "index.html"
$html = [System.IO.File]::ReadAllText($htmlFile, [System.Text.Encoding]::UTF8)

# 找到 brand-logo div 中的 img 标签，替换 src
$pattern = '<img src="data:image/jpeg;base64,[^"]*"'
$replacement = '<img src="logo.jpg"'
$newHtml = [regex]::Replace($html, $pattern, $replacement)

if ($newHtml -eq $html) {
    Write-Host "NO MATCH - trying alt pattern"
    # 可能已经是 logo.jpg 了
    $pattern2 = 'src="logo\.jpg"'
    if ($newHtml -match $pattern2) {
        Write-Host "Already using logo.jpg"
    }
} else {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($htmlFile, $newHtml, $utf8)
    Write-Host "OK: replaced base64 with logo.jpg"
}
