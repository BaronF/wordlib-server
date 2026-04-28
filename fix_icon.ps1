$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 找到词条审核tab中的文本，用正则替换图标
# 匹配 data-at="word" 那行中 gap:6px"> 后面到 词条审核 之间的内容
$pattern = '(data-at="word"[^>]*gap:6px">)(.+?)(词条审核)'
$replacement = "`${1}`u{1F4DA} `${3}"
$content = [regex]::Replace($content, $pattern, $replacement)

# 同样修复待审核词条标题中的图标
$pattern2 = '(align-items:center;gap:8px">).+?(待审核词条)'
$replacement2 = "`${1}`u{1F4DA} `${2}"
$content = [regex]::Replace($content, $pattern2, $replacement2)

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Done: fixed word audit icons"
