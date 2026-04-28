$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 找到审核中心Tab区域的两个div，交换顺序
# 匹配 word tab 的 div
$wordTabPattern = '(<div class="audit-tab) on(" data-at="word")'
$rootTabPattern = '(<div class="audit-tab)(" data-at="root")'

# 把 word tab 的 on 去掉，给 root tab 加上 on
$content = $content.Replace('audit-tab on" data-at="word"', 'audit-tab" data-at="word"')
$content = $content.Replace('audit-tab" data-at="root"', 'audit-tab on" data-at="root"')

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Done: swapped audit tab default selection"
