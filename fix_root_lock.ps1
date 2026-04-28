$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 词根卡片下线按钮：图标被损坏，只剩变体选择器 U+FE0F
# 需要在 offlineRoot 附近找到 title="下线"> 后面的残留字符，替换为锁图标
$lockStr = [char]::ConvertFromUtf32(0x1F512)

# 精确匹配：title="下线"> 后跟变体选择器残留 + </span>
$pattern = '(title="' + [regex]::Escape('"') + ')([^<]{0,5})(</span>)'
# 更简单：直接找 offlineRoot 所在行，替换 title="下线">X</span> 中的 X
$lines = $content -split "`n"
for ($i = 0; $i -lt $lines.Length; $i++) {
    if ($lines[$i] -match 'offlineRoot' -and $lines[$i] -match 'rcard-actions|title=') {
        # 替换 title="下线">任意短内容</span> 为 title="下线">锁图标</span>
        $old = $lines[$i]
        $lines[$i] = $lines[$i] -replace '(title=\\?"下线\\?">)[^<]{0,10}(</span>)', "`${1}$lockStr`${2}"
        if ($lines[$i] -ne $old) {
            Write-Host "Fixed line $($i+1)"
        }
    }
}
$content = $lines -join "`n"

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding $true))
Write-Host "Done"
