$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$bytes = [System.IO.File]::ReadAllBytes($file)
$content = [System.Text.Encoding]::UTF8.GetString($bytes)

# 方法：找到 data-at="word" 那行，定位 gap:6px"> 后面的字符
$marker = 'data-at="word" onclick="switchAuditTab'
$idx = $content.IndexOf($marker)
Write-Host "word tab at char index: $idx"

if ($idx -gt 0) {
    $gapStr = 'gap:6px">'
    $gapIdx = $content.IndexOf($gapStr, $idx)
    $afterGap = $gapIdx + $gapStr.Length
    
    # 找到这之后第一个中文字符（词条审核的"词"）
    # 扫描后面的字符，跳过非中文字符（emoji/空格等）
    $scanIdx = $afterGap
    $found = $false
    for ($i = $scanIdx; $i -lt $scanIdx + 20; $i++) {
        $c = [int]$content[$i]
        Write-Host "  char[$i] = U+$($c.ToString('X4')) '$($content[$i])'"
        # 中文字符范围 0x4E00-0x9FFF
        if ($c -ge 0x4E00 -and $c -le 0x9FFF) {
            Write-Host "  Found CJK at index $i"
            # 替换从 afterGap 到 $i 之间的内容为新图标+空格
            $newIcon = [char]0xD83D, [char]0xDCDA, ' ' -join ''  # 📚
            $content = $content.Remove($afterGap, $i - $afterGap).Insert($afterGap, $newIcon)
            $found = $true
            break
        }
    }
    if (-not $found) { Write-Host "CJK char not found near word tab" }
}

# 同样处理待审核词条的h3标题
$marker2 = 'id="pendingCountLabel"'
$idx2 = $content.IndexOf($marker2)
if ($idx2 -gt 0) {
    $lineStart = $content.LastIndexOf("`n", $idx2)
    $gapStr2 = 'gap:8px">'
    $gapIdx2 = $content.IndexOf($gapStr2, $lineStart)
    if ($gapIdx2 -gt 0) {
        $afterGap2 = $gapIdx2 + $gapStr2.Length
        for ($i = $afterGap2; $i -lt $afterGap2 + 20; $i++) {
            $c = [int]$content[$i]
            if ($c -ge 0x4E00 -and $c -le 0x9FFF) {
                $newIcon2 = [char]0xD83D, [char]0xDCDA, ' ' -join ''
                $content = $content.Remove($afterGap2, $i - $afterGap2).Insert($afterGap2, $newIcon2)
                Write-Host "Fixed h3 icon"
                break
            }
        }
    }
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($file, $content, $utf8NoBom)
Write-Host "Done"
