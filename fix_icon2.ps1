$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 找到词条审核tab - 匹配 gap:6px"> 到 词条审核 之间的所有字符并替换
$idx = $content.IndexOf('data-at="word" onclick="switchAuditTab')
if ($idx -gt 0) {
    # 找到这行中 gap:6px"> 的位置
    $lineEnd = $content.IndexOf("`n", $idx)
    $line = $content.Substring($idx, $lineEnd - $idx)
    $gapIdx = $line.IndexOf('gap:6px">')
    if ($gapIdx -gt 0) {
        $afterGap = $gapIdx + 'gap:6px">'.Length
        $wordIdx = $line.IndexOf([char]0x8BCD + [char]0x6761 + [char]0x5BA1 + [char]0x6838)
        # 尝试直接搜索 "词条审核"
        $wordIdx2 = $line.IndexOf("`u{8BCD}`u{6761}`u{5BA1}`u{6838}")
        Write-Host "afterGap=$afterGap wordIdx=$wordIdx wordIdx2=$wordIdx2"
        if ($wordIdx2 -gt 0) {
            # 替换 gap:6px"> 到 词条审核 之间的内容为新图标
            $before = $line.Substring(0, $afterGap)
            $after = $line.Substring($wordIdx2)
            $newLine = $before + "`u{1F4DA} " + $after
            $content = $content.Remove($idx, $lineEnd - $idx).Insert($idx, $newLine)
            Write-Host "Replaced word tab icon successfully"
        }
    }
}

# 同样修复待审核词条标题
$idx2 = $content.IndexOf('id="pendingCountLabel"')
if ($idx2 -gt 0) {
    # 往前找到 h3 标签
    $lineStart = $content.LastIndexOf("`n", $idx2) + 1
    $lineEnd2 = $content.IndexOf("`n", $idx2)
    $line2 = $content.Substring($lineStart, $lineEnd2 - $lineStart)
    $gapIdx2 = $line2.IndexOf('gap:8px">')
    if ($gapIdx2 -gt 0) {
        $afterGap2 = $gapIdx2 + 'gap:8px">'.Length
        $wordIdx3 = $line2.IndexOf("`u{5F85}`u{5BA1}`u{6838}`u{8BCD}`u{6761}")
        Write-Host "h3 afterGap=$afterGap2 wordIdx=$wordIdx3"
        if ($wordIdx3 -gt 0) {
            $before2 = $line2.Substring(0, $afterGap2)
            $after2 = $line2.Substring($wordIdx3)
            $newLine2 = $before2 + "`u{1F4DA} " + $after2
            $content = $content.Remove($lineStart, $lineEnd2 - $lineStart).Insert($lineStart, $newLine2)
            Write-Host "Replaced h3 icon successfully"
        }
    }
}

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Done"
