$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 找到导航栏词根库那行，图标已损坏（U+FFFD）
# 替换为 🔤 (U+1F524)
$targetIcon = [char]::ConvertFromUtf32(0x1F524)
$replacement = [char]0xFFFD

# 找 data-t="roots" 附近的损坏图标
$idx = $content.IndexOf('data-t="roots"')
$count = 0
while ($idx -ge 0) {
    # 往前搜索最近的 U+FFFD
    $searchStart = [Math]::Max(0, $idx - 80)
    $region = $content.Substring($searchStart, $idx - $searchStart)
    $fffdPos = $region.LastIndexOf($replacement)
    if ($fffdPos -ge 0) {
        $absPos = $searchStart + $fffdPos
        $content = $content.Substring(0, $absPos) + $targetIcon + $content.Substring($absPos + 1)
        $count++
        Write-Host "Fixed icon near position $absPos"
    }
    $idx = $content.IndexOf('data-t="roots"', $idx + 1)
}

Write-Host "Replaced $count icons"
[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding $true))
Write-Host "Done"
