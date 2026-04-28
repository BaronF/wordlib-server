$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$bytes = [System.IO.File]::ReadAllBytes($file)
$content = [System.Text.Encoding]::UTF8.GetString($bytes)

$lockStr = [char]::ConvertFromUtf32(0x1F512)

# 直接搜索包含 offlineRoot 的行中 title="下线"> 和 </span> 之间的内容
$searchStr = 'title=' + [char]0x5C + '"' + [char]0x4E0B + [char]0x7EBF + [char]0x5C + '">'
Write-Host "Search approach 2: looking for exact pattern near offlineRoot"

# 找 offlineRoot 位置
$pos = $content.IndexOf('offlineRoot')
$found = $false
while ($pos -ge 0 -and -not $found) {
    # 从这个位置往后找 </span>
    $spanEnd = $content.IndexOf('</span>', $pos)
    if ($spanEnd -gt 0 -and ($spanEnd - $pos) -lt 500) {
        # 往回找 > (title 属性结束后的 >)
        $gtPos = $content.LastIndexOf('>', $spanEnd - 1)
        if ($gtPos -gt $pos -and ($spanEnd - $gtPos) -lt 20) {
            $between = $content.Substring($gtPos + 1, $spanEnd - $gtPos - 1)
            Write-Host "Found content between > and </span>: length=$($between.Length) chars"
            for ($c = 0; $c -lt $between.Length; $c++) {
                Write-Host ("  char[$c] = U+{0:X4}" -f [int]$between[$c])
            }
            # 替换这段内容为锁图标
            $content = $content.Substring(0, $gtPos + 1) + $lockStr + $content.Substring($spanEnd)
            $found = $true
            Write-Host "Replaced with lock icon"
        }
    }
    $pos = $content.IndexOf('offlineRoot', $pos + 1)
}

if (-not $found) {
    Write-Host "Pattern not found, trying alternative"
}

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding $true))
Write-Host "Done"
