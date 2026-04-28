$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# Replace pause icon (U+23F8 + U+FE0F) near offlineRoot with lock icon (U+1F512)
$pauseChar = [char]0x23F8
$lockStr = [char]::ConvertFromUtf32(0x1F512)

# Find all occurrences near offlineRoot
$idx = 0
$count = 0
while (($idx = $content.IndexOf('offlineRoot', $idx)) -ge 0) {
    # Search within next 300 chars for the pause icon
    $searchEnd = [Math]::Min($idx + 300, $content.Length)
    $searchRegion = $content.Substring($idx, $searchEnd - $idx)
    $pausePos = $searchRegion.IndexOf($pauseChar)
    if ($pausePos -ge 0) {
        $absPos = $idx + $pausePos
        # Check if followed by variation selector U+FE0F
        $removeLen = 1
        if ($absPos + 1 -lt $content.Length -and $content[$absPos + 1] -eq [char]0xFE0F) {
            $removeLen = 2
        }
        $content = $content.Substring(0, $absPos) + $lockStr + $content.Substring($absPos + $removeLen)
        $count++
    }
    $idx += 10
}

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding $true))
Write-Host "Replaced $count pause icons with lock icons near offlineRoot"
