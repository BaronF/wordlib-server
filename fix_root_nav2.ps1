$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

$targetIcon = [char]::ConvertFromUtf32(0x1F524)
$fffd = [char]0xFFFD

# 找 data-t="roots" 后面的 U+FFFD
$idx = $content.IndexOf('data-t="roots"')
$count = 0
while ($idx -ge 0) {
    $searchEnd = [Math]::Min($idx + 80, $content.Length)
    $region = $content.Substring($idx, $searchEnd - $idx)
    $fffdPos = $region.IndexOf($fffd)
    if ($fffdPos -ge 0) {
        $absPos = $idx + $fffdPos
        $content = $content.Substring(0, $absPos) + $targetIcon + $content.Substring($absPos + 1)
        $count++
        Write-Host "Fixed icon at position $absPos"
    }
    $idx = $content.IndexOf('data-t="roots"', $idx + 10)
}

Write-Host "Replaced $count icons"
[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding $true))
Write-Host "Done"
