$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 1. Remove the delete button span from root cards
# Match the delete span: '<span style="cursor:pointer;...deleteRoot...title="删除">🗑️</span>'
# It starts after the offline span's closing condition
$deletePattern = "(?s)(<span[^>]*onclick=""event\.stopPropagation\(\);deleteRoot\(\d*'\+r\.id\+'\d*\)""[^>]*>).{1,10}(</span>'\s*\+)"
# Simpler approach: find the line with deleteRoot and remove it
$lines = $content -split "`n"
$newLines = @()
$skipNext = $false
foreach ($line in $lines) {
    if ($line -match 'deleteRoot\(' -and $line -match 'title=') {
        # Skip this delete button line
        continue
    }
    $newLines += $line
}
$content = $newLines -join "`n"

# 2. Replace pause icon with lock icon for root offline button
# Find the offline button in root cards (near offlineRoot)
$pauseIcon = [char]0x23F8 # ⏸
$lockIcon = [char]0x1F512 # 🔒
# Replace only near offlineRoot context
$content = $content -replace "($([regex]::Escape('offlineRoot')).*?)$pauseIcon", "`$1$lockIcon"

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding $true))
Write-Host "Done: removed delete button from root cards and changed offline icon to lock"
