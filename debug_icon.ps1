$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 找导航栏词根库那行
$idx = $content.IndexOf('data-t="roots"')
if ($idx -ge 0) {
    # 取前面50个字符
    $start = [Math]::Max(0, $idx - 50)
    $region = $content.Substring($start, 50)
    Write-Host "Region before data-t=roots:"
    for ($i = 0; $i -lt $region.Length; $i++) {
        $ch = $region[$i]
        $code = [int]$ch
        if ($code -gt 127 -or $code -lt 32) {
            Write-Host ("  [$i] U+{0:X4} (non-ASCII)" -f $code)
        }
    }
    Write-Host "Text: [$region]"
}
