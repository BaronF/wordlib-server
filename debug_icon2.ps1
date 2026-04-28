$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

$idx = $content.IndexOf('data-t="roots"')
if ($idx -ge 0) {
    # 取后面60个字符（包含 onclick 之后的内容）
    $region = $content.Substring($idx, [Math]::Min(120, $content.Length - $idx))
    Write-Host "Region after data-t=roots:"
    for ($i = 0; $i -lt $region.Length; $i++) {
        $ch = $region[$i]
        $code = [int]$ch
        if ($code -gt 127) {
            Write-Host ("  [$i] U+{0:X4}" -f $code)
        }
    }
    Write-Host "---"
    Write-Host $region
}
