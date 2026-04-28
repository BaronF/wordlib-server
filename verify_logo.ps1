$ErrorActionPreference = "Stop"
try {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $b64File = Join-Path $scriptDir "logo_base64.txt"
    $b64 = [System.IO.File]::ReadAllText($b64File).Trim()
    $decoded = [Convert]::FromBase64String($b64)
    $outFile = Join-Path $scriptDir "logo_verify.jpg"
    [System.IO.File]::WriteAllBytes($outFile, $decoded)
    $original = [System.IO.File]::ReadAllBytes('C:\Users\Baron\Desktop\logo.jpg')
    Write-Host ("Decoded: " + $decoded.Length + " Original: " + $original.Length)
    if ($decoded.Length -eq $original.Length) { Write-Host "MATCH" } else { Write-Host "MISMATCH" }
} catch {
    Write-Host ("ERROR: " + $_.Exception.Message)
}
