# log_error.ps1 — Append error output to errors/run_log.txt and push to GitHub.
#
# Usage (pipe):
#   python scripts/generate_workbook.py 2>&1 | .\scripts\log_error.ps1
#
# Usage (interactive paste):
#   .\scripts\log_error.ps1
#   [paste error, then press Ctrl+Z and Enter to finish]

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$logFile  = Join-Path $repoRoot "errors\run_log.txt"

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$divider   = "=" * 60
$header    = "$divider`n[$timestamp]`n$divider"

if ($MyInvocation.ExpectingInput) {
    $content = $input | Out-String
} else {
    Write-Host "Paste error output below. Press Ctrl+Z then Enter when done:"
    $content = $input | Out-String
}

Add-Content -Path $logFile -Value "$header`n$content`n" -Encoding UTF8
Write-Host "Logged to errors/run_log.txt"

Push-Location $repoRoot
git add errors/run_log.txt
git commit -m "wip: error log $timestamp"
git push
Pop-Location

Write-Host "Pushed. Switch to personal machine and pull to have Claude read it."
