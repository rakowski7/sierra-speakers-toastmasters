# PowerShell helper: copies the real screenshots Claude captured into the images folder
# Run from the docs folder:  powershell -ExecutionPolicy Bypass -File .\copy-screenshots.ps1
#
# Claude saves screenshots to the per-session outputs folder while it works.
# On your computer that folder lives under the Claude AppData LocalCache path.
# This script finds the latest screenshots and renames them to the filenames
# the tutorial markdown and PDF expect.

$ErrorActionPreference = "Stop"

$repoImages = Join-Path $PSScriptRoot "..\images"
$repoImages = (Resolve-Path $repoImages).Path

# The 11 screenshots in the exact order Claude captured them, with target names.
$map = @(
    @{ Epoch = 1776129292991; Name = "01-sheet-overview.jpg" }
    @{ Epoch = 1776129418412; Name = "02-menu-dropdown.jpg" }
    @{ Epoch = 1776129434973; Name = "03-rc-confirm-sheet.jpg" }
    @{ Epoch = 1776129453647; Name = "04-rc-date.jpg" }
    @{ Epoch = 1776129498670; Name = "05-rc-status.jpg" }
    @{ Epoch = 1776129532668; Name = "06-ag-confirm-sheet.jpg" }
    @{ Epoch = 1776129563499; Name = "07-ag-date.jpg" }
    @{ Epoch = 1776129652559; Name = "08-ag-word.jpg" }
    @{ Epoch = 1776129703329; Name = "09-ag-speeches.jpg" }
    @{ Epoch = 1776129797018; Name = "10-email-dialog.jpg" }
    @{ Epoch = 1776129932387; Name = "11-schedule-grid.jpg" }
)

# Search common locations for the Claude outputs folder.
$candidates = @(
    (Join-Path $env:USERPROFILE "Documents\Claude\outputs"),
    (Join-Path $env:USERPROFILE "OneDrive\Documents\Claude\outputs"),
    (Join-Path $env:LOCALAPPDATA "Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude")
)

$found = @()
foreach ($root in $candidates) {
    if (Test-Path $root) {
        $found += Get-ChildItem -Path $root -Filter "screenshot-*.jpg" -Recurse -ErrorAction SilentlyContinue
    }
}

if (-not $found) {
    Write-Warning "No screenshot-*.jpg files found in common Claude outputs locations."
    Write-Host "Searched:"
    foreach ($c in $candidates) { Write-Host "  $c" }
    Write-Host ""
    Write-Host "Find your Claude outputs folder manually, then run:"
    Write-Host "  `$claudeOut = 'PATH\\TO\\outputs'"
    Write-Host "  ls `$claudeOut\\screenshot-*.jpg"
    exit 1
}

$copied = 0
foreach ($entry in $map) {
    $pattern = "screenshot-$($entry.Epoch).jpg"
    $match = $found | Where-Object { $_.Name -eq $pattern } | Select-Object -First 1
    if ($match) {
        $dest = Join-Path $repoImages $entry.Name
        Copy-Item $match.FullName $dest -Force
        Write-Host "copied $($match.Name) -> $($entry.Name)"
        $copied++
    } else {
        Write-Warning "missing: $pattern (expected target $($entry.Name))"
    }
}

Write-Host ""
Write-Host "Done. $copied / $($map.Count) screenshots copied into $repoImages"
Write-Host "Now rebuild the PDF if you want it to include the real screenshots."
