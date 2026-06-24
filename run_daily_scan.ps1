$ErrorActionPreference = "Stop"

$InputWorkbook = "C:\Users\chris\Downloads\JOB APPLICATION TEMPLATE.xlsx"
$OutputDir = "C:\Users\chris\Documents\Codex\2026-06-23\files-mentioned-by-the-user-job\outputs\daily_scans"
$Scanner = Join-Path $PSScriptRoot "summer_2027_scanner.py"
$BundledPython = "C:\Users\chris\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if (Test-Path $BundledPython) {
    $Python = $BundledPython
} else {
    $Python = "python"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

& $Python $Scanner --input $InputWorkbook --output-dir $OutputDir

