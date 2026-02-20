# ═══════════════════════════════════════════════════════════════
# Test Runner for generate_screenshot_docx
# ═══════════════════════════════════════════════════════════════
# This script runs the DOCX generator using the test data in this folder.
# Run from the generate_screenshot_docx/ parent folder:
#   .\test\run-test.ps1

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Testing generate_screenshot_docx                 " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Paths (relative to generate_screenshot_docx/ folder)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir

Push-Location $ProjectDir

try {
    # Check npm dependencies
    if (-not (Test-Path "node_modules")) {
        Write-Host "[setup] Installing npm dependencies..." -ForegroundColor Yellow
        npm install
        Write-Host ""
    }

    # Run the generator
    Write-Host "[test] Generating PRE Transfer Control Study DOCX..." -ForegroundColor Green
    Write-Host ""

    node generate_screenshot_docx.js `
        --config "./test/Test_Step_Mapping.xlsx" `
        --module "MEDS_REPORTER" `
        --testing-phase "PRE" `
        --study-type "Control" `
        --section "4.1.6" `
        --study-id "TEST_STUDY_001" `
        --images "./test/images_PRE_Control" `
        --module-display-name "MEDS Reporter" `
        --logo "./test/Test_Logo.jpg" `
        --date "18-FEB-2026" `
        --tested-by "Test User" `
        --vars "./test/test_vars.json" `
        --output "./test"

    Write-Host ""
    Write-Host "[test] Check the output DOCX in the test/ folder!" -ForegroundColor Green

    # Open the output
    $OutputFile = Get-ChildItem "./test/*.docx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($OutputFile) {
        Write-Host "[test] Opening: $($OutputFile.Name)" -ForegroundColor Cyan
        Start-Process $OutputFile.FullName
    }
}
finally {
    Pop-Location
}
