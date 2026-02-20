# ═══════════════════════════════════════════════════════════════
# Test: Section 4.4.6 — POST Transfer, Transfer Study, MEDS Reporter
# 33 steps: 16 export/compare + 1 add-user + 16 re-export/verify
# ═══════════════════════════════════════════════════════════════

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Test: 4.4.6 POST Transfer - Transfer Study      " -ForegroundColor Cyan
Write-Host "  33 steps, ~50 screenshots                        " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir

Push-Location $ProjectDir

try {
    if (-not (Test-Path "node_modules")) {
        Write-Host "[setup] Installing npm dependencies..." -ForegroundColor Yellow
        npm install
        Write-Host ""
    }

    Write-Host "[test] Generating POST Transfer Transfer Study DOCX (4.4.6)..." -ForegroundColor Green
    Write-Host ""

    node generate_screenshot_docx.js `
        --config "./test/Test_Step_Mapping.xlsx" `
        --module "MEDS_REPORTER" `
        --module-display-name "MEDS Reporter" `
        --testing-phase "POST" `
        --study-type "Transfer" `
        --section "4.4.6" `
        --study-id "TEST_STUDY_002" `
        --images "./test/images_POST_Transfer_4.4.6" `
        --logo "./test/Test_Logo.jpg" `
        --date "18-FEB-2026" `
        --tested-by "Test User" `
        --vars "./test/test_vars.json" `
        --output "./test"

    Write-Host ""

    $OutputFile = Get-ChildItem "./test/*4.4.6*Screenshots*.docx" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($OutputFile) {
        Write-Host "[test] Opening: $($OutputFile.Name)" -ForegroundColor Cyan
        Start-Process $OutputFile.FullName
    }
}
finally {
    Pop-Location
}
