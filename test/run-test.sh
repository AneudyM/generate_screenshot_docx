#!/bin/bash
# Test Runner for generate_screenshot_docx
# Run from the generate_screenshot_docx/ parent folder:
#   bash test/run-test.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

cd "$PROJECT_DIR"

# Check npm dependencies
if [ ! -d "node_modules" ]; then
    echo "[setup] Installing npm dependencies..."
    npm install
    echo ""
fi

echo "[test] Generating PRE Transfer Control Study DOCX..."
echo ""

node generate_screenshot_docx.js \
    --config "./test/Test_Step_Mapping.xlsx" \
    --module "MEDS_REPORTER" \
    --testing-phase "PRE" \
    --study-type "Control" \
    --section "4.1.6" \
    --study-id "TEST_STUDY_001" \
    --images "./test/images_PRE_Control" \
    --logo "./test/Test_Logo.jpg" \
    --date "18-FEB-2026" \
    --tested-by "Test User" \
    --vars "./test/test_vars.json" \
    --output "./test"

echo ""
echo "[test] Check the output DOCX in the test/ folder!"
