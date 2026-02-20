# generate_screenshot_docx

Standalone Node.js CLI that generates **STS Screenshot Evidence DOCX** files from normalized pipeline artifacts — identical formatting to the Playwright automation (`handleWordDoc_1.js`).

## Installation

```bash
cd generate_screenshot_docx
npm install
```

## Quick Start

```bash
node generate_screenshot_docx.js \
    --config "./MEDS_REPORTER_Step_Mapping.xlsx" \
    --module "MEDS_REPORTER" \
    --testing-phase "POST" \
    --study-type "Control" \
    --section "4.3.6" \
    --study-id "D8313C00001" \
    --images "C:\Evidence\Artifacts\POST_Control\MEDS_Reporter" \
    --logo "C:\STS\assets\Medidata_Logo.jpg" \
    --date "30-JAN-2026" \
    --tested-by "Aneudy Mota"
```

**Output:** `POST Transfer QC_Control Study_MEDS Reporter_D8313C00001_Screenshots.docx`

## All CLI Arguments

| Argument | Required | Default | Description |
| --- | --- | --- | --- |
| `--config` | ✅ | — | Path to Excel config with step definitions |
| `--sheet` | ❌ | `MEDS_Reporter_Step_Mapping` | Sheet name in config Excel |
| `--module` | ✅ | — | Module key (e.g., `MEDS_REPORTER`) |
| `--testing-phase` | ✅ | — | `PRE` or `POST` |
| `--study-type` | ✅ | — | `Control` or `Transfer` |
| `--section` | ✅ | — | Section number (e.g., `4.3.6`) |
| `--study-id` | ✅ | — | Study identifier (e.g., `D8313C00001`) |
| `--images` | ✅ | — | Folder with normalized screenshot images |
| `--logo` | ✅ | — | Path to `Medidata_Logo.jpg` |
| `--date` | ✅ | — | Date for metadata (DD-MMM-YYYY) |
| `--tested-by` | ✅ | — | Tester name |
| `--reviewed-by` | ❌ | `NA` | Reviewer name |
| `--pass-fail` | ❌ | `Pass` | Pass/Fail value |
| `--output` | ❌ | `.` (current dir) | Output folder |
| `--vars` | ❌ | — | JSON file with `{Key}`→`Value` text replacements |
| `--section-key` | ❌ | Auto-constructed | Override the Step_Mapping lookup key |

## Image Naming Conventions

Place all images in the `--images` folder. The script recognizes these patterns:

### Pipeline-produced images (from existing automation)

| Pattern | Source Script | What It Is |
| --- | --- | --- |
| `UseCase_{Section}_{StudyID}_Step_{N}_Shot_2.png` | `Medidata-OpenExcelAndScreenshot.ps1` + Normalize | Excel report screenshot |
| `UseCase_{Section}_{StudyID}_Compare_Step_{N}_Shot_1.png` | `Medidata-SpreadsheetCompare-Auto.ps1` + Normalize | Spreadsheet Compare screenshot |

### Manual images (you provide these)

| Pattern | What It Is |
| --- | --- |
| `Jira_Shot.png` | Global Jira request screenshot (reused for all export steps as Shot 1) |
| `Step_{N}_Shot_1.png` | Step-specific Jira screenshot (overrides `Jira_Shot.png` for that step) |
| `Step_{N}_Shot_{M}.png` | Any step-specific screenshot (for special multi-shot steps) |

### Matching Priority

For export steps (Shot 1 = Jira, Shot 2 = Excel):

1. **Shot 1**: `Step_{N}_Shot_1.png` → fallback → `Jira_Shot.png`
2. **Shot 2**: `UseCase_{Section}_{StudyID}_Step_{N}_Shot_2.png`

For compare steps (Shot 1 = Spreadsheet Compare):

1. **Shot 1**: `UseCase_{Section}_{StudyID}_Compare_Step_{N}_Shot_1.png`

## Variable Replacements

If your step definitions contain placeholders like `{Source Study Group}`, create a JSON file:

```json
{
    "Source Study Group": "FORTREA",
    "Control Study": "D8313C00001",
    "Transfer Study 1": "D7060C00003"
}
```

Then pass it: `--vars ./my_vars.json`

## Examples

### PRE Transfer — Control Study — MEDS Reporter (8 steps, 16 screenshots)

```bash
node generate_screenshot_docx.js \
    --config "./MEDS_REPORTER_Step_Mapping.xlsx" \
    --module "MEDS_REPORTER" \
    --testing-phase "PRE" \
    --study-type "Control" \
    --section "4.1.6" \
    --study-id "D8313C00001" \
    --images "./artifacts/PRE_Control/MEDS_Reporter" \
    --logo "./assets/Medidata_Logo.jpg" \
    --date "26-AUG-2025" \
    --tested-by "Aneudy Mota"
```

### POST Transfer — Transfer Study — MEDS Reporter (33 steps, \~50 screenshots)

```bash
node generate_screenshot_docx.js \
    --config "./MEDS_REPORTER_Step_Mapping.xlsx" \
    --module "MEDS_REPORTER" \
    --testing-phase "POST" \
    --study-type "Transfer" \
    --section "4.4.6" \
    --study-id "D7060C00003" \
    --images "./artifacts/POST_Transfer/MEDS_Reporter" \
    --logo "./assets/Medidata_Logo.jpg" \
    --date "16-SEP-2025" \
    --tested-by "Aneudy Mota" \
    --vars "./fortrea_vars.json"
```

## Full Pipeline (MEDS Reporter)

```
1. Developer exports SQL query results (Excel files)
2. Medidata-Normalize-MEDSArtifacts.ps1  →  Rename files to standard names
3. Medidata-SortForCompare.ps1           →  Sort rows for comparison
4. Medidata-OpenExcelAndScreenshot.ps1   →  Screenshot each Excel file
5. Medidata-SpreadsheetCompare-Auto.ps1  →  Compare PRE vs POST + screenshot
6. Medidata-Normalize-MEDSArtifacts.ps1  →  Rename screenshots to standard names
7. You: Save Jira screenshot as Jira_Shot.png in the artifacts folder
8. generate_screenshot_docx.js           →  Build the DOCX ← THIS SCRIPT
```

## Dependencies

- [docx](https://www.npmjs.com/package/docx) `^8.5.0` — DOCX generation
- [xlsx](https://www.npmjs.com/package/xlsx) `^0.18.5` — Excel reading
- [yargs](https://www.npmjs.com/package/yargs) `^17.7.2` — CLI parsing