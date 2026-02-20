/**
 * config_reader.js
 * Reads step definitions from an Excel config file (Step_Mapping sheet).
 */

const XLSX = require("xlsx");
const path = require("path");

/**
 * Read step mapping from Excel config.
 *
 * @param {string} configPath  - Path to the .xlsx config file
 * @param {string} sectionKey  - Section key to filter (e.g., "POST Transfer QC_Control Study_MEDS_REPORTER")
 * @param {string} [sheetName] - Sheet name (default: "MEDS_Reporter_Step_Mapping")
 * @returns {Array<{stepNumber: number, stepText: string, expectedResult: string}>}
 */
function readStepMapping(configPath, sectionKey, sheetName) {
    if (!require("fs").existsSync(configPath)) {
        throw new Error(`Config file not found: ${configPath}`);
    }

    const workbook = XLSX.readFile(configPath);

    // Determine sheet name
    const targetSheet = sheetName || "MEDS_Reporter_Step_Mapping";
    if (!workbook.Sheets[targetSheet]) {
        const available = workbook.SheetNames.join(", ");
        throw new Error(
            `Sheet "${targetSheet}" not found in ${path.basename(configPath)}.\n` +
            `Available sheets: ${available}`
        );
    }

    const sheet = workbook.Sheets[targetSheet];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Skip header row (row 0), filter by Section_No (col 0)
    const steps = [];
    const allKeys = new Set();

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || !row[0]) continue;

        const key = String(row[0]).trim();
        allKeys.add(key);

        if (key === sectionKey) {
            steps.push({
                stepNumber: parseInt(row[1], 10),
                stepText: String(row[2] || "").trim(),
                expectedResult: String(row[3] || "").trim(),
            });
        }
    }

    if (steps.length === 0) {
        const available = [...allKeys].join("\n  - ");
        throw new Error(
            `No steps found for section key: "${sectionKey}"\n` +
            `Available keys in "${targetSheet}":\n  - ${available}`
        );
    }

    // Sort by step number
    steps.sort((a, b) => a.stepNumber - b.stepNumber);

    console.log(`[config_reader] Found ${steps.length} steps for "${sectionKey}"`);
    return steps;
}

/**
 * Apply variable replacements to step text.
 * Replaces {Key} patterns with values from the vars object.
 *
 * @param {Array} steps - Array of step definitions
 * @param {Object} vars - Key-value pairs for replacement
 * @returns {Array} Steps with replaced text
 */
function applyVars(steps, vars) {
    if (!vars || Object.keys(vars).length === 0) return steps;

    return steps.map((step) => {
        let text = step.stepText;
        let expected = step.expectedResult;

        for (const [key, value] of Object.entries(vars)) {
            const pattern = new RegExp(`\\{${key}\\}`, "g");
            text = text.replace(pattern, value);
            expected = expected.replace(pattern, value);
        }

        return { ...step, stepText: text, expectedResult: expected };
    });
}

module.exports = { readStepMapping, applyVars };
