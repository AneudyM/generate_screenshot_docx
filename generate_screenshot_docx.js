#!/usr/bin/env node

/**
 * generate_screenshot_docx.js
 *
 * Standalone CLI to generate STS Screenshot Evidence DOCX files
 * from normalized pipeline artifacts + manual Jira screenshots.
 *
 * Produces output identical to handleWordDoc_1.js (Playwright automation).
 *
 * Usage:
 *   node generate_screenshot_docx.js \
 *     --config "./MEDS_REPORTER_Step_Mapping.xlsx" \
 *     --module "MEDS_REPORTER" \
 *     --testing-phase "POST" \
 *     --study-type "Control" \
 *     --section "4.3.6" \
 *     --study-id "D8313C00001" \
 *     --images "./artifacts/D8313C00001/MEDS_Reporter" \
 *     --logo "./assets/Medidata_Logo.jpg" \
 *     --date "30-JAN-2026" \
 *     --tested-by "Aneudy Mota"
 */

const yargs = require("yargs/yargs");
const { hideBin } = require("yargs/helpers");
const path = require("path");
const fs = require("fs");

const { readStepMapping, applyVars } = require("./lib/config_reader");
const { scanImages, resolveStepImages } = require("./lib/image_scanner");
const { buildDocx } = require("./lib/docx_builder");

// ─── CLI argument definitions ────────────────────────────────
const argv = yargs(hideBin(process.argv))
    .usage("Usage: $0 [options]")
    .option("config", {
        describe: "Path to Excel config with Step_Mapping sheet",
        type: "string",
        demandOption: true,
    })
    .option("sheet", {
        describe: "Sheet name in config Excel (default: MEDS_Reporter_Step_Mapping)",
        type: "string",
        default: "MEDS_Reporter_Step_Mapping",
    })
    .option("module", {
        describe: "Module key (e.g., MEDS_REPORTER)",
        type: "string",
        demandOption: true,
    })
    .option("testing-phase", {
        describe: "Testing phase: PRE or POST",
        type: "string",
        choices: ["PRE", "POST"],
        demandOption: true,
    })
    .option("study-type", {
        describe: "Study type: Control or Transfer",
        type: "string",
        choices: ["Control", "Transfer"],
        demandOption: true,
    })
    .option("section", {
        describe: "Section number from Use Cases doc (e.g., 4.3.6)",
        type: "string",
        demandOption: true,
    })
    .option("study-id", {
        describe: "Study identifier / study code (e.g., D8313C00001)",
        type: "string",
        demandOption: true,
    })
    .option("images", {
        describe: "Folder containing normalized screenshot images",
        type: "string",
        demandOption: true,
    })
    .option("logo", {
        describe: "Path to Medidata_Logo.jpg",
        type: "string",
        demandOption: true,
    })
    .option("date", {
        describe: "Date for metadata tables (DD-MMM-YYYY)",
        type: "string",
        demandOption: true,
    })
    .option("tested-by", {
        describe: 'Name for "Tested By" field',
        type: "string",
        demandOption: true,
    })
    .option("reviewed-by", {
        describe: 'Name for "Reviewed By" field',
        type: "string",
        default: "NA",
    })
    .option("pass-fail", {
        describe: "Default Pass/Fail value",
        type: "string",
        default: "Pass",
    })
    .option("output", {
        describe: "Output folder (default: current directory)",
        type: "string",
        default: ".",
    })
    .option("vars", {
        describe: "JSON file with {Key}→Value replacements for step text",
        type: "string",
    })
    .option("module-display-name", {
        describe: 'Exact module name for cover page (e.g., "MEDS Reporter", "Coder"). If not set, derives from --module by replacing underscores with spaces.',
        type: "string",
    })
    .option("section-key", {
        describe: "Override the auto-constructed section key for Step_Mapping lookup",
        type: "string",
    })
    .example(
        '$0 --config ./config.xlsx --module MEDS_REPORTER --testing-phase POST --study-type Control --section 4.3.6 --study-id D8313C00001 --images ./artifacts --logo ./logo.jpg --date "30-JAN-2026" --tested-by "Aneudy Mota"',
        "Generate POST Transfer Control Study MEDS Reporter evidence DOCX"
    )
    .wrap(100)
    .help()
    .alias("h", "help")
    .strict()
    .parseSync();

// ─── Main execution ──────────────────────────────────────────
async function main() {
    const {
        config,
        sheet,
        module: moduleName,
        testingPhase,
        studyType,
        section,
        studyId,
        images: imagesDir,
        logo,
        date,
        testedBy,
        reviewedBy,
        passFail,
        output: outputDir,
        vars: varsFile,
        sectionKey: sectionKeyOverride,
    } = argv;

    console.log("\n╔══════════════════════════════════════════════════╗");
    console.log("║   STS Screenshot Evidence DOCX Generator         ║");
    console.log("╚══════════════════════════════════════════════════╝\n");

    // 1. Section key = just the section number (matches Section_No column)
    const sectionKey = sectionKeyOverride || section;
    console.log(`[main] Section key: "${sectionKey}"`);

    // 2. Construct output filename
    //    Format: {Phase} Transfer QC_{StudyType} Study_{Module}_{StudyID}_Screenshots.docx
    const moduleDisplayName = argv.moduleDisplayName || moduleName.replace(/_/g, " ");
    const tag = `${testingPhase} Transfer QC_${studyType} Study_${moduleDisplayName}_${studyId}`;
    const outputFileName = `${tag}_Screenshots.docx`;
    const outputPath = path.join(outputDir, outputFileName);

    console.log(`[main] Output: ${outputPath}`);

    // 3. Read step definitions from config
    let steps = readStepMapping(config, sectionKey, sheet);

    // 4. Apply variable replacements if --vars provided
    if (varsFile) {
        if (!fs.existsSync(varsFile)) {
            throw new Error(`Vars file not found: ${varsFile}`);
        }
        const vars = JSON.parse(fs.readFileSync(varsFile, "utf-8"));
        console.log(
            `[main] Applying ${Object.keys(vars).length} variable replacements from ${varsFile}`
        );
        steps = applyVars(steps, vars);
    }

    // 5. Scan images folder
    const { stepImages, jiraGlobal } = scanImages(imagesDir, section, studyId);

    // 6. Build the pages array — one entry per screenshot (not per step)
    const pages = [];
    const warnings = [];

    for (const step of steps) {
        // Determine if this is an export step (needs Jira Shot 1)
        // Export steps contain "Request the developer" or "Run the SQL Query"
        // Compare steps contain "Compare the export"
        // Special steps (like add-user) are anything else
        const textLower = step.stepText.toLowerCase();
        const isExportStep =
            textLower.includes("request the developer") ||
            textLower.includes("run the sql query") ||
            textLower.includes("export the query result");
        const isCompareStep = textLower.includes("compare the export");

        const images = resolveStepImages(
            stepImages,
            jiraGlobal,
            step.stepNumber,
            isExportStep
        );

        if (images.length === 0) {
            warnings.push(
                `⚠️  Step ${step.stepNumber}: No images found — step will be SKIPPED`
            );
            continue;
        }

        const totalShots = images.length;
        for (const img of images) {
            if (!fs.existsSync(img.filePath)) {
                warnings.push(
                    `⚠️  Step ${step.stepNumber}, Shot ${img.shot}: File not found: ${img.filePath}`
                );
                continue;
            }

            pages.push({
                stepNumber:     step.stepNumber,
                stepText:       step.stepText,
                expectedResult: step.expectedResult,
                imagePath:      img.filePath,
                shotCurrent:    img.shot,
                shotTotal:      totalShots,
            });
        }
    }

    // 7. Print warnings
    if (warnings.length > 0) {
        console.log(`\n${"─".repeat(60)}`);
        console.log("WARNINGS:");
        for (const w of warnings) {
            console.log(`  ${w}`);
        }
        console.log(`${"─".repeat(60)}\n`);
    }

    if (pages.length === 0) {
        throw new Error(
            "No screenshot pages to generate! Check your --images folder and step mapping."
        );
    }

    console.log(`[main] Building DOCX with ${pages.length} screenshot pages...`);

    // 8. Build the DOCX
    await buildDocx({
        section,
        tag,
        moduleName: moduleDisplayName,
        logoPath: logo,
        date,
        testedBy,
        passFail,
        reviewedBy,
        pages,
        outputPath,
    });

    // 9. Summary
    console.log(`\n${"═".repeat(60)}`);
    console.log(`✅ DOCX generated successfully!`);
    console.log(`   File: ${outputPath}`);
    console.log(`   Steps: ${steps.length}`);
    console.log(`   Screenshots: ${pages.length}`);
    if (warnings.length > 0) {
        console.log(`   Warnings: ${warnings.length}`);
    }
    console.log(`${"═".repeat(60)}\n`);
}

main().catch((err) => {
    console.error(`\n❌ ERROR: ${err.message}\n`);
    process.exit(1);
});
