/**
 * image_scanner.js
 * Scans an image folder and matches files to step numbers using naming conventions.
 *
 * Supported naming patterns:
 *   - Jira_Shot.png                                          → global Jira fallback (Shot 1 for all export steps)
 *   - Step_{N}_Shot_{M}.png                                  → step-specific screenshot
 *   - UseCase_{Section}_{StudyID}_Step_{N}_Shot_2.png        → pipeline Excel screenshot
 *   - UseCase_{Section}_{StudyID}_Compare_Step_{N}_Shot_1.png → pipeline Compare screenshot
 */

const fs = require("fs");
const path = require("path");

const IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".bmp", ".gif"]);

/**
 * Scan image folder and build a step→images map.
 *
 * @param {string} imagesDir   - Path to the images folder
 * @param {string} section     - Section number (e.g., "4.3.6")
 * @param {string} studyId     - Study ID (e.g., "D8313C00001")
 * @returns {{ stepImages: Map<number, Array<{shot: number, filePath: string}>>, jiraGlobal: string|null }}
 */
function scanImages(imagesDir, section, studyId) {
    if (!fs.existsSync(imagesDir)) {
        throw new Error(`Images folder not found: ${imagesDir}`);
    }

    const files = fs.readdirSync(imagesDir);
    const stepImages = new Map(); // step_number → [{ shot, filePath }]
    let jiraGlobal = null;

    // Regex patterns
    const reStepShot = /^Step_(\d+)_Shot_(\d+)\.(png|jpg|jpeg|bmp|gif)$/i;
    const rePipelineExport = new RegExp(
        `^UseCase_${escapeRegex(section)}_${escapeRegex(studyId)}_Step_(\\d+)_Shot_(\\d+)\\.(png|jpg|jpeg|bmp|gif)$`,
        "i"
    );
    const rePipelineCompare = new RegExp(
        `^UseCase_${escapeRegex(section)}_${escapeRegex(studyId)}_Compare_Step_(\\d+)_Shot_(\\d+)\\.(png|jpg|jpeg|bmp|gif)$`,
        "i"
    );
    const reJiraGlobal = /^Jira_Shot\.(png|jpg|jpeg|bmp|gif)$/i;

    for (const file of files) {
        const ext = path.extname(file).toLowerCase();
        if (!IMAGE_EXTENSIONS.has(ext)) continue;

        const fullPath = path.join(imagesDir, file);
        let match;

        // 1. Jira global fallback
        if (reJiraGlobal.test(file)) {
            jiraGlobal = fullPath;
            continue;
        }

        // 2. Step-specific: Step_{N}_Shot_{M}.png
        match = file.match(reStepShot);
        if (match) {
            const stepNum = parseInt(match[1], 10);
            const shotNum = parseInt(match[2], 10);
            addToMap(stepImages, stepNum, shotNum, fullPath);
            continue;
        }

        // 3. Pipeline compare: UseCase_{Section}_{StudyID}_Compare_Step_{N}_Shot_{M}.png
        match = file.match(rePipelineCompare);
        if (match) {
            const stepNum = parseInt(match[1], 10);
            const shotNum = parseInt(match[2], 10);
            addToMap(stepImages, stepNum, shotNum, fullPath);
            continue;
        }

        // 4. Pipeline export: UseCase_{Section}_{StudyID}_Step_{N}_Shot_{M}.png
        match = file.match(rePipelineExport);
        if (match) {
            const stepNum = parseInt(match[1], 10);
            const shotNum = parseInt(match[2], 10);
            addToMap(stepImages, stepNum, shotNum, fullPath);
            continue;
        }
    }

    // Sort each step's images by shot number
    for (const [stepNum, images] of stepImages) {
        images.sort((a, b) => a.shot - b.shot);
    }

    const totalImages = [...stepImages.values()].reduce((sum, arr) => sum + arr.length, 0);
    console.log(
        `[image_scanner] Found ${totalImages} images across ${stepImages.size} steps` +
        (jiraGlobal ? " + 1 Jira global fallback" : "")
    );

    return { stepImages, jiraGlobal };
}

/**
 * Resolve images for a specific step, applying Jira fallback logic.
 *
 * @param {Map} stepImages     - The step→images map
 * @param {string|null} jiraGlobal - Path to Jira_Shot.png (or null)
 * @param {number} stepNumber  - The step number
 * @param {boolean} isExportStep - Whether this is an export step (needs Jira as Shot 1)
 * @returns {Array<{shot: number, filePath: string}>}
 */
function resolveStepImages(stepImages, jiraGlobal, stepNumber, isExportStep) {
    const images = stepImages.get(stepNumber) || [];

    if (isExportStep && images.length > 0) {
        // Check if Shot 1 exists
        const hasShot1 = images.some((img) => img.shot === 1);
        if (!hasShot1 && jiraGlobal) {
            // Insert Jira global as Shot 1
            images.unshift({ shot: 1, filePath: jiraGlobal });
        }
    } else if (isExportStep && images.length === 0 && jiraGlobal) {
        // No images at all — only add Jira as Shot 1 (Shot 2 is missing, will warn)
        images.push({ shot: 1, filePath: jiraGlobal });
    }

    return images;
}

function addToMap(map, stepNum, shotNum, filePath) {
    if (!map.has(stepNum)) {
        map.set(stepNum, []);
    }
    // Avoid duplicates for same shot number (step-specific overrides pipeline)
    const existing = map.get(stepNum);
    const existingIdx = existing.findIndex((img) => img.shot === shotNum);
    if (existingIdx >= 0) {
        // Step-specific (Step_{N}_Shot_{M}) takes priority — check if new one is step-specific
        if (/^Step_/i.test(require("path").basename(filePath))) {
            existing[existingIdx] = { shot: shotNum, filePath };
        }
    } else {
        existing.push({ shot: shotNum, filePath });
    }
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

module.exports = { scanImages, resolveStepImages };
