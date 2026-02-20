/**
 * docx_builder.js
 * Generates a Screenshot Evidence DOCX identical to handleWordDoc_1.js output.
 *
 * Formatting spec (exact match to Playwright automation):
 *   - Font: Verdana, 11pt (22 half-points)
 *   - Header: Medidata logo, right-aligned, 150×80px
 *   - Footer: TEMPLATE-SDLC-018-05 + page numbers
 *   - Cover: bold 35pt, centered, top/bottom borders
 *   - Step header: blue #0070C0, not bold
 *   - Image: 600×320px, centered
 *   - Table: 9-row metadata, gray borders, gray shading on labels
 */

const fs = require("fs");
const path = require("path");
const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    ImageRun,
    Table,
    TableRow,
    TableCell,
    Header,
    Footer,
    PageNumber,
    Tab,
    PageBreak,
    AlignmentType,
    BorderStyle,
    WidthType,
    ShadingType,
    TabStopType,
    VerticalAlign,
} = require("docx");

// ─── Constants (exact match to handleWordDoc_1.js) ───────────
const FONT = "Verdana";
const SIZE_BODY = 22;       // 11pt in half-points
const SIZE_COVER = 70;      // 35pt in half-points
const SIZE_FOOTER = 18;     // 9pt in half-points
const COLOR_BLUE = "0070C0";
const COLOR_GRAY_BORDER = "808080";
const COLOR_GRAY_SHADING = "D9D9D9";
const IMG_WIDTH = 600;
const IMG_HEIGHT = 320;
const LOGO_WIDTH = 150;
const LOGO_HEIGHT = 80;
const COL1_WIDTH = 2500;    // DXA (~1.7 inches)
const COL2_WIDTH = 5000;    // DXA (~3.5 inches)
const FOOTER_TEXT =
    "TEMPLATE-SDLC-018-05              PROPRIETARY - LIMITED DISTRIBUTION           ";

// ─── Table border helper ─────────────────────────────────────
const TABLE_BORDERS = {
    top:              { style: BorderStyle.SINGLE, size: 2, color: COLOR_GRAY_BORDER },
    bottom:           { style: BorderStyle.SINGLE, size: 2, color: COLOR_GRAY_BORDER },
    left:             { style: BorderStyle.SINGLE, size: 2, color: COLOR_GRAY_BORDER },
    right:            { style: BorderStyle.SINGLE, size: 2, color: COLOR_GRAY_BORDER },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: COLOR_GRAY_BORDER },
    insideVertical:   { style: BorderStyle.SINGLE, size: 2, color: COLOR_GRAY_BORDER },
};

// ─── Create a single metadata table row ──────────────────────
function metadataRow(label, value) {
    return new TableRow({
        children: [
            new TableCell({
                width: { size: COL1_WIDTH, type: WidthType.DXA },
                shading: { type: ShadingType.CLEAR, fill: COLOR_GRAY_SHADING },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: label, font: FONT, size: SIZE_BODY }),
                        ],
                    }),
                ],
            }),
            new TableCell({
                width: { size: COL2_WIDTH, type: WidthType.DXA },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: value, font: FONT, size: SIZE_BODY }),
                        ],
                    }),
                ],
            }),
        ],
    });
}

// ─── Create the 9-row metadata table ─────────────────────────
function metadataTable(opts) {
    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        alignment: AlignmentType.CENTER,
        borders: TABLE_BORDERS,
        rows: [
            metadataRow("Step Name:",                opts.stepName),
            metadataRow("Expected Result:",          opts.expectedResult),
            metadataRow("Step Number:",              String(opts.stepNumber)),
            metadataRow("Screenshot Number:",        `${opts.shotCurrent} of ${opts.shotTotal}`),
            metadataRow("Comment(s):",               opts.comment),
            metadataRow("Pass / Fail step status:",  opts.passFail),
            metadataRow("Date:",                     opts.date),
            metadataRow("Tested By:",                opts.testedBy),
            metadataRow("Reviewed By:",              opts.reviewedBy),
        ],
    });
}

// ─── Build a single screenshot page (header + image + table) ─
function buildScreenshotPage(opts) {
    const elements = [];

    // 1. Blue step header — only on the FIRST screenshot of the step (e.g., "1 of 4")
    if (opts.isFirstShot) {
        elements.push(
            new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 0, after: 0 },
                children: [
                    new TextRun({
                        text: `Step ${opts.stepNumber}:  ${opts.stepText}`,
                        font: FONT,
                        size: SIZE_BODY,
                        bold: false,
                        color: COLOR_BLUE,
                    }),
                ],
            })
        );
    }

    // 2. Screenshot image (centered, 600×320)
    const imageBuffer = fs.readFileSync(opts.imagePath);
    elements.push(
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200 },
            children: [
                new ImageRun({
                    data: imageBuffer,
                    transformation: { width: IMG_WIDTH, height: IMG_HEIGHT },
                }),
            ],
        })
    );

    // 3. Spacer before table
    elements.push(
        new Paragraph({
            spacing: { before: 700 },
            children: [],
        })
    );

    // 4. Metadata table
    elements.push(
        metadataTable({
            stepName:       opts.stepText,
            expectedResult: opts.expectedResult,
            stepNumber:     opts.stepNumber,
            shotCurrent:    opts.shotCurrent,
            shotTotal:      opts.shotTotal,
            comment:        opts.stepText,
            passFail:       opts.passFail,
            date:           opts.date,
            testedBy:       opts.testedBy,
            reviewedBy:     opts.reviewedBy || "NA",
        })
    );

    return elements;
}

// ─── Build the cover page paragraph ──────────────────────────
function buildCoverPage(section, tag) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 3000, after: 3000 },
        border: {
            top:    { style: BorderStyle.SINGLE, size: 4 },
            bottom: { style: BorderStyle.SINGLE, size: 4 },
        },
        children: [
            new TextRun({
                text: "Screenshot Document for",
                font: FONT,
                size: SIZE_COVER,
                bold: true,
            }),
            new TextRun({
                text: `Section: ${section}`,
                font: FONT,
                size: SIZE_COVER,
                bold: true,
                break: 1,
            }),
            new TextRun({
                text: tag,
                font: FONT,
                size: SIZE_COVER,
                bold: true,
                break: 1,
            }),
        ],
    });
}

// ─── Build the complete DOCX ─────────────────────────────────
/**
 * @param {Object} opts
 * @param {string} opts.section       - Section number (e.g., "4.3.6")
 * @param {string} opts.tag           - Document tag (used for filename only)
 * @param {string} opts.moduleName    - Module display name for cover page (e.g., "MEDS Reporter", "Coder")
 * @param {string} opts.logoPath      - Path to Medidata_Logo.jpg
 * @param {string} opts.date          - Date string (DD-MMM-YYYY)
 * @param {string} opts.testedBy      - Tester name
 * @param {string} opts.passFail      - Default Pass/Fail value
 * @param {string} opts.reviewedBy    - Reviewer name (default "NA")
 * @param {Array}  opts.pages         - Array of page objects:
 *   { stepNumber, stepText, expectedResult, imagePath, shotCurrent, shotTotal }
 * @param {string} opts.outputPath    - Full output file path
 */
async function buildDocx(opts) {
    const { section, tag, moduleName, logoPath, date, testedBy, passFail, reviewedBy, pages, outputPath } = opts;

    // Read logo
    if (!fs.existsSync(logoPath)) {
        throw new Error(`Logo file not found: ${logoPath}`);
    }
    const logoBuffer = fs.readFileSync(logoPath);

    // Build all document children
    const children = [];

    // Cover page (third line = module display name only, not the full tag)
    children.push(buildCoverPage(section, moduleName));

    // Page break after cover
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // Screenshot pages
    for (let i = 0; i < pages.length; i++) {
        const page = pages[i];

        const pageElements = buildScreenshotPage({
            stepText:       page.stepText,
            expectedResult: page.expectedResult,
            stepNumber:     page.stepNumber,
            imagePath:      page.imagePath,
            shotCurrent:    page.shotCurrent,
            shotTotal:      page.shotTotal,
            isFirstShot:    page.shotCurrent === 1,
            passFail:       passFail,
            date:           date,
            testedBy:       testedBy,
            reviewedBy:     reviewedBy || "NA",
        });

        children.push(...pageElements);

        // Page break after each screenshot page (except the last)
        if (i < pages.length - 1) {
            children.push(new Paragraph({ children: [new PageBreak()] }));
        }
    }

    // Assemble document
    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: { font: FONT, size: SIZE_BODY },
                },
            },
        },
        sections: [
            {
                headers: {
                    default: new Header({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.RIGHT,
                                children: [
                                    new ImageRun({
                                        data: logoBuffer,
                                        transformation: {
                                            width: LOGO_WIDTH,
                                            height: LOGO_HEIGHT,
                                        },
                                    }),
                                ],
                            }),
                        ],
                    }),
                },
                footers: {
                    default: new Footer({
                        children: [
                            new Paragraph({
                                tabStops: [
                                    {
                                        type: TabStopType.RIGHT,
                                        position: 9000,
                                    },
                                ],
                                children: [
                                    new TextRun({
                                        text: FOOTER_TEXT,
                                        font: FONT,
                                        size: SIZE_FOOTER,
                                        bold: false,
                                    }),
                                    new TextRun({
                                        children: [new Tab(), PageNumber.CURRENT],
                                        font: FONT,
                                        size: SIZE_FOOTER,
                                    }),
                                    new TextRun({
                                        text: " / ",
                                        font: FONT,
                                        size: SIZE_FOOTER,
                                    }),
                                    new TextRun({
                                        children: [PageNumber.TOTAL_PAGES],
                                        font: FONT,
                                        size: SIZE_FOOTER,
                                    }),
                                ],
                            }),
                        ],
                    }),
                },
                children: children,
            },
        ],
    });

    // Write to file
    const buffer = await Packer.toBuffer(doc);
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }
    fs.writeFileSync(outputPath, buffer);

    console.log(`[docx_builder] DOCX written: ${outputPath}`);
    console.log(`[docx_builder] Total pages: ${pages.length + 1} (cover + ${pages.length} screenshots)`);
}

module.exports = { buildDocx };
