import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";

// "Erm, Jair, this used to have emojis in them, why did you remove them?"
// Because apparently now using them is a "sign AI wrote the code, not you" (mfw I add emojis so user gets cozy lines and it gets called "AI code"),
// so to prevent people reading the code and thinking "oh no, this is AI generated code", bye emojis.

// Get cmd args
const args = process.argv.slice(2);

// Default values for the folders
let inputFolder = "./Raws";
let outputFolder = "./Extracted";

function showHelp() {
    console.log(`
Usage: {node | bun run} extract.ts [options]

    Options:
    -i, --input <folder>    Input folder (default: ./Raws)
    -o, --output <folder>   Output folder (default: ./Extracted)
    -h, --help              Show this help

Example:
    bun run extract.ts -i "./Raw files" -o "./XLSX files"
    `);
}

// Process args
for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === "-i" || arg === "--input") {
        inputFolder = args[i + 1] || inputFolder;
        i++;
    } else if (arg === "-o" || arg === "--output") {
        outputFolder = args[i + 1] || outputFolder;
        i++;
    } else if (arg === "-h" || arg === "--help") {
        showHelp();
        process.exit(0);
    }
}

console.log(`Reading JSON files from: ${inputFolder}`);
console.log(`Output Excel files to: ${outputFolder}`);

function escapeNewlines(text) {
    if (typeof text !== "string") return text;
    return text.replace(/\n/g, "\\n");
}

function parseJSON(fileContent) {
    const safe = fileContent.replace(/"m_Id"\s*:\s*(\d+)/g, '"m_Id":"$1"'); // wrap m_Id in quotes because big numbers make fucky wucky
    return JSON.parse(safe);
}

async function processAndCreateExcel(filePath) {
    try {
        const fileContent = fs.readFileSync(filePath, "utf8");
        const jsonData = parseJSON(fileContent);

        if (!jsonData.m_TableData || !Array.isArray(jsonData.m_TableData)) {
            console.warn(`Warning: ${filePath} doesn't have valid m_TableData array`);
            return false;
        }

        const totalRows = jsonData.m_TableData.length;

        const fileName = path.basename(filePath, ".json");
        console.log(`Processing ${fileName}: Found ${totalRows} entries`);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Translation Data");

        // Define column properties, also yes, I tabbed stuff because I want to read it properly.
        const columnStyles = [
            { header: "ID",                 key: "id",          width: 30,      font: { size: 12, color: { argb: "FFC73A46" /* Red */ } },        wrap: false },
            { header: "Original Line",      key: "orig",        width: 70,      font: { size: 12, color: { argb: "FF2A2AD4" /* Blue */ } },       wrap: true  },
            { header: "Translated Line",    key: "trans",       width: 70,      font: { size: 12, color: { argb: "FF000000" /* Black */ } },      wrap: true  },
            { header: "TL Notes",           key: "tlNote",      width: 30,      font: { size: 12, color: { argb: "FF693AC7" /* Purple */ } },     wrap: false },
            { header: "Editor Notes",       key: "edNote",      width: 30,      font: { size: 12, color: { argb: "FFC73A46" /* Red */ } },        wrap: false },
            { header: "Progress",           key: "progress",    width: 10,      font: { size: 12, color: { argb: "FF1C9128" /* Green */ } },      wrap: false },
            { header: "0/0",                key: "tlLines",     width: 15,      font: { size: 12, color: { argb: "FF1C9128" /* Green */ } },      wrap: false }
        ];

        // Freeze first row (headers)
        worksheet.views = [
            { state: "frozen", xSplit: 0, ySplit: 1 }
        ];

        // Set columns
        worksheet.columns = columnStyles.map(col => ({
            header: col.header,
            key: col.key,
            width: col.width
        }));

        // Header borders
        worksheet.getRow(1).border = {
            left: { style: 'medium', color: { argb: 'FF000000' } },
            bottom: { style: 'medium', color: { argb: 'FF000000' } },
            right: { style: 'medium', color: { argb: 'FF000000' } }
        };

        // Add progress stuff
        worksheet.getCell("G1").value = {
            formula: `COUNTA(C2:C${totalRows + 1}) & "/" & ${totalRows}`
        };

        // Add rows
        jsonData.m_TableData.forEach((item: any) => {
            worksheet.addRow({
                id: item.m_Id || "", // IDs we converted to strings before
                orig: escapeNewlines(item.m_Localized || ""), // Escape the original json file's newlines to avoid screwing up Excel cell formatting
                trans: "",
                tlNote: "",
                edNote: ""
            });
        });

        // Apply styles to each cell based on column
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                const style = columnStyles[colNumber - 1];
                if (style) {
                    cell.font = style.font;
                    cell.alignment = { wrapText: style.wrap };
                }
            });
        });

        const outputFile = path.join(outputFolder, `${fileName}.xlsx`);
        await workbook.xlsx.writeFile(outputFile);

        console.log(`Created: ${outputFile} (${jsonData.m_TableData.length} entries)`);
        return true;
    } catch (error: any) {
        console.error(`Error processing ${fileName}:`, error.message);
        return false;
    }
}

// Main function
function main() {
    try {
        if (!fs.existsSync(inputFolder)) {
            console.error(`Error: Input folder "${inputFolder}" does not exist`);
            process.exit(1);
        }

        if (!fs.existsSync(outputFolder)) {
            fs.mkdirSync(outputFolder, { recursive: true });
            console.log(`Created output folder: ${outputFolder}`);
        }

        // Get all JSON files from the folder
        const files = fs
            .readdirSync(inputFolder)
            .filter((file) => file.toLowerCase().endsWith(".json"))
            .map((file) => path.join(inputFolder, file));

        if (files.length === 0) {
            console.error(`Error: No JSON files found in "${inputFolder}"`);
            process.exit(1);
        }

        console.log(`Found ${files.length} JSON files`);

        let successCount = 0;
        const totalEntries = 0;

        // Process 'em
        files.forEach((file) => {
            if (processAndCreateExcel(file)) {
                successCount++;
            }
        });

        console.log(`\nSuccess! Created ${successCount} Excel files in: ${outputFolder}`);
        console.log(`Processed ${successCount}/${files.length} JSON files`);
    } catch (error) {
        console.error("Error:", error.message);
        process.exit(1);
    }
}

main();

