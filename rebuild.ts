import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";

// Get cmd args
const args = process.argv.slice(2);

// Default values for the folders
let inputFolder = "./Extracted";
let outputFolder = "./Rebuilt";

function showHelp() {
    console.log(`
Usage: {node | bun run} rebuild.ts [-i input-folder] [-o output-folder]

Options:
    -i, --input <folder>    Input folder (default: "./Extracted")
    -o, --output <folder>   Output folder (default: "./Rebuilt")
    -h, --help              Show this help

Example:
    bun run rebuild.ts -i "./XLSX files" -o "./Processed files"
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

console.log(`Reading Excel files from: ${inputFolder}`);
console.log(`Output JSON files to: ${outputFolder}`);

async function rebuildJsonFromExcel(filePath: string) {
    try {
        const fileName = path.basename(filePath, ".xlsx");
        console.log(`Processing ${fileName}...`);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        const worksheet = workbook.worksheets[0]; // first sheet
        if (!worksheet) {
            console.warn(`Warning: ${fileName} has no sheets`);
            return false;
        }

        const data: any[] = [];

        // Convert rows to JSON-like object
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber === 1) return; // skip header
            const rowValues = row.values as any[];
            data.push({
                ID: rowValues[1]?.toString().trim() || "",
                translated_line: rowValues[3]?.toString().trim() || "",
            });
        });

        if (data.length === 0) {
            console.warn(`Warning: ${fileName} has no data`);
            return false;
        }

        // Build translations object
        const translations: Record<string, string> = {};
        let translatedCount = 0;

        data.forEach((row) => {
            const id = row.ID;
            const translatedLine = row.translated_line;
            if (id && translatedLine) {
                translations[id] = translatedLine.replace(/\\n/g, "\n");
                translatedCount++;
            }
        });

        if (translatedCount === 0) {
            console.warn(`Warning: ${fileName} has no translated entries`);
            return false;
        }

        let outputFileName = fileName;
        if (fileName.includes("_en")) {
            outputFileName = fileName.replace("_en", "_es");
        } else {
            outputFileName = fileName + "_es";
        }

        const outputFile = path.join(outputFolder, `${outputFileName}.json`);
        fs.writeFileSync(outputFile, JSON.stringify(translations, null, 2), "utf8");
        console.log(`Created: ${outputFile} (${translatedCount} translations)`);

        return true;
    } catch (error: any) {
        console.error(`Error processing ${filePath}:`, error.message);
        return false;
    }
}

// We use process.exit(1) to exit the app btw
function main() {
    try {
        // Check if input folder exists
        if (!fs.existsSync(inputFolder!)) {
            console.error(`Error: Input folder "${inputFolder}" does not exist`);
            process.exit(1);
        }

        if (!fs.existsSync(outputFolder!)) {
            fs.mkdirSync(outputFolder!, { recursive: true });
            console.log(`Created output folder: ${outputFolder}`);
        }

        // Get all Excel files from the folder
        const files = fs
            .readdirSync(inputFolder!)
            .filter((file) => file.toLowerCase().endsWith(".xlsx"))
            .map((file) => path.join(inputFolder!, file));

        if (files.length === 0) {
            console.error(`Error: No Excel files found in "${inputFolder}"`);
            process.exit(1);
        }

        console.log(`\nFound ${files.length} Excel files. \n`);

        let successCount = 0;

        files.forEach((file) => {
            if (rebuildJsonFromExcel(file)) {
                successCount++;
            }
        });

        console.log(`\nProcessed ${successCount}/${files.length} Excel files. \n`);
    } catch (error) {
        console.error("Error:", error.message);
        process.exit(1);
    }
}

main();
