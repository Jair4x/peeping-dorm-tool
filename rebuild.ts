import fs from "fs";
import path from "path";
import XLSX from "xlsx";

// Get cmd args
const args = process.argv.slice(2);

// Default values for the folders
let inputFolder = "./Extracted";
let outputFolder = "./Rebuilt";

function showHelp() {
    console.log(`
        Usage: {node | bun run} rebuild.ts [-i input-folder] [-o output-folder]

        Options:
        -i, --input <folder>    Input folder (default: ./Extracted)
        -o, --output <folder>   Output folder (default: ./Rebuilt)
        -h, --help              Show this help
    `);
}

// I hate that this is the solution I came out with
if (args.length === 0) {
    showHelp();
    process.exit(0);
} else {
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
}

console.log(`Reading Excel files from: ${inputFolder}`);
console.log(`Output JSON files to: ${outputFolder}`);

function rebuildJsonFromExcel(filePath) {
    try {
        const fileName = path.basename(filePath, ".xlsx");
        console.log(`Processing ${fileName}...`);

        // Read Excel file
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const data = XLSX.utils.sheet_to_json(worksheet);

        if (data.length === 0) {
            console.warn(`Warning: ${fileName} has no data`);
            return false;
        }

        // Create the translation object
        const translations = {};
        let translatedCount = 0;

        data.forEach((row) => {
            const id = String(row.ID || "").trim();
            const translatedLine = String(row.translated_line || "").trim();

            if (id && translatedLine) {
                // Unescape newlines back to actual newlines for JSON
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

        // Write JSON file in the format we want
        fs.writeFileSync(outputFile, JSON.stringify(translations, null, 2), "utf8");
        console.log(`✅ Created: ${outputFile} (${translatedCount} translations)`);

        return true;
    } catch (error) {
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

        console.log(`Found ${files.length} Excel files`);

        let successCount = 0;

        files.forEach((file) => {
            if (rebuildJsonFromExcel(file)) {
                successCount++;
            }
        })

        console.log(`\n✅ Success! Created ${successCount} JSON files in: ${outputFolder}`);
        console.log(`📁 Processed ${successCount}/${files.length} Excel files`);
    } catch (error) {
        console.error("Error:", error.message);
        process.exit(1);
    }
}

main();
