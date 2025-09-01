import fs from "fs";
import path from "path";
import XLSX from "xlsx";

// Get cmd args
const args = process.argv.slice(2);

// Default values for the folders
let inputFolder = "./Raws";
let outputFolder = "./Extracted";

function showHelp() {
    console.log(`
        Usage: {node | bun run} extract.ts [-i input-folder] [-o output-folder]

        Options:
        -i, --input <folder>    Input folder (default: ./Raws)
        -o, --output <folder>   Output folder (default: ./Extracted)
        -h, --help              Show this help
    `);
}

// I hate that this is the solution I came out with
if (args.length !== 0) {
    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        if (arg === "-i" || arg === "--input") {
            inputFolder = args[i + 1] || inputFolder;
            i++;
        } else if (arg === "-o" || arg === "--output") {
            outputFolder = args[i + 1] || outputFolder;
            i++;
        } else if (arg === "-h" || arg === "--help") {
            // Why would you use -h if you already set input/output?
            // Whatever, just show it. Don't process at all if you use the flag.
            showHelp();
            process.exit(0);
        }
    }
} else {
    showHelp();
    process.exit(0);
}

console.log(`Reading JSON files from: ${inputFolder}`);
console.log(`Output Excel files to: ${outputFolder}`);

// Function to escape newlines
function escapeNewlines(text) {
    if (typeof text !== "string") return text;
    return text.replace(/\n/g, "\\n");
}

// Wrap large numeric m_Id literals in quotes BEFORE JSON.parse to avoid screwing things up when making the Excel files
function parseWithStringIds(fileContent) {
    const safe = fileContent.replace(/"m_Id"\s*:\s*(\d+)/g, '"m_Id":"$1"');
    return JSON.parse(safe);
}

// Function to process and create separate Excel file for each JSON
function processAndCreateExcel(filePath) {
    try {
        const fileContent = fs.readFileSync(filePath, "utf8");
        const jsonData = parseWithStringIds(fileContent);

        if (!jsonData.m_TableData || !Array.isArray(jsonData.m_TableData)) {
            console.warn(`Warning: ${filePath} doesn't have valid m_TableData array`);
            return false;
        }

        const fileName = path.basename(filePath, ".json");
        console.log(`Processing ${fileName}: Found ${jsonData.m_TableData.length} entries`);

        const data = jsonData.m_TableData.map((item) => ({
            ID: item.m_Id || "", // already a precise string
            original_line: escapeNewlines(item.m_Localized || ""),
            translated_line: "",
            notes: "",
        }));

        // Create Excel workbook for this file
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Translation Data");

        // Set ID column to text format to prevent scientific notation
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 0 }) // Column A (ID)
            if (worksheet[cellAddress]) {
                worksheet[cellAddress].t = 's' // Force cell type to string
            }
        };

        // Set column widths for better readability
        worksheet["!cols"] = [
            { width: 20 }, // ID - increased width for long IDs
            { width: 100 }, // original_line
            { width: 100 }, // translated_line
            { width: 30 }, // notes
        ];

        // Create output filename
        const outputFile = path.join(outputFolder, `${fileName}.xlsx`);

        // Write Excel file
        XLSX.writeFile(workbook, outputFile);
        console.log(`✅ Created: ${outputFile} (${data.length} entries)`);

        return true;
    } catch (error) {
        console.error(`Error processing ${filePath}:`, error.message);
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

        files.forEach((file) => {
            if (processAndCreateExcel(file)) {
                successCount++;
            }
        });

        console.log(`\n✅ Success! Created ${successCount} Excel files in: ${outputFolder}`);
        console.log(`📁 Processed ${successCount}/${files.length} JSON files`);
    } catch (error) {
        console.error("Error:", error.message);
        process.exit(1);
    }
}

main();
