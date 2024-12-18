const fs = require("fs");
const path = require("path");
const { XMLParser } = require("fast-xml-parser"); // Import XMLParser

// Initialize the parser
const parser = new XMLParser({ ignoreAttributes: false });
const { create } = require("xmlbuilder2");
const xlsx = require("xlsx");

/**
 * Utility function to convert "Requirement Name" into a valid string for NAMEOFFLOW
 */
function generateNameOfFlow(requirementName) {
    return requirementName
        .replace(/[&/,.:()\-]/g, "") // Remove symbols
        .split(/\s+/) // Split into words
        .filter(word => !["and", "as","or", "in", "to", "of"].includes(word.toLowerCase())) // Ignore small first-letter words
        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()) // Capitalize each word
        .join(""); // Combine into one string
}

function convertRequirementNameToFile(requirementName) {
    return requirementName
        .replace(/[^a-zA-Z0-9\s]/g, "")
        .replace(/\s+/g, "_") 
        .replace(/([a-z])([A-Z])/g, "$1_$2") // Add underscores between lowercase and uppercase letters
        .replace(/(\d+)/g, "_($1)") 
        .toLowerCase() // Convert the whole string to lowercase
        .replace(/_+/g, "_") // Clean up multiple underscores
        .concat(".xml");
}

/**
 * Function to find the exact XML file in the directory
 */
function findExactFile(templateDirectory, nameOfFlow) {
    const fileNameToMatch = convertRequirementNameToFile(nameOfFlow);
    console.log("Looking for file:", fileNameToMatch);

    const files = fs.readdirSync(templateDirectory); // List all files in the directory

    for (const file of files) {
        if (file.toLowerCase() === fileNameToMatch) {
            return path.join(templateDirectory, file); // Return the matched file's full path
        }
    }

    throw new Error(`File matching '${fileNameToMatch}' not found in ${templateDirectory}`);
}

/**
 * Main function to process the Excel file and generate XML files
 */
async function processRequirements(excelFilePath, templateDirectory, outputDirectory, outputListFilePath) {
    // Load Excel file
    const workbook = xlsx.readFile(excelFilePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);
    // console.log("data - ", workbook, sheet, data);
    if (!fs.existsSync(outputDirectory)) {
        fs.mkdirSync(outputDirectory, { recursive: true });
    }

    const generatedFiles = [];

    for (const row of data) {
        const requirementName = row["Requirements Name"];
        
        if (!requirementName) continue;

        const nameOfFlow = generateNameOfFlow(requirementName);
        const exactFilePath = findExactFile(templateDirectory, requirementName);
        if (!fs.existsSync(exactFilePath)) {
            console.error(`Template file not found: ${exactFilePath}`);
            continue;
        }

        const templateContent = fs.readFileSync(exactFilePath, "utf-8");
        const templateXml = parser.parse(templateContent, { ignoreAttributes: false });

        console.log("template xml ", templateXml);
        // Extract <template name="string">
        const templateName = templateXml.template?.['@_name'];
        if (!templateName) {
            console.error(`Template name not found in: ${exactFilePath}`);
            continue;
        }

        // Generate XML content
        const newXmlContent = `<workflow name="ChangeTONeed.Keys.${nameOfFlow}"
          start="main">
    <sequence id="main">
        <activity id="Workflow.${nameOfFlow}"
                  type="Comp.Logic.Workflow.Tasks.RunTemplate">
            {
                TemplateName: "${templateName}",
                DueDate: "0d",
                DueAfter: "@@End",
                AssignTo: "@@Self"
            }
        </activity>
    </sequence>
</workflow>`;

        // Save XML file
        const outputFileName = `ChangeToNeed.Keys.${nameOfFlow}.xml`;
        const outputFilePath = path.join(outputDirectory, outputFileName);
        fs.writeFileSync(outputFilePath, newXmlContent, "utf-8");

        generatedFiles.push(outputFileName);
    }

    // Save list of generated files
    const csvContent = generatedFiles.map(fileName => fileName).join("\n");
    fs.writeFileSync(outputListFilePath, csvContent, "utf-8");

    console.log("Processing completed. Generated files:", generatedFiles);
}

// Example usage
const excelFilePath = "./worksheet.xlsx"; // Path to the Excel file
const templateDirectory = "./templates"; // Directory containing XML templates
const outputDirectory = "./output"; // Directory to save generated XML files
const outputListFilePath = "./output/generated_files.csv"; // File to save the list of generated XML files

processRequirements(excelFilePath, templateDirectory, outputDirectory, outputListFilePath).catch(console.error);
