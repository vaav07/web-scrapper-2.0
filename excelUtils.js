const xlsx = require("xlsx");

function readInputFile(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    if (data.length === 0) {
      throw new Error("No data found in the Excel file");
    }

    console.log("Columns found in the Excel file:");
    console.log(Object.keys(data[0]));

    // Try to find column names (case-insensitive)
    const companyNameKey = Object.keys(data[0]).find(
      (key) =>
        key.toLowerCase().includes("company") ||
        key.toLowerCase().includes("name") ||
        key.toLowerCase() === "exhibitor"
    );
    const boothNumberKey = Object.keys(data[0]).find(
      (key) =>
        key.toLowerCase().includes("booth") ||
        key.toLowerCase().includes("stand") ||
        key.toLowerCase().includes("number")
    );

    if (!companyNameKey) {
      throw new Error(
        'Could not find a column for company names. Please ensure there is a column with "Company", "Name", or "Exhibitor" in its header.'
      );
    }
    if (!boothNumberKey) {
      throw new Error(
        'Could not find a column for booth numbers. Please ensure there is a column with "Booth", "Stand", or "Number" in its header.'
      );
    }

    console.log(`Using "${companyNameKey}" as the company name column`);
    console.log(`Using "${boothNumberKey}" as the booth number column`);

    return data
      .map((row) => ({
        companyName: row[companyNameKey],
        boothNumber: row[boothNumberKey],
      }))
      .filter((item) => item.companyName && item.boothNumber);
  } catch (error) {
    console.error(`Error reading input file: ${error.message}`);
    process.exit(1);
  }
}

function saveToExcel(data, filePath) {
  const worksheet = xlsx.utils.json_to_sheet(data);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Results");
  xlsx.writeFile(workbook, filePath);
  console.log(`Results saved to ${filePath}`);
}

module.exports = { readInputFile, saveToExcel };
