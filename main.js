const config = require("./config");
const { readInputFile, saveToExcel } = require("./excelUtils");
const {
  setupBrowser,
  searchCompany,
  getFirstResult,
} = require("./googleSearch");

async function main() {
  let companies;
  try {
    companies = readInputFile(config.INPUT_FILE);
  } catch (error) {
    console.error(
      "Failed to read input file. Please check the file and column names."
    );
    console.error("Error details:", error.message);
    return;
  }

  if (!companies || companies.length === 0) {
    console.error("No valid company data found in the input file.");
    return;
  }

  console.log(
    `Successfully read ${companies.length} companies from the input file.`
  );

  const results = [];
  let browser, page;

  try {
    ({ browser, page } = await setupBrowser());

    for (const company of companies) {
      try {
        console.log(`Processing: ${company.companyName}`);
        await searchCompany(page, company.companyName);
        console.log(`Search completed for: ${company.companyName}`);

        const companyUrl = await getFirstResult(page);
        console.log(
          `First result URL for ${company.companyName}: ${companyUrl}`
        );

        results.push({
          "Company Name": company.companyName,
          "Booth Number": company.boothNumber,
          Website: companyUrl || "Not found",
        });
      } catch (error) {
        console.error(
          `Error processing ${company.companyName}: ${error.message}`
        );
        results.push({
          "Company Name": company.companyName,
          "Booth Number": company.boothNumber,
          Website: "Error",
        });
      }

      await new Promise((resolve) => setTimeout(resolve, 3000));
    }
  } catch (error) {
    console.error("An unexpected error occurred:", error.message);
  } finally {
    if (browser) await browser.close();
  }

  if (results.length > 0) {
    saveToExcel(results, config.OUTPUT_FILE);
  } else {
    console.error("No results to save. Please check for errors above.");
  }
}

main().catch(console.error);
