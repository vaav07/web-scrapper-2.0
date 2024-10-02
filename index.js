const puppeteer = require("puppeteer");
const xlsx = require("xlsx");
const fs = require("fs");

// Function to read company names from Excel file
function readInputFile(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    if (data.length === 0) {
      throw new Error("No data found in the Excel file");
    }

    const companyNameKey = Object.keys(data[0]).find(
      (key) =>
        key.toLowerCase().includes("company") &&
        key.toLowerCase().includes("name")
    );

    if (!companyNameKey) {
      throw new Error("Could not find a column for company names");
    }

    return data.map((row) => row[companyNameKey]).filter((name) => name);
  } catch (error) {
    console.error(`Error reading input file: ${error.message}`);
    process.exit(1);
  }
}

// Updated searchCompany function
async function searchCompany(page, companyName) {
  await page.goto("https://www.google.com", { waitUntil: "networkidle0" });

  // Reject cookies if the prompt appears
  try {
    await page.waitForSelector('button[id="W0wltc"]', { timeout: 5000 });
    await page.click('button[id="W0wltc"]');
  } catch (error) {
    console.log("No cookie prompt found or unable to reject cookies.");
  }

  try {
    // Wait for the search input to be available
    await page.waitForSelector('textarea[name="q"]', { timeout: 10000 });

    // Type the company name into the search box
    await page.type('textarea[name="q"]', companyName);

    // Wait for a short time to ensure the input is registered
    await new Promise((resolve) => setTimeout(resolve, 300));

    // Press Enter and wait for navigation
    await Promise.all([
      page.keyboard.press("Enter"),
      page.waitForNavigation({ waitUntil: "networkidle0" }),
    ]);
  } catch (error) {
    throw new Error(`Error searching for ${companyName}: ${error.message}`);
  }
}

// Function to get the first result URL
async function getFirstResult(page) {
  try {
    await page.waitForSelector("div.g div.r a, div.yuRUbf a", {
      timeout: 5000,
    });
    return await page.$eval("div.g div.r a, div.yuRUbf a", (el) => el.href);
  } catch (error) {
    console.warn(`No results found for ${await page.url()}`);
    return null;
  }
}

// Function to extract contact information from a page
async function extractContactInfo(page, url) {
  await page.goto(url, { waitUntil: "networkidle0" });

  // Attempt to reject cookies on the company website
  try {
    const cookieSelectors = [
      'button[id*="reject"]',
      'button[id*="decline"]',
      'button[class*="reject"]',
      'button[class*="decline"]',
      'a[id*="reject"]',
      'a[id*="decline"]',
      'a[class*="reject"]',
      'a[class*="decline"]',
    ];
    for (const selector of cookieSelectors) {
      const [button] = await page.$x(
        `//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'reject') or contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'decline')]`
      );
      if (button) {
        await button.click();
        break;
      }
    }
  } catch (error) {
    console.log("Unable to find or click cookie rejection button.");
  }

  let contactInfo = {
    "Company Name": "",
    "Contact Person": "",
    Designation: "",
    Country: "",
    Email: "",
    Website: url,
    Telephone: "",
    Mobile: "",
    Address: "",
  };

  contactInfo["Email"] = await findEmail(page);
  contactInfo["Telephone"] = await findPhoneNumber(page);
  contactInfo["Address"] = await findAddress(page);
  // contactInfo["Country"] = await findCountry(page);

  if (
    !contactInfo["Email"] ||
    !contactInfo["Telephone"] ||
    !contactInfo["Address"]
  ) {
    const contactPage = await findContactPage(page);
    if (contactPage) {
      await page.goto(contactPage, { waitUntil: "networkidle0" });
      if (!contactInfo["Email"]) contactInfo["Email"] = await findEmail(page);
      if (!contactInfo["Telephone"])
        contactInfo["Telephone"] = await findPhoneNumber(page);
      if (!contactInfo["Address"])
        contactInfo["Address"] = await findAddress(page);
      // if (!contactInfo["Country"])
      //   contactInfo["Country"] = await findCountry(page);
    }
  }

  return contactInfo;
}

// Helper function to find phone number
async function findPhoneNumber(page) {
  return await page.evaluate(() => {
    const phoneRegex =
      /(\+?1?[ -]?)?(\(?[0-9]{3}\)?|[0-9]{3})[ -]?[0-9]{3}[ -]?[0-9]{4}/g;
    const text = document.body.innerText;
    const matches = text.match(phoneRegex);
    return matches ? matches.join(", ") : null;
  });
}

// Helper function to find email
async function findEmail(page) {
  return await page.evaluate(() => {
    const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/;
    const element = document.querySelector("body");
    const match = element.innerText.match(emailRegex);
    return match ? match[0] : null;
  });
}

// Helper function to find address
async function findAddress(page) {
  return await page.evaluate(() => {
    const addressKeywords = [
      "address",
      "location",
      "street",
      "avenue",
      "boulevard",
    ];
    const paragraphs = document.querySelectorAll("p");
    for (let p of paragraphs) {
      if (
        addressKeywords.some((keyword) =>
          p.innerText.toLowerCase().includes(keyword)
        )
      ) {
        return p.innerText.trim();
      }
    }
    return null;
  });
}

// Helper function to find contact page
async function findContactPage(page) {
  return await page.evaluate(() => {
    const links = Array.from(document.querySelectorAll("a"));
    const contactLink = links.find(
      (link) =>
        link.innerText.toLowerCase().includes("contact") ||
        link.innerText.toLowerCase().includes("about")
    );
    return contactLink ? contactLink.href : null;
  });
}

// Updated main function with more logging
async function main() {
  const inputFile = "company_list.xlsx";
  const outputFile = "company_contact_info.xlsx";

  const companies = readInputFile(inputFile);
  const results = [];

  const browser = await puppeteer.launch({
    headless: false,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const page = await browser.newPage();

  // Set a longer default navigation timeout
  page.setDefaultNavigationTimeout(60000);

  for (const company of companies) {
    try {
      console.log(`Processing: ${company}`);
      await searchCompany(page, company);
      console.log(`Search completed for: ${company}`);

      const companyUrl = await getFirstResult(page);
      console.log(`First result URL for ${company}: ${companyUrl}`);

      if (companyUrl) {
        const contactInfo = await extractContactInfo(page, companyUrl);
        contactInfo["Company Name"] = company;
        results.push(contactInfo);
        console.log(`Processed ${company} successfully`);
      } else {
        console.warn(`Could not find a website for ${company}`);
        results.push({ "Company Name": company, Error: "No website found" });
      }
    } catch (error) {
      console.error(`Error processing ${company}: ${error.message}`);
      results.push({ "Company Name": company, Error: error.message });
    }

    // Add a delay between requests to avoid overwhelming the servers
    await new Promise((resolve) => setTimeout(resolve, 6000));
  }

  await browser.close();

  // Save results to Excel
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(results);
  xlsx.utils.book_append_sheet(wb, ws, "Results");
  xlsx.writeFile(wb, outputFile);

  console.log(`Results saved to ${outputFile}`);
}

main().catch(console.error);
