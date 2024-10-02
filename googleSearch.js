const puppeteer = require("puppeteer");
const config = require("./config");

async function setupBrowser() {
  const browser = await puppeteer.launch({
    headless: false,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const page = await browser.newPage();
  page.setDefaultNavigationTimeout(60000);
  return { browser, page };
}

async function rejectCookies(page) {
  try {
    await page.waitForSelector(config.COOKIE_REJECT_SELECTOR, {
      timeout: 5000,
    });
    await page.click(config.COOKIE_REJECT_SELECTOR);
  } catch (error) {
    console.log("No cookie prompt found or unable to reject cookies.");
  }
}

async function searchCompany(page, companyName) {
  await page.goto(config.GOOGLE_SEARCH_URL, { waitUntil: "networkidle0" });
  await rejectCookies(page);

  try {
    await page.waitForSelector(config.SEARCH_SELECTOR, { timeout: 10000 });
    await page.type(config.SEARCH_SELECTOR, companyName);
    await new Promise((resolve) => setTimeout(resolve, 300));
    await Promise.all([
      page.keyboard.press("Enter"),
      page.waitForNavigation({ waitUntil: "networkidle0" }),
    ]);
  } catch (error) {
    throw new Error(`Error searching for ${companyName}: ${error.message}`);
  }
}

async function getFirstResult(page) {
  try {
    await page.waitForSelector(config.RESULT_SELECTOR, { timeout: 5000 });
    return await page.$eval(config.RESULT_SELECTOR, (el) => el.href);
  } catch (error) {
    console.warn(`No results found for ${await page.url()}`);
    return null;
  }
}

module.exports = { setupBrowser, searchCompany, getFirstResult };
