// https://blog.risingstack.com/pdf-from-html-node-js-puppeteer/ Used starter code from
const puppeteer = require('puppeteer');

async function convertHtmlToPng(htmlFilePath, pngFilePath) {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  // Open the HTML file
  await page.goto(`file://${htmlFilePath}`, { waitUntil: 'networkidle0' });

  // Wait for the chart to render (adjust as needed)
  await page.waitForSelector('canvas');

  // Take a screenshot and save it as PNG
  await page.screenshot({ path: pngFilePath, fullPage: true });

  await browser.close();
}

const [,, htmlFilePath, pngFilePath] = process.argv;
convertHtmlToPng(htmlFilePath, pngFilePath);
