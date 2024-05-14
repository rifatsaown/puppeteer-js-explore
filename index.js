import dotenv from 'dotenv';
dotenv.config();
import puppeteer from 'puppeteer';
import fs from 'fs';

(async () => {
  // Launch the browser with signed-in user
  const browser = await puppeteer.launch({
    headless: false,
    args: ['--start-maximized']
  });
  const page = await browser.newPage();
  await page.goto(process.env.URL);
  await page.type('#Username', process.env.USERNAME);
  await page.type('#Password', process.env.PASSWORD);
  // Click in a checkbox
  await page.click('#chkbxAgree');
  await page.click('#btnAgreeLogin');

  /* ------------------------------------------------------------ */
  // Define Selector
  const requestPageSelector = "#ctl00_Menu1_linkEligibilityRequest";
  // Wait for the navigation link to appear
  await page.waitForSelector(requestPageSelector);
  await page.click(requestPageSelector);

  const clientIDSelector = "#ctl00_ContentPlaceHolder1_textBoxClientID";
  await page.waitForSelector(clientIDSelector);

  // Fill the form
  await page.type(clientIDSelector, 'RG75840H');
  // Click on the submit button
  await page.click('#ctl00_ContentPlaceHolder1_buttonSubmit');

  // Wait for the response to appear
  await page.waitForSelector('#ctl00_ContentPlaceHolder1_lblSummary');

  // Click on the navigation link
  await page.click('#ctl00_Menu1_linkEligibilityResponse');

  // Wait for the response to appear
  await page.waitForSelector('#ctl00_ContentPlaceHolder1_RadGrid1_ctl00__0');

  // Click on the response
  await page.click('#ctl00_ContentPlaceHolder1_RadGrid1_ctl00__0 a');

  // Wait until the text "Eligibility Response Details" appears on the page
  await page.waitForSelector('td.pageTitle');
  await page.waitForFunction(() => {
    const pageTitle = document.querySelector('td.pageTitle');
    return pageTitle && pageTitle.innerText.includes("Eligibility Response Details");
  });

  // Get the text of the page
  const text = await page.evaluate(() => document.querySelector('#mainBody').innerText);
  // Save the text to a xlsx file and without rewriting the file, just append the text
  fs.appendFileSync('output.xlsx', text);
  fs.appendFileSync('output.xlsx', '\n \n----------------------------------------------------\n \n');

  /* ------------------------------------------------------------ */

})();