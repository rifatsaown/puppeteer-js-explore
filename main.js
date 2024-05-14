import dotenv from 'dotenv';
dotenv.config();
import puppeteer from 'puppeteer';
import ExcelJS from 'exceljs';

(async () => {
  const browser = await puppeteer.launch(
    {
      headless: false,
      args: ['--start-maximized']
    }
  );
  const page = await browser.newPage();
  await page.goto(process.env.URL);
  await page.type('#Username', process.env.USERNAME);
  await page.type('#Password', process.env.PASSWORD);
  await page.click('#chkbxAgree');
  await page.click('#btnAgreeLogin');

  const requestPageSelector = "#ctl00_Menu1_linkEligibilityRequest";
  await page.waitForSelector(requestPageSelector);
  await page.click(requestPageSelector);

  const clientIDSelector = "#ctl00_ContentPlaceHolder1_textBoxClientID";
  await page.waitForSelector(clientIDSelector);
  await page.type(clientIDSelector, 'RG75840H');
  await page.click('#ctl00_ContentPlaceHolder1_buttonSubmit');

  await page.waitForSelector('#ctl00_ContentPlaceHolder1_lblSummary');
  await page.click('#ctl00_Menu1_linkEligibilityResponse');
  await page.waitForSelector('#ctl00_ContentPlaceHolder1_RadGrid1_ctl00__0');
  await page.click('#ctl00_ContentPlaceHolder1_RadGrid1_ctl00__0 a');

  await page.waitForSelector('td.pageTitle');
  await page.waitForFunction(() => {
    const pageTitle = document.querySelector('td.pageTitle');
    return pageTitle && pageTitle.innerText.includes("Eligibility Response Details");
  });

  // Scrape structured data from the page
  const data = await page.evaluate(() => {
    const rows = Array.from(document.querySelectorAll('#mainBody table tr'));
    return rows.map(row => {
      const cells = Array.from(row.querySelectorAll('td'));
      // cells that contain "clntColmnLbl" class are the labels
        // const label = cells.find(cell => cell.classList.contains('labelOnMed', 'clntColmnLbl'));
        const value = cells.find(cell => cell.classList.contains('textOnMed', 'clntColmnData'));
        return [value?.innerText ? value.innerText.trim() : ''];
    });
  });

  console.log(data);

  // Create a new workbook and worksheet
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');

  // make the header bold
    worksheet.getRow(1).font = { bold: true };
    // in the first row, write the column names
    worksheet.columns = [{ header: 'Data', key: 'data' }];
    worksheet.columns = data;
    // write the data
    worksheet.addRows(data);

  // Save the workbook to a file
  await workbook.xlsx.writeFile('output.xlsx');

  await browser.close();
})();
