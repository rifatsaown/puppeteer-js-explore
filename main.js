import dotenv from 'dotenv';
import puppeteer from 'puppeteer';
import ExcelJS from 'exceljs';
dotenv.config();

(async () => {
  // Launch the browser in headless mode and set the window size to 1920x1080
  const browser = await puppeteer.launch(
    {
      headless: false,
      args: ['--start-maximized']
    }
  );
  const page = await browser.newPage();
  await page.setViewport({ width: 1920, height: 1080 });

  // Go to the site and login
  await page.goto(process.env.SITE_URL);
  await page.type('#Username', process.env.SITE_USERNAME);
  await page.type('#Password', process.env.SITE_PASSWORD);
  await page.click('#chkbxAgree');
  await page.click('#btnAgreeLogin');

  // Navigate to the request page
  const requestPageSelector = "#ctl00_Menu1_linkEligibilityRequest";
  await page.waitForSelector(requestPageSelector);
  // await page.click(requestPageSelector);

  // // Fill out the form and submit
  // const clientIDSelector = "#ctl00_ContentPlaceHolder1_textBoxClientID";
  // await page.waitForSelector(clientIDSelector);
  // await page.type(clientIDSelector, 'RG75840H');
  // await page.click('#ctl00_ContentPlaceHolder1_buttonSubmit');
  // await page.waitForSelector('#ctl00_ContentPlaceHolder1_lblSummary');

  // Navigate to the response page
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
    // Get the client All information from the page
    const clientID = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientID')?.innerText;
    const clientName = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientName')?.innerText;
    const clientGender = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientGender')?.innerText;
    const clientSSN = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientSSN')?.innerText;
    const clientDOB = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientDOB')?.innerText;
    const clientAnniversaryDate = document.querySelector('#ctl00_ContentPlaceHolder1_labelAnniversary')?.innerText;
    const clientRecertification = document.querySelector('#ctl00_ContentPlaceHolder1_labelRecertification')?.innerText;
    const clientAddress1 = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientAddress1')?.innerText;
    const clientAddress2 = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientAddress2')?.innerText;
    const clientCityStateZip = document.querySelector('#ctl00_ContentPlaceHolder1_labelClientCityStateZip')?.innerText;
    const clientCounty = document.querySelector('#ctl00_ContentPlaceHolder1_labelCounty')?.innerText;
    const clientOffice = document.querySelector('#ctl00_ContentPlaceHolder1_labelOffice')?.innerText;
    const clientDateOfService = document.querySelector('#ctl00_ContentPlaceHolder1_labelDateOfService')?.innerText;
    const clientPlanDate = document.querySelector('#ctl00_ContentPlaceHolder1_labelPlanDate')?.innerText;


    const data = [{
      clientID,
      clientName,
      clientGender,
      clientSSN,
      clientDOB,
      clientAnniversaryDate,
      clientRecertification,
      clientAddress1,
      clientAddress2,
      clientCityStateZip,
      clientCounty,
      clientOffice,
      clientDateOfService,
      clientPlanDate,
      
    }];

    return data;
  });

  console.log(data);

  // Create a new workbook and worksheet
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');

  // make the header bold
    worksheet.getRow(1).font = { bold: true };
    // in the next columns, write the Column Names
    const header = ['Client ID', 'Client Name', 'Gender', 'SSN', 'Date of Birth', 'Anniversary Date', 'Recertification', 'Address 1', 'Address 2', 'City, State Zip', 'County', 'Office', 'Date of Service', 'Plan Date' 
    ];
    worksheet.columns = header.map((key, i) => {
      return { header: key, key: key, width: 20 };
    });

  // Add rows to the worksheet
  data.forEach((row, index) => {
    const worksheetRow = worksheet.getRow(index + 2); // index + 2 because the first row is the header
    worksheetRow.getCell('Client ID').value = row.clientID;
    worksheetRow.getCell('Client Name').value = row.clientName;
    worksheetRow.getCell('Gender').value = row.clientGender;
    worksheetRow.getCell('SSN').value = row.clientSSN;
    worksheetRow.getCell('Date of Birth').value = row.clientDOB;
    worksheetRow.getCell('Anniversary Date').value = row.clientAnniversaryDate;
    worksheetRow.getCell('Recertification').value = row.clientRecertification;
    worksheetRow.getCell('Address 1').value = row.clientAddress1;
    worksheetRow.getCell('Address 2').value = row.clientAddress2;
    worksheetRow.getCell('City, State Zip').value = row.clientCityStateZip;
    worksheetRow.getCell('County').value = row.clientCounty;
    worksheetRow.getCell('Office').value = row.clientOffice;
    worksheetRow.getCell('Date of Service').value = row.clientDateOfService;
    worksheetRow.getCell('Plan Date').value = row.clientPlanDate;

  });

  // Save the workbook to a file
  await workbook.xlsx.writeFile('data.xlsx');

  await browser.close();
})();
