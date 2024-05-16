import dotenv from 'dotenv';
import ExcelJS from 'exceljs';
import puppeteer from 'puppeteer';

dotenv.config();

// Function to grab data for multiple client IDs
const grabData = async (clientIDs) => {
    //start a timer
    console.time('Total processing time');

    // Launch a new browser instance
    const browser = await puppeteer.launch({
        // headless: false, // Set headless to false for debugging
        // args: ['--start-maximized'], // Maximize the browser window
        // // // slowMo: 50, // Slow down Puppeteer operations by 50ms
    });

    // Create a new page in the browser
    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080 }); // Set viewport size

    // Define retry settings
    const maxRetries = 1;
    const retryDelay = 3000; // 3 seconds

    try {
        // Navigate to the login page
        await page.goto(process.env.SITE_URL);
        await page.type('#Username', process.env.SITE_USERNAME); // Type username
        await page.type('#Password', process.env.SITE_PASSWORD); // Type password
        await page.click('#chkbxAgree'); // Click on agree checkbox
        await page.click('#btnAgreeLogin'); // Click on login button

        // Create a new Excel workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Data');

        // Iterate over each client ID
        for (let index = 0; index < clientIDs.length; index++) {
            const id = clientIDs[index];
            console.log(`Processing client ID: ${id}`);
            // Start measuring time
            console.time(`Processing time for client ID: ${id}`);


            let retryCount = 0;
            let success = false;

            while (!success && retryCount < maxRetries) {
                try {
                    // Navigate to the request page
                    const requestPageSelector = "#ctl00_Menu1_linkEligibilityRequest";
                    await page.waitForSelector(requestPageSelector);
                    await page.click(requestPageSelector);

                    // Fill out the form and submit
                    const clientIDSelector = "#ctl00_ContentPlaceHolder1_textBoxClientID";
                    await page.waitForSelector(clientIDSelector);
                    await page.type(clientIDSelector, id);
                    await page.waitForSelector('#ctl00_ContentPlaceHolder1_buttonSubmit');
                    await page.click('#ctl00_ContentPlaceHolder1_buttonSubmit');
                    await page.waitForSelector('#ctl00_ContentPlaceHolder1_lblSummary');

                    // Navigate to the response page
                    await page.click('#ctl00_Menu1_linkEligibilityResponse');
                    await page.waitForSelector('#ctl00_ContentPlaceHolder1_RadGrid1_ctl00__0');
                    await page.click('#ctl00_ContentPlaceHolder1_RadGrid1_ctl00__0 a');

                    // Wait for response page to load
                    await page.waitForSelector('td.pageTitle');
                    await page.waitForFunction(() => {
                        const pageTitle = document.querySelector('td.pageTitle');
                        return pageTitle && pageTitle.innerText.includes("Eligibility Response Details");
                    });

                    // Scrape structured data from the page
                    const clientData = await page.evaluate(() => {
                        // Get the client information from the page
                        return {
                            clientID: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientID')?.innerText,
                            clientName: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientName')?.innerText,
                            clientGender: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientGender')?.innerText,
                            clientSSN: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientSSN')?.innerText,
                            clientDOB: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientDOB')?.innerText,
                            clientAnniversaryDate: document.querySelector('#ctl00_ContentPlaceHolder1_labelAnniversary')?.innerText,
                            clientRecertification: document.querySelector('#ctl00_ContentPlaceHolder1_labelRecertification')?.innerText,
                            clientAddress1: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientAddress1')?.innerText,
                            clientAddress2: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientAddress2')?.innerText,
                            clientCityStateZip: document.querySelector('#ctl00_ContentPlaceHolder1_labelClientCityStateZip')?.innerText,
                            clientCounty: document.querySelector('#ctl00_ContentPlaceHolder1_labelCounty')?.innerText,
                            clientOffice: document.querySelector('#ctl00_ContentPlaceHolder1_labelOffice')?.innerText,
                            clientDateOfService: document.querySelector('#ctl00_ContentPlaceHolder1_labelDateOfService')?.innerText,
                            clientPlanDate: document.querySelector('#ctl00_ContentPlaceHolder1_labelPlanDate')?.innerText
                        };
                    });

                    console.log(clientData);

                    // Add headers to the worksheet
                    if (index === 0) {
                        worksheet.getRow(1).font = { bold: true };
                        const header = ['Client ID', 'Client Name', 'Gender', 'SSN', 'Date of Birth', 'Anniversary Date', 'Recertification', 'Address 1', 'Address 2', 'City, State Zip', 'County', 'Office', 'Date of Service', 'Plan Date'];
                        worksheet.columns = header.map((key) => {
                            return { header: key, key, width: 20 };
                        });
                    }

                    // Add data to the worksheet
                    const worksheetRow = worksheet.getRow(index + 2); // index + 2 because the first row is the header
                    worksheetRow.getCell('Client ID').value = clientData.clientID;
                    worksheetRow.getCell('Client Name').value = clientData.clientName;
                    worksheetRow.getCell('Gender').value = clientData.clientGender;
                    worksheetRow.getCell('SSN').value = clientData.clientSSN;
                    worksheetRow.getCell('Date of Birth').value = clientData.clientDOB;
                    worksheetRow.getCell('Anniversary Date').value = clientData.clientAnniversaryDate;
                    worksheetRow.getCell('Recertification').value = clientData.clientRecertification;
                    worksheetRow.getCell('Address 1').value = clientData.clientAddress1;
                    worksheetRow.getCell('Address 2').value = clientData.clientAddress2;
                    worksheetRow.getCell('City, State Zip').value = clientData.clientCityStateZip;
                    worksheetRow.getCell('County').value = clientData.clientCounty;
                    worksheetRow.getCell('Office').value = clientData.clientOffice;
                    worksheetRow.getCell('Date of Service').value = clientData.clientDateOfService;
                    worksheetRow.getCell('Plan Date').value = clientData.clientPlanDate;

                    success = true;

                    console.log(`Data for client ID: ${id} processed successfully!`);

                    // Stop measuring time
                    console.timeEnd(`Processing time for client ID: ${id}`);
                } catch (error) {
                    console.error(`Attempt ${retryCount + 1} failed:`, error);
                    retryCount++;
                    console.log(`Retrying (${retryCount}/${maxRetries})...`);
                    await new Promise(resolve => setTimeout(resolve, retryDelay)); // Wait before retrying
                }
            }

            if (!success) {
                console.error(`Maximum retry attempts reached for client ID: ${id}. Skipping...`);
            }
        };

        // Save the workbook to a file
        await workbook.xlsx.writeFile('data.xlsx');
        console.log("Data saved successfully!");

        // Stop measuring total time
        console.timeEnd('Total processing time');
        // Close the browser
        await browser.close();
    } catch (error) {
        // Handle errors
        console.error('Error occurred:', error);
        await browser.close();
    }
};

// Example usage
grabData([
    'RG75840H',
    'TM58421K',
    'PX18869R',
    'MV35570G',
    'XK81995H', 'MV35570G', 'PX18869R', 'TM58421K', 'ZF26004Z', 'RV39782E', 'SC14165Q', 'XX50738B', 'ZJ69393M', 'ZE52322D',
    'YR52344J', 'ZT66618E', 'KT19352G', 'QG07774T', 'UK20379D', 'MJ17213H', 'RK54434S', 'XW37131M', 'PR58391S', 'NZ93716M',
    'KN36667W', 'RT23712P', 'XF22996A', 'PF38721H', 'RX43458M', 'MZ66537J', 'UT06895P', 'YC88670E', 'UP71463N', 'PX05909V',
    'WY18234Y', 'VW49879Z', 'WC76039C', 'PG89090E', 'NW23358B', 'NZ66539R', 'YA70018D', 'YQ23082D', 'ZB29259A', 'RJ70817D',
    'SA40463X', 'PV41139W', 'WE85258X', 'SN88456B', 'YY39849R', 'NB44045D', 'ZE06067F', 'KB79956F', 'WB29733W', 'VQ57556H',
    'VD56059H', 'YY84283B', 'PJ79768K', 'US66130X', 'UW93018S', 'QA02131Y', 'PU88211X', 'XK49162X', 'YA57996T', 'QP57102P',
    'VY74468J', 'VW81327U', 'SK64442T', 'ZD66427Z', 'NQ79885R', 'NX88165F', 'XS45110A', 'NX39533V', 'VX34696S', 'NK87964W',
    'KH17971J', 'TB08369J', 'QA44562Z', 'NG16244K', 'TC72213U', 'ND07383N', 'XV38088K', 'QM23878X', 'KF17467H', 'ST41035K',
    'PK81381M', 'NK96954H', 'WC07792G', 'QQ64206U', 'VJ43988H', 'PW39987K', 'YZ96310R', 'KZ69688P', 'ME73764E', 'KJ88756U',
    'TX73419U', 'KP65953B', 'UC37343K', 'ZZ54206A', 'TD10562X', 'ZR35417K', 'YT52141U', 'ZN92123T', 'WZ24862C'
]);
