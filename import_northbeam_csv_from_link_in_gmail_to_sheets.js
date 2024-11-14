// created by <3 with chatgp

// Replace 'YOUR_SHEET_ID' with your actual Google Sheets ID.
const SHEET_ID = '';
// Sheet name if you want exports to override previous one
const SHEET_NAME = 'Sheet1';
// Specify if you want to override data in one sheet or create new sheet for every export
const OVERRIDE_SHEET = true;
// Define the subject line filter to detect emails
const EMAIL_SUBJECT_FILTER = 'Your daily data export is ready';
// Define the specific sender's email to filter for
const SENDER_EMAIL = 'support@northbeam.io';

// entrypoint
function doGet(e) {
  try {
    // Call the CSV import function
    importCSVFromEmail();
    return ContentService.createTextOutput('CSV Import Triggered Successfully');
  } catch (error) {
    // Handle any errors and return the message
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

// Function to find emails with specific subject and specific sender, and process the CSV link
function importCSVFromEmail() {
  // Search for unread emails with the specified subject and sender
  const threads = GmailApp.search(`subject:/.*${EMAIL_SUBJECT_FILTER}.*/ from:${SENDER_EMAIL} is:unread`);

  // Process each email thread
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // Mark the email as read after processing
      message.markRead();

      // Extract the message body and search for the CSV link
      const body = message.getBody();
      const csvLink = extractCSVLink(body);
      
      if (csvLink) {
        // Fetch and parse the CSV data, then import into the Google Sheets
        const csvData = fetchCSVData(csvLink);
        if (csvData) {
          insertDataIntoSheet(csvData);
        }
      }
    });
  });
}

// Helper function to extract the CSV link from email content
function extractCSVLink(body) {
  const urlPattern = /(https:\/\/storage\.googleapis\.com\/[^\s]+\.csv(\?[^<\s]*)?)/g;
  const matches = body.match(urlPattern);
  
  if (matches) {
    // Decode any `&amp;` to `&` for compatibility
    return matches[0].replace(/&amp;/g, '&');
  }
  return null;
}

// Fetch and parse CSV data from the link
function fetchCSVData(url) {
  const fullUrl = prepareFullUrl(url);  

  try {
    const response = UrlFetchApp.fetch(fullUrl, { muteHttpExceptions: true });
    const csvContent = response.getContentText();
    return Utilities.parseCsv(csvContent);
  } catch (error) {
    console.error('Error fetching or parsing CSV:', error.toString());
    return null;
  }
}

// Insert parsed CSV data into the Google Sheet
function insertDataIntoSheet(data) {
  if (OVERRIDE_SHEET) {
    console.log('override sheet');
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  
    // Clear existing content in the sheet
    sheet.clear();

    // Insert data into the sheet
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  } else {
    console.log('new sheet');
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);

    // Get the current date in `YYYY-MM-DD` format
    const date = new Date();
    const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd H:MM:ss");

    // Base sheet name with date
    let sheetName = `${formattedDate}`;
    let sheet = spreadsheet.getSheetByName(sheetName);

    // If a sheet with this name already exists, add a unique suffix
    let counter = 1;
    while (sheet) {
      sheetName = `${formattedDate}_${counter}`;
      sheet = spreadsheet.getSheetByName(sheetName);
      counter++;
    }

    // Create the new sheet with the unique name
    sheet = spreadsheet.insertSheet(sheetName);

    // Insert data into the new sheet
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

// Call createTimeTrigger manually once to set up the automated trigger

function prepareFullUrl(url) {
  const parts = url.split('?');
  url = parts[0];
  const query = parts[1].split('&');
  const params = {};
  query.forEach(param => { params[param.split('=')[0]] = param.split('=')[1] });

  var queryString = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + (params[key]);
  }).join('&');

  // Append the query string to the URL
  var fullUrl = (url + '?' + queryString);
  console.log(fullUrl);
  // for some reason there is `"` character at the end of last parameter, so just remove last char and should be good
  return fullUrl.substring(0, fullUrl.length-1);
}
