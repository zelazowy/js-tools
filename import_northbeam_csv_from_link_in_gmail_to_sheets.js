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
// Define receiver email that is used for finding email with export data, like name.surname+alias@gmail.com
const RECEIVER_EMAIL = '';
// created by <3 with chatgp

// Replace 'YOUR_SHEET_ID' with your actual Google Sheets ID.
const SHEET_ID = '1oDpZf5oxWcV5Z91TRS3hgJGFaG4UmaDF2alWbnmIxvI';
// Sheet name if you want exports to override previous one
const SHEET_NAME = 'Sheet1';
// Specify if you want to override data in one sheet or create new sheet for every export
const OVERRIDE_SHEET = false;
// Define the subject line filter to detect emails
const EMAIL_SUBJECT_FILTER = 'Your daily data export is ready';
// Define the specific sender's email to filter for
const SENDER_EMAIL = 'viktoras@moerie.com';
// Define export name from northbeam
const EXPORT_NAME = 'MO_DE_L7D_Daily';

// Entry point for the web app
function doGet() {
  try {
    importCSVFromEmail();
    return ContentService.createTextOutput('CSV Import Triggered Successfully');
  } catch (error) {
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

// Main function to find relevant emails and import CSV data from email link
function importCSVFromEmail() {
  const threads = GmailApp.search(`subject:"${EMAIL_SUBJECT_FILTER}" from:${SENDER_EMAIL} is:unread`);
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const body = message.getBody();
      const csvLink = extractCSVLink(body);
      
      if (csvLink) {
        const csvData = fetchCSVData(csvLink);
        if (csvData) {
          insertDataIntoSheet(csvData);
          message.markRead();  // Mark as read after successful processing
        }
      }
    });
  });
}

// Extracts CSV link from email body content based on specific export name
function extractCSVLink(body) {
  const urlPattern = new RegExp(`https:\/\/storage\\.googleapis\\.com\/[^\\s]+${EXPORT_NAME}[^\\s]*\\.csv(\\?[^<\\s]*)?`, 'g');
  const matches = body.match(urlPattern);
  if (matches) {
    let url = matches[0].replace(/&amp;/g, '&');
    return url.endsWith('"') ? url.slice(0, -1) : url;
  }
  return null;
}

// Fetches and parses CSV data from a provided URL
function fetchCSVData(url) {
  const sanitizedUrl = sanitizeUrl(url);

  try {
    const response = UrlFetchApp.fetch(sanitizedUrl, { muteHttpExceptions: true });
    return Utilities.parseCsv(response.getContentText());
  } catch (error) {
    console.error('Error fetching or parsing CSV:', error.toString());
    return null;
  }
}

// Inserts CSV data into Google Sheet, either in a new or existing sheet
function insertDataIntoSheet(data) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  
  if (OVERRIDE_SHEET) {
    const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
    sheet.clear();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  } else {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    let sheetName = `Imported_${timestamp}`;
    
    // Ensure unique sheet name in case of duplicates
    let sheet = spreadsheet.getSheetByName(sheetName);
    let counter = 1;
    while (sheet) {
      sheetName = `Imported_${timestamp}_${counter++}`;
      sheet = spreadsheet.getSheetByName(sheetName);
    }
    
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

// Helper function to sanitize and reconstruct URL if needed
function sanitizeUrl(url) {
  const [baseUrl, queryString] = url.split('?');
  if (!queryString) return baseUrl;
  
  const queryParams = queryString.split('&').reduce((params, param) => {
    const [key, value] = param.split('=');
    params[decodeURIComponent(key)] = value ? decodeURIComponent(value) : '';
    return params;
  }, {});

  const sanitizedQueryString = Object.entries(queryParams)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join('&');

  return `${baseUrl}?${sanitizedQueryString}`.replace(/"$/, ''); // Remove trailing quote if present
}
