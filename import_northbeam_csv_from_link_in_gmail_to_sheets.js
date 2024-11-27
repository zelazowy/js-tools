// created by <3 with chatgp

// Mapping of export names to spreadsheet IDs
const EXPORT_MAPPING = {
  'MO_SPRAY_L7D_Daily': 'sheet_1',
  'MO_PILLS_L7D_Daily': 'sheet_2',
  'MO_DE_L7D_Daily': 'sheet_3'
};

// Specify if data should override existing content or create a new sheet
const OVERRIDE_SHEET = true;

// Email filters
const EMAIL_SUBJECT_FILTER = 'Your daily data export is ready';
const SENDER_EMAIL = 'support@northbeam.io';

// Entry point for the web app
function doGet() {
  try {
    importCSVFromEmail();
    return ContentService.createTextOutput('CSV Import Triggered Successfully');
  } catch (error) {
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

// Main function to process all relevant messages in threads
function importCSVFromEmail() {
  const threads = GmailApp.search(`subject:"${EMAIL_SUBJECT_FILTER}" from:${SENDER_EMAIL} is:unread`);

  threads.forEach(thread => {
    let threadProcessed = false; // Track if the thread was processed

    const messages = thread.getMessages();
    messages.forEach(message => {
      const body = message.getBody();

      Object.keys(EXPORT_MAPPING).forEach(exportName => {
        const csvLink = extractCSVLink(body, exportName);
        if (csvLink) {
          const csvData = fetchCSVData(csvLink);
          if (csvData) {
            const targetSheetId = EXPORT_MAPPING[exportName];
            insertDataIntoSpreadsheet(csvData, targetSheetId, exportName);
            threadProcessed = true;
          }
        }
      });
    });

    // Mark the thread as read if any message was processed
    if (threadProcessed) {
      thread.markRead();
    }
  });
}

// Extracts CSV link matching a specific export name
function extractCSVLink(body, exportName) {
  const urlPattern = new RegExp(`https:\/\/storage\\.googleapis\\.com\/[^\\s]+${exportName}[^\\s]*\\.csv(\\?[^<\\s]*)?`, 'g');
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

// Inserts CSV data into a target spreadsheet
function insertDataIntoSpreadsheet(data, sheetId, exportName) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);

  if (OVERRIDE_SHEET) {
    const sheet = spreadsheet.getActiveSheet(); // Use the default active sheet
    sheet.clear();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  } else {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    let sheetName = `${exportName}_${timestamp}`;
    let sheet = spreadsheet.getSheetByName(sheetName);

    // Ensure unique sheet name
    let counter = 1;
    while (sheet) {
      sheetName = `${exportName}_${timestamp}_${counter++}`;
      sheet = spreadsheet.getSheetByName(sheetName);
    }

    // Create a new sheet and insert data
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
