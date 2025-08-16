const REQUESTED_SHEET_NAME = "Requested";
const APPROVED_SHEET_NAME = "Approved";
const ORDERED_SHEET_NAME = "Ordered";
const APPROVAL_CHECKBOX_COLUMN = 1; // Column A

/**
 * Creates a custom menu with direct actions for approved items.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Ordering')
      .addItem('Approve Checked Items', 'approveSelectedItems')
      .addSeparator() // Adds a visual line in the menu
      .addItem('Download Approved Items as CSV', 'downloadApprovedItems')
      .addItem('Send Approved Items to Server', 'sendApprovedItems')
      .addToUi();
}

function approveSelectedItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestedSheet = ss.getSheetByName(REQUESTED_SHEET_NAME);
  const approvedSheet = ss.getSheetByName(APPROVED_SHEET_NAME);
  const responseSheet = ss.getSheetByName("Raw Responses"); // Get the source sheet
  const ui = SpreadsheetApp.getUi();

  if (requestedSheet.getLastRow() < 2) {
    ui.alert('No data to approve.', 'The "Requested" sheet is empty.', ui.ButtonSet.OK);
    return;
  }

  const dataRange = requestedSheet.getRange("A1").getDataRegion();
  const allValues = dataRange.getValues();
  
  const headers = allValues.shift();
  const originalNumDataRows = allValues.length;

  const rowsToMove = [];
  const timestamp = new Date();

  allValues.forEach(row => {
    if (row[0] === true) {
      rowsToMove.push([timestamp, ...row.slice(1)]);
    }
  });

  if (rowsToMove.length === 0) {
    ui.alert('No Items Approved', 'No checked boxes were found.', ui.ButtonSet.OK);
    return;
  }

  // 1. Append moved items to the "Approved" sheet
  approvedSheet.getRange(
    approvedSheet.getLastRow() + 1,
    1,
    rowsToMove.length,
    rowsToMove[0].length
  ).setValues(rowsToMove);

    // 3. NEW STEP: Clear all checkboxes in column A to reset the sheet for the next use.
  // This prevents "phantom" checked boxes from remaining.
  requestedSheet.getRange("A2:A").clearContent();

  // 2. Delete the corresponding rows from the "Raw Responses" sheet
  for (let i = originalNumDataRows - 1; i >= 0; i--) {
    if (allValues[i][0] === true) {
      responseSheet.deleteRow(i + 2);
    }
  }
  
  // 4. Force the sheet to update everything
  SpreadsheetApp.flush();
  
  
  ui.alert('Approval Complete', `${rowsToMove.length} item(s) have been moved.`, ui.ButtonSet.OK);
}

/**
 * Fetches data from the "Approved" sheet and prepares it for processing.
 * This is a helper function used by both download and send actions.
 * @returns {object|null} An object containing the items and sheet dimensions, or null if empty.
 */
function getAndPrepareApprovedItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvedSheet = ss.getSheetByName(APPROVED_SHEET_NAME);

  if (approvedSheet.getLastRow() <= 1) {
    return null; // Return null if there are no items to process
  }

  const numDataRows = approvedSheet.getLastRow() - 1;
  const numColumns = approvedSheet.getLastColumn();
  const itemsToOrder = approvedSheet.getRange(2, 1, numDataRows, numColumns).getValues();

  // Create a new timestamp for when the items were processed
  const timestamp = new Date();
  const itemsWithNewTimestamp = itemsToOrder.map(row => {
    row[0] = timestamp; // The first column is replaced with the new timestamp
    return row;
  });
  
  return {
    items: itemsWithNewTimestamp,
    numRows: numDataRows,
    numCols: numColumns
  };
}


/**
 * Moves items to the "Ordered" sheet and clears the "Approved" sheet.
 * This is a helper function used after a successful download or send.
 * @param {object} dataToFinalize The object returned by getAndPrepareApprovedItems.
 */
function finalizeOrder(dataToFinalize) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvedSheet = ss.getSheetByName(APPROVED_SHEET_NAME);
  const orderedSheet = ss.getSheetByName(ORDERED_SHEET_NAME);
  
  // Append the processed items to the "Ordered" sheet
  orderedSheet.getRange(orderedSheet.getLastRow() + 1, 1, dataToFinalize.items.length, dataToFinalize.numCols).setValues(dataToFinalize.items);
  
  // Clear the "Approved" sheet
  approvedSheet.getRange(2, 1, dataToFinalize.numRows, dataToFinalize.numCols).clearContent();
}


/**
 * Triggered by the menu. Creates a CSV for download and finalizes the order.
 */
function downloadApprovedItems() {
  const ui = SpreadsheetApp.getUi();
  const dataPayload = getAndPrepareApprovedItems();

  if (!dataPayload) {
    ui.alert('No items to download.', 'The "Approved" sheet is empty.', ui.ButtonSet.OK);
    return;
  }
  
  // Create the CSV content and a download link
  const csvContent = dataPayload.items.map(row => row.join("||")).join("\n");
  const fileName = `Order_${new Date().toISOString().slice(0,10)}.csv`;
  const dataUri = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csvContent);
  
  // Show a simple dialog with the download link
  const html = `<html><body><p>Your file is ready. <a href="${dataUri}" download="${fileName}">Click here to download</a>.</p></body></html>`;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(80);
  ui.showModalDialog(htmlOutput, 'Download CSV');

  // Move data from "Approved" to "Ordered"
  finalizeOrder(dataPayload);
  
  ui.alert('Process Complete', 'The "Approved" items have been moved to the "Ordered" sheet.', ui.ButtonSet.OK);
}


/**
 * Triggered by the menu. Sends data to the server and finalizes the order upon success.
 */
function sendApprovedItems() {
  const ui = SpreadsheetApp.getUi();
  const dataPayload = getAndPrepareApprovedItems();

  if (!dataPayload) {
    ui.alert('No items to send.', 'The "Approved" sheet is empty.', ui.ButtonSet.OK);
    return;
  }

  const csvContent = dataPayload.items.map(row => row.join("||")).join("\n");
  
  try {
    // This calls your existing, debugged function to send the data
    const serverResponse = sendDataToFlask(csvContent);
    
    // IMPORTANT: Only finalize the order if the send was successful
    finalizeOrder(dataPayload);
    
    ui.alert('Success!', 'Data sent to server. Server responded: ' + serverResponse, ui.ButtonSet.OK);
  } catch (e) {
    // If the send fails, we show an error and do NOT move the items,
    // so the user can try again.
    ui.alert('Error', 'Failed to send data to server: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Sends CSV data as a file to a specified URL with enhanced debugging.
 * @param {string} csvContent The CSV data to be sent.
 * @returns {string} The response text from the server.
 */
function sendDataToFlask(csvContent) {
  // Ensure this URL is correct and the tunnel is running.
  const FLASK_APP_URL = "https://add.orionsoftware.systems/add"; 

  console.log('test test loc2')
  
  const fileName = `Order_${new Date().toISOString().slice(0,10)}.csv`;
  const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
  
  const options = {
    'method': 'post',
    'payload': {
      'file': blob 
    },
    // This is crucial for debugging. It prevents Apps Script from
    // throwing an error on HTTP codes like 404 or 502, allowing
    // us to inspect the response ourselves.
    'muteHttpExceptions': true 
  };
  
  Logger.log(`Attempting to send POST request to: ${FLASK_APP_URL}`);
  
  // We wrap this in a try/catch for network-level errors (e.g., DNS failure)
  try {
    const response = UrlFetchApp.fetch(FLASK_APP_URL, options);
    
    // Log everything we receive from the server
    const responseCode = response.getResponseCode();
    const responseHeaders = response.getAllHeaders();
    const responseText = response.getContentText();
    
    Logger.log(`Response Code: ${responseCode}`);
    Logger.log(`Response Headers: ${JSON.stringify(responseHeaders, null, 2)}`);
    // Only log the first 500 characters of the response to avoid flooding the log
    Logger.log(`Response Body (first 500 chars): ${responseText.substring(0, 500)}`);

    // Now, we analyze the response
    if (responseCode === 200) {
      // SUCCESS! The request reached your Flask app and it responded correctly.
      Logger.log('Success! Flask app responded with 200 OK.');
      return responseText;
    } else {
      // FAILURE! The server responded with an error.
      // The log will contain the HTML of the error page from Cloudflare.
      const errorMessage = `Request failed. Server responded with HTTP status ${responseCode}. Check the Apps Script logs for the full error page content.`;
      Logger.log(errorMessage);
      throw new Error(errorMessage);
    }
    
  } catch (e) {
    // This block catches catastrophic failures, like if the URL is completely wrong or the network is down.
    Logger.log('FATAL ERROR during UrlFetchApp.fetch: ' + e.toString());
    throw new Error('Failed to send request. Is the URL correct and the server online? Error: ' + e.message);
  }
}
