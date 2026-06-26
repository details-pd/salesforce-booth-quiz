/**
 * Salesforce Booth Quiz — raffle email capture endpoint
 * ---------------------------------------------------------------
 * Deploy this as a Web App bound to your Google Sheet:
 *   1. Open your Google Sheet.
 *   2. Extensions ▸ Apps Script.
 *   3. Paste this whole file in (replace anything there). Save.
 *   4. Deploy ▸ New deployment ▸ type "Web app".
 *        - Description:  Raffle capture
 *        - Execute as:   Me
 *        - Who has access: Anyone
 *   5. Authorize when prompted.
 *   6. Copy the Web app URL (ends in /exec) and send it to Claude.
 *
 * To change columns, edit HEADERS + the row array in doPost.
 */

var SHEET_NAME = 'Entries';
var HEADERS = ['Timestamp', 'Email', 'Score'];

/** Receives the POST from the quiz and appends a row. */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var email = (data.email || '').toString().trim();

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json_({ ok: false, error: 'invalid_email' });
    }

    var sheet = getSheet_();
    sheet.appendRow([new Date(), email, data.score == null ? '' : data.score]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Lets you sanity-check the deployment in a browser (GET request). */
function doGet() {
  return json_({ ok: true, status: 'Raffle capture endpoint is live.' });
}

/** Returns the target sheet, creating it + header row if needed. */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** JSON response helper. */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
