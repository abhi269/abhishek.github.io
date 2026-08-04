/**
 * ============================================================
 *  Sagar & Parul — Wedding RSVP Backend
 *  Google Apps Script (bound to your RSVP spreadsheet)
 * ============================================================
 *
 *  What it does:
 *   - Receives RSVP submissions from the wedding website
 *   - Saves one row per guest into the sheet
 *   - Auto-creates the header row the first time it runs
 *   - Returns a JSON response the site can read
 *
 *  Sheet columns:
 *   Timestamp | Party ID | Guest # | Full Name | WhatsApp Number | Date of Arrival
 * ============================================================
 */

// Name of the tab that RSVPs are written to. It is created automatically
// if it does not already exist.
var SHEET_NAME = 'RSVPs';

// Header row written automatically on first use.
var HEADERS = [
  'Timestamp',
  'Party ID',
  'Guest #',
  'Full Name',
  'WhatsApp Number',
  'Date of Arrival'
];

/**
 * Handles POST requests from the website's RSVP form.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    // Prevent two submissions writing at the same time.
    lock.waitLock(20000);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'No data received' });
    }

    var data = JSON.parse(e.postData.contents);
    var guests = data.guests;

    if (!Array.isArray(guests) || guests.length === 0) {
      return jsonResponse({ status: 'error', message: 'No guests in submission' });
    }

    var sheet = getSheet_();
    var timestamp = new Date();

    // A short unique id so you can see which guests came in together.
    var partyId = Utilities.getUuid().substring(0, 8);

    var rows = [];
    for (var i = 0; i < guests.length; i++) {
      var g = guests[i] || {};
      var name = (g.name || '').toString().trim();
      var whatsapp = (g.whatsapp || '').toString().trim();
      var arrival = (g.arrivalDate || '').toString().trim();

      // Skip completely empty guests just in case.
      if (!name && !whatsapp && !arrival) continue;

      rows.push([
        timestamp,
        partyId,
        i + 1,
        name,
        // Keep the leading + and spaces; store as text so Sheets doesn't mangle it.
        whatsapp,
        arrival
      ]);
    }

    if (rows.length === 0) {
      return jsonResponse({ status: 'error', message: 'All guest entries were empty' });
    }

    // Write all guest rows in one batch.
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);

    // Force the WhatsApp column to render as text (avoids number formatting).
    sheet.getRange(startRow, 5, rows.length, 1).setNumberFormat('@');

    return jsonResponse({
      status: 'success',
      partyId: partyId,
      saved: rows.length
    });

  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handles GET requests — used to quickly test the deployment in a browser.
 * Visiting the web-app URL should show: {"status":"ok",...}
 */
function doGet() {
  return jsonResponse({
    status: 'ok',
    message: 'Sagar & Parul RSVP backend is live. Send RSVPs via POST.'
  });
}

/**
 * Returns the RSVP sheet, creating it (with headers) if needed.
 */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Add headers if the sheet is empty.
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#d5ddd2');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(4, 200); // Full Name
    sheet.setColumnWidth(5, 170); // WhatsApp
    sheet.setColumnWidth(6, 140); // Arrival
  }

  return sheet;
}

/**
 * Builds a JSON HTTP response.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
