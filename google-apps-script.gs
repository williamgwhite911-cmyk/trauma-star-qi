/**
 * Trauma Star QI Report - Google Sheets endpoint
 *
 * Receives posts from the Trauma Star Documentation QI app and writes them to
 * the spreadsheet this script is bound to. Two tabs are used:
 *
 *   "Flight Log"  - one row per Trauma Star North / South flight: OCA, date,
 *                   nurse, medic, pilot, the dispatch-to-depart clock, sending
 *                   and receiving hospital, diagnosis, and each crew member's
 *                   running flight number.
 *   "QI Review"   - the full documentation audit row for the same call.
 *
 * Rows are keyed by OCA number: re-sending a call overwrites its row instead of
 * adding a duplicate, so a corrected review replaces the original.
 *
 * SETUP
 *  1. Open the Google Sheet named "Trauma Star QI Report".
 *  2. Extensions > Apps Script. Delete anything there and paste this file.
 *  3. Deploy > New deployment > type "Web app".
 *       Execute as: Me
 *       Who has access: Anyone
 *     (Anyone is required because the app posts from the browser with no login.
 *      The URL is the only credential - treat it like a password and do not
 *      publish it. No patient identifiers or chart text are ever sent.)
 *  4. Copy the /exec URL and paste it into the app under
 *     "Central QI Log - Apps Script URL".
 *
 * After changing this file, use Deploy > Manage deployments > edit > New
 * version, or the old code keeps running.
 */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var tabName = body.sheet || 'QI Review';
    var header = body.header || [];
    var rows = body.rows || [];
    var keyColumn = body.keyColumn || null;

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // The app's "Test Connection" button sends this. It proves the deployment
    // is reachable and correctly configured without writing a junk row.
    if (body.action === 'ping') {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, ping: true, spreadsheet: ss.getName() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);

    // Write or refresh the header row, widening the sheet if columns were added.
    if (header.length) {
      if (sh.getMaxColumns() < header.length) {
        sh.insertColumnsAfter(sh.getMaxColumns(), header.length - sh.getMaxColumns());
      }
      var current = sh.getLastRow() > 0
        ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
        : [];
      if (current.join('') !== header.join('')) {
        sh.getRange(1, 1, 1, header.length).setValues([header]);
        sh.getRange(1, 1, 1, header.length)
          .setFontWeight('bold')
          .setBackground('#1a3a5c')
          .setFontColor('#ffffff');
        sh.setFrozenRows(1);
      }
    }

    // Index existing rows by key so a resend updates in place.
    var keyIdx = keyColumn ? header.indexOf(keyColumn) : -1;
    var existing = {};
    if (keyIdx >= 0 && sh.getLastRow() > 1) {
      var keys = sh.getRange(2, keyIdx + 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        var k = String(keys[i][0]).trim();
        if (k) existing[k] = i + 2; // sheet row number
      }
    }

    var appended = 0, updated = 0;
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      while (row.length < header.length) row.push('');
      var key = keyIdx >= 0 ? String(row[keyIdx]).trim() : '';
      if (key && existing[key]) {
        sh.getRange(existing[key], 1, 1, row.length).setValues([row]);
        updated++;
      } else {
        sh.appendRow(row);
        if (key) existing[key] = sh.getLastRow();
        appended++;
      }
    }

    if (appended || updated) sh.autoResizeColumns(1, Math.min(header.length, 12));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, sheet: tabName, appended: appended, updated: updated }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Trauma Star QI Report endpoint is running. Post JSON to this URL.');
}
