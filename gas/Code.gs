// ============================================================
// Smart 5G Dashboard — Google Apps Script Web App (Code.gs)
// Handles sync, read, delete, and append operations for all
// Google Sheets used by the dashboard.
//
// Deploy as a Web App:
//   Execute as: Me
//   Who has access: Anyone
//
// ── CONFIGURATION ────────────────────────────────────────────
// Set this to your Google Spreadsheet ID (the long string in
// the sheet URL between /d/ and /edit).
// This is the FALLBACK used when the request does NOT include
// a spreadsheetId in the POST body.
// ============================================================
var FALLBACK_SPREADSHEET_ID = '15HggDixs1lC0rer1msSOyuQZJXSgqC0zielcMbCDgCE';

/**
 * GET handler – health check / sanity ping.
 */
function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'Smart 5G GAS Web App is running.' });
}

/**
 * POST handler – main entry point.
 * Expected body: { sheet, action, data [, spreadsheetId] }
 *   spreadsheetId (optional) – if supplied, targets that specific spreadsheet so every
 *                               role's browser writes to the same sheet regardless of
 *                               which GAS deployment URL is cached locally.
 *   action = 'sync'   → replace all rows with data[]
 *   action = 'read'   → return all rows as JSON array
 *   action = 'delete' → delete row where id === data.id
 *   action = 'append' → append a single row
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheetName = body.sheet;
    var action    = body.action;
    var data      = body.data;

    // Resolve the target spreadsheet:
    //   1. Use the spreadsheetId from the POST body (sent by app.js).
    //   2. Fall back to the hardcoded FALLBACK_SPREADSHEET_ID constant above.
    //   3. Last resort: the script's bound spreadsheet (getActiveSpreadsheet).
    var targetId = body.spreadsheetId || FALLBACK_SPREADSHEET_ID;
    var ss;
    if (targetId) {
      try {
        ss = SpreadsheetApp.openById(targetId);
      } catch (openErr) {
        console.error('openById failed for id=' + targetId + ': ' + openErr.toString());
        ss = SpreadsheetApp.getActiveSpreadsheet();
      }
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }

    if (action === 'sync')      return syncSheetData(ss, sheetName, data);
    if (action === 'read')      return readSheetData(ss, sheetName);
    if (action === 'delete')    return deleteRow(ss, sheetName, data ? data.id : null);
    if (action === 'append')    return appendRow(ss, sheetName, data);
    if (action === 'updateRow') return updateRow(ss, sheetName, data);

    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ── Helpers ───────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Return the named sheet, creating it (with a bold header placeholder)
 * if it does not exist yet.
 * Auto-creation ensures that a brand-new spreadsheet is bootstrapped
 * on the very first request without any manual setup.
 */
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

/**
 * Convert a single field value to a cell-safe scalar.
 *   - null / undefined  → '' (empty string)
 *   - plain object / array → JSON string
 *   - all other values → returned as-is (string, number, boolean)
 *
 * This helper centralises the null / object check so it is applied
 * consistently when building rows for syncSheetData and appendRow.
 */
function toSheetValue(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

/**
 * Read column headers from row 1.
 * Returns an array of non-empty string header names.
 */
function getHeaders(sheet) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
  var values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headers = [];
  for (var i = 0; i < values.length; i++) {
    var h = String(values[i]).trim();
    if (h) headers.push(h);
  }
  return headers;
}

/**
 * Ensure every key found across all data records has a corresponding
 * column header in row 1.  New columns are appended to the right.
 * Returns the complete, ordered headers array.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object|Object[]} data - A single record object or an array of records.
 *   Both forms are accepted so callers can pass either without extra wrapping.
 * @returns {string[]} The full ordered list of column header names after update.
 */
function ensureHeaders(sheet, data) {
  var records  = Array.isArray(data) ? data : [data];
  var existing = getHeaders(sheet);
  var keySet   = {};
  existing.forEach(function(h) { keySet[h] = true; });

  var newKeys = [];
  records.forEach(function(record) {
    if (!record) return;
    Object.keys(record).forEach(function(k) {
      if (!keySet[k]) {
        keySet[k] = true;
        newKeys.push(k);
      }
    });
  });

  var headers = existing.concat(newKeys);

  if (headers.length > 0) {
    // Write / extend header row and make it bold for readability
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  return headers;
}

// ── Actions ───────────────────────────────────────────────

/**
 * sync – Replace all data rows (rows 2+) with the supplied array.
 * Column headers are auto-created / extended as needed.
 */
function syncSheetData(ss, sheetName, data) {
  var sheet = getOrCreateSheet(ss, sheetName);

  if (!Array.isArray(data) || data.length === 0) {
    // Clear data rows but keep header row
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    return jsonResponse({ status: 'ok' });
  }

  var headers = ensureHeaders(sheet, data);
  if (headers.length === 0) {
    return jsonResponse({ status: 'error', message: 'No headers could be determined.' });
  }

  // Delete existing data rows (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  // Build rows using the shared toSheetValue helper
  var rows = data.map(function(record) {
    return headers.map(function(h) { return toSheetValue(record[h]); });
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  return jsonResponse({ status: 'ok' });
}

/**
 * read – Return all data rows as a JSON array of objects keyed by header name.
 * Auto-creates the sheet tab if it does not exist yet (returns [] immediately).
 * This bootstraps a blank sheet in a brand-new spreadsheet so subsequent
 * sync / append calls can target the correct tab right away.
 */
function readSheetData(ss, sheetName) {
  // Auto-create the tab if missing so a new spreadsheet is bootstrapped on first read.
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ss.insertSheet(sheetName);
    return jsonResponse([]);
  }

  if (sheet.getLastRow() < 2) return jsonResponse([]);

  var headers = getHeaders(sheet);
  if (headers.length === 0) return jsonResponse([]);

  var lastRow = sheet.getLastRow();
  var values  = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var result = [];
  values.forEach(function(row) {
    var obj      = {};
    var hasValue = false;
    headers.forEach(function(h, i) {
      var raw = row[i];
      var v;
      if (raw === null || raw === undefined) {
        v = '';
      } else if (raw instanceof Date) {
        // Google Sheets may auto-convert YYYY-MM-DD strings to Date objects.
        // Format them back to YYYY-MM-DD so date comparisons in the app work correctly.
        v = isNaN(raw.getTime()) ? '' : Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        v = String(raw);
      }
      obj[h] = v;
      if (v !== '') hasValue = true;
    });
    if (hasValue) result.push(obj);
  });

  return jsonResponse(result);
}

/**
 * delete – Remove the row whose 'id' column matches data.id.
 * Returns ok (not an error) when the sheet does not exist yet —
 * there is simply nothing to delete in a fresh spreadsheet.
 */
function deleteRow(ss, sheetName, id) {
  if (!id) return jsonResponse({ status: 'error', message: 'No id provided.' });

  var sheet = ss.getSheetByName(sheetName);
  // Auto-create the tab so the sheet is bootstrapped; nothing to delete yet.
  if (!sheet) {
    ss.insertSheet(sheetName);
    return jsonResponse({ status: 'ok', message: 'Sheet was empty (auto-created).' });
  }

  var headers  = getHeaders(sheet);
  var idColIdx = headers.indexOf('id');
  if (idColIdx < 0) return jsonResponse({ status: 'error', message: 'No "id" column in sheet.' });

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'ok' });

  var idValues = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < idValues.length; i++) {
    if (String(idValues[i][0]) === String(id)) {
      sheet.deleteRow(i + 2); // +2: 1-based index + header row offset
      return jsonResponse({ status: 'ok' });
    }
  }

  return jsonResponse({ status: 'ok', message: 'Row not found.' });
}

/**
 * append – Append a single record row to the sheet.
 * Column headers are auto-created / extended as needed.
 */
function appendRow(ss, sheetName, data) {
  if (!data) return jsonResponse({ status: 'error', message: 'No data provided.' });

  var sheet   = getOrCreateSheet(ss, sheetName);
  var headers = ensureHeaders(sheet, data);
  if (headers.length === 0) return jsonResponse({ status: 'error', message: 'No headers.' });

  // Build the row using the shared toSheetValue helper
  var row = headers.map(function(h) { return toSheetValue(data[h]); });
  sheet.appendRow(row);

  return jsonResponse({ status: 'ok' });
}

/**
 * updateRow – Update specific fields of the row whose 'id' column matches data.id.
 * Only the keys present in data (other than 'id') are written; all other columns
 * are left untouched.  New columns referenced in data are auto-created as needed.
 * Auto-creates the sheet tab if missing (returns error since there is no row to update).
 *
 * @param {GoogleAppsScript.Spreadsheet} ss
 * @param {string} sheetName
 * @param {Object} data - Must contain 'id' plus whichever fields to update.
 */
function updateRow(ss, sheetName, data) {
  if (!data || !data.id) return jsonResponse({ status: 'error', message: 'No id provided.' });

  var sheet = ss.getSheetByName(sheetName);
  // Auto-create the tab so the sheet is bootstrapped; nothing to update yet.
  if (!sheet) {
    ss.insertSheet(sheetName);
    return jsonResponse({ status: 'error', message: 'Sheet was empty (auto-created). Row not found: ' + data.id });
  }

  // Ensure any new columns referenced in data exist in the header row
  var headers = ensureHeaders(sheet, data);
  var idColIdx = headers.indexOf('id');
  if (idColIdx < 0) return jsonResponse({ status: 'error', message: 'No "id" column in sheet.' });

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'error', message: 'No data rows in sheet.' });

  var idValues = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < idValues.length; i++) {
    if (String(idValues[i][0]) === String(data.id)) {
      var rowNumber = i + 2; // +2: 1-based index + header row offset
      // Read the full existing row so we can write all columns back in one batch call.
      // This avoids multiple round-trips and prevents a partial-write state if the
      // script times out mid-loop.
      var fullRow = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      // Overlay only the columns present in data (skip 'id')
      Object.keys(data).forEach(function(k) {
        if (k === 'id') return;
        var colIdx = headers.indexOf(k);
        if (colIdx >= 0) {
          fullRow[colIdx] = toSheetValue(data[k]);
        }
      });
      sheet.getRange(rowNumber, 1, 1, headers.length).setValues([fullRow]);
      return jsonResponse({ status: 'ok' });
    }
  }

  return jsonResponse({ status: 'error', message: 'Row not found: ' + data.id });
}
