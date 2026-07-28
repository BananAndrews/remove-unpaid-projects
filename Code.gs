// ============================================================
//  Blue Belle Weddings – Relocate "Not paid" Projects
//  Server-side Apps Script (Code.gs)
//
//  Scans the last SCAN_ROWS filled rows of the "invoices" sheet,
//  finds rows whose status (column R or S) is "Not paid", clears
//  columns C, D, E, G on those rows, appends them after the true
//  last filled row, then deletes the originals.
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin Menu')
    .addItem('Remove Unpaid Projects', 'moveNotPaidRows')
    .addItem('Sort Payouts', 'sortPayouts')
    .addToUi();
}

// ── Configuration ─────────────────────────────────────────────
var SHEET_NAME    = 'invoices';
var STATUS_COLS   = [18, 19];        // R, S  (1-based)
var CLEAR_COLS    = [3, 4, 5, 7];   // C, D, E, G  (1-based)
var TARGET_STATUS = 'not paid';
var SCAN_ROWS     = 300;             // scan only the last N filled rows
var DATA_LAST_COL = 20;              // T – columns U-Y (21+) are ignored
                                     // entirely (they hold unrelated data
                                     // that runs far below the projects)
var KEY_COL       = 10;              // J (Company Name) – a column that is
                                     // always filled for a real project;
                                     // used to find the true last data row

// ── Entry point ───────────────────────────────────────────────
function moveNotPaidRows() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    ui.alert('Sheet "' + SHEET_NAME + '" not found.');
    return;
  }

  // True last project row, ignoring the "infinite" data in columns U-Y.
  var lastFilledRow = getLastDataRow(sheet);
  if (lastFilledRow === 0) {
    ui.alert('Sheet "' + SHEET_NAME + '" is empty.');
    return;
  }

  var lastColumn = DATA_LAST_COL;    // only ever read/write columns A-T
  var scanStart  = Math.max(1, lastFilledRow - SCAN_ROWS + 1);
  var numRows    = lastFilledRow - scanStart + 1;

  // Single bulk read – fast even for large sheets.
  var data = sheet.getRange(scanStart, 1, numRows, lastColumn).getValues();

  var rowsToMove = [];
  var sourceRows = [];

  for (var i = 0; i < data.length; i++) {
    if (!isNotPaid(data[i])) continue;

    var row = data[i].slice();
    for (var c = 0; c < CLEAR_COLS.length; c++) {
      row[CLEAR_COLS[c] - 1] = '';
    }
    rowsToMove.push(row);
    sourceRows.push(scanStart + i);
  }

  if (rowsToMove.length === 0) {
    ui.alert(
      'No "' + TARGET_STATUS + '" rows found.\n\n' +
      'Scanned rows ' + scanStart + '–' + lastFilledRow + ' (' + numRows + ' rows).\n\n' +
      'If rows exist above row ' + scanStart + ', increase SCAN_ROWS in the script.'
    );
    return;
  }

  // Append after the true last filled row – never overwrites existing data.
  sheet.getRange(lastFilledRow + 1, 1, rowsToMove.length, lastColumn)
       .setValues(rowsToMove);

  // Delete originals bottom-up so row indices don't shift.
  deleteRowsByNumbers(sheet, sourceRows);

  ss.toast(
    'Moved ' + rowsToMove.length + ' row(s) to row ' + (lastFilledRow + 1) + '.',
    'Done', 6
  );
}

// ── Helpers ───────────────────────────────────────────────────

function isNotPaid(rowValues) {
  for (var i = 0; i < STATUS_COLS.length; i++) {
    var v = String(rowValues[STATUS_COLS[i] - 1] || '').trim().toLowerCase();
    if (v === TARGET_STATUS) return true;
  }
  return false;
}

// Find the last row that holds a real project, based on KEY_COL only, so
// the unrelated "infinite" data in columns U-Y can never inflate it.
function getLastDataRow(sheet) {
  var maxRows = sheet.getMaxRows();
  var col = sheet.getRange(1, KEY_COL, maxRows, 1).getValues();
  for (var r = col.length - 1; r >= 0; r--) {
    if (String(col[r][0]).trim() !== '') return r + 1;
  }
  return 0;
}

// Delete the given absolute row numbers, collapsing consecutive ones
// into a single deleteRows() call and working top-down (descending) so
// earlier deletions never shift rows that are still to be removed.
function deleteRowsByNumbers(sheet, rowNumbers) {
  var sorted = rowNumbers.slice().sort(function (a, b) { return b - a; });

  var i = 0;
  while (i < sorted.length) {
    var end = sorted[i];   // highest row number in this block
    var start = end;
    while (i + 1 < sorted.length && sorted[i + 1] === start - 1) {
      start = sorted[i + 1];
      i++;
    }
    sheet.deleteRows(start, end - start + 1);
    i++;
  }
}

function sortPayouts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('For Batch Payments');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet named "For Batch Payments" not found.');
    return;
  }

  const startRow = 2;
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) return;

  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 5);
  const data = dataRange.getValues();

  // Run the transformation logic twice in one execution
  for (let run = 0; run < 2; run++) {
    for (let i = 0; i < data.length; i++) {
      let tempA = data[i][0];
      let tempB = data[i][1];

      data[i][0] = data[i][4]; // A = E
      data[i][1] = tempA;      // B = A
      data[i][4] = tempB;      // E = B
    }
  }

  // Update the sheet once with the final double-processed data
  dataRange.setValues(data);
}
