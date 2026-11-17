/**
 * Chess Ladder — live feed of the form responses.
 *
 * Unlike "Publish to web" (which serves a copy cached for several minutes),
 * this reads the sheet at the moment it is asked, so new games and resets
 * show up on the site immediately.
 *
 * Setup (once):
 *   1. Open the response Sheet → Extensions → Apps Script.
 *   2. Replace the contents of Code.gs with this file and click Save.
 *   3. Deploy → New deployment → gear icon → Web app.
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Click Deploy and approve the permission prompt.
 *   4. Copy the Web app URL (ends in /exec) into SCRIPT_URL in index.html.
 *
 * After editing this script later, use Deploy → Manage deployments → Edit →
 * Version: New version, so the same URL keeps working.
 */

// Tab holding the form responses. Leave empty to use the first tab.
const SHEET_NAME = "";

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = (SHEET_NAME && ss.getSheetByName(SHEET_NAME)) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();

  // Columns A–D: Timestamp, Player 1 name, Player 2 name, Game Result.
  const rows = lastRow < 2 ? [] : sheet
    .getRange(2, 1, lastRow - 1, 4)
    .getDisplayValues()
    .filter(r => r.some(c => c !== ""));

  return ContentService
    .createTextOutput(JSON.stringify({ rows: rows, at: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}
