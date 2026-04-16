const SCRIPT_CATALOG_SHEET_NAME = 'SCRIPT_CATALOG';

const SCRIPT_CATALOG = [
  { file: 'menu.js', name: 'onOpen', purpose: 'Builds the custom menu in the sheet UI.', trigger: 'Spreadsheet open', output: 'Menu items', timing: 'Fast' },
  { file: 'menu.js', name: 'runMasterSearch', purpose: 'Scans MAIN and LABS against flags, deactivated NPIs, and QPP feedback.', trigger: 'Manual menu action', output: 'Columns A-C on MAIN and LABS', timing: 'Heavy' },
  { file: 'menu.js', name: 'runNewLabsMasterSearch', purpose: 'Scans NEWLABS against flags, deactivated NPIs, QPP feedback, and MAIN/LABS duplicates.', trigger: 'Manual menu action', output: 'Columns A-C on NEWLABS', timing: 'Heavy' },
  { file: 'menu.js', name: 'capitalizeHeadersBatch', purpose: 'Uppercases name columns on the active sheet.', trigger: 'Manual menu action', output: 'Edited values in active sheet', timing: 'Fast' },
  { file: 'menu.js', name: 'reformatPhoneNumbers', purpose: 'Normalizes MAIN and LABS phone columns to 10-digit format.', trigger: 'Manual menu action', output: 'Formatted phone columns', timing: 'Fast' },
  { file: 'menu.js', name: 'unmergeAllCells', purpose: 'Breaks apart merged cells on the active sheet.', trigger: 'Manual menu action', output: 'Merged ranges removed', timing: 'Fast' },
  { file: 'menu.js', name: 'clearCurrentTabFormatting', purpose: 'Clears conditional formatting from the active sheet.', trigger: 'Manual menu action', output: 'Conditional rules removed', timing: 'Fast' },
  { file: 'menu.js', name: 'normalizeActiveSheetTimestamps', purpose: 'Normalizes timestamp-like text across the active sheet.', trigger: 'Manual menu action', output: 'Normalized timestamps', timing: 'Medium' },
  { file: 'menu.js', name: 'normalizeAndReverseTimestamps', purpose: 'Sorts and deduplicates timestamp history in target tabs.', trigger: 'Manual menu action', output: 'Column Q history cleaned', timing: 'Medium' },
  { file: 'menu.js', name: 'smartNormalizer', purpose: 'Parses and standardizes timestamp text.', trigger: 'Helper function', output: 'Normalized timestamp string', timing: 'Fast' },
  { file: 'onedit.js', name: 'onEditInstallable', purpose: 'Single installable edit trigger for comment and routing logic.', trigger: 'Spreadsheet edit', output: 'Delegates to handlers', timing: 'Fast' },
  { file: 'onedit.js', name: '_handleCommentEdit', purpose: 'Adds auto-date and timestamp history for comment edits.', trigger: 'Column F edit', output: 'Column F and Q updates', timing: 'Medium' },
  { file: 'onedit.js', name: '_handleLeadRouting', purpose: 'Routes leads to flag tabs, meeting log, or email automation.', trigger: '"Send Lead to" edit', output: 'Row copied or exported', timing: 'Medium' },
  { file: 'onedit.js', name: 'logAction', purpose: 'Writes event rows into the LOGS tab.', trigger: 'Called by handlers', output: 'LOGS sheet entry', timing: 'Fast' },
  { file: 'add filter.js', name: 'createLeadFilterViews', purpose: 'Builds filter views for MAIN and LABS from ReferenceToCancel keywords.', trigger: 'Manual menu action', output: 'Filter views', timing: 'Medium' },
  { file: 'DailyTriggersAppscript.js', name: 'setupDailyTriggers', purpose: 'Creates daily time-based project triggers.', trigger: 'Manual run', output: 'Script triggers', timing: 'Fast' },
  { file: 'DailyTriggersAppscript.js', name: 'codexListTriggers', purpose: 'Returns trigger metadata for inspection.', trigger: 'Manual run', output: 'Trigger summary object', timing: 'Fast' },
  { file: 'DailyTriggersAppscript.js', name: 'codexProjectSummary', purpose: 'Summarizes the project trigger state.', trigger: 'Manual run', output: 'Project summary object', timing: 'Fast' },
  { file: 'DailyTriggersAppscript.js', name: 'safeCall', purpose: 'Wrapper for guarded function calls.', trigger: 'Helper function', output: 'Safe return value', timing: 'Fast' },
  { file: 'DailyTriggersAppscript.js', name: 'safeEnum', purpose: 'Safely converts enum-like values to strings.', trigger: 'Helper function', output: 'String or null', timing: 'Fast' },
  { file: 'DailyTriggersAppscript.js', name: 'safeString', purpose: 'Safely converts values to strings.', trigger: 'Helper function', output: 'String or null', timing: 'Fast' },
  { file: 'DailyTriggersAppscript.js', name: 'uniqueValues_', purpose: 'Deduplicates arrays while preserving order.', trigger: 'Helper function', output: 'Unique array', timing: 'Fast' },
];

function buildScriptCatalog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SCRIPT_CATALOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SCRIPT_CATALOG_SHEET_NAME);
  }

  const headers = ['File', 'Function', 'Purpose', 'Trigger', 'Output', 'Timing'];
  const rows = SCRIPT_CATALOG.map(item => [
    item.file,
    item.name,
    item.purpose,
    item.trigger,
    item.output,
    item.timing,
  ]);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return { ok: true, sheet: SCRIPT_CATALOG_SHEET_NAME, count: rows.length };
}

function refreshExecutionLogSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EXECUTION_LOG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('EXEC_LOG does not exist yet. Run a script once first.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const status = lastRow > 1 ? `Rows logged: ${lastRow - 1}` : 'No execution rows yet';
  SpreadsheetApp.getUi().alert(status);
}
