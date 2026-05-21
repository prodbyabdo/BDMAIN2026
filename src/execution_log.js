const EXECUTION_LOG_SHEET_NAME = 'EXEC_LOG';
const EXECUTION_LOG_HEADERS = [
  'Execution ID',
  'Timestamp',
  'Status',
  'Script',
  'Sheet',
  'Row',
  'Column',
  'Trigger',
  'Duration Ms',
  'User',
  'Message',
  'Context',
];

function runWithExecutionLog_(scriptName, context, fn) {
  const startedAt = new Date();
  const startedMs = Date.now();
  const entry = beginExecutionLog_(scriptName, context, startedAt);

  try {
    const result = fn();
    endExecutionLog_(entry, 'SUCCESS', Date.now() - startedMs, 'Completed', context);
    return result;
  } catch (error) {
    endExecutionLog_(
      entry,
      'ERROR',
      Date.now() - startedMs,
      error && error.message ? error.message : String(error),
      context
    );
    throw error;
  }
}

function beginExecutionLog_(scriptName, context, startedAt) {
  const sheet = ensureExecutionLogSheet_();
  const executionId = Utilities.getUuid();
  const user = getExecutionUser_();
  const row = [
    executionId,
    startedAt,
    'RUNNING',
    scriptName,
    getContextValue_(context, 'sheet'),
    getContextValue_(context, 'row'),
    getContextValue_(context, 'column'),
    getContextValue_(context, 'trigger'),
    '',
    user,
    'Started',
    stringifyContext_(context),
  ];
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
  return { sheet: sheet, row: nextRow };
}

function endExecutionLog_(entry, status, durationMs, message, context) {
  if (!entry || !entry.sheet || !entry.row) return;
  const sheet = entry.sheet;
  const values = [
    [
      sheet.getRange(entry.row, 1).getValue(),
      sheet.getRange(entry.row, 2).getValue(),
      status,
      sheet.getRange(entry.row, 4).getValue(),
      getContextValue_(context, 'sheet'),
      getContextValue_(context, 'row'),
      getContextValue_(context, 'column'),
      getContextValue_(context, 'trigger'),
      durationMs,
      sheet.getRange(entry.row, 10).getValue(),
      message,
      stringifyContext_(context),
    ],
  ];
  sheet.getRange(entry.row, 1, 1, values[0].length).setValues(values);
}

function ensureExecutionLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EXECUTION_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EXECUTION_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, EXECUTION_LOG_HEADERS.length).setValues([EXECUTION_LOG_HEADERS]);
    sheet.getRange(1, 1, 1, EXECUTION_LOG_HEADERS.length).setFontWeight('bold');

    const protection = sheet.protect().setDescription('Only owner can edit EXEC_LOG');
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  }
  return sheet;
}

function getExecutionUser_() {
  try {
    return Session.getActiveUser().getEmail() || 'Unknown User';
  } catch (error) {
    return 'Unknown User';
  }
}

function getContextValue_(context, key) {
  if (!context || typeof context !== 'object') return '';
  const value = context[key];
  return value === undefined || value === null ? '' : String(value);
}

function stringifyContext_(context) {
  if (!context || typeof context !== 'object') return '';
  try {
    const json = JSON.stringify(context);
    return json.length > 3000 ? json.slice(0, 3000) : json;
  } catch (error) {
    return String(context);
  }
}

function logExecutionEvent_(scriptName, status, message, context) {
  const sheet = ensureExecutionLogSheet_();
  const user = getExecutionUser_();
  const row = [
    Utilities.getUuid(),
    new Date(),
    status,
    scriptName,
    getContextValue_(context, 'sheet'),
    getContextValue_(context, 'row'),
    getContextValue_(context, 'column'),
    getContextValue_(context, 'trigger'),
    0,
    user,
    message,
    stringifyContext_(context),
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}
