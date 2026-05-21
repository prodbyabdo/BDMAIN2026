/**
 * ═══════════════════════════════════════════════════════════════
 * COMMENT → GOOGLE TASK  —  BD DME 2026
 * ═══════════════════════════════════════════════════════════════
 *
 * Modular handler called from onEditInstallable(e).
 * Creates a Google Task in the user's @default list whenever a
 * value is entered in the "Comments" column (detected from the
 * header row; falls back to Column D if the header is not found).
 *
 * HOW TO CALL:
 *   Add the following line inside onEditInstallable(e), BEFORE
 *   any early-return routing blocks:
 *
 *     processCommentTask(e);
 *
 * PREREQUISITE:
 *   Enable the Google Tasks API advanced service in your project:
 *   Apps Script editor → Extensions → Services → Google Tasks API (v1)
 *   This also requires the entry in appsscript.json (already added).
 *
 * CONCURRENCY:
 *   Uses CacheService for a 90-second per-row dedup window so that
 *   rapid multi-keystroke saves don't spam task creation.
 */

// ── Fallback column index if "Comments" header is not found ──────────────────
const TASK_FALLBACK_COL = 4; // Column D

// ── Dedup window in seconds ───────────────────────────────────────────────────
const TASK_COOLDOWN_SECONDS = 90;

// ── Task list to target ───────────────────────────────────────────────────────
const TASK_LIST_ID = '@default';

// =============================================================================
// PUBLIC ENTRY POINT — called from onEditInstallable(e)
// =============================================================================

/**
 * Creates a Google Task when the edited cell is in the "Comments" column
 * and contains a non-empty value.
 *
 * Self-guarded: performs its own column check and dedup, so it is safe to
 * call unconditionally at the top of onEditInstallable before any routing.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e  The onEdit event object.
 */
function processCommentTask(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const col = e.range.getColumn();
    const row = e.range.getRow();
    const sheetName = sheet.getName();

    // ── GUARD 1: Must be a data row (not header) ─────────────────────────────
    if (row <= 1) return;

    // ── GUARD 2: Must be the "Comments" column ────────────────────────────────
    const commentsCol = _getCommentsColumnIndex(sheet);
    if (col !== commentsCol) return;

    // ── GUARD 3: Cell must have a value ───────────────────────────────────────
    let cellText = '';
    if (e.value !== undefined && e.value !== null) {
      cellText = String(e.value).trim();
    } else {
      cellText = String(e.range.getValue() ?? '').trim();
    }
    if (!cellText) return;

    // ── GUARD 4: Dedup — skip if a task was already created for this row recently
    const cache = CacheService.getScriptCache();
    const cacheKey = `task_${sheetName}_${row}_${commentsCol}`;
    if (cache.get(cacheKey)) {
      console.log(`[processCommentTask] Dedup skip — ${sheetName} R${row} (within ${TASK_COOLDOWN_SECONDS}s window)`);
      return;
    }

    // ── BUILD TASK ─────────────────────────────────────────────────────────────
    const titleLine = cellText.split('\n')[0].trim(); // first line only for title
    const dueDate = _parseDueDate(cellText);

    /** @type {GoogleAppsScript.Tasks.Schema.Task} */
    const task = {
      title: titleLine || '(no title)',
      notes: cellText,
    };

    // Due date must be an RFC 3339 timestamp (time portion is ignored by Tasks)
    if (dueDate) {
      task.due = dueDate.toISOString();
    }

    // ── INSERT TASK ────────────────────────────────────────────────────────────
    const created = Tasks.Tasks.insert(task, TASK_LIST_ID);

    // ── SET DEDUP CACHE ────────────────────────────────────────────────────────
    cache.put(cacheKey, 'created', TASK_COOLDOWN_SECONDS);

    // ── LOG SUCCESS ────────────────────────────────────────────────────────────
    const dueSuffix = dueDate ? ` | due: ${dueDate.toDateString()}` : ' | no due date';
    logAction(
      'TASK_CREATED',
      sheetName,
      row,
      `Task created: "${created.title}"${dueSuffix} (taskId: ${created.id})`,
      Session.getActiveUser().getEmail() || 'Unknown'
    );

  } catch (err) {
    console.error('[processCommentTask] Error creating task: ' + err.message);
    // Non-fatal — do not rethrow; let onEditInstallable continue normally.
    try {
      logAction(
        'TASK_ERROR',
        e && e.range ? e.range.getSheet().getName() : '?',
        e && e.range ? e.range.getRow() : 0,
        err.toString(),
        Session.getActiveUser().getEmail() || 'Unknown'
      );
    } catch (_) { /* swallow logging failures */ }
  }
}

// =============================================================================
// PRIVATE HELPERS
// =============================================================================

/**
 * Returns the 1-based column index of the "Comments" header in row 1.
 * Falls back to TASK_FALLBACK_COL if the header is not found.
 *
 * @param  {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {number}
 */
function _getCommentsColumnIndex(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return TASK_FALLBACK_COL;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toLowerCase() === 'comments') {
        return i + 1; // convert to 1-based
      }
    }
  } catch (err) {
    console.error('[_getCommentsColumnIndex] ' + err.message);
  }
  console.warn(`[_getCommentsColumnIndex] "Comments" header not found — defaulting to column ${TASK_FALLBACK_COL}`);
  return TASK_FALLBACK_COL;
}

/**
 * Scans `text` for a recognisable date and returns a Date object, or null.
 *
 * Supported formats (in priority order):
 *   YYYY-MM-DD   →  "2026-05-01"
 *   M/D variants →  "5/1", "05/01", "5/01", "05/1"
 *                   (year is assumed to be the current calendar year)
 *
 * Note: M/D will also match the auto-date suffix appended by _handleCommentEdit
 * (e.g. "Follow up 4/24"), which is intentional — that becomes the task due date.
 *
 * @param  {string} text
 * @return {Date|null}
 */
function _parseDueDate(text) {
  // ── Priority 1: ISO date  YYYY-MM-DD ──────────────────────────────────────
  const isoMatch = text.match(/\b(\d{4})-(\d{2})-(\d{2})\b/);
  if (isoMatch) {
    const d = new Date(`${isoMatch[0]}T00:00:00`);
    if (!isNaN(d.getTime())) return d;
  }

  // ── Priority 2: M/D or MM/DD (current year) ───────────────────────────────
  const slashMatch = text.match(/\b(\d{1,2})\/(\d{1,2})\b/);
  if (slashMatch) {
    const year = new Date().getFullYear();
    const month = parseInt(slashMatch[1], 10) - 1; // 0-indexed
    const day = parseInt(slashMatch[2], 10);
    const d = new Date(year, month, day);
    if (!isNaN(d.getTime()) && month >= 0 && month <= 11 && day >= 1 && day <= 31) {
      return d;
    }
  }

  return null;
}

// =============================================================================
// MANUAL TEST — run from the Apps Script editor to verify end-to-end
// =============================================================================

/**
 * Simulate a "Comments" cell edit and verify a task is created.
 * Adjust SHEET_NAME / CELL_ADDRESS before running.
 *
 * Run via: Apps Script editor → select _testProcessCommentTask → Run
 */
function _testProcessCommentTask() {
  const SHEET_NAME = 'MAIN';   // ← change if needed
  const CELL_ADDRESS = 'D2';     // ← change to a cell in your Comments column

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    console.error(`Test aborted: sheet "${SHEET_NAME}" not found.`);
    return;
  }

  const fakeEvent = {
    range: sheet.getRange(CELL_ADDRESS),
    value: 'Test task — follow up on billing 2026-05-15',
    source: ss,
  };

  console.log('[_testProcessCommentTask] Firing processCommentTask with fake event…');
  processCommentTask(fakeEvent);
  console.log('[_testProcessCommentTask] Done — check Google Tasks and the LOGS tab.');
}
