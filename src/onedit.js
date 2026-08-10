/**
 * ═══════════════════════════════════════════════════════════════
 * UNIFIED INSTALLABLE onEdit — BD DME 2026
 * ═══════════════════════════════════════════════════════════════
 *
 * Single entry point for ALL edit-triggered logic:
 *   Col F            → Auto-date + timestamp   (formerly onEdit.js)
 *   "Send Lead to"   → Lead routing            (formerly unfunctional.js)
 *
 * ⚠️  TRIGGER SETUP (do this ONCE in Apps Script → Triggers):
 *   Function   : onEditInstallable
 *   Event type : From spreadsheet → On edit
 *
 * ⚠️  IMPORTANT — avoid double-firing:
 *   There must be NO other function named "onEdit", "onEditt", etc.
 *   Delete ALL old triggers before creating the single one above.
 *
 * Logs are written to a tab named "LOGS".
 *
 * CONCURRENCY: Uses getScriptLock() (cross-user) + CacheService
 * for per-row cooldown. Safe for 4+ simultaneous users.
 */

// ── External Spreadsheet IDs ──────────────────────────────────────────────────
const MEETING_LOG_SS_ID = '1uicpBruuFeno2ES4hNw-TIAwNkGEI37gw8Z-A4yMpC8';
const EMAIL_AUTO_SS_ID = '10tWkyiVrYYgVtkguYuw1NLBhO3k041_RoykCd_mXuso';

// ── Internal flag-tab mapping ─────────────────────────────────────────────────
const FLAG_TAB_MAP = {
  'Ben Flags': 'Ben Flags',
  'Jimmy Flags': 'Jimmy Flags',
  'Jane Flags': 'Jane Flags',
  'Selene Flags': 'Selene Flags',
};

// ── Header names (must match Row 1 exactly) ───────────────────────────────────
const HEADER_TRIGGER = 'Send Lead to';
const HEADER_FIRST_DATA = 'NAME';

// ── Config ────────────────────────────────────────────────────────────────────
const EXCLUDED_TABS = ["LOGS", "IMPORT_DATA", "Deactivated", "SCRIPT_CATALOG", "EXECUTION_LOG"];
const COMMENT_COL = 6;   // Column F — auto-date + timestamp trigger
const TIMESTAMP_COL = 17;  // Column Q
const NAME_COL = 4;   // Column D
const COOLDOWN_SECONDS = 2;   // 2s dedup window to allow consecutive user edits

// =============================================================================
// UNIFIED onEdit HANDLER
// =============================================================================
function onEditInstallable(e) {
  // ── GUARD 1: Valid event ──────────────────────────────────────────────────
  if (!e || !e.range) return;
  if (e.range.getWidth() > 1 || e.range.getHeight() > 1) return;

  const col = e.range.getColumn();
  if (col !== 6 && col !== 16 && col !== 17) return;

  const row = e.range.getRow();
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // ── GUARD 2: Exclude system tabs, must be a data row ──────────────────────
  if (EXCLUDED_TABS.includes(sheetName)) return;
  if (row <= 1) return;

  // ── GUARD 3: Ignore programmatic write-back columns ───────────────────────
  // _handleCommentEdit writes to TIMESTAMP_COL (17). Block it here.
  if (col === TIMESTAMP_COL) return;

  // ── GUARD 4: Skip if this edit was a programmatic self-write ─────────────
  const selfWriteKey = `selfwrite_${sheetName}_${row}_${col}`;
  const selfWriteCache = CacheService.getScriptCache();
  if (selfWriteCache.get(selfWriteKey)) {
    selfWriteCache.remove(selfWriteKey);
    return;
  }

  // ── ROUTE: Col F → Auto-date + Timestamp ─────────────────────────────────
  if (col === COMMENT_COL) {
    // ── ACTIVATE TASKS INTEGRATION (PAUSED) ───────────────────────────────────
    // // processCommentTask(e); // Paused to prevent lock contention / API overhead
    _handleCommentEdit(e, sheet, sheetName, col, row);
    return;
  }

  // ── ROUTE: "Send Lead to" column → Lead routing ──────────────────────────
  // Resolve the column dynamically from headers (it may shift)
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = {};
  headerRow.forEach((h, i) => { if (h) headerMap[String(h).trim()] = i + 1; });

  const leadCol = headerMap[HEADER_TRIGGER];
  if (leadCol && col === leadCol) {
    _handleLeadRouting(e, sheet, sheetName, row, headerMap, leadCol);
    return;
  }

  // ── Not a relevant column — exit silently ────────────────────────────────
}

// =============================================================================
// HANDLER A — Col F: Auto-date + Timestamp
// =============================================================================
/**
 * Processes Col F edits: appends today's date to the comment and logs a
 * timestamp in Col Q. Uses getScriptLock() for cross-user safety.
 *
 * KEY DESIGN DECISIONS (concurrency-safe for 4 simultaneous users):
 *  1. getScriptLock() — single lock across ALL users (not per-user)
 *  2. e.value — immutable snapshot from edit time (not re-read from sheet)
 *  3. Explicit ordered writes — timestamp first, then comment
 *  4. Script-level cache cooldown — prevents any user from re-triggering
 *     same row within COOLDOWN_SECONDS
 */
function _handleCommentEdit(e, sheet, sheetName, col, row) {
  const user = Session.getActiveUser().getEmail() || "Unknown User";
  const now = new Date();

  // ── 1. COOLDOWN CHECK BEFORE LOCK (fail fast without blocking) ──────────
  const cache = CacheService.getScriptCache();
  const cooldownKey = `edit_${sheetName}_${row}_${col}`;
  const lastRunISO = cache.get(cooldownKey);

  if (lastRunISO) {
    const lastRun = new Date(lastRunISO);
    if (!isNaN(lastRun.getTime())) {
      const secDiff = (now.getTime() - lastRun.getTime()) / 1000;
      if (secDiff < COOLDOWN_SECONDS) {
        logAction("COOLDOWN_SKIP", sheetName, row, `Skipped (diff: ${Math.round(secDiff)}s, by: ${user})`, user);
        return;
      }
    }
  }

  // ── 2. SINGLE BATCH READ: Grab Col D and Col Q in one API roundtrip ───────
  const rowVals = sheet.getRange(row, NAME_COL, 1, TIMESTAMP_COL - NAME_COL + 1).getValues()[0];
  const colDRaw = rowVals[0]; // Column D (first col of the narrow read)
  const colDValue = String(colDRaw ?? "").trim().toLowerCase();
  const existingTSRaw = rowVals[TIMESTAMP_COL - NAME_COL]; // Column Q offset
  const existingTS = String(existingTSRaw ?? "").trim();

  // ── 3. GET CELL VALUE FROM EVENT (immutable, not re-read) ──────────────
  let cellValue = "";
  if (e.value !== undefined && e.value !== null) {
    cellValue = String(e.value).trim();
  } else {
    // Fallback: read from range (less safe but covers edge cases like paste/formulas)
    const rawValue = e.range.getValue();
    if (rawValue instanceof Date) {
      const m = rawValue.getMonth() + 1;
      const d = rawValue.getDate();
      const y = rawValue.getFullYear();
      cellValue = `${m}/${d}/${y}`;
    } else {
      cellValue = String(rawValue ?? "").trim();
    }
  }

  // ── 4. ROW-LEVEL LOCK (prevents lock starvation by locking specific row via CacheService) ────
  const rowLockKey = `lock_${sheetName}_${row}`;
  let lockAcquired = false;
  let isDateWritten = false;
  let logMessage = "";
  let logStatus = "SUCCESS";

  try {
    // Try to acquire row-level lock from Cache. If locked, wait up to 1s before giving up.
    const lockVal = cache.get(rowLockKey);
    if (lockVal) {
      logAction("LOCK_TIMEOUT", sheetName, row, `Could not acquire row lock (held by: ${lockVal})`, user);
      return;
    }
    cache.put(rowLockKey, user, 15); // Lock for 15 seconds
    lockAcquired = true;

    // Use the timestamp retrieved from the single batch read outside the lock
    const currentTS = existingTS;

    // Write Timestamp
    const formattedTS = Utilities.formatDate(now, Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
    const newTS = currentTS ? formattedTS + "\n" + currentTS : formattedTS;
    sheet.getRange(row, TIMESTAMP_COL).setValue(newTS);

    // Auto-date in Column F
    const isDeletion = cellValue === "" || cellValue === "-";

    if (!isDeletion) {
      const month = now.getMonth() + 1;
      const day = now.getDate();
      const dateSuffix = ` ${month}/${day}`;

      const lines = cellValue.split("\n");
      const lastLine = lines[lines.length - 1].trim();

      const todayRegex = new RegExp(`(?:^|\\s)${month}\\/${day}(?:\\s|$)`);
      const lastLineHasToday = todayRegex.test(lastLine);

      if (!lastLineHasToday) {
        // Mark that WE are writing Col F — so the re-trigger can skip
        cache.put(`selfwrite_${sheetName}_${row}_${COMMENT_COL}`, '1', 10);
        lines[lines.length - 1] = lines[lines.length - 1].trimEnd() + dateSuffix;
        sheet.getRange(row, COMMENT_COL).setValue(lines.join("\n"));
        isDateWritten = true;
      }
    }

    // Set cooldown cache
    cache.put(cooldownKey, now.toISOString(), COOLDOWN_SECONDS + 10);

    if (isDateWritten) {
      logMessage = `Edit processed in Col F (user: ${user})`;
      logStatus = "SUCCESS";
    } else {
      logMessage = `Skipped date formatting — deletion, dash, or today already appended (user: ${user})`;
      logStatus = "DATE_SKIP";
    }

  } catch (err) {
    logStatus = "ERROR";
    logMessage = err.toString();
  } finally {
    if (lockAcquired) {
      cache.remove(rowLockKey);
    }
  }

  // ── 5. LOG WRITE DEFERRED OUTSIDE THE LOCK (prevents lock starvation) ───
  if (lockAcquired) {
    logAction(logStatus, sheetName, row, logMessage, user);
  }
}

// =============================================================================
// HANDLER B — "Send Lead to": Lead Routing
// =============================================================================
function _handleLeadRouting(e, sheet, sheetName, row, headerMap, leadCol) {
  // Batch read everything from Col 1 to leadCol in one single hit
  const fullRowData = sheet.getRange(row, 1, 1, leadCol).getValues()[0];
  const selectedValue = String(fullRowData[leadCol - 1] || "").trim();

  if (!selectedValue) return;

  Logger.log(`Lead routing: "${selectedValue}" on ${sheetName} row ${row}`);

  // Duplicate-run guard (cache-based)
  const lockKey = `${sheetName}_${row}_${leadCol}_${selectedValue}`;
  const cache = CacheService.getScriptCache();
  if (cache.get(lockKey)) {
    Logger.log(`Duplicate trigger blocked for key: "${lockKey}"`);
    return;
  }
  cache.put(lockKey, 'running', 300); // 5 min dedup window (was 60s)

  // Helper: lookup from local batch data
  const val = (headerName) => {
    const idx = headerMap[headerName];
    if (!idx) return '';
    return fullRowData[idx - 1]; // 0-indexed adjustment
  };

  // ── Path A: Copy row to internal flag tab ──────────────────────────────────
  if (FLAG_TAB_MAP.hasOwnProperty(selectedValue)) {
    const destSheet = e.source.getSheetByName(selectedValue);
    if (!destSheet) {
      SpreadsheetApp.getUi().alert(`Tab "${selectedValue}" not found in this spreadsheet.`);
      return;
    }
    const startCol = headerMap[HEADER_FIRST_DATA];
    if (!startCol) return;

    const sliceStart = startCol - 1;
    const sliceEnd = leadCol; // non-inclusive in slice, but leadCol is 1-indexed so it works
    const rowValues = fullRowData.slice(sliceStart, sliceEnd);

    _appendRow(destSheet, rowValues);
    Logger.log(`Path A: Copied row to "${selectedValue}"`);
    return;
  }

  // ── Path B: Schedule Meeting → external Meeting Log ───────────────────────
  if (selectedValue === 'Schedule Meeting') {
    const meetingRow = [
      (sheetName === 'LABS') ? 'IMMUNE' : '',
      val('Owner'),
      'New',
      new Date(),
      val('Legalbusinessname'),
      val('AuthOfficialName'),
      val('OfficePhone'),
      val('Email'),
    ];
    const meetingSS = SpreadsheetApp.openById(MEETING_LOG_SS_ID);
    const meetingSheet = meetingSS.getSheetByName('New Meetings');
    if (!meetingSheet) {
      SpreadsheetApp.getUi().alert('Tab "New Meetings" not found in Meeting Log.');
      return;
    }
    _appendRow(meetingSheet, meetingRow);
    Logger.log('Path B: Meeting row appended');
    return;
  }

  // ── Path C: Send Lab Email → external Email Automation SS ─────────────────
  if (selectedValue === 'Send Lab Email') {
    const emailRow = [
      new Date(),
      val('Legalbusinessname'),
      val('AuthOfficialName'),
      val('Email'),
    ];
    const emailSS = SpreadsheetApp.openById(EMAIL_AUTO_SS_ID);
    const emailSheet = emailSS.getSheetByName('Sheet1');
    if (!emailSheet) {
      SpreadsheetApp.getUi().alert('Tab "Sheet1" not found in Email Automation.');
      return;
    }
    _appendRow(emailSheet, emailRow);
    Logger.log('Path C: Email row appended');
    return;
  }

  Logger.log(`No path matched for value: "${selectedValue}"`);
}

// =============================================================================
// PRIVATE HELPERS
// =============================================================================

/** Appends a values array to the next available row in a sheet. */
function _appendRow(sheet, values) {
  const destRow = sheet.getLastRow() + 1;
  sheet.getRange(destRow, 1, 1, values.length).setValues([values]);
}

/** Writes a log entry to the "LOGS" tab. */
function logAction(status, sheetName, row, message, user) {
  try {
    // Always log to console (visible in Apps Script execution logs)
    console.log(`[${status}] ${sheetName} R${row}: ${message} (${user})`);

    // Only write to the Sheet "LOGS" tab for critical/actionable statuses
    // This avoids flooding the sheet and causing lock contention on fast typing
    const importantStatuses = ["SUCCESS", "ERROR", "LOCK_TIMEOUT", "TASK_CREATED", "TASK_ERROR"];
    if (!importantStatuses.includes(status)) {
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName("LOGS");

    if (!logSheet) {
      logSheet = ss.insertSheet("LOGS");
      const headers = ["Timestamp", "Status", "Sheet", "Row", "Message", "User"];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#f3f3f3");
      logSheet.setFrozenRows(1);
    }

    const lastRow = logSheet.getLastRow();
    const values = [[new Date(), status, sheetName, row, message, user]];
    logSheet.getRange(lastRow + 1, 1, 1, 6).setValues(values);
  } catch (e) {
    console.error("Logging failed: " + e.message);
  }
}
