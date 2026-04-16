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
const TARGET_TABS = ["MAIN", "LABS", "Ben Flags"];
const COMMENT_COL = 6;   // Column F — auto-date + timestamp trigger
const TIMESTAMP_COL = 17;  // Column Q
const NAME_COL = 4;   // Column D
const COOLDOWN_SECONDS = 60;

// =============================================================================
// UNIFIED onEdit HANDLER
// =============================================================================
function onEditInstallable(e) {
  // ── GUARD 1: Valid event ──────────────────────────────────────────────────
  if (!e || !e.range) return;
  if (e.range.getWidth() > 1 || e.range.getHeight() > 1) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // ── GUARD 2: Must be a target tab, must be a data row ────────────────────
  if (!TARGET_TABS.includes(sheetName)) return;
  if (row <= 1) return;

  // ── ROUTE: Col F → Auto-date + Timestamp ─────────────────────────────────
  if (col === COMMENT_COL) {
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
function _handleCommentEdit(e, sheet, sheetName, col, row) {
  const user = Session.getActiveUser().getEmail() || "Unknown User";
  const lock = LockService.getUserLock();

  try {
    if (!lock.tryLock(10000)) {
      logAction("LOCK_TIMEOUT", sheetName, row, "Could not acquire lock in 10s", user);
      return;
    }

    const now = new Date();

    // ── Batch read: grab Col D through Col Q in one call ──────────────────
    const batchData = sheet.getRange(row, NAME_COL, 1, TIMESTAMP_COL - NAME_COL + 1).getValues()[0];
    const colDValue = String(batchData[0] || "").trim().toLowerCase();
    const existingTS = String(batchData[TIMESTAMP_COL - NAME_COL] || "").trim();

    // ── COOLDOWN CHECK ────────────────────────────────────────────────────
    if (existingTS) {
      const entries = existingTS.split("\n");
      const lastEntry = entries[entries.length - 1];
      const lastTSDate = new Date(lastEntry);

      if (!isNaN(lastTSDate.getTime())) {
        const secDiff = (now.getTime() - lastTSDate.getTime()) / 1000;
        if (secDiff < COOLDOWN_SECONDS) {
          logAction("COOLDOWN_SKIP", sheetName, row, `Skipped (diff: ${Math.round(secDiff)}s)`, user);
          return;
        }
      }
    }

    const updates = {}; // Register of columns to update: { colIndex: newValue }

    // ── PART 1: AUTO-DATE COLUMN F ──────────────────────
    if (colDValue) {
      const month = now.getMonth() + 1;
      const day = now.getDate();
      const dateSuffix = ` ${month}/${day}`;

      const rawValue = e.range.getValue();
      let cellValue;

      if (rawValue instanceof Date) {
        cellValue = sheet.getRange(row, col).getDisplayValue().trim();
      } else {
        cellValue = String(rawValue).trim();
      }

      if (cellValue === "") {
        updates[COMMENT_COL] = `${month}/${day}`;
      } else {
        const lines = cellValue.split("\n");
        const lastLine = lines[lines.length - 1].trim();

        const todayRegex = new RegExp(`\\b${month}\\/${day}\\b`);
        const lastLineHasToday = todayRegex.test(lastLine);

        if (!lastLineHasToday) {
          lines[lines.length - 1] = lines[lines.length - 1].trimEnd() + dateSuffix;
          updates[COMMENT_COL] = lines.join("\n");
        }
      }
    }

    // ── PART 2: TIMESTAMP IN COLUMN Q (EVERYONE) ──────────────────────────
    const formattedTS = Utilities.formatDate(now, Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
    updates[TIMESTAMP_COL] = existingTS ? formattedTS + "\n" + existingTS : formattedTS;

    // ── EXECUTE BATCH UPDATE ──────────────────────────────────────────────
    // Apply updates if any exist
    for (const [colIdx, val] of Object.entries(updates)) {
      sheet.getRange(row, Number(colIdx)).setValue(val);
    }

    logAction("SUCCESS", sheetName, row, "Edit processed in Col F", user);

  } catch (err) {
    logAction("ERROR", sheetName, row, err.toString(), user);
  } finally {
    lock.releaseLock();
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
  cache.put(lockKey, 'running', 60);

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