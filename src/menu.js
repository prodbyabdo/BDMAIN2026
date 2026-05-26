/**
 * 
 *  
 * Custom Menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
    .addItem('Run Master Search', 'runMasterSearch')
    .addItem('Run NEWLABS Master Search', 'runNewLabsMasterSearch')
    .addItem('Verify Selected Phones', 'verifySelectedPhones')
    .addSeparator()
    .addItem('Capitalize Business & Names', 'capitalizeHeadersBatch')
    .addItem('Clean & Format Phone Numbers', 'reformatPhoneNumbers')
    .addItem('Unmerge All Cells', 'unmergeAllCells')
    .addItem('Clear Formatting on TIMESTAMP', 'normalizeActiveSheetTimestamps')
    .addItem('Fix Time 1', 'normalizeAndReverseTimestamps')
    .addItem('Clear Formatting on Current Tab', 'clearCurrentTabFormatting')
    .addSeparator()
    .addItem('Create Filters', 'createLeadFilterViews')
    .addItem('Row Height Ben ', 'setRowHeightForBen')
    .addToUi();
}
// =============================================================================
// SINGLE CONFIG POINT - rename the tab here and everything follows
// QPP_FEEDBACK tab layout: col A = NPI | col B = CLASS | col C = FEEDBACK
// =============================================================================
const QPP_FEEDBACK_TAB = "QPP_FEEDBACK";

// =============================================================================
// MASTER SEARCH (Optimized - Set-based O(1) lookups)
// Col A priority per row:
//   DEAC + QPP  -> "DEAC | EL"  (coexist)
//   DEAC only   -> "DEAC"
//   QPP + flag  -> "EL | Ben"
//   QPP only    -> "EL"
//   No NPI/QPP  -> normal flag search
// =============================================================================
function runMasterSearch() {
  if (typeof runWithExecutionLog_ === 'function') {
    return runWithExecutionLog_('runMasterSearch', { trigger: 'menu' }, () => {
      runMasterSearchCore_({
        targetTabs: ["MAIN", "LABS"],
        lookupTabs: [
          { name: "Ben", tab: "Ben Flags" },
          { name: "Jimmy", tab: "Jimmy Flags" },
          { name: "Selene", tab: "Selene Flags" },
          { name: "Jane", tab: "Jane Flags" },
          { name: "NI", tab: "NI / Not Eligible" },
          { name: "Dis/Wn", tab: "Disconnected" },
          { name: "AI", tab: "DMEDesk Booked" }
        ],
        toastSuffix: "MAIN & LABS"
      });
    });
  }

  runMasterSearchCore_({
    targetTabs: ["MAIN", "LABS"],
    lookupTabs: [
      { name: "Ben", tab: "Ben Flags" },
      { name: "Jimmy", tab: "Jimmy Flags" },
      { name: "Selene", tab: "Selene Flags" },
      { name: "Jane", tab: "Jane Flags" },
      { name: "NI", tab: "NI / Not Eligible" },
      { name: "Dis/Wn", tab: "Disconnected" },
      { name: "AI", tab: "DMEDesk Booked" }
    ],
    toastSuffix: "MAIN & LABS"
  });
}

function runNewLabsMasterSearch() {
  runMasterSearchCore_({
    targetTabs: ["NEWDME"],
    lookupTabs: [
      { name: "Ben", tab: "Ben Flags" },
      { name: "Jimmy", tab: "Jimmy Flags" },
      { name: "Selene", tab: "Selene Flags" },
      { name: "Jane", tab: "Jane Flags" },
      { name: "NI", tab: "NI / Not Eligible" },
      { name: "Dis/Wn", tab: "Disconnected" },
      { name: "MAIN", tab: "MAIN" },
      { name: "LABS", tab: "LABS" }
    ],
    toastSuffix: "NEWLABS & NEWDME"
  });
}

function runMasterSearchCore_(options) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Another search is running. Try again in 30s.', 'Busy', 5
    );
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const QPP_TAB_NAME = typeof QPP_FEEDBACK_TAB !== 'undefined' ? QPP_FEEDBACK_TAB : "QPP_FEEDBACK";
    const targetTabs = options.targetTabs || [];
    const lookupTabs = options.lookupTabs || [];
    const toastSuffix = options.toastSuffix || targetTabs.join(" & ");

    console.time("MasterSearch_Total");

    // ---------------------------------------------------------------------------
    // 1. Config & Data Loading
    // ---------------------------------------------------------------------------
    ss.toast("Loading lookup data...", "Search", 2);

    const lookupData = {};
    lookupTabs.forEach(config => {
      const sheet = ss.getSheetByName(config.tab);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          // Fetch only cols H to L (max width needed) using config constant
          lookupData[config.name] = sheet.getRange(2, COL_SEARCH_START, lastRow - 1, COL_LOOKUP_WIDTH_MAX).getValues();
        } else {
          lookupData[config.name] = [];
        }
      } else {
        lookupData[config.name] = [];
      }
    });

    // PRE-BUILD LOOKUP SETS
    const lookupSets = {};
    lookupTabs.forEach(config => {
      const data = lookupData[config.name];
      const nameSet = new Set();
      const phone1Set = new Set();
      const phone2Set = new Set();
      for (let r = 0; r < data.length; r++) {
        const row = data[r];
        if (row[COL_LOOKUP_NAME_OFFSET]) {
          nameSet.add(String(row[COL_LOOKUP_NAME_OFFSET]).toLowerCase().trim());
        }
        if (row.length > COL_LOOKUP_PH1_OFFSET && row[COL_LOOKUP_PH1_OFFSET]) {
          phone1Set.add(String(row[COL_LOOKUP_PH1_OFFSET]).toLowerCase().trim());
        }
        if (row.length > COL_LOOKUP_PH2_OFFSET && row[COL_LOOKUP_PH2_OFFSET]) {
          phone2Set.add(String(row[COL_LOOKUP_PH2_OFFSET]).toLowerCase().trim());
        }
      }
      lookupSets[config.name] = { name: nameSet, phone1: phone1Set, phone2: phone2Set };
    });

    // Build meeting data lookups (IMPORT_DATA)
    let meetingNameSet = new Set();
    let meetingPhoneSet = new Set();
    const importSheet = ss.getSheetByName("IMPORT_DATA");
    if (importSheet) {
      const lastRow = importSheet.getLastRow();
      if (lastRow >= 2) {
        const meetingData = importSheet.getRange(2, COL_IMPORT_NAME, lastRow - 1, COL_IMPORT_WIDTH).getValues();
        for (let r = 0; r < meetingData.length; r++) {
          if (meetingData[r][0]) meetingNameSet.add(String(meetingData[r][0]).toLowerCase().trim());
          if (meetingData[r][1]) meetingPhoneSet.add(String(meetingData[r][1]).toLowerCase().trim());
        }
      }
    }

    // Load Deactivated NPIs
    const deactivatedNPIs = new Set();
    const deacSheet = ss.getSheetByName("Deactivated");
    if (deacSheet) {
      const lastRow = deacSheet.getLastRow();
      if (lastRow >= 2) {
        console.time("LoadDeactivated");
        const deacValues = deacSheet.getRange(2, COL_DEAC_NPI, lastRow - 1, 1).getValues();
        for (let r = 0; r < deacValues.length; r++) {
          const sNpi = String(deacValues[r][0] || "").trim();
          if (sNpi && sNpi !== "0") deactivatedNPIs.add(sNpi);
        }
        console.timeEnd("LoadDeactivated");
      }
    }

    // Initialize caching maps to avoid duplicate loops for the same names/phones
    const nameLookupCache = new Map();
    const phoneLookupCache = new Map();

    function searchNameFast_(term) {
      if (term === "") return "";
      const cached = nameLookupCache.get(term);
      if (cached !== undefined) return cached;

      const matches = [];
      lookupTabs.forEach(config => {
        const sets = lookupSets[config.name];
        if (sets && sets.name.has(term)) matches.push(config.name);
      });
      if (meetingNameSet.has(term)) matches.push("Meeting");
      const result = matches.join(", ");
      nameLookupCache.set(term, result);
      return result;
    }

    function searchPhoneFast_(term) {
      if (term === "") return "";
      const cached = phoneLookupCache.get(term);
      if (cached !== undefined) return cached;

      const matches = [];
      lookupTabs.forEach(config => {
        const sets = lookupSets[config.name];
        if (sets && (sets.phone1.has(term) || sets.phone2.has(term))) matches.push(config.name);
      });
      if (meetingPhoneSet.has(term)) matches.push("Meeting");
      const result = matches.join(", ");
      phoneLookupCache.set(term, result);
      return result;
    }

    // ---------------------------------------------------------------------------
    // 2. Determine Tabs to Scan
    // ---------------------------------------------------------------------------
    ss.toast(`Scanning ${toastSuffix}...`, "Search");

    // ---------------------------------------------------------------------------
    // 3. Process Each Tab
    // ---------------------------------------------------------------------------
    const processedCache = new Map();

    targetTabs.forEach(tabName => {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet) return;
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      const numRows = lastRow - 1;
      console.time(`Process_${tabName}`);

      _backupColumns(ss, sheet, tabName, numRows);

      // PERFORM A SINGLE BULK READ for columns H to O (8 columns starting from H)
      const searchValues = sheet.getRange(2, COL_SEARCH_START, numRows, COL_SEARCH_WIDTH).getValues();
      const outputRows = [];

      const offsetNpi = COL_NPI - COL_SEARCH_START;         // 15 - 8 = 7
      const offsetName = COL_LOOKUP_NAME_OFFSET;           // 0
      const offsetPhone1 = COL_PHONE1 - COL_SEARCH_START;   // 10 - 8 = 2
      const offsetPhone2 = COL_PHONE2 - COL_SEARCH_START;   // 12 - 8 = 4

      for (let i = 0; i < numRows; i++) {
        const npi = String(searchValues[i][offsetNpi] || "").trim();
        const hasNpi = npi && npi !== "0";

        const termH = clean(searchValues[i][offsetName]);
        const termJ = clean(searchValues[i][offsetPhone1]);
        const termL = clean(searchValues[i][offsetPhone2]);

        // Cache Check: If already processed this exact lead, skip to save CPU
        const cacheKey = (npi || "no_npi") + "|" + termH + "|" + termJ + "|" + termL;
        if ((npi || termH || termJ || termL) && processedCache.has(cacheKey)) {
          outputRows.push(processedCache.get(cacheKey));
          continue;
        }

        const isDeac = hasNpi && deactivatedNPIs.has(npi);

        const flagMatch = searchNameFast_(termH);
        const phone1Match = searchPhoneFast_(termJ);
        const phone2Match = searchPhoneFast_(termL);

        let colA = "";
        if (isDeac) {
          colA = "DEAC";
        } else {
          colA = flagMatch || "";
        }

        const rowResult = [colA, phone1Match, phone2Match];

        // Save to cache
        if (npi || termH || termJ || termL) {
          processedCache.set(cacheKey, rowResult);
        }

        outputRows.push(rowResult);
      }

      // Verify that the row count has not changed since the start of execution
      const currentLastRow = sheet.getLastRow();
      if (currentLastRow - 1 !== numRows) {
        ss.toast(
          `Row count changed (expected ${numRows}, got ${currentLastRow - 1}). Aborting write on ${tabName}.`,
          '⚠️ Aborted', 10
        );
        return; // skip write
      }

      // HIGH-SPEED BULK WRITE: Write all rows in a single API call for maximum performance!
      sheet.getRange(2, COL_OUTPUT_START, numRows, COL_OUTPUT_WIDTH).setValues(outputRows);
      SpreadsheetApp.flush();

      console.timeEnd(`Process_${tabName}`);
    });

    console.timeEnd("MasterSearch_Total");
    ss.toast(`Search complete on ${toastSuffix}`, "Done");
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// SHARED HELPERS
// =============================================================================
function _backupColumns(ss, sheet, tabName, numRows) {
  const backupName = '_SEARCH_BACKUP';
  let backup = ss.getSheetByName(backupName);
  if (!backup) {
    backup = ss.insertSheet(backupName);
    backup.hideSheet();
  }
  // Clear previous backup
  backup.clearContents();
  // Write tab name as header
  backup.getRange(1, 1).setValue(`Backup of ${tabName} @ ${new Date().toISOString()}`);

  if (numRows > 0) {
    // High-speed atomic copy using config.js constants
    const sourceRange = sheet.getRange(2, COL_BACKUP_START, numRows, COL_BACKUP_WIDTH);
    const destRange = backup.getRange(2, COL_BACKUP_START);
    sourceRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    SpreadsheetApp.flush();
  }
}

function clean(val) {
  if (!val) return "";
  if (val instanceof Date) return "";
  return String(val).toLowerCase().trim();
}

// =============================================================================
// REMAINING MENU FUNCTIONS (unchanged)
// =============================================================================
function capitalizeHeadersBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const ALLOWED_TABS = ["MAIN", "LABS", "NEWLABS", "NEWDME", "Ben Flags"];
  const sheetName = sheet.getName();
  if (!ALLOWED_TABS.includes(sheetName)) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Capitalize not allowed on "${sheetName}". Use on: ${ALLOWED_TABS.join(", ")}`);
    return;
  }

  const range = sheet.getDataRange();
  const data = range.getValues();
  const headers = data[0];

  const targets = ["Legalbusinessname", "AuthOfficialName"];
  const columnIndices = [];
  targets.forEach(target => {
    const index = headers.indexOf(target);
    if (index !== -1) columnIndices.push(index);
  });

  if (columnIndices.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Target columns not found on this tab.');
    return;
  }
  for (let i = 1; i < data.length; i++) {
    columnIndices.forEach(colIdx => {
      const val = data[i][colIdx];
      if (val && typeof val === 'string') data[i][colIdx] = val.trim().toUpperCase();
    });
  }
  range.setValues(data);
  SpreadsheetApp.getActiveSpreadsheet().toast('Capitalization complete!', 'Success');
}

function reformatPhoneNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetTabs = ["MAIN", "LABS"];
  const columnIndices = [10, 12];
  targetTabs.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    columnIndices.forEach(colIndex => {
      const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
      const newValues = range.getValues().map(row => {
        const num = String(row[0] || '').replace(/\D/g, '');
        return (num.length === 10)
          ? [`${num.slice(0, 3)}-${num.slice(3, 6)}-${num.slice(6, 10)}`]
          : [row[0]];
      });
      range.setValues(newValues);
    });
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Phone numbers formatted on MAIN and LABS.');
}

function unmergeAllCells() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    .getMergedRanges().forEach(r => r.breakApart());
}

function clearCurrentTabFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear all conditional rules?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    sheet.clearConditionalFormatRules();
  }
}
/**
 * ACTION 1: Scans every cell in the CURRENT sheet and fixes timestamp formats.
 */
function normalizeActiveSheetTimestamps() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  // OPTIMIZATION 1: Only check the columns that actually hold timestamps.
  // Update this array with the column numbers where timestamps live.
  // E.g., 6 is Col F, 17 is Col Q.
  const targetColumns = [6, 17];

  let totalChanges = 0;

  targetColumns.forEach(colIndex => {
    // Check if the column actually exists in the sheet's current scope
    if (colIndex > sheet.getLastColumn()) return;

    const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
    const values = range.getValues();
    let colChangesMade = false;

    for (let r = 0; r < values.length; r++) {
      const cellValue = values[r][0];

      // OPTIMIZATION 2: Fast-fail on empty cells without casting to string
      if (!cellValue) continue;

      const strValue = String(cellValue);

      // OPTIMIZATION 3: Refined check to avoid processing URLs or random text
      if (strValue.includes('/') && (strValue.includes(':') || /gmt/i.test(strValue))) {

        const normalized = smartNormalizer(strValue);

        if (normalized !== strValue) {
          values[r][0] = normalized;
          colChangesMade = true;
          totalChanges++;
        }
      }
    }

    // OPTIMIZATION 4: Only write back the specific column, and only if changes occurred
    if (colChangesMade) {
      range.setValues(values);
    }
  });

  const ui = SpreadsheetApp.getActiveSpreadsheet();
  if (totalChanges > 0) {
    ui.toast(`${totalChanges} cells normalized.`, 'Task Complete', 5);
  } else {
    ui.toast('No timestamps needed normalization.', 'Status', 3);
  }
}

/**
 * ACTION 2: Targeted logic for MAIN, LABS, etc.
 * Normalizes, Sorts Newest-to-Oldest, and keeps 1 per hour.
 */
/**
 * ACTION 2: Targeted logic for MAIN, LABS, etc.
 * Normalizes, Sorts Newest-to-Oldest, and keeps 1 per hour.
 * Note: Assumes America/New_York or user's appsscript.json timezone (Africa/Cairo)
 */
function normalizeAndReverseTimestamps() {
  if (typeof runWithExecutionLog_ === 'function') {
    return runWithExecutionLog_('normalizeAndReverseTimestamps', { trigger: 'menu' }, _normalizeAndReverseTimestampsCore);
  }
  return _normalizeAndReverseTimestampsCore();
}

function _normalizeAndReverseTimestampsCore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetTabs = ["MAIN", "LABS", "Ben Flags", "NI / Not Eligible"];
  const timestampCol = 17; // Column Q

  targetTabs.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, timestampCol, lastRow - 1, 1);
    const values = range.getValues();

    values.forEach((row, i) => {
      const cellVal = String(row[0]).trim();
      if (!cellVal) return;

      // Split lines, normalize each, and filter out failures
      const lines = cellVal.split("\n").map(l => l.trim()).filter(l => l !== "");

      const mapped = lines.map(line => {
        const normalized = smartNormalizer(line);
        // Create a date object for sorting logic
        const dateObj = new Date(normalized.replace(/(\d{1,2})\/(\d{1,2})\/(\d{4})/, "$3-$1-$2").replace(" ", "T"));
        return { line: normalized, date: dateObj };
      });

      const parseable = mapped.filter(item => !isNaN(item.date.getTime()));
      const unparseable = mapped.filter(item => isNaN(item.date.getTime())).map(item => item.line);

      // Sort Newest First
      parseable.sort((a, b) => b.date - a.date);

      // Dedupe: Keep only the latest per hour
      const seenDateHour = new Set();
      const filteredLines = [];

      parseable.forEach(item => {
        const dateHourKey = item.line.split(":")[0]; // e.g., "4/9/2026 22"
        if (!seenDateHour.has(dateHourKey)) {
          seenDateHour.add(dateHourKey);
          filteredLines.push(item.line);
        }
      });

      const newVal = [...filteredLines, ...unparseable].join("\n");
      if (newVal !== cellVal) {
        values[i][0] = newVal;
      }
    });

    range.setValues(values);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Deduplication and Reversal complete.', 'Process Status', 5);
}

/**
 * THE CORE ENGINE: Extracts and normalizes a timestamp from a string.
 * Handles row 312 style: "Some text 4/9/2026 22:17:59"
 */
function smartNormalizer(input) {
  if (!input || typeof input !== "string") return input;

  // 1. Try to find a standard date pattern: M/D/YYYY H:MM:SS AM/PM
  const regex = /(\d{1,2}\/\d{1,2}\/\d{4}),?\s+(\d{1,2}:\d{2}(?::\d{2})?)\s*(AM|PM)?/i;
  const match = input.match(regex);

  if (match) {
    const datePart = match[1];
    const timePart = match[2];
    const meridiem = match[3] ? match[3].toUpperCase() : null;

    let [h, m, s] = timePart.split(":").map(Number);
    s = s || 0; // default to 0 if seconds missing

    if (meridiem === "AM" && h === 12) h = 0;
    else if (meridiem === "PM" && h !== 12) h += 12;

    return `${datePart} ${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  }

  // 2. Fallback: If it's a JS Date string (e.g. Wed Apr 01...)
  const parsed = new Date(input);
  if (!isNaN(parsed.getTime())) {
    return `${parsed.getMonth() + 1}/${parsed.getDate()}/${parsed.getFullYear()} ${parsed.getHours()}:${String(parsed.getMinutes()).padStart(2, '0')}:${String(parsed.getSeconds()).padStart(2, '0')}`;
  }

  return input; // Return original if no date found
}
function setRowHeightForBen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const targetColumn = 3; // Column D is index 3 (0-indexed)
  const targetValue = "Ben";

  // Loop through the data (starting from row 2 to skip headers if necessary)
  for (let i = 0; i < data.length; i++) {
    if (data[i][targetColumn] === targetValue) {
      // +1 because sheet rows are 1-indexed
      sheet.autoResizeRow(i + 1);
    }
  }
}
