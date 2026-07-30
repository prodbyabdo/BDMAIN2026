/** 
 * Custom Menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
    .addItem('Run Main Search', 'runMasterSearch')
    .addItem('Run NEWLEADS Search', 'runNewLabsMasterSearch')
    .addItem('Check NEWDME vs MAIN/LABS', 'runNewDmeVsMainLabsSearch')
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
    .addItem('Reset NEWLEADS Scan', 'resetNewLabsScan')
    .addSeparator()
    .addItem('📥 Download XLSX Copy', 'exportSheetAsXlsxPrompt')
    .addToUi();
}

/**
 * Direct Zero-CPU XLSX Download Helper
 */
function exportSheetAsXlsxPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
  const html = `<div style="font-family: sans-serif; padding: 15px; text-align: center;">
    <p style="font-size: 14px; margin-bottom: 20px;">Click below to download the entire workbook as <strong>.xlsx</strong>:</p>
    <a href="${url}" target="_blank" style="display: inline-block; padding: 12px 24px; background: #1a73e8; color: white; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px;">📥 Download .xlsx File</a>
  </div>`;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(420).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Workbook as XLSX');
}
// =============================================================================
// MASTER SEARCH (Optimized - Set-based O(1) lookups)
// Col A priority per row:
//   DEAC only   -> "DEAC"
//   No DEAC     -> normal flag search
// =============================================================================
function runMasterSearch() {
  const tabs = ["MAIN" , "LABS"];
  const lookups = [
    { name: "Ben", tab: "Ben Flags" },
    { name: "Jimmy", tab: "Jimmy Flags" },
    { name: "Selene", tab: "Selene Flags" },
    { name: "Jane", tab: "Jane Flags" },
    { name: "NI", tab: "NI / Not Eligible" },
    { name: "Dis/Wn", tab: "Disconnected" },
    { name: "Cleads", tab: "mirror for chasers" },
    { name: "DNC", tab: "DNC" },
    { name: "AI", tab: "DMEDesk Booked" }
    
  ];
  if (typeof runWithExecutionLog_ === 'function') {
    return runWithExecutionLog_('runMasterSearch', { trigger: 'menu' }, () => {
      runMasterSearchCore_({
        targetTabs: tabs,
        lookupTabs: lookups,
        toastSuffix: tabs.join(' & ')
      });
    });
  }

  runMasterSearchCore_({
    targetTabs: tabs,
    lookupTabs: lookups,
    toastSuffix: tabs.join(' & ')
  });
}

function runNewLabsMasterSearch() {
  const props = PropertiesService.getScriptProperties();
  const startRow = parseInt(props.getProperty('NEWDME_NEXT_ROW') || '2', 10);

  const tabs = ["NEWDME"];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEWDME");
  const totalRows = sheet ? sheet.getLastRow() : 2;

  if (startRow > totalRows) {
    props.deleteProperty('NEWDME_NEXT_ROW');
    SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME scan already complete. Reset and re-run to start over.', 'Done', 5);
    return;
  }

  const CHUNK = 3800;
  const endRow = Math.min(startRow + CHUNK - 1, totalRows);

  runMasterSearchCore_({
    targetTabs: tabs,
    lookupTabs: [
      { name: "Ben", tab: "Ben Flags" },
      { name: "LABS", tab: "LABS" },
      { name: "Jimmy", tab: "Jimmy Flags" },
      { name: "Selene", tab: "Selene Flags" },
      { name: "Jane", tab: "Jane Flags" },
      { name: "NI", tab: "NI / Not Eligible" },
      { name: "Dis/Wn", tab: "Disconnected" },
      { name: "DNC", tab: "DNC" },
      { name: "AI", tab: "DMEDesk Booked" },
      { name: "Cleads", tab: "mirror for chasers" }
    ],
    toastSuffix: `NEWDME rows ${startRow}–${endRow} of ${totalRows}`,
    startRow: startRow,
    endRow: endRow,
    chunkSize: CHUNK
  });

  if (endRow < totalRows) {
    props.setProperty('NEWDME_NEXT_ROW', String(endRow + 1));
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `NEWLABChunk done (rows ${startRow}–${endRow}). Run again to continue from row ${endRow + 1}.`,
      'Paused — run again', 8
    );
  } else {
    props.deleteProperty('NEWDME_NEXT_ROW');
    SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME scan complete.', 'Done', 5);
  }
}
function runMasterSearchCore_(options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetTabs = options.targetTabs || [];
    const lookupTabs = options.lookupTabs || [];
    const toastSuffix = options.toastSuffix || targetTabs.join(" & ");

    console.time("MasterSearch_Total");

    // ---------------------------------------------------------------------------
    // 1. Config & Data Loading via batchGet
    // ---------------------------------------------------------------------------
    ss.toast("Loading lookup data in batch...", "Search", 2);

    const ssId = ss.getId();
    const ranges = [];
    const rangeMetadata = [];

    lookupTabs.forEach(config => {
      if (ss.getSheetByName(config.tab)) {
        ranges.push(`'${config.tab}'!H2:L`);
        rangeMetadata.push({ type: 'lookup', name: config.name });
      }
    });

    if (ss.getSheetByName("IMPORT_DATA")) {
      ranges.push(`'IMPORT_DATA'!O2:P`);
      rangeMetadata.push({ type: 'import' });
    }

    if (ss.getSheetByName("Deactivated")) {
      ranges.push(`'Deactivated'!O2:O`);
      rangeMetadata.push({ type: 'deac' });
    }

    const lookupData = {};
    const deactivatedNPIs = new Set();
    let meetingRows = [];

    if (ranges.length > 0) {
      console.time("batchGet_API_Call");
      const response = Sheets.Spreadsheets.Values.batchGet(ssId, { ranges: ranges });
      const valueRanges = response.valueRanges || [];
      console.timeEnd("batchGet_API_Call");

      valueRanges.forEach((vr, index) => {
        const meta = rangeMetadata[index];
        const values = vr.values || [];

        if (meta.type === 'lookup') {
          lookupData[meta.name] = values;
        } else if (meta.type === 'import') {
          meetingRows = values;
        } else if (meta.type === 'deac') {
          for (let r = 0; r < values.length; r++) {
            const sNpi = String(values[r][0] || "").trim();
            if (sNpi && sNpi !== "0") deactivatedNPIs.add(sNpi);
          }
        }
      });
    }

    // BUILD UNIFIED HASH INDEX MAPS (O(1) lookups)
    const nameMap = new Map();
    const phoneMap = new Map();

    const addMatch = (map, key, label) => {
      if (!key) return;
      let set = map.get(key);
      if (!set) {
        set = new Set();
        map.set(key, set);
      }
      set.add(label);
    };

    // Populate maps from lookup tabs
    lookupTabs.forEach(config => {
      const data = lookupData[config.name] || [];
      for (let r = 0; r < data.length; r++) {
        const row = data[r];
        if (row.length > COL_LOOKUP_NAME_OFFSET && row[COL_LOOKUP_NAME_OFFSET]) {
          addMatch(nameMap, clean(row[COL_LOOKUP_NAME_OFFSET]), config.name);
        }
        if (row.length > COL_LOOKUP_PH1_OFFSET && row[COL_LOOKUP_PH1_OFFSET]) {
          addMatch(phoneMap, cleanPhone(row[COL_LOOKUP_PH1_OFFSET]), config.name);
        }
        if (row.length > COL_LOOKUP_PH2_OFFSET && row[COL_LOOKUP_PH2_OFFSET]) {
          addMatch(phoneMap, cleanPhone(row[COL_LOOKUP_PH2_OFFSET]), config.name);
        }
      }
    });

    // Populate maps from IMPORT_DATA (meetings)
    if (meetingRows.length > 0) {
      for (let r = 0; r < meetingRows.length; r++) {
        const row = meetingRows[r];
        if (row[0]) {
          addMatch(nameMap, clean(row[0]), "Meeting");
        }
        if (row.length > 1 && row[1]) {
          addMatch(phoneMap, cleanPhone(row[1]), "Meeting");
        }
      }
    }

    // ---------------------------------------------------------------------------
    // 2. Determine Tabs to Scan & Process Each Tab (Scoped Lock)
    // ---------------------------------------------------------------------------
    ss.toast(`Scanning ${toastSuffix}...`, "Search");
    const processedCache = new Map();
    const labelCache = new WeakMap();
    const getLabel_ = (matches) => {
      if (!matches) return "";
      let s = labelCache.get(matches);
      if (s === undefined) {
        s = Array.from(matches).join(", ");
        labelCache.set(matches, s);
      }
      return s;
    };

    targetTabs.forEach(tabName => {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet || sheet.getLastRow() < 2) return;

      const lastRow = sheet.getLastRow();
      const startRow = options.startRow || 2;
      const endRow = options.endRow || lastRow;
      const numRows = Math.min(endRow, lastRow) - startRow + 1;      if (numRows <= 0) return;

      console.time(`Process_${tabName}`);

      // Acquire lock only during processing and write back of the active target tab
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(15000)) {
        ss.toast(`Skip processing ${tabName} - locked by another run.`, 'Busy', 5);
        return;
      }

        try {
        // Get values for the entire sheet (Cols A to O)
        const maxCol = 15;
        const targetValues = sheet.getRange(startRow, 1, numRows, maxCol).getValues();

        // OPTIMIZATION: Extract existing results directly from targetValues in-memory to eliminate a redundant API read
        const existingResults = targetValues.map(row => [row[0], row[1], row[2]]);

          
        let hasChanges = false;

        for (let i = 0; i < numRows; i++) {
          const npi = String(targetValues[i][COL_NPI - 1] || "").trim();
          const hasNpi = npi && npi !== "0";

          const termName = clean(targetValues[i][COL_SEARCH_START - 1]); // Col H
          const termPhone1 = cleanPhone(targetValues[i][COL_PHONE1 - 1]); // Col J
          const termPhone2 = cleanPhone(targetValues[i][COL_PHONE2 - 1]); // Col L

          // Cache Check: If already processed this exact lead, skip to save CPU
          // Cache Check: If already processed this exact lead, skip to save CPU
          const cacheKey = (npi || "no_npi") + "|" + termName + "|" + termPhone1 + "|" + termPhone2;
          if ((npi || termName || termPhone1 || termPhone2) && processedCache.has(cacheKey)) {
            const cachedResult = processedCache.get(cacheKey);
            const cur = existingResults[i];
            if (cur[0] !== cachedResult[0] || cur[1] !== cachedResult[1] || cur[2] !== cachedResult[2]) {
              existingResults[i] = cachedResult;
              hasChanges = true;
            }
            continue;
          }

          const isDeac = hasNpi && deactivatedNPIs.has(npi);

          const nameMatches = nameMap.get(termName);
          const flagMatch = getLabel_(nameMatches);

          const phone1Matches = phoneMap.get(termPhone1);
          const phone1Match = getLabel_(phone1Matches);

          const phone2Matches = phoneMap.get(termPhone2);
          const phone2Match = getLabel_(phone2Matches);

          let colA = "";
          if (isDeac) {
            colA = "DEAC";
          } else {
            colA = flagMatch;
          }
          const current = existingResults[i];

          const rowResult = [
            mergeLabels(current[0], colA),
            mergeLabels(current[1], phone1Match),
            mergeLabels(current[2], phone2Match)
          ];

          // Save to cache
          if (npi || termName || termPhone1 || termPhone2) {
            processedCache.set(cacheKey, rowResult);
          }

          if (
            current[0] !== rowResult[0] ||
            current[1] !== rowResult[1] ||
            current[2] !== rowResult[2]
          ) {
            existingResults[i] = rowResult;
            hasChanges = true;
          }
        }

        // Check if target sheet size changed during processing
        if (sheet.getLastRow() !== lastRow) {
          ss.toast(
            `Row count changed (expected ${lastRow}, got ${sheet.getLastRow()}). Aborting write on ${tabName}.`,
            '⚠️ Aborted', 10
          );
          return;
        }

                // Only write rows if something actually changed
        if (hasChanges) {
          _backupColumns(ss, sheet, tabName, startRow, numRows);
          sheet
            .getRange(startRow, 1, numRows, 3)
            .setValues(existingResults);

          SpreadsheetApp.flush();
        }
      } finally {
        lock.releaseLock();
      }

      console.timeEnd(`Process_${tabName}`);
    });

    console.timeEnd("MasterSearch_Total");
    ss.toast(`Search complete on ${toastSuffix}`, "Done");
  } catch (error) {
    console.error("Master Search Error: " + error.message);
    throw error;
  }
}

// =============================================================================
// SHARED HELPERS
// =============================================================================
function _backupColumns(ss, sheet, tabName, startRow, numRows) {
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
    // High-speed atomic copy instead of slow chunked getValues/setValues
    const sourceRange = sheet.getRange(startRow, 1, numRows, 3);
    const destRange = backup.getRange(2, 1);
    sourceRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    SpreadsheetApp.flush();
  }
}

function clean(val) {
  if (!val) return "";
  if (val instanceof Date) return "";
  return String(val).toLowerCase().trim();
}

function cleanPhone(val) {
  if (!val) return "";
  if (val instanceof Date) return "";
  return String(val).replace(/\D/g, "");
}
function mergeLabels(existing, incoming) {
  // OPTIMIZATION: Short-circuit early to avoid heavy Set/Split allocations for empty or identical values
  if (!incoming) return existing || "";
  if (!existing) return incoming || "";
  if (existing === incoming) return existing;

  const set = new Set();

  String(existing || "")
    .split(",")
    .map(v => v.trim())
    .filter(Boolean)
    .forEach(v => set.add(v));

  String(incoming || "")
    .split(",")
    .map(v => v.trim())
    .filter(Boolean)
    .forEach(v => set.add(v));

  return Array.from(set).join(", ");
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
  const targetTabs = ["NEWDME", "DNC"];
  const columnIndices = [10, 12];

  ss.toast(`Formatting phone numbers on ${targetTabs.join(' & ')}...`, "Clean & Format", 2);

  targetTabs.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const lastRow = Math.min(sheet.getLastRow(), 25000); // cap safety
    if (lastRow < 2) return;

    columnIndices.forEach(colIndex => {
      // Process in chunks to avoid timeout
      const CHUNK = 5000;
      for (let startR = 2; startR <= lastRow; startR += CHUNK) {
        const count = Math.min(CHUNK, lastRow - startR + 1);
        const range = sheet.getRange(startR, colIndex, count, 1);
        const newValues = range.getValues().map(row => {
          const num = String(row[0] || '').replace(/\D/g, '');
          return (num.length === 10)
            ? [`${num.slice(0,3)}-${num.slice(3,6)}-${num.slice(6,10)}`]
            : [row[0]];
        });
        range.setValues(newValues);
        SpreadsheetApp.flush();
      }
    });
  });
  ss.toast(`Phone numbers formatted on ${targetTabs.join(' & ')}.`, "Done", 5);
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
function runNewDmeVsMainLabsSearch() {
  const props = PropertiesService.getScriptProperties();
  const startRow = parseInt(props.getProperty('NEWDME_MAINLABS_NEXT_ROW') || '2', 10);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEWDME");
  const totalRows = sheet ? sheet.getLastRow() : 2;

  if (startRow > totalRows) {
    props.deleteProperty('NEWDME_MAINLABS_NEXT_ROW');
    SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME vs MAIN/LABS scan already complete. Reset to re-run.', 'Done', 5);
    return;
  }

  const CHUNK = 4500;
  const endRow = Math.min(startRow + CHUNK - 1, totalRows);

  runMasterSearchCore_({
    targetTabs: ["NEWDME"],
    lookupTabs: [
      { name: "MAIN", tab: "MAIN" },
      { name: "LABS", tab: "LABS" }
    ],
    toastSuffix: `Wheelchairs vs MAIN/LABS rows ${startRow}–${endRow} of ${totalRows}`,
    startRow: startRow,
    endRow: endRow,
    chunkSize: CHUNK
  });

  if (endRow < totalRows) {
    props.setProperty('NEWDME_MAINLABS_NEXT_ROW', String(endRow + 1));
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Chunk done (rows ${startRow}–${endRow}). Run again to continue from row ${endRow + 1}.`,
      'Paused — run again', 8
    );
  } else {
    props.deleteProperty('NEWDME_MAINLABS_NEXT_ROW');
    SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME vs MAIN/LABS scan complete.', 'Done', 5);
  }
}

function resetNewDmeVsMainLabsScan() {
  PropertiesService.getScriptProperties().deleteProperty('NEWDME_MAINLABS_NEXT_ROW');
  SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME vs MAIN/LABS scan reset. Next run starts from row 2.', 'Reset', 4);
}
function resetNewLabsScan() {
  PropertiesService.getScriptProperties().deleteProperty('NEWDME_NEXT_ROW');
  SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME scan reset. Next run starts from row 2.', 'Reset', 4);
}
