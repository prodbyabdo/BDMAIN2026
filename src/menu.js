/** 
 * Custom Menu
 */
//  
//   on spreadsheet open ->
//     build "Custom Tools" menu
//     wire each menu item to its handler function by name (string, not reference)
//   no data processing happens here — just UI wiring
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
    .addToUi();
}
// =============================================================================
// MASTER SEARCH (Optimized - Set-based O(1) lookups)
// Col A priority per row:
//   DEAC only   -> "DEAC"
//   No DEAC     -> normal flag search
// =============================================================================

//  
//   define target tabs = [MAIN, LABS]
//   define lookup tabs = [Ben/Jimmy/Jasmine/Nora/Selene/Jane Flags, NI, Disconnected, DNC, AI]
//   if execution-logging wrapper exists -> run core search wrapped in logger
//   else -> run core search directly
//   (this is the "full scan" entry point — no row chunking/resuming)
function runMasterSearch() {
  const tabs = ["MAIN" , "LABS"];
  const lookups = [
    { name: "Ben", tab: "Ben Flags" },
    { name: "Jimmy", tab: "Jimmy Flags" },
    { name: "Selene", tab: "Selene Flags" },
    { name: "Jasmine", tab: "Jasmine Flags" },
    { name: "Jane", tab: "Jane Flags" },
    { name: "Nora", tab: "Nora Flags" },
    { name: "NI", tab: "NI / Not Eligible" },
    { name: "Dis/Wn", tab: "Disconnected" },
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

//  
//   read saved "next row to process" from script properties (default: row 2)
//   if saved row > total rows on NEWDME -> scan already done, clear property, toast, exit
//   compute this chunk's end row = min(startRow + CHUNK - 1, totalRows)
//   run core search scoped to [startRow, endRow] on NEWDME vs Ben/LABS/Jimmy/etc.
//   if more rows remain -> save new "next row" property, toast "run again to continue"
//   else -> clear property, toast "scan complete"
//   (this is the CHUNKED/RESUMABLE entry point for the large NEWDME tab)
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

  const CHUNK = 8950;
  const endRow = Math.min(startRow + CHUNK - 1, totalRows);

  runMasterSearchCore_({
    targetTabs: tabs,
    lookupTabs: [
      { name: "Ben", tab: "Ben Flags" },
      { name: "LABS", tab: "LABS" },
      { name: "Jimmy", tab: "Jimmy Flags" },
      { name: "Selene", tab: "Selene Flags" },
      { name: "Jasmine", tab: "Jasmine Flags" },
      { name: "Jane", tab: "Jane Flags" },
      { name: "Nora", tab: "Nora Flags" },
      { name: "NI", tab: "NI / Not Eligible" },
      { name: "Dis/Wn", tab: "Disconnected" },
      { name: "DNC", tab: "DNC" },
      { name: "AI", tab: "DMEDesk Booked" }
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

// PSEUDOCODE (the shared engine used by all three "search" menu items):
//   STEP 1 — LOAD LOOKUP DATA (single batched read, not one call per tab)
//     for each configured lookup tab that exists -> queue its H:L range
//     if IMPORT_DATA tab exists -> queue its O:P range (meeting names/phones)
//     if Deactivated tab exists -> queue its O column (deactivated NPIs)
//     batchGet all queued ranges in ONE API call
//     split results back out into: lookupData{}, meetingRows[], deactivatedNPIs Set
//
//   STEP 2 — BUILD O(1) LOOKUP INDEXES
//     nameMap:  cleaned business name -> Set of tab labels that mention it
//     phoneMap: cleaned phone digits  -> Set of tab labels that mention it
//     populate both maps from every lookup tab (name @ offset 0, phone1 @ offset 2, phone2 @ offset 4)
//     also populate both maps from IMPORT_DATA meeting rows (label = "Meeting")
//
//   STEP 3 — SCAN EACH TARGET TAB (MAIN/LABS/NEWDME), ONE AT A TIME, LOCKED
//     for each target tab:
//       determine [startRow, endRow] — full tab, or a caller-specified chunk
//       acquire a script-wide lock (skip this tab if another run holds it)
//       read the tab's row data in ONE getRange call (cols A..O)
//       extract existing col A/B/C values into an in-memory working copy
//       for each row in the chunk:
//         read NPI (col O), search name (col H), phone1 (col J), phone2 (col L)
//         build a cache key from (npi|name|phone1|phone2)
//         IF this exact key was already computed earlier in this same run:
//           reuse the cached result; if it differs from what's currently in the
//           sheet's working copy, mark that row changed
//           (this is the "duplicate lead" fast path — skips recomputation AND
//            still applies the correct result to every duplicate occurrence)
//         ELSE (first time seeing this key):
//           is this NPI in the deactivated set? -> col A becomes "DEAC"
//           else -> look up name in nameMap, phone1/phone2 in phoneMap,
//                   turn each Set of matching labels into a joined string
//                   (join is memoized per-Set via labelCache so identical
//                    match-sets across many rows don't re-stringify each time)
//           merge each new value with whatever's already in that cell
//           (merge = union of comma-separated labels, dedup'd)
//           store this row's result in the cache for future duplicates
//           if the merged result differs from what's currently there -> mark changed
//       after the loop: verify the sheet's row count didn't change mid-run
//         (guards against a concurrent import/delete corrupting the write)
//       if anything actually changed in this chunk:
//         snapshot cols A-C to a hidden backup tab BEFORE writing
//         write the whole updated A:C block back in one setValues call
//         flush so the write is visible immediately
//       release the lock
//   STEP 4 — final toast: "search complete"
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
    // processedCache: dedupes IDENTICAL leads seen more than once IN THIS RUN
    // (key = npi|name|phone1|phone2) so repeated rows skip full recomputation
    const processedCache = new Map();
    // labelCache: memoizes the (Array.from(set).join(", ")) string for a given
    // Set object, since the SAME Set (e.g. everyone flagged only by "Ben") can
    // be looked up thousands of times across a scan — build the string once.
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

        // PSEUDOCODE (per-row loop):
        //   for each row in this chunk:
        //     pull npi / termName / termPhone1 / termPhone2 from the raw row data
        //     build cacheKey from those four values
        //     if cacheKey already seen this run -> reuse cached result, apply if different, skip rest
        //     else -> compute DEAC/name/phone matches fresh, merge with existing
        //             cell values, cache the result, mark changed if different
        for (let i = 0; i < numRows; i++) {
          const npi = String(targetValues[i][COL_NPI - 1] || "").trim();
          const hasNpi = npi && npi !== "0";

          const termName = clean(targetValues[i][COL_SEARCH_START - 1]); // Col H
          const termPhone1 = cleanPhone(targetValues[i][COL_PHONE1 - 1]); // Col J
          const termPhone2 = cleanPhone(targetValues[i][COL_PHONE2 - 1]); // Col L

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

//                               
//   get or create a hidden "_SEARCH_BACKUP" tab
//   clear it, write a header noting which tab/timestamp this backup is for
//   fast-copy (copyTo, values only) the current A:C block into the backup
//   (this is a "just in case" undo snapshot taken right before every write)
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

//   lowercase + trim a cell value for name matching; Dates and empties -> ""
function clean(val) {
  if (!val) return "";
  if (val instanceof Date) return "";
  return String(val).toLowerCase().trim();
}

//   strip everything but digits for phone matching; Dates and empties -> ""
function cleanPhone(val) {
  if (!val) return "";
  if (val instanceof Date) return "";
  return String(val).replace(/\D/g, "");
}

//  
//   if incoming is empty -> return existing unchanged
//   if existing is empty -> return incoming unchanged
//   if they're identical strings -> return as-is (skip allocation)
//   otherwise -> split both on commas, union into a Set (dedupes), rejoin
//   (this is how col A/B/C accumulate multiple flag labels over repeated runs)
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

//  
//   only allow this on a whitelisted set of tabs (MAIN/LABS/NEWLABS/NEWDME/Ben Flags)
//   find the "Legalbusinessname" and "AuthOfficialName" columns by header name
//   for every data row -> uppercase + trim the value in those columns
//   write the whole sheet back in one setValues call
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

//  
//   for each target tab (NEWDME, DNC):
//     for each chunk of up to CHUNK rows:
//       read cols J:L in one getRange call (J and L are the two phone columns,
//       K rides along untouched in between so only one read/write is needed)
//       for each row -> strip non-digits from J and L; if exactly 10 digits,
//       reformat as XXX-XXX-XXXX; leave anything else as-is
//       only write the chunk back if something in it actually changed
//   flush once at the very end (not per-chunk) and toast completion
function reformatPhoneNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetTabs = ["NEWDME"];
  const startCol = 10; // J
  const width = 3;     // J, K, L — reads both phone cols in one pass, K rides along untouched

  ss.toast(`Formatting phone numbers on ${targetTabs.join(' & ')}...`, "Clean & Format", 2);

  targetTabs.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const lastRow = Math.min(sheet.getLastRow(), 50000); // cap safety
    if (lastRow < 2) return;

    const CHUNK = 19990;
    for (let startR = 2; startR <= lastRow; startR += CHUNK) {
      const count = Math.min(CHUNK, lastRow - startR + 1);
      const range = sheet.getRange(startR, startCol, count, width);
      const values = range.getValues();
      let changed = false;

      const newValues = values.map(row => {
        const newRow = row;
        [0, 2].forEach(offset => {
          const num = String(row[offset] || '').replace(/\D/g, '');
          if (num.length === 10) {
            const formatted = `${num.slice(0,3)}-${num.slice(3,6)}-${num.slice(6,10)}`;
            if (formatted !== row[offset]) {
              newRow[offset] = formatted;
              changed = true;
            }
          }
        });
        return newRow;
      });

      if (changed) {
        range.setValues(newValues);
      }
    }
  });

  SpreadsheetApp.flush();
  ss.toast(`Phone numbers formatted on ${targetTabs.join(' & ')}.`, "Done", 5);
}

//   unmerge every merged cell range on the currently active sheet
function unmergeAllCells() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    .getMergedRanges().forEach(r => r.breakApart());
}

//   confirm with the user, then clear all conditional formatting rules on the active sheet
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
//  
//   only check the columns known to hold timestamps (F and Q)
//   for each of those columns -> read the whole column's values in one call
//   for each cell -> skip empties; only bother normalizing if it "looks like"
//     a timestamp (contains "/" and either ":" or "gmt")
//   if normalizing actually changed the string -> mark it and count it
//   write each changed column back in one setValues call (only if it changed)
//   toast how many cells were normalized (or that none needed it)
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
//   thin wrapper — run the real logic wrapped in the execution logger if it exists, else run it directly
function normalizeAndReverseTimestamps() {
  if (typeof runWithExecutionLog_ === 'function') {
    return runWithExecutionLog_('normalizeAndReverseTimestamps', { trigger: 'menu' }, _normalizeAndReverseTimestampsCore);
  }
  return _normalizeAndReverseTimestampsCore();
}

//  
//   for each target tab (MAIN, LABS, Ben Flags, NI / Not Eligible):
//     read the whole timestamp column (Q) in one call
//     for each cell:
//       split its multi-line content into individual timestamp lines
//       normalize each line's format via smartNormalizer
//       build a Date object from each normalized line (for sorting)
//       separate lines into "parseable" (valid date) vs "unparseable" (junk/text)
//       sort parseable lines NEWEST first
//       dedupe: keep only the latest line per (date+hour) bucket
//       rebuild the cell as [deduped, sorted parseable lines] + [unparseable lines]
//     write the whole column back in one setValues call
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
//  
//   1. try to regex-match a "M/D/YYYY H:MM(:SS)? (AM/PM)?" pattern anywhere in the string
//      if found -> convert to 24-hour time, zero-pad, return "M/D/YYYY H:MM:SS"
//   2. else, try parsing the whole string as a native JS Date (fallback for
//      Date-object-style strings like "Wed Apr 01...")
//      if that parses -> reformat to the same "M/D/YYYY H:MM:SS" shape
//   3. else -> give up, return the original string unchanged
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

//   on the active sheet, find every row where col D === "Ben" and auto-resize that row's height
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

//  
//   read saved "next row to process" from script properties (default: row 2)
//   if saved row > total NEWDME rows -> scan already done, clear property, toast, exit
//   compute this chunk's end row = min(startRow + CHUNK - 1, totalRows)
//   run core search scoped to [startRow, endRow], but ONLY vs MAIN and LABS
//   (this is a narrower cross-check than runNewLabsMasterSearch — just NEWDME
//    against the two main tabs, not against every flag/status tab)
//   if more rows remain -> save new "next row" property, toast "run again"
//   else -> clear property, toast "scan complete"
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
const CHUNK = 8800;
const endRow = Math.min(startRow + CHUNK - 1, totalRows);

// 1. Declare targetTabs outside the configuration object so it can be referenced
const targetTabs = ["NEWDME"];

runMasterSearchCore_({
  targetTabs: targetTabs,
  lookupTabs: [
    { name: "MAIN", tab: "MAIN" },
    { name: "LABS", tab: "LABS" }
  ],
  toastSuffix: `${targetTabs[0]} vs MAIN/LABS rows ${startRow}–${endRow} of ${totalRows}`,
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

//   clear the saved resume-point for the MAIN/LABS-vs-NEWDME scan so the next run starts at row 2
function resetNewDmeVsMainLabsScan() {
  PropertiesService.getScriptProperties().deleteProperty('NEWDME_MAINLABS_NEXT_ROW');
  SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME vs MAIN/LABS scan reset. Next run starts from row 2.', 'Reset', 4);
}

//   clear the saved resume-point for the NEWDME full-lookup scan so the next run starts at row 2
function resetNewLabsScan() {
  PropertiesService.getScriptProperties().deleteProperty('NEWDME_NEXT_ROW');
  SpreadsheetApp.getActiveSpreadsheet().toast('NEWDME scan reset. Next run starts from row 2.', 'Reset', 4);
}