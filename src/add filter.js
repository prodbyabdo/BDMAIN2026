/**
 * Creates "Checklist-style" Filter Views for DME and Labs.
 * EXPANDS multi-word keywords (e.g., "Adult Daycare" -> "Adult", "Daycare")
 * Ensures all hidden values are passed as Strings to avoid API errors.
 */
function createLeadFilterViews() {
  if (typeof runWithExecutionLog_ === 'function') {
    return runWithExecutionLog_('createLeadFilterViews', { trigger: 'menu' }, _createLeadFilterViewsCore);
  }
  return _createLeadFilterViewsCore();
}

function _createLeadFilterViewsCore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("MAIN");
  var labsSheet = ss.getSheetByName("LABS");
  var refSheet = ss.getSheetByName("ReferenceToCancel");

  if (!mainSheet || !labsSheet || !refSheet) {
    SpreadsheetApp.getUi().alert("Error: One of the tabs (MAIN, LABS, or ReferenceToCancel) is missing.");
    return;
  }

  var ssId = ss.getId();
  var mainSheetId = mainSheet.getSheetId();
  var labsSheetId = labsSheet.getSheetId();

  // --- 1. CLEANUP: Delete existing filter views with these names ---
  var spreadsheet = Sheets.Spreadsheets.get(ssId);
  var deleteRequests = [];
  spreadsheet.sheets.forEach(function (sheet) {
    if (sheet.filterViews) {
      sheet.filterViews.forEach(function (fv) {
        if (fv.title === "FILTER: DME ONLY" || fv.title === "FILTER: LABS ONLY") {
          deleteRequests.push({ "deleteFilterView": { "filterId": fv.filterViewId } });
        }
      });
    }
  });
  if (deleteRequests.length > 0) Sheets.Spreadsheets.batchUpdate({ "requests": deleteRequests }, ssId);

  // --- 2. KEYWORD EXPANSION: Split 2-worded cells ---
  function getExpandedKeywords(columnNumber) {
    var lastRow = refSheet.getLastRow();
    if (lastRow < 2) return [];
    var rawValues = refSheet.getRange(2, columnNumber, lastRow - 1, 1).getValues().flat();
    var expanded = [];

    rawValues.forEach(function (val) {
      if (!val) return;
      var str = String(val).trim().toUpperCase();
      expanded.push(str); // Add full phrase

      var parts = str.split(/\s+/);
      if (parts.length > 1) {
        expanded = expanded.concat(parts); // Add individual words (Adult, Daycare, etc)
      }
    });
    return [...new Set(expanded)].filter(Boolean);
  }

  var dmeKeywords = getExpandedKeywords(1); // Col A
  var labKeywords = getExpandedKeywords(2); // Col B

  // --- 3. LOGIC: Determine what to HIDE ---
  function calculateHiddenValues(sheet, keywords) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Get unique values from Col H
    var values = [...new Set(sheet.getRange("H2:H" + lastRow).getValues().flat())];

    return values.filter(function (val) {
      if (val === "" || val === null) return false;
      var cellString = String(val).toUpperCase();

      // Keep row if it matches ANY of the expanded keywords
      var isMatch = keywords.some(function (kw) {
        var regex = new RegExp("\\b" + kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "\\b", "i");
        return regex.test(cellString);
      });

      return !isMatch; // If it's NOT a match, we add it to the 'hidden' list
    }).map(String); // CRITICAL: Force all hidden values to be Strings for the API
  }

  var hideDme = calculateHiddenValues(mainSheet, dmeKeywords);
  var hideLabs = calculateHiddenValues(labsSheet, labKeywords);

  // --- 4. CREATE FILTER VIEWS ---
  var requests = [
    {
      "addFilterView": {
        "filter": {
          "title": "FILTER: DME ONLY",
          "range": { "sheetId": mainSheetId, "startRowIndex": 0, "endRowIndex": mainSheet.getLastRow(), "startColumnIndex": 0, "endColumnIndex": 15 },
          "criteria": { "7": { "hiddenValues": hideDme } } // Col H is Index 7
        }
      }
    },
    {
      "addFilterView": {
        "filter": {
          "title": "FILTER: LABS ONLY",
          "range": { "sheetId": labsSheetId, "startRowIndex": 0, "endRowIndex": labsSheet.getLastRow(), "startColumnIndex": 0, "endColumnIndex": 15 },
          "criteria": { "7": { "hiddenValues": hideLabs } }
        }
      }
    }
  ];

  try {
    Sheets.Spreadsheets.batchUpdate({ "requests": requests }, ssId);
    SpreadsheetApp.getUi().alert("Filters Recreated successfully with split-word logic.");
  } catch (e) {
    Logger.log(e.message);
    SpreadsheetApp.getUi().alert("Error: " + e.message);
  }
}