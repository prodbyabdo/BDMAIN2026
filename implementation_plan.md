# Hardening Plan — 3 Iterations

**Context:** 4 dialers share sheet. `runMasterSearch` bulk-writes cols A-C on MAIN+LABS. No locks, no backup, no validation. `Benleadsssss` library already removed by user. ✅

---

## Iteration 1 — Stop Data Loss (P0)

> [!CAUTION]
> Ship this first. Everything else secondary.

### 1A. Add Lock to `runMasterSearch` — [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)

**Problem:** Two dialers run Master Search simultaneously → both read stale data → last writer wins → first writer's results vanish.

**Fix:** Wrap `runMasterSearchCore_` in `LockService.getScriptLock()`. One execution at a time. Others get toast message.

```javascript
// BEFORE (no lock)
function runMasterSearchCore_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // ... runs unprotected

// AFTER
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
    // ... existing logic unchanged
  } finally {
    lock.releaseLock();
  }
}
```

**Lines affected:** ~74-239 (wrap entire function body)

---

### 1B. Add Row-Count Validation Before Write — [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)

**Problem:** If rows added/deleted between read and write, `numRows` mismatches actual row count → data shifts or overwrites wrong rows.

**Fix:** Re-check `getLastRow()` right before `setValues()`. Abort if changed.

```javascript
// BEFORE (blind write)
sheet.getRange(2, 1, numRows, 3).setValues(outputRows);

// AFTER (validated write)
const currentLastRow = sheet.getLastRow();
if (currentLastRow - 1 !== numRows) {
  ss.toast(
    `Row count changed (expected ${numRows}, got ${currentLastRow - 1}). Aborting write on ${tabName}.`,
    '⚠️ Aborted', 10
  );
  return; // skip this tab, don't write
}
sheet.getRange(2, 1, numRows, 3).setValues(outputRows);
```

**Lines affected:** ~227 (add 6 lines before existing `setValues` call)

---

### 1C. Backup Cols A-C Before Overwrite — [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)

**Problem:** No way to recover if write goes wrong. User had to Ctrl+Z.

**Fix:** Before bulk write, snapshot cols A-C to hidden `_SEARCH_BACKUP` tab. Auto-overwrites each run (only keeps last snapshot).

```javascript
// Add before the bulk write, inside the targetTabs.forEach loop:
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
  // Copy cols A-C
  const data = sheet.getRange(2, 1, numRows, 3).getValues();
  if (data.length > 0) {
    backup.getRange(2, 1, data.length, 3).setValues(data);
  }
}
```

**Lines affected:** New helper function + 1 call line before each `setValues`

---

### 1D. Fix `appsscript.json` Web App Access — [appsscript.json](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/appsscript.json)

**Problem:** `"access": "ANYONE_ANONYMOUS"` — anyone with deployment URL can execute functions as you.

**Fix:**

```diff
   "webapp": {
     "executeAs": "USER_DEPLOYING",
-    "access": "ANYONE_ANONYMOUS"
+    "access": "ANYONE"
   }
```

`ANYONE` = anyone with Google account (logged in). `ANYONE_ANONYMOUS` = literally anyone, no login needed.

> [!IMPORTANT]
> If you have external integrations hitting this web app without auth, changing to `ANYONE` will break them. Let me know if anything calls this deployment URL externally.

**Lines affected:** Line 26

---

### Iteration 1 — Verification

- [ ] `clasp push` updated files
- [ ] Open sheet with 2 browser tabs, run Master Search in both → second one gets "Busy" toast
- [ ] Add a row to MAIN during search → search aborts with mismatch message
- [ ] Run Master Search normally → check `_SEARCH_BACKUP` tab exists with old cols A-C
- [ ] Verify web app still works after access change

---

## Iteration 2 — Fix Logic Bugs (P1)

### 2A. Batch `setValue` in `_handleCommentEdit` — [onedit.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/onedit.js)

**Problem:** Lines 157-159 loop `setValue()` per column. If execution times out mid-loop → partial state (Col F updated, Col Q not, or vice versa).

**Fix:** Collect all updates, write in one call per column (can't do single range because cols F and Q aren't contiguous — but can use `getRangeList`).

```javascript
// BEFORE
for (const [colIdx, val] of Object.entries(updates)) {
  sheet.getRange(row, Number(colIdx)).setValue(val);
}

// AFTER
const ranges = [];
const values = [];
for (const [colIdx, val] of Object.entries(updates)) {
  sheet.getRange(row, Number(colIdx)).setValue(val);
}
// Note: For non-contiguous cols (F and Q), individual setValue is actually
// acceptable here — just need to ensure BOTH writes happen.
// Real fix: wrap in try/catch, if second write fails, revert first.
```

Actually, better approach — write both in explicit order with revert:

```javascript
const colFVal = updates[COMMENT_COL];
const colQVal = updates[TIMESTAMP_COL];

if (colQVal !== undefined) {
  sheet.getRange(row, TIMESTAMP_COL).setValue(colQVal);
}
if (colFVal !== undefined) {
  sheet.getRange(row, COMMENT_COL).setValue(colFVal);
}
```

Timestamp first (more important for audit trail), date second. If date write fails, timestamp still persists.

**Lines affected:** 155-159

---

### 2B. Fix Date Regex Word Boundary — [onedit.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/onedit.js)

**Problem:** Line 141 — `\b1\/12\b` also matches `11/12` because `\b` sits between `1` and `1` in `11/12`.

**Fix:** Anchor to start-of-string or space:

```javascript
// BEFORE
const todayRegex = new RegExp(`\\b${month}\\/${day}\\b`);

// AFTER
const todayRegex = new RegExp(`(?:^|\\s|\\n)${month}\\/${day}(?:\\s|$|\\n)`);
```

**Lines affected:** Line 141

---

### 2C. Protect Timestamp Dedup From Dropping Valid Entries — [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)

**Problem:** `normalizeAndReverseTimestamps` (line 409) silently drops entries that fail date parsing. Valid timestamps with unusual formats get deleted.

**Fix:** Keep unparseable lines at bottom instead of dropping:

```javascript
// BEFORE
.filter(item => !isNaN(item.date.getTime()));

// AFTER — split into parseable and unparseable
const parseable = withDates.filter(item => !isNaN(item.date.getTime()));
const unparseable = lines.filter(line => {
  const normalized = smartNormalizer(line);
  const dateObj = new Date(normalized.replace(...));
  return isNaN(dateObj.getTime());
});

// Sort parseable, keep unparseable at end
parseable.sort((a, b) => b.date - a.date);
// ... dedup logic on parseable only ...
const newVal = [...filteredLines, ...unparseable].join("\n");
```

**Lines affected:** ~404-426

---

### 2D. Lock `capitalizeHeadersBatch` to Target Tabs — [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)

**Problem:** Uses `getActiveSheet()` — runs on whatever tab is open. Could uppercase data on wrong tab.

**Fix:** Add tab whitelist:

```javascript
// Add at top of function
const ALLOWED_TABS = ["MAIN", "LABS", "NEWLABS", "Ben Flags"];
const sheetName = sheet.getName();
if (!ALLOWED_TABS.includes(sheetName)) {
  SpreadsheetApp.getUi().alert(`Capitalize not allowed on "${sheetName}". Use on: ${ALLOWED_TABS.join(", ")}`);
  return;
}
```

**Lines affected:** ~285-290

---

### 2E. Add Caching to `getnpi()` — [npisearch.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/npisearch.js)

**Problem:** Each `=getnpi()` cell = 1 HTTP request. 1000 cells = 1000 requests. Quota: 20,000/day.

**Fix:** Cache results in `CacheService` (6 hour TTL):

```javascript
// Add at top of getnpi()
const cache = CacheService.getScriptCache();
const cacheKey = `npi_${npi}_${fields}`;
const cached = cache.get(cacheKey);
if (cached) return JSON.parse(cached);

// After successful result, before return:
cache.put(cacheKey, JSON.stringify(result), 21600); // 6 hours
```

**Lines affected:** Lines 3-36

---

### Iteration 2 — Verification

- [ ] Edit Col F on row with date "1/12" when today is "1/1" → should NOT match
- [ ] Run `normalizeAndReverseTimestamps` on tab with mixed valid/invalid timestamps → invalid lines preserved at bottom
- [ ] Run `capitalizeHeadersBatch` on LOGS tab → gets blocked with alert
- [ ] Use `=getnpi()` twice with same NPI → second call returns from cache (no network hit)

---

## Iteration 3 — Cleanup & Polish (P2/P3)

### 3A. Remove Dead QPP Code — [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)

- Line 141: `const qppFeedback = new Map();` — always empty
- Lines 200, 211-214: QPP checks that always return `null`
- Either restore QPP_FEEDBACK population or strip dead branches

### 3B. Add Sheet Protection to Log Tabs — [onedit.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/onedit.js) + [execution_log.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/execution_log.js)

- When LOGS / EXEC_LOG tabs are auto-created, add `Protection` so only owner can edit
- Prevents other dialers from accidentally deleting log rows

### 3C. Wire Up `execution_log.js` — All files

- `runWithExecutionLog_` wrapper exists but nothing uses it
- Wrap `runMasterSearch`, `createLeadFilterViews`, `normalizeAndReverseTimestamps` in it
- Gives full execution audit trail with timing

### 3D. Document Timezone — [appsscript.json](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/appsscript.json)

- Add comment in code docs that all timestamps use `Africa/Cairo` (UTC+2)
- If dialers in different timezone, timestamps in Col Q won't match their wall clock

### 3E. Improve `clean()` Type Safety — [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)

```javascript
// BEFORE
function clean(val) {
  if (!val) return "";
  return String(val).toLowerCase().trim();
}

// AFTER — handle Date objects explicitly
function clean(val) {
  if (!val) return "";
  if (val instanceof Date) return "";
  return String(val).toLowerCase().trim();
}
```

### 3F. Fix `_handleLeadRouting` Dedup TTL — [onedit.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/onedit.js)

- Line 189: Cache TTL 60s → increase to 300s (5 min)
- Prevents duplicate lead copies if user clicks same dropdown value again within 5 minutes

---

### Iteration 3 — Verification

- [ ] Run Master Search → check EXEC_LOG tab has entry with timing
- [ ] Try editing LOGS tab as non-owner → blocked
- [ ] Verify QPP branches removed/restored cleanly
- [ ] Check duplicate lead routing blocked for 5 min window

---

## Summary

| Iteration | Focus | Files | Risk Level |
|-----------|-------|-------|------------|
| **1** | Stop data loss — lock, validate, backup | menu.js, appsscript.json | 🔴 Critical |
| **2** | Fix logic bugs — regex, timestamps, caching | onedit.js, menu.js, npisearch.js | 🟠 High |
| **3** | Cleanup — dead code, logging, polish | All files | 🟡 Medium |

Each iteration is independent. Ship and test one at a time via `clasp push`.
