# Codebase & Worksheet Index — BD DME 2026

This index maps the structural components of the BD DME 2026 spreadsheet automation system. It catalogues worksheets (tabs), columns, triggers, functions, and external links.

---

## 📊 Worksheet (Tab) Catalog

The spreadsheet workbook contains multiple tabs serving specific structural, operational, or logging functions:

### 1. Active Lead Directories
* **`MAIN`**: Primary lead directory for DME cold calling. Scanned by Master Search.
* **`LABS`**: Primary lead directory for Clinical Medical Laboratory cold calling. Scanned by Master Search.
* **`NEWDME`**: Incoming DME leads. Target of the `NEWLABS` Master Search.
* **`NEWLABS`**: Target sheet containing new incoming laboratory leads.

### 2. Owner Lead Tabs (Internal Flags)
When an owner is selected for a lead, or lead routing dropdown values are selected, the lead details are appended to these tabs:
* **`Ben Flags`**: Flags assigned to Ben.
* **`Jimmy Flags`**: Flags assigned to Jimmy.
* **`Selene Flags`**: Flags assigned to Selene.
* **`Jane Flags`**: Flags assigned to Jane.

### 3. Exclusions & Status Directories
* **`NI / Not Eligible`**: Leads qualified as Not Interested (NI) or medically ineligible.
* **`Disconnected`**: Phone numbers flagged as disconnected or invalid ("Dis/Wn"). Used to filter out dead numbers during searches.
* **`DMEDesk Booked`**: Leads successfully routed as booked appointments.
* **`Deactivated`**: Holds deactivated NPI numbers (checked in Column O) to exclude dead healthcare provider records.

### 4. System Control & Logging Sheets
* **`ReferenceToCancel`**: Holds filter view keywords for DME (Col A) and LABS (Col B) to build dynamic filters.
* **`IMPORT_DATA`**: Reference tab containing imported third-party meeting logs (names/phones) for cross-referencing.
* **`Verified_Valid`**: Local cache sheet of verified, active phone numbers to minimize outbound API calls.
* **`_SEARCH_BACKUP`**: Hidden worksheet containing the snapshotted pre-search state of Columns A-C.
* **`LOGS`**: Automatically generated log sheet recording comments history edits and script execution statuses.
* **`EXEC_LOG`**: Automatically generated and locked audit sheet recording function durations, triggers, and execution IDs.
* **`SCRIPT_CATALOG`**: Tabulated summary of code files, functions, triggers, and execution weights.

---

## 🗂️ Global Column Configuration (`src/config.js`)

All column references are 1-based indices (matching standard Google Apps Script range references):

| Index | Column Letter | Config Variable Name | Purpose / Contents |
|:---:|:---:|:---|:---|
| **1** | **A** | `COL_STATUS` | Master Search result state / matching Flag tab name |
| **2** | **B** | `COL_PHONE1_MATCH` | Matches detected on Phone 1 (Column J) |
| **3** | **C** | `COL_PHONE2_MATCH` | Matches detected on Phone 2 (Column L) |
| **4** | **D** | `COL_OWNER` | Lead owner (e.g., "Ben", "Jimmy") |
| **6** | **F** | `COL_COMMENTS` | Comments cell (triggers date stamps & tasks) |
| **8** | **H** | `COL_SEARCH_START` | First column of lookup block (Business Name) |
| **10** | **J** | `COL_PHONE1` | Primary Office Phone number |
| **12** | **L** | `COL_PHONE2` | Cell / Alternate Phone number |
| **15** | **O** | `COL_NPI` | 10-digit National Provider Identifier (NPI) |
| **16** | **P** | `COL_SEND_LEAD` | Dropdown trigger for lead routing paths |
| **17** | **Q** | `COL_TIMESTAMP` | Log record of chronological comments history |
| **19** | **S** | `COL_VERIFY_CHECKBOX` | Checkbox to request phone verification |

---

## 🛠️ Code Module Directory

### 📄 `menu.js`
* **`onOpen()`**: Registers the custom menu **Custom Tools** in the Sheets UI.
* **`runMasterSearch()`**: Coordinates bulk scanning for `MAIN` and `LABS` sheets.
* **`runNewLabsMasterSearch()`**: Coordinates bulk scanning for `NEWDME` and `NEWLABS` sheets.
* **`runMasterSearchCore_(options)`**: Lock-guarded scanning engine that performs O(1) set-based match computations.
* **`capitalizeHeadersBatch()`**: Uppercases names on whitelisted active tabs (`Legalbusinessname` / `AuthOfficialName`).
* **`reformatPhoneNumbers()`**: Standardizes phone number grids to `XXX-XXX-XXXX` format.
* **`unmergeAllCells()`**: Safe utility to break apart merged cells.
* **`normalizeActiveSheetTimestamps()`**: Standardizes formatting of all timestamps on the active tab.
* **`normalizeAndReverseTimestamps()`**: Deduplicates timestamps to 1 per hour and sorts them newest-first.
* **`smartNormalizer(input)`**: Advanced timestamp regex standardization parser.

### 📄 `onedit.js`
* **`onEditInstallable(e)`**: Single entry hook for Sheet edits, routing events based on column triggers.
* **`_handleCommentEdit(e, sheet, sheetName, col, row)`**: Lock-secured handler appending date-stamps and recording timestamps.
* **`_handleLeadRouting(e, sheet, sheetName, row)`**: Cooldown-protected router copying rows to flag tabs or external Sheets.
* **`logAction(status, sheetName, row, message, user)`**: Writes general log records to the `LOGS` sheet.

### 📄 `phone_verification.js`
* **`verifyPhoneStacked(phoneNumber)`**: Custom function executing stacked API validation (Antideo → APITier → IPQS).
* **`verifySelectedPhones()`**: Evaluates Checked checkboxes in Column S, queries verification APIs, writes "Dis/Wn" markers, caches responses, and unchecks checkboxes.

### 📄 `add filter.js`
* **`createLeadFilterViews()`**: Recreates DME and LABS filter views by parsing and splitting keywords from `ReferenceToCancel`.

### 📄 `execution_log.js`
* **`runWithExecutionLog_(scriptName, context, fn)`**: Wrapper that runs tasks and logs structural runtime metadata.
* **`beginExecutionLog_(...)` / `endExecutionLog_(...)`**: Generates and updates log rows inside the protected `EXEC_LOG` worksheet.

### 📄 `commentTask.js`
* **`processCommentTask(e)`**: Creates a Google Task on a comment containing `@task(...)` for user `Ben` (currently paused).

### 📄 `npisearch.js`
* **`bgetnpi(range, fields)`**: Custom spreadsheet function returning batch NPI information with 6-hour caching.

### 📄 `DailyTriggersAppscript.js`
* **`setupDailyTriggers()`**: Installs clean daily project summaries triggers.

---

## 🔗 External Integration Points

The system interacts dynamically with two primary external spreadsheets, configured via global IDs:

1. **Meeting Log Spreadsheet**
   * **ID:** `1uicpBruuFeno2ES4hNw-TIAwNkGEI37gw8Z-A4yMpC8`
   * **Tab:** `New Meetings`
   * **Trigger:** Selecting 'Schedule Meeting' in Column P dropdown.
2. **Email Automation Spreadsheet**
   * **ID:** `10tWkyiVrYYgVtkguYuw1NLBhO3k041_RoykCd_mXuso`
   * **Tab:** `Sheet1`
   * **Trigger:** Selecting 'Send Lab Email' in Column P dropdown.
