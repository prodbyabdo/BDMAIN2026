# Implementation Plan: Pull Sheet & Auth Official Name Duplicate Checker

## Overview
1. **Pull Code & Sheet Data**:
   - Apps Script code was successfully pulled via `clasp pull` (all 12 files synchronized, including latest changes in `src/menu.js`).
   - The local sheet export (`sync-sheet.js`) encountered a stale Google Drive OAuth token (`invalid_grant`), which cleared `auth/token.json`. Re-authenticating requires generating a fresh token through Google's OAuth consent flow.
2. **Authorized Official Name (`AuthOfficialName`) Duplicate Checker**:
   - Currently, the Master Search engine (`runMasterSearchCore_` in [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)) checks:
     - **Col H** (`Legalbusinessname`) &rarr; output to **Col A** (`NAME DUP`)
     - **Col J** (`OfficePhone`) &rarr; output to **Col B** (`NUM DUP`)
     - **Col L** (`AuthPhone`) &rarr; output to **Col C** (`Auth Num`)
   - **Col K** (`AuthOfficialName`) is completely unindexed and unchecked across all sheets.
   - We need to establish the exact architecture for checking `AuthOfficialName` duplicates across sheets (`MAIN`, `LABS`, `NEWDME`, `NEWLABS`, `Ben Flags`, `Jimmy Flags`, etc.).

---

## User Review Required

> [!IMPORTANT]
> **Sheet Sync Authentication**:
> To pull the latest live spreadsheet (`data/BD DME 2026.xlsx`), run `sync-sheet.js` in a PowerShell terminal or authorize via the OAuth URL when launched. Since `auth/token.json` was purged due to expiration, visiting the Google OAuth URL is required to obtain a new refresh token.

> [!IMPORTANT]
> **Design Decision on Duplicate Output Location**:
> The workbook layout currently has 3 output columns:
> - Col A: `NAME DUP`
> - Col B: `NUM DUP`
> - Col C: `Auth Num`
> - Col D: `Owner`
> - Col E: `Email`
> - Col F: `Comments` (critical trigger column for auto-dating)
> 
> How would you like the `AuthOfficialName` duplicate results displayed?
> 
> - **Option 1 (Merged into Col A `NAME DUP`)**: Col A checks both `Legalbusinessname` and `AuthOfficialName`. If `AuthOfficialName` matches, it labels the match in Col A (e.g. `Ben (Auth)` or `Ben`). **Zero column shifts** required across existing sheets.
> - **Option 2 (Dedicated Column - e.g., Insert New Col D `Auth Name DUP`)**: Inserts a new column specifically for Authorized Official Name matches. *Note: Shifting columns from D onwards requires updating `src/config.js` and all column offset references.*
> - **Option 3 (Separate Audit / Checker Utility)**: A dedicated menu item (e.g., *Check Auth Name Duplicates*) that scans and highlights or writes results to an audit tab or reports all cross-sheet matches.

---

## Open Questions

> [!WARNING]
> 1. **Matching Scope**: Should `AuthOfficialName` be checked:
>    - **Only against `AuthOfficialName`** in other tabs? (i.e. person name to person name)
>    - **Or against both `AuthOfficialName` and `Legalbusinessname`**?
> 2. **Cross-Tab Target Scope**:
>    - Which tabs should be checked as **targets** (e.g. `NEWDME`, `MAIN`, `LABS`, `NEWLABS`)?
>    - Which tabs should serve as **lookup sources** (all flag tabs, `MAIN`, `LABS`, `DNC`, etc.)?
> 3. **Output Destination**: Do you prefer **Option 1** (tagging inside Col A without shifting sheet columns) or **Option 2** (creating a dedicated `AUTH NAME DUP` column)?

---

## Proposed Changes

### Configuration Updates

#### [MODIFY] [config.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/config.js)
- Add constant `COL_AUTH_NAME = 11;` (Col K: `AuthOfficialName`).
- Add `COL_LOOKUP_AUTH_NAME_OFFSET = 3;` (Offset from Col H in the `H:L` lookup range).
- If a new column is added, adjust subsequent column constants (`COL_OWNER`, `COL_COMMENTS`, `COL_NPI`, `COL_SEND_LEAD`, `COL_TIMESTAMP`, etc.) accordingly.

---

### Search Engine Updates

#### [MODIFY] [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)
- In `runMasterSearchCore_`:
  1. Expand lookup reading from `H2:L` to read `AuthOfficialName` (Col K, index offset 3).
  2. Maintain an `authNameMap` (or integrate into `nameMap` with an `(Auth)` tag).
  3. In the per-row scan loop, extract `termAuthName = clean(targetValues[i][COL_AUTH_NAME - 1])`.
  4. Query the index for `termAuthName` matches.
  5. Include `termAuthName` in the row's `cacheKey` to ensure duplicate cache accuracy.
  6. Output the match result to the chosen destination (Col A or designated column).

---

### Sync & Automation Verification

#### [MODIFY] [sync-sheet.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/sync-sheet.js)
- Ensure clean error handling when starting the OAuth callback listener.
- Verify download to `data/BD DME 2026.xlsx`.

---

## Verification Plan

### Automated / Syntax Verification
- Run local node syntax checks:
  ```powershell
  & "C:\Users\ben.arthur\node-v24.14.1-win-x64\node.exe" -c "src/config.js"
  & "C:\Users\ben.arthur\node-v24.14.1-win-x64\node.exe" -c "src/menu.js"
  ```
- Run python data inspection script to verify column offsets against `data/BD DME 2026.xlsx`.

### Clasp Deployment
- Push code to Apps Script project via clasp:
  ```powershell
  & "C:\Users\ben.arthur\node-v24.14.1-win-x64\node.exe" `
    "C:\Users\ben.arthur\node-v24.14.1-win-x64\node_modules\@google\clasp\build\src\index.js" push
  ```

### Manual Verification
- Ask user to run the menu action (e.g. *Run Main Search* or *Run NEWLEADS Search*) in Google Sheets and verify that duplicates on Authorized Official Name are properly flagged.
