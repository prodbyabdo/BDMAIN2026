# Redesign CLIA Matching Logic

The recent integration of CLIA matching in `runMasterSearchCore_` uses an `OR` condition between `Legal Name`, `OfficePhone`, and `AuthPhone`. While downloading and analyzing the latest `BD DME 2026.xlsx` Google Sheet data, I discovered a crucial data constraint:

1. **Shared Phone Numbers**: Many clinics, labs, and providers (especially in major hospital networks) share "mainline" phone numbers. 
2. **False Positives**: When the CLIA dataset flags one of these mainline phone numbers, the script then applies the `CLIA` tag to **every lead** in the `LABS` and `MAIN` tabs that shares that generic phone number, regardless of whether it's actually the registered lab. (Our testing script showed 81 leads flagged via phone match vs only 24 via exact name match).

## Proposed Changes

To fix this issue without fully abandoning phone-based searches, we need a stricter confidence threshold for CLIA tag application.

### `src/menu.js`
We will rewrite the `isClia` boolean logic to avoid the false positive cascade. Here are three potential approaches depending on your preference:

**Option 1: Name-Only Matching (Strictest & Safest)**
- [MODIFY] `src/menu.js`
  - Remove `cliaPhoneSet` entirely.
  - Rely exclusively on an exact `clean(termH)` match against `cliaNameSet`. This guarantees no cross-pollination from shared hospital networks.

**Option 2: Require Dual Matches (High Confidence)**
- [MODIFY] `src/menu.js`
  - Change the OR condition to an AND condition for phone numbers: A lead is flagged as CLIA only if `cliaNameSet.has(termH)` **OR** _(it matches both name AND phone)_.

**Option 3: Phone Matching with Exclusion for Duplicated Phones**
- [MODIFY] `src/menu.js`
  - Disable CLIA phone matching when the search encounters a heavily duplicated phone number in `LABS`.

## User Review Required

> [!IMPORTANT]
> **Which approach do you prefer for fixing the duplicate phone issue?**
> - **Option 1** is highly recommended: It avoids the shared phone number trap entirely and is the simplest and fastest robust solution for O(1) matching.
> - Please let me know how aggressively you want to match CLIA records.

## Verification Plan

### Automated Tests
- I will run a local Python script simulating the newly selected `isClia` logic against the newly synced `BD DME 2026.xlsx` to confirm that the number of `LABS` leads flagged matches exactly the expected subset.

### Manual Verification
- After updating and pushing `menu.js` via clasp, you can run `Run NEWLABS Master Search` or `Run Master Search` directly in the spreadsheet and verify that `CLIA` no longer inappropriately flags groups of non-lab leads.
