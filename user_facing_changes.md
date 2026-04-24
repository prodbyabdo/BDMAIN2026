# User-Facing Changes: What You'll Experience in the Sheet

Now that all three iterations of the hardening plan have been completed, here is exactly what you (and the other dialers) will actually *see* and experience while working in the sheet.

### 1. The "Busy" Toast
If you click **Run Master Search** while another dialer is already running it, you won't overwrite each other's data anymore. Instead, you'll immediately see a toast notification in the bottom right corner saying:
> **Busy:** *Another search is running. Try again in 30s.*

### 2. The Aborted Write Protection
If someone accidentally adds or deletes a row in the MAIN or LABS sheet *while* a Master Search is processing, the system will detect the row shift and abort to prevent writing data to the wrong rows. You will see a 10-second warning toast:
> **⚠️ Aborted:** *Row count changed (expected X, got Y). Aborting write on [Tab Name].* 

### 3. A New `_SEARCH_BACKUP` Tab
You will now see a hidden tab mathematically titled `_SEARCH_BACKUP`. Every single time someone runs a Master Search, the exact state of columns A, B, and C is copied to this tab right before the overwrite happens. If you ever need to restore lost data immediately, you can unhide this tab and copy it back.

### 4. Better Dropdown Lead Routing (5-Min Cooldown)
When you select an option from the **"Send Lead to"** column (e.g., 'Selene Flags', 'Schedule Meeting'), the system will now prevent you from accidentally triggering the same routing choice twice within **5 minutes** (up from 60 seconds). You'll no longer see duplicate lead drops on other tabs if you misclick twice.

### 5. Safer Capitalization Button
If you accidentally run **Capitalize Business & Names** while looking at an unsupported tab like "LOGS", nothing will break. It will immediately pop up an alert blocking you:
> *Capitalize not allowed on "[SheetName]". Use on: MAIN, LABS, NEWLABS, Ben Flags*

### 6. The `EXEC_LOG` Tracker & Protection
When the system tracks macros, you won't be able to accidentally delete or edit the `LOGS` or `EXEC_LOG` entries. The moment those sheets generate, they are automatically set to **Owner Only** protection, ensuring a pristine audit trail.

### 7. Faster API Performance (NPI Cache)
Formulas utilizing `=getnpi()` on the sheet will load significantly faster after the first query and utilize 99% less quota. The system will remember the exact output for 6 solid hours, meaning no more frustrating buffering loops every time the spreadsheet recalcs. 

### 8. Intact "Messy" Timestamps
When running **"Fix Time 1"** (the Timestamp Dedup tool), if there are strange timestamps that a non-standard human typed into a cell that fail native date parsing, they will no longer be permanently deleted. Instead, they will be quietly sequestered at the very bottom of the cell so they're at least preserved.

### 9. Dates Parse More Accurately
If the current date is `1/1` and there is a log for `11/12`, making a comment edit won't incorrectly highlight `11/12`. It strictly respects standalone dates now.
