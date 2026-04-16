# Implementation Plan - Universal Execution Logging (2 Iterations)

This plan implements the `execution_log.js` system across all major functions. To ensure stability, the deployment is divided into two iterations with manual verification steps.

## User Review Required

> [!IMPORTANT]
> **Performance Optimization**: The logger will be optimized to be memory-resident until the final write, ensuring it doesn't slow down the spreadsheet.
> 
> **Data Migration**: The old `LOGS` tab will be overwritten/replaced by the new `EXEC_LOG` system as requested.

---

## Iteration 1: Core Logger & Menu Scripts
**Goal**: Optimize the engine and track manual search/utility scripts.

### [Component] Execution Logger Optimization
#### [MODIFY] [execution_log.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/execution_log.js)
- Refactor `endExecutionLog_` to remove `getValue()` calls.
- Store initial capture data (ID, name, user) in the `entry` object and pass it back to the end-log call.

### [Component] Menu Functions
#### [MODIFY] [menu.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/menu.js)
- Wrap `runMasterSearch` and `runNewLabsMasterSearch`.
- Wrap utility routines: `capitalizeHeadersBatch`, `reformatPhoneNumbers`, `unmergeAllCells`.

### Manual Check 1
- **Action**: I will ask you to run "Capitalize Business & Names" from the menu.
- **Verification**: We will confirm a `SUCCESS` entry appears in the new `EXEC_LOG` sheet with the correct duration and user email.

---

## Iteration 2: Trigger Consolidation
**Goal**: Replace legacy logging in the high-frequency `onEdit` handler.

### [Component] Triggers
#### [MODIFY] [onedit.js](file:///c:/Users/ben.arthur/Desktop/BD%20MAIN%202026/src/onedit.js)
- Wrap `onEditInstallable` in the logger.
- Remove the legacy `logAction` function definition.
- Update `_handleCommentEdit` to use `logExecutionEvent_` for specific status logs (COOLDOWN, LOCK_TIMEOUT).

### Manual Check 2
- **Action**: I will ask you to make a comment edit in Column F.
- **Verification**: Confirm that the `onEdit` event is logged in `EXEC_LOG` and that the auto-dating/timestamping still functions correctly with the added overhead.

---

## Verification Plan
### Manual Verification
1.  **Iter 1 Run**: Check menu execution in `EXEC_LOG`.
2.  **Iter 2 Run**: Check cell edit execution in `EXEC_LOG`.

