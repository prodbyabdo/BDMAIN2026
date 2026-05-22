/**
 * ═══════════════════════════════════════════════════════════════
 * COLUMN CONFIG — BD DME 2026
 * ═══════════════════════════════════════════════════════════════
 *
 * Single source of truth for ALL column indices used across the project.
 * All values are 1-based (as required by Apps Script getRange calls).
 *
 * ⚠️  If a column is added or moved in the spreadsheet, update
 *     THIS FILE ONLY — every script will pick up the change.
 *
 * Column letter reference:
 *   A=1  B=2  C=3  D=4  E=5  F=6  G=7  H=8  I=9  J=10
 *   K=11 L=12 M=13 N=14 O=15 P=16 Q=17
 * ═══════════════════════════════════════════════════════════════
 */

// ── Main data sheet columns (MAIN, LABS, NEWDME, Ben Flags, etc.) ─────────────
const COL_STATUS = 1;   // A — Search result / flag status written by runMasterSearch
const COL_PHONE1_MATCH = 2;   // B — Phone 1 match result written by runMasterSearch
const COL_PHONE2_MATCH = 3;   // C — Phone 2 match result written by runMasterSearch
const COL_OWNER = 4;   // D — Lead owner name (e.g. "Ben", "Jimmy")
const COL_COMMENTS = 6;   // F — Comments cell (auto-date + task trigger)
const COL_SEARCH_START = 8;   // H — First column of the search/lookup block
const COL_SEARCH_WIDTH = 8;   // H–O — Number of columns in the search block (H to O inclusive)
const COL_NPI = 15;  // O — NPI number
const COL_SEND_LEAD = 16;  // P — "Send Lead to" routing dropdown
const COL_TIMESTAMP = 17;  // Q — Timestamp history (written by _handleCommentEdit)

// ── Phone columns ─────────────────────────────────────────────────────────────
const COL_PHONE1 = 10;  // J — Office / primary phone
const COL_PHONE2 = 12;  // L — Cell / alternate phone

// ── Backup and output block (cols A–C) ───────────────────────────────────────
const COL_BACKUP_START = 1;   // A — First column of the pre-search backup range
const COL_BACKUP_WIDTH = 3;   // A–C — Width of the backup range
const COL_OUTPUT_START = 1;   // A — First column written by runMasterSearch
const COL_OUTPUT_WIDTH = 3;   // A–C — Width of the output range

// ── Lookup sheet columns (Ben Flags, Jimmy Flags, etc.) ──────────────────────
// Data in flag tabs starts at COL_SEARCH_START (col H).
// The offsets below are 0-based relative to that start column.
const COL_LOOKUP_START = COL_SEARCH_START; // H (same as main sheet search start)
const COL_LOOKUP_NAME_OFFSET = 0;   // → Col H — business name
const COL_LOOKUP_PH1_OFFSET = 2;   // → Col J — phone 1
const COL_LOOKUP_PH2_OFFSET = 4;   // → Col L — phone 2
const COL_LOOKUP_WIDTH_MAX = 5;   // H–L — max width when all 5 lookup cols exist

// ── IMPORT_DATA sheet (meeting log import) ────────────────────────────────────
const COL_IMPORT_NAME = 15;  // O — Meeting / business name
const COL_IMPORT_PHONE = 16;  // P — Meeting phone number
const COL_IMPORT_WIDTH = 2;   // O–P — width of the range to read

// ── Deactivated sheet ─────────────────────────────────────────────────────────
const COL_DEAC_NPI = 15;  // O — NPI number in the Deactivated tab

// ── Composite helpers used across multiple files ──────────────────────────────

/**
 * Columns that onEditInstallable will act on.
 * Any edit outside these columns is ignored immediately.
 */
const ONEDIT_TRIGGER_COLS = [COL_COMMENTS, COL_SEND_LEAD, COL_TIMESTAMP];

/**
 * Columns scanned by normalizeActiveSheetTimestamps.
 */
const TIMESTAMP_TARGET_COLS = [COL_COMMENTS, COL_TIMESTAMP];
