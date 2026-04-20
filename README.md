# attendance-register-script

A Google Apps Script for managing learner attendance registers in Google Sheets. It adds custom menus to the spreadsheet UI for setup, attendance checking, and generating per-partner report spreadsheets.

---

## Spreadsheet structure

The script expects the following named sheets to exist (or be created during setup):

| Sheet name | Purpose |
|---|---|
| `SUMMARY` | One row per learner: first name, last name, funding partner, and additional metadata (e.g. current week in column 13) |
| `RECORDS` | Attendance data laid out in repeating weekly blocks (16 weeks, each block starting at row `21 + weekIndex * (numStudents + 7)`) |
| `PARTNER_REPORTS` | Log of all generated partner report spreadsheets (date, partner name, week, URL) |
| `TEMPLATE_PARTNER_REPORTS` | Template sheet copied to create `PARTNER_REPORTS` if it does not yet exist |
| Various `TEMPLATE_*` sheets | Used by the setup functions to create the working sheets |

---

## Menus

Three custom menus are added to the spreadsheet UI when it is opened.

### Setup

| Menu item | Function | Description |
|---|---|---|
| Create New Setup Sheet | `createNewSetupSheet` | Creates a fresh setup configuration sheet |
| Manual Setup Steps > Step 1 | `createSheetsFromTemplates` | Copies template sheets into the spreadsheet |
| Manual Setup Steps > Step 2 | `setupDatabase` | Populates the database sheet |
| Manual Setup Steps > Step 3 | `setupRecords` | Builds the attendance records structure |
| Manual Setup Steps > Step 4 | `setupSummary` | Builds the summary sheet |
| Manual Setup Steps > Step 5 | `hideAllUnusedSheets` | Hides sheets not needed by end users |
| Auto Setup Steps | `setupEverything` | Runs all five steps above in sequence |
| Update Database | `updateDatabase` | Updates the database after initial setup |

### Auto Attendance

| Menu item | Function | Description |
|---|---|---|
| Check Today | `checkAttendanceToday` | Checks attendance entries for today |
| Check All | `checkAllAttendance` | Re-checks attendance entries across all dates |

### Reports

| Menu item | Function | Description |
|---|---|---|
| Generate Single Partner Report | `generatePartnerReport` | Prompts for a partner name and generates one report |
| Generate Multiple Partner Reports | `autoGeneratePartnerReports` | Generates a report for every unique partner automatically |

---

## Functions

### `onOpen()`
Runs automatically when the spreadsheet is opened. Registers the three custom menus described above.

### `setupEverything()`
Runs all five setup steps in sequence and alerts the user when complete, including the elapsed time.

### `generatePartnerReport()`
Prompts the user for a funding partner name, then:
1. Reads all learner rows from `SUMMARY` and splits them into matching (kept) and non-matching (removed) groups.
2. Shows a confirmation dialog listing which learners will be kept and removed.
3. Creates a new Google Spreadsheet named `Partner Report - <week> - <spreadsheet name>`.
4. Copies `SUMMARY` and `RECORDS` into the new spreadsheet, removes non-partner learner rows, and compacts the remaining rows upward.
5. Appends a log entry (date, partner name, current week, URL) to `PARTNER_REPORTS`.
6. Navigates the user to the `PARTNER_REPORTS` sheet and shows a success alert.

### `autoGeneratePartnerReports()`
Same report generation logic as `generatePartnerReport()`, but runs automatically for every unique funding partner found in the `SUMMARY` sheet — no user prompts. Logs total elapsed time to the Apps Script logger when finished.

---

## Known issues

- **Row calculation bug** — in both report functions, the RECORDS row deletion logic uses absolute `SUMMARY` sheet row numbers as offsets into each weekly block, which is likely incorrect. The offset should be the student's 0-based index within the block, not the absolute row number.
- **Duplicated logic** — `generatePartnerReport` and `autoGeneratePartnerReports` share ~80 lines of identical report-building code. This should be extracted into a shared helper function.
- **Loop-invariant work inside the partner loop** — in `autoGeneratePartnerReports`, the current week is fetched and the `PARTNER_REPORTS` sheet is resolved on every iteration when both should happen once before the loop.
- **Debug logging left in** — several `Logger.log` calls in `autoGeneratePartnerReports` were left in from development.

---

## Dependencies

This script relies on several helper functions defined elsewhere in the project (not yet included in this file):

- `getNumOfStudents()` — returns the total number of learner rows in the `SUMMARY` sheet
- `formatDate(dateString)` — formats a `YYYY-MM-DD` string for display in `PARTNER_REPORTS`
- `createSheetsFromTemplates()`, `setupDatabase()`, `setupRecords()`, `setupSummary()`, `hideAllUnusedSheets()` — the five setup step functions
- `checkAttendanceToday()`, `checkAllAttendance()` — attendance checking functions
- `updateDatabase()`, `createNewSetupSheet()` — database and setup utilities
