# Attendance Register — User Guide

This Google Apps Script manages learner attendance registers inside Google Sheets. There are two versions:

- **Upskill script** — for upskill courses (80-day schedule, 16 weekly blocks)
- **Short course script** — for short courses (setup and reports not yet implemented)

Both scripts add custom menus to the spreadsheet UI. There is no installation step — paste the script files into the Apps Script editor (Extensions → Apps Script) and the menus appear the next time the spreadsheet is opened.

---

## Upskill script

### Sheets overview

| Sheet | Purpose |
|---|---|
| `SETUP` | Where cohort configuration is entered before running setup |
| `DATABASE` | Internal data store — populated by setup, read at runtime. Hidden after setup |
| `RECORDS` | Weekly attendance blocks — one block per week, 16 weeks total |
| `SUMMARY` | One row per learner — GLH totals, hackathon attendance, status |
| `PARTNER_REPORTS` | Log of all generated report spreadsheets (date, partner, week, URL) |

---

### Initial setup

Before running setup, fill in the `SETUP` sheet with:

- **Student names and email addresses** (first name, last name, email in columns A–C, one row per learner, starting at row 3)
- **Per-student funding partner** (column D, optional): a dropdown populated from the partner list in column N. Defaults to the first entry in N3 when a student name is entered. Leave blank to use the cohort default for all learners — only needed when a cohort has students from more than one funding partner
- **Calendar names** for each session type (column G): Stand Up, Stand Down, Guest Speaker, SME, Career Coach
- **Delivery team details** (columns J–L): name, email, and Google Drive folder URL for the Facilitator, SME, and Career Coach
- **Cohort metadata** (columns N–P): funding partner list (N3:N7 — add one partner per row, up to five), location (e.g. `IE` for Ireland, or any value for England & Wales), and start date

Once the SETUP sheet is complete, open the **Setup** menu and click **Auto Setup Steps**. This runs all five steps in sequence:

1. Creates the working sheets (DATABASE, RECORDS, SUMMARY, PARTNER_REPORTS) from templates
2. Populates DATABASE from the SETUP sheet
3. Builds the 16 weekly attendance blocks in RECORDS
4. Builds the SUMMARY sheet with GLH and hackathon formulas
5. Hides all sheets that end users don't need to see

The alert at the end confirms success and shows how long it took. If you need to run any step individually, use **Setup → Manual Setup Steps**.

---

### Checking attendance

Attendance data is pulled automatically from Google Meet records stored in each team member's Drive folder.

**Auto Attendance → Check Today**
Checks attendance for today's sessions only. Use this at the end of each day.

**Auto Attendance → Check All**
Re-checks attendance across all past dates in the schedule.

**Auto Attendance → Check Between Dates**
Opens a date-picker dialog where you select a start and end date. Only session days within that range are re-checked. The end date defaults to the same day as the start date and can be adjusted before confirming.

**What gets preserved vs overwritten**

All three options follow the same rule: a cell is only preserved if it holds a positive number (assumed to be a manually entered GLH value). Blank cells, zeros, and dash (`-`) values are all overwritten with fresh data from the attendance records.

---

### Changing a learner's status

Learner statuses are managed in the `SUMMARY` sheet, column D. The available values are:

- `Active` — learner is attending
- `Withdrawn` — learner has left the programme
- `Non Starter` — learner never attended

When you edit the status cell in SUMMARY, a confirmation dialog appears asking you to verify or correct the learner's last attended date. After confirming:

- Setting to **Withdrawn** or **Non Starter** writes `X` into all attendance cells after the last attended date in RECORDS.
- Setting back to **Active** writes `-` (not applicable) into all cells from the start of the cohort.

This action affects the RECORDS sheet and is difficult to undo manually — confirm carefully before proceeding.

---

### Adding email addresses to a learner record

A learner may join a Google Meet session using a different email address from the one registered. To ensure their attendance is captured:

1. Enter the additional email address in column C of the SETUP sheet, on the same row as the learner.
2. Open **Setup → Update Database**.

The script anonymises and appends the new email to the learner's existing record. The SETUP email column is cleared automatically once processed, so it is ready for the next batch. Repeat as needed — a learner can have any number of email addresses on record.

---

### Adding a returning learner

A learner who completed a previous cohort and is joining this cohort mid-programme can be added via **Setup → Add Returning Learner**. You will be prompted for:

1. **Full name** (First Last)
2. **Funding partner**
3. **Email address**
4. **URL of their prior cohort spreadsheet**

The script opens the prior cohort spreadsheet, finds the learner by email, and reads their final GLH total and hackathon attendance from the prior SUMMARY sheet. A confirmation dialog shows what will be carried over. After confirming:

- The learner is added to the end of the DATABASE, SUMMARY, and all 16 RECORDS weekly blocks.
- Their prior GLH is added to their running total in SUMMARY automatically — it appears as a carry-over in the GLH column even before they attend any sessions in the current cohort.
- If they completed either hackathon in the prior cohort, the relevant hackathon column in SUMMARY shows `Yes` immediately.

If the prior spreadsheet is inaccessible or the learner cannot be matched by email, you are given the option to add them without prior data.

**Setup → Refresh Prior GLH**
Re-reads all prior cohort spreadsheets and updates the carried-over GLH and hackathon data for all returning learners. Useful if the prior cohort's attendance was updated after the learner was added here. A summary alert reports how many learners were updated and lists any warnings (e.g. inaccessible prior spreadsheet).

---

### Generating partner reports

Partner reports are separate Google Spreadsheets containing only the learners belonging to a given funding partner. They include a filtered copy of SUMMARY and RECORDS and are used for submission to funding bodies.

**Reports → Generate Single Partner Report**
Prompts for a funding partner name. Shows a confirmation dialog listing which learners will be included and which removed. Generates the report spreadsheet and logs it to PARTNER_REPORTS.

**Reports → Generate Multiple Partner Reports**
Automatically generates one report per unique funding partner found in SUMMARY — no confirmation prompts. All reports are logged to PARTNER_REPORTS.

Both options name the generated spreadsheet `Partner Report - <current week> - <cohort name>` and log the date, partner name, current week, and URL to the PARTNER_REPORTS sheet.

> **Note for Excel export:** Reports must be exported to `.xlsx` for partners who cannot open Google Sheets. Use File → Download → Microsoft Excel in the generated report spreadsheet. Signature images are not preserved in Excel export — these must be added manually after download.

---

## Short course script

The short course script adds two menus — **Setup** and **Report** — to the spreadsheet UI.

**Setup → Setup Sheets**
First run creates a SETUP sheet from the template. Fill in the cohort configuration, then run again to build the ATTENDANCE and SUMMARY sheets. SETUP rows 1–13 (col B) hold cohort metadata; rows 14+ hold student first names (col A) and last names (col B).

**Report → Generate Student Report**
Opens a checkbox dialog listing all students. Check the students to include and submit to generate a report spreadsheet containing only their SUMMARY and ATTENDANCE rows.

**Report → Generate Partner Report**
Prompts for a funding partner name, opens the same checkbox dialog with that partner's students pre-checked, then generates a filtered report spreadsheet.

---

## WECA reports

The WECA reports script generates individual Training Completion Declaration Google Docs for each learner from a spreadsheet of GLH data. It adds a single **Custom Menu → Generate Reports** item to the UI.

### Setup

Paste the script into the Apps Script editor of a Google Spreadsheet containing learner GLH data. The active sheet must have the following columns (row 1 = headers, row 2+ = learner data):

| Col | Value |
|---|---|
| A | First name |
| B | Last name |
| C | Total GLH |
| D | % GLH (decimal, e.g. 0.9427) |
| E | Total Employability GLH |
| F | Total Careers GLH |
| G | Guided Behaviours & Workplace Skills GLH |
| H | Pastoral Support Hours |

### Usage

1. Open **Custom Menu → Generate Reports**
2. Paste the URL of the Training Completion Declaration template Google Doc
3. One copy of the template is created per learner row, with all `{{ Placeholder }}` fields substituted with the learner's data
4. Generated docs are saved to the root of My Drive. A success alert confirms completion; any failures are listed by learner name.
