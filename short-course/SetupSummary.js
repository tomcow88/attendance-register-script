/**
 * Populates the SUMMARY sheet with cohort metadata, GLH formulas, hackathon/project dates,
 * delivery team details, and per-student rows.
 *
 * Steps:
 *  1. Sets the current week display and cohort start/end dates.
 *  2. Writes summary totals (student count, active count, retention rate) as formulas.
 *  3. Writes GLH target, threshold, and cumulative total formulas.
 *  4. Writes hackathon and project date ranges.
 *  5. Writes delivery team names, emails, and session count formulas.
 *  6. Copies the first student row template down for all students, then populates names.
 */
function setupSummary() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");

    const currentWeekData = setCurrentWeek(false);
    const startDate = getStartOrEndDate("start", "date");
    const endDate = getStartOrEndDate("end", "date");

    summarySheet.getRange(5, 11, 1, 1).setValue(formatDate(startDate));
    summarySheet.getRange(5, 12, 1, 1).setValue(formatDate(endDate));

    const numOfStudents = databaseSheet.getRange(3, 22, 1, 1).getValue();

    // Cohort overview formulas: total learners, active count, and retention rate.
    summarySheet.getRange(8, 11, 1, 1).setValue("=COUNTA(A:A) - 1");
    summarySheet.getRange(8, 12, 1, 1).setValue('=COUNTIF(INDIRECT("D2:D" & $K$8 + 1), "Active")');
    summarySheet.getRange(8, 13, 1, 1).setValue('=(1 - (COUNTIF(INDIRECT("D2:D" & $K$8 + 1), "Withdrawn") / $K$8))');

    // GLH targets: average GLH per active learner, 80% threshold, contracted total, and cumulative sum across all weeks.
    summarySheet.getRange(11, 11, 1, 1).setValue(
        `=SUMIF(INDIRECT("D2:D" & $K$8 + 1), "Active", INDIRECT("G2:G" & $K$8 + 1)) / COUNTIF(INDIRECT("D2:D" & $K$8 + 1), "Active")`,
    );
    summarySheet.getRange(11, 12, 1, 1).setValue("=M11 * 0.8");
    summarySheet.getRange(11, 13, 1, 1).setValue("308");
    // Sums column 11 of each weekly block in RECORDS across all 16 weeks.
    summarySheet.getRange(11, 14, 1, 1).setValue(
        `=SUM(INDIRECT("'RECORDS'!" & ADDRESS(21 + (0 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (1 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (2 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (3 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (4 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (5 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (6 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (7 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (8 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (9 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (10 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (11 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (12 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (13 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (14 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (15 * ($K$8 + 7)), 11)))`,
    );

    // Hackathon and project date ranges read from DATABASE and displayed in SUMMARY.
    const proj1StartDate = databaseSheet.getRange(7, 21, 1, 1).getValue();
    const proj1EndDate = databaseSheet.getRange(8, 21, 1, 1).getValue();
    const hack1StartDate = databaseSheet.getRange(11, 21, 1, 1).getValue();
    const hack1EndDate = databaseSheet.getRange(12, 21, 1, 1).getValue();
    const proj2StartDate = databaseSheet.getRange(15, 21, 1, 1).getValue();
    const proj2EndDate = databaseSheet.getRange(16, 21, 1, 1).getValue();
    const hack2StartDate = databaseSheet.getRange(19, 21, 1, 1).getValue();
    const hack2EndDate = databaseSheet.getRange(20, 21, 1, 1).getValue();

    summarySheet.getRange(14, 11, 1, 1).setValue(formatDate(hack1StartDate));
    summarySheet.getRange(14, 12, 1, 1).setValue(formatDate(hack1EndDate));
    summarySheet.getRange(14, 13, 1, 1).setValue(formatDate(hack2StartDate));
    summarySheet.getRange(14, 14, 1, 1).setValue(formatDate(hack2EndDate));
    summarySheet.getRange(17, 11, 1, 1).setValue(formatDate(proj1StartDate));
    summarySheet.getRange(17, 12, 1, 1).setValue(formatDate(proj1EndDate));
    summarySheet.getRange(17, 13, 1, 1).setValue(formatDate(proj2StartDate));
    summarySheet.getRange(17, 14, 1, 1).setValue(formatDate(proj2EndDate));

    // Delivery team names and emails — unpacked from the JSON arrays stored in DATABASE.
    const facName = JSON.parse(databaseSheet.getRange(3, 15, 1, 1).getValue())[0];
    const smeName = JSON.parse(databaseSheet.getRange(4, 15, 1, 1).getValue())[0];
    const ccName = JSON.parse(databaseSheet.getRange(5, 15, 1, 1).getValue())[0];

    summarySheet.getRange(20, 12, 1, 1).setValue(facName);
    summarySheet.getRange(21, 12, 1, 1).setValue(smeName);
    summarySheet.getRange(22, 12, 1, 1).setValue(ccName);

    const facEmail = JSON.parse(databaseSheet.getRange(3, 16, 1, 1).getValue())[0];
    const smeEmail = JSON.parse(databaseSheet.getRange(4, 16, 1, 1).getValue())[0];
    const ccEmail = JSON.parse(databaseSheet.getRange(5, 16, 1, 1).getValue())[0];

    summarySheet.getRange(20, 13, 1, 1).setValue(facEmail);
    summarySheet.getRange(21, 13, 1, 1).setValue(smeEmail);
    summarySheet.getRange(22, 13, 1, 1).setValue(ccEmail);

    // COUNTIF formulas count how many times each delivery team member's name appears in RECORDS
    // (used to track their session delivery count for signature/audit purposes).
    summarySheet.getRange(20, 14, 1, 1).setValue('=COUNTIF(INDIRECT("RECORDS!A:BO"), O20)');
    summarySheet.getRange(21, 14, 1, 1).setValue('=COUNTIF(INDIRECT("RECORDS!A:BO"), O21)');
    summarySheet.getRange(22, 14, 1, 1).setValue('=COUNTIF(INDIRECT("RECORDS!A:BO"), O22)');
    summarySheet.getRange(23, 14, 1, 1).setValue("=SUM(N20:N22)");

    summarySheet.getRange(20, 15, 1, 1).setValue(facName);
    summarySheet.getRange(21, 15, 1, 1).setValue(smeName);
    summarySheet.getRange(22, 15, 1, 1).setValue(ccName);

    // Set up the first student row with default status and per-student GLH/hackathon formulas,
    // then copy it down for all remaining students.
    summarySheet.getRange(2, 4, 1, 1).setValue("Active");
    // Per-student cumulative GLH: sums column 11 of each weekly block for this student's row offset.
    summarySheet.getRange(2, 6, 1, 1).setValue(
        `=SUM(INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (0 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (1 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (2 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (3 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (4 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (5 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (6 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (7 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (8 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (9 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (10 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (11 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (12 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (13 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (14 * ($K$8 + 7)), 11)), INDIRECT("'RECORDS'!" & ADDRESS((23 + ROW() - 2) + (15 * ($K$8 + 7)), 11)))`,
    );
    // Attendance rate: proportion of GLH achieved, with a stepped formula to handle partial progress.
    summarySheet.getRange(2, 7, 1, 1).setValue(
        "=IF(F2 < $N$11, (IF($N$11 > $M$11, IF(F2 > $L$11, ((F2 - $L$11) / ($N$11 - $L$11)) * 0.2, 0), 0) + IF($N$11 > $M$11, IF(F2 > $L$11, 0.8, (F2 / $L$11) * 0.8), F2 / $N$11)), 1)",
    );
    // Hackathon 1 attendance: checks whether the student has any positive value in the hack1 RECORDS columns.
    summarySheet.getRange(2, 8).setFormula(
        `=IF(SUM(INDIRECT("'RECORDS'!" & ADDRESS(21 + (7 * ($K$8 + 7)), 52)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (7 * ($K$8 + 7)), 63)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (8 * ($K$8 + 7)), 19))) <= 0, "-----", IF(SUM(INDIRECT("'RECORDS'!" & ADDRESS((21 + ROW()) + (7 * ($K$8 + 7)), 52)), INDIRECT("'RECORDS'!" & ADDRESS((21 + ROW()) + (7 * ($K$8 + 7)), 63)), INDIRECT("'RECORDS'!" & ADDRESS((21 + ROW()) + (8 * ($K$8 + 7)), 19))) > 0, "Yes", "No"))`,
    );
    // Hackathon 2 attendance: same logic for the second hackathon window.
    summarySheet.getRange(2, 9, 1, 1).setValue(
        `=IF(SUM(INDIRECT("'RECORDS'!" & ADDRESS(21 + (15 * ($K$8 + 7)), 30)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (15 * ($K$8 + 7)), 41)), INDIRECT("'RECORDS'!" & ADDRESS(21 + (15 * ($K$8 + 7)), 52))) <= 0, "-----", IF(SUM(INDIRECT("'RECORDS'!" & ADDRESS((21 + ROW()) + (15 * ($K$8 + 7)), 30)), INDIRECT("'RECORDS'!" & ADDRESS((21 + ROW()) + (15 * ($K$8 + 7)), 41)), INDIRECT("'RECORDS'!" & ADDRESS((21 + ROW()) + (15 * ($K$8 + 7)), 52))) > 0, "Yes", "No"))`,
    );

    // Copy the first student row down for all remaining students, then overwrite names from DATABASE.
    const studentRange = summarySheet.getRange(2, 1, 1, 9);
    for (let i = 1; i < numOfStudents; i++) {
        studentRange.copyTo(summarySheet.getRange(2 + i, 1, 1, 9));
    }

    summarySheet.getRange(2, 1, numOfStudents, 1).setValues(databaseSheet.getRange(3, 2, numOfStudents, 1).getValues());
    summarySheet.getRange(2, 2, numOfStudents, 1).setValues(databaseSheet.getRange(3, 3, numOfStudents, 1).getValues());
}

/**
 * Formats a YYYY-MM-DD date string to DD/MM/YYYY for display in the spreadsheet.
 */
function formatDate(date) {
    let [year, month, day] = date.split("-");
    return `${day}/${month}/${year}`;
}

/**
 * Trigger that fires whenever a cell is edited in the spreadsheet.
 * Intercepts edits to column 4 (Status) in the SUMMARY sheet and prompts the user
 * to confirm the change before applying it, since status changes affect RECORDS data
 * and are difficult to reverse.
 */
function onEdit(e) {
    const ui = SpreadsheetApp.getUi();
    const sheet = e.range.getSheet();

    if (sheet.getName() !== "SUMMARY") return;

    const cell = e.range;
    const col = cell.getColumn();

    if (col !== 4) return;

    const newValue = cell.getValue();
    const oldValue = e.oldValue || "";

    let response = ui.alert(
        `Change Status`,
        `Are you sure you want to change the status to "${newValue}"?\nThis will change the RECORDS sheet and be hard to undo.`,
        ui.ButtonSet.YES_NO,
    );

    if (response !== ui.Button.YES) {
        cell.setValue(oldValue);
    } else {
        let isStatusChanged = changeStatus(ui, cell, newValue);
        if (!isStatusChanged) {
            cell.setValue(oldValue);
        }
    }
}

/**
 * Applies a learner status change to the RECORDS sheet.
 *
 * For a status of "Active", writes "-" (not applicable) into all session cells from the
 * start of the cohort. For any other status (e.g. "Withdrawn"), writes "X" into all session
 * cells after the learner's last attended date.
 *
 * Prompts the user to confirm or correct the last attended date before writing.
 * Returns true if the change was applied, false if the user cancelled.
 */
function changeStatus(ui, cell, status) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const numOfStudents = databaseSheet.getRange(3, 22, 1, 1).getValue();
    const scheduleData = JSON.parse(databaseSheet.getRange(3, 23, 1, 1).getValue());
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;
    const startDate = schedule[0][0].date;

    const lastAttendedRange = summarySheet.getRange(cell.getRow(), cell.getColumn() + 1);
    let lastAttendedValue = lastAttendedRange.getDisplayValue();
    if (!lastAttendedValue || status == "Non Starter") {
        lastAttendedValue = "Non Starter (No Last Date)";
    }

    const response = ui.prompt(
        "Confirm Last Attended Date:",
        `Current last attended date is: "${lastAttendedValue}"\nIf this is incorrect, enter the correct date \ndd/mm/yyyy (e.g., 20/11/2024):`,
        ui.ButtonSet.OK_CANCEL,
    );

    if (response.getSelectedButton() === ui.Button.CANCEL) return false;

    const datePrompt = response.getResponseText();
    if (datePrompt && lastAttendedValue !== datePrompt) {
        lastAttendedValue = datePrompt;
    }
    lastAttendedRange.setValue(lastAttendedValue);

    let lastAttendedDateTime;
    if (lastAttendedValue == "Non Starter (No Last Date)" || status == "Active") {
        // For active or non-starter learners, treat as if attendance starts from the day before cohort start.
        lastAttendedDateTime = new Date(startDate);
        lastAttendedDateTime.setDate(lastAttendedDateTime.getDate() - 1);
    } else {
        const [day, month, year] = lastAttendedValue.split("/");
        lastAttendedDateTime = new Date(`${year}-${month}-${day}`);
    }

    const recordsSymbol = status == "Active" ? "-" : "X";
    const recordsValuesArr = Array.from({ length: 9 }, () => recordsSymbol);
    const cols = JSON.parse(databaseSheet.getRange(3, 10, 1, 1).getValue());
    const startWeekRow = 21;

    for (let i = 0; i < weeks; i++) {
        let week = schedule[i];
        // Row offset for this student in this week's block.
        let weekRow = startWeekRow + cell.getRow() + i * (numOfStudents + 7);

        for (let j = 0; j < week.length; j++) {
            let dateTime = new Date(week[j].date);
            // Skip days on or before the last attended date — only future sessions are affected.
            if (dateTime <= lastAttendedDateTime) continue;

            let col = cols[j + 1];
            recordsSheet.getRange(weekRow, col, 1, 9).setValues([recordsValuesArr]);
        }
    }

    ui.alert(`Finished`);
    return true;
}
