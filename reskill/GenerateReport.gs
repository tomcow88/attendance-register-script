/**
 * Opens a checkbox dialog listing all students with their funding partners.
 * The user selects which students to include, then submits to generate the report.
 * All students are unchecked by default — the report includes only checked students.
 */
function generateStudentReport() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const numOfStudents = summarySheet.getRange(5, 6, 1, 1).getValue();
    const firstNames = summarySheet
        .getRange(14, 1, numOfStudents, 1)
        .getValues()
        .flat();
    const lastNames = summarySheet
        .getRange(14, 2, numOfStudents, 1)
        .getValues()
        .flat();
    const partners = summarySheet
        .getRange(14, 3, numOfStudents, 1)
        .getValues()
        .flat();
    const fullNames = firstNames.map((first, i) => `${first} ${lastNames[i]}`);

    const template = HtmlService.createTemplateFromFile("ReportSelection");
    template.fullNames = fullNames;
    template.partners = partners;
    template.type = "Student";
    const height = Math.min(80 + fullNames.length * 24, 500);
    const html = template.evaluate().setHeight(height).setWidth(300);

    SpreadsheetApp.getUi().showModalDialog(html, "Choose Students");
}

/**
 * Prompts for a funding partner name, then opens the same checkbox dialog as
 * generateStudentReport but with only that partner's students pre-checked.
 * The partner name is passed to ReportSelection as the `type` value, which
 * the template uses to pre-check matching rows.
 */
function generatePartnerReport() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const numOfStudents = summarySheet.getRange(5, 6, 1, 1).getValue();
    const firstNames = summarySheet
        .getRange(14, 1, numOfStudents, 1)
        .getValues()
        .flat();
    const lastNames = summarySheet
        .getRange(14, 2, numOfStudents, 1)
        .getValues()
        .flat();
    const partners = summarySheet
        .getRange(14, 3, numOfStudents, 1)
        .getValues()
        .flat();
    const fullNames = firstNames.map((first, i) => `${first} ${lastNames[i]}`);

    const ui = SpreadsheetApp.getUi();
    const partnerPrompt = ui.prompt("Type in what partner you want");
    const response = partnerPrompt.getResponseText();

    const template = HtmlService.createTemplateFromFile("ReportSelection");
    template.fullNames = fullNames;
    template.partners = partners;
    template.type = response;
    const height = Math.min(80 + fullNames.length * 24, 500);
    const html = template.evaluate().setHeight(height).setWidth(300);

    SpreadsheetApp.getUi().showModalDialog(html, "Choose Students");
}

/**
 * Creates the report spreadsheet from the checkbox selection submitted by ReportSelection.html.
 *
 * `data` is an array of 0-based student indices that were UNCHECKED in the dialog —
 * these students are excluded from the report. Students not in `data` are included.
 *
 * Steps:
 *  1. Creates a new spreadsheet, copies SUMMARY and ATTENDANCE into it.
 *  2. Deletes rows for excluded students from both sheets.
 *  3. Logs the report to the REPORTS sheet.
 *  4. Re-sets SUMMARY formulas to force recalculation against the trimmed ATTENDANCE.
 */
function processCheckboxSelection(data, type) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const attendanceSheet = spreadsheet.getSheetByName("ATTENDANCE");
    const numOfStudents = summarySheet.getRange(5, 6, 1, 1).getValue();
    const numOfWeeksOnly = summarySheet.getRange(1, 6, 1, 1).getValue();
    const numOfExtraSessions = summarySheet.getRange(4, 6, 1, 1).getValue();
    const numOfWeeks = Number(numOfWeeksOnly) + Number(numOfExtraSessions);

    const date = new Date();
    const formattedDate = `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;

    const firstNames = summarySheet
        .getRange(14, 1, numOfStudents, 1)
        .getValues()
        .flat();
    const lastNames = summarySheet
        .getRange(14, 2, numOfStudents, 1)
        .getValues()
        .flat();
    const fullNames = firstNames.map((first, i) => `${first} ${lastNames[i]}`);
    let selectedStudents = [];

    // Build the list of included students (those not in the unchecked `data` array).
    for (let i = 0; i < fullNames.length; i++) {
        if (!data.includes(i.toString())) {
            selectedStudents.push(fullNames[i]);
        }
    }

    let studentsString = selectedStudents.join(", ");
    let formattedType;

    if (type != "Student") {
        formattedType = `Partner Report - ${type}`;
    } else {
        formattedType = `Student Report - ${studentsString}`;
    }

    const newSpreadSheet = SpreadsheetApp.create(
        `${formattedType} - ${formattedDate}`,
    );
    const newSummarySheet = summarySheet.copyTo(newSpreadSheet);
    newSummarySheet.setName("SUMMARY");
    const newAttendanceSheet = attendanceSheet.copyTo(newSpreadSheet);
    newAttendanceSheet.setName("ATTENDANCE");

    newSpreadSheet.deleteSheet(newSpreadSheet.getSheets()[0]);
    const newSpreadSheetURL = newSpreadSheet.getUrl();

    let summaryRemovedRows = [];
    let attendanceRemovedRows = [];

    for (let i = numOfStudents; i >= 0; i--) {
        if (data.includes(i.toString())) {
            // SUMMARY: student i is at row 14 + i (rows 1–13 are cohort metadata).
            let summaryRow = i + 14;
            summaryRemovedRows.push(summaryRow);
            newSummarySheet.deleteRow(summaryRow);

            for (let j = numOfWeeks - 1; j >= 0; j--) {
                // ATTENDANCE: each weekly block is (3 + numOfStudents) rows tall.
                // Student i within week j sits at block start + 1 header row + i.
                let attendanceRow = (3 + numOfStudents) * j + (4 + i);
                attendanceRemovedRows.push(attendanceRow);
            }
        }
    }

    // Sort descending so later rows are deleted first — deleting from the bottom
    // up prevents earlier row numbers from shifting as rows are removed.
    attendanceRemovedRows.sort((a, b) => b - a);

    for (let i = 0; i < attendanceRemovedRows.length; i++) {
        newAttendanceSheet.deleteRow(attendanceRemovedRows[i]);
    }

    const reportsSheet = spreadsheet.getSheetByName("REPORTS");
    const lastReportRow = reportsSheet.getLastRow() + 1;
    reportsSheet.getRange(lastReportRow, 1, 1, 1).setValue(formattedDate);
    reportsSheet.getRange(lastReportRow, 2, 1, 1).setValue(formattedType);
    reportsSheet.getRange(lastReportRow, 3, 1, 1).setValue(newSpreadSheetURL);

    // Re-set formulas to force recalculation now that excluded student rows are gone.
    const refreshSummaryRange = newSummarySheet.getRange(
        14,
        5,
        selectedStudents.length,
        2,
    );
    const refreshSummaryFormulas = refreshSummaryRange.getFormulas();
    refreshSummaryRange.setFormulas(refreshSummaryFormulas);

    spreadsheet.setActiveSheet(reportsSheet);
    SpreadsheetApp.getUi().alert(`Report Created!`);
}
