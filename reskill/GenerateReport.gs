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
    const html = template.evaluate().setHeight(200).setWidth(300);

    SpreadsheetApp.getUi().showModalDialog(html, "Choose Students");
}

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
    const html = template.evaluate().setHeight(200).setWidth(300);

    SpreadsheetApp.getUi().showModalDialog(html, "Choose Students");
}

function processCheckboxSelection(data, type) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const attendanceSheet = spreadsheet.getSheetByName("ATTENDANCE");
    const numOfStudents = summarySheet.getRange(5, 6, 1, 1).getValue();
    const numOfWeeksOnly = summarySheet.getRange(1, 6, 1, 1).getValue();
    const numOfExtraSessions = summarySheet.getRange(4, 6, 1, 1).getValue();
    const numOfWeeks = Number(numOfWeeksOnly) + Number(numOfExtraSessions);

    const date = new Date(); // today's date
    const day = String(date.getDate()).padStart(2, "0"); // dd
    const month = String(date.getMonth() + 1).padStart(2, "0"); // mm (months are 0-indexed)
    const year = date.getFullYear(); // yyyy
    const formattedDate = `${day}/${month}/${year}`;

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

    let defaultSheet = newSpreadSheet.getSheetByName("Sheet1");
    newSpreadSheet.deleteSheet(defaultSheet);
    const newSpreadSheetURL = newSpreadSheet.getUrl();

    let summaryRemovedRows = [];
    let attendanceRemovedRows = [];

    for (let i = numOfStudents; i >= 0; i--) {
        if (data.includes(i.toString())) {
            let summaryRow = i + 14;
            summaryRemovedRows.push(summaryRow);
            newSummarySheet.deleteRow(summaryRow);

            for (let j = numOfWeeks - 1; j >= 0; j--) {
                let attendanceRow = (3 + numOfStudents) * j + (4 + i);
                attendanceRemovedRows.push(attendanceRow);
            }
        }
    }

    attendanceRemovedRows.sort((a, b) => b - a);

    for (let i = 0; i < attendanceRemovedRows.length; i++) {
        newAttendanceSheet.deleteRow(attendanceRemovedRows[i]);
    }

    const reportsSheet = spreadsheet.getSheetByName("REPORTS");
    const lastReportRow = reportsSheet.getLastRow() + 1;
    reportsSheet.getRange(lastReportRow, 1, 1, 1).setValue(formattedDate);
    reportsSheet.getRange(lastReportRow, 2, 1, 1).setValue(formattedType);
    reportsSheet.getRange(lastReportRow, 3, 1, 1).setValue(newSpreadSheetURL);

    const refreshSummaryRange = newSummarySheet.getRange(
        14,
        5,
        numOfStudents,
        2,
    );
    const refreshSummaryFormulas = refreshSummaryRange.getFormulas();
    refreshSummaryRange.setFormulas(refreshSummaryFormulas);

    spreadsheet.setActiveSheet(reportsSheet);
    const reportUi = SpreadsheetApp.getUi();
    reportUi.alert(`Report Created!`);
}
