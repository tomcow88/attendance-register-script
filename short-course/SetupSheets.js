function createSheetsFromTemplates() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    let setupSheet = spreadsheet.getSheetByName("SETUP");

    if (!setupSheet) {
        setupSheet = createNewSetupSheet();
    }

    setupSheet.showSheet();
    spreadsheet.setActiveSheet(setupSheet);

    const oldDatabaseSheet = spreadsheet.getSheetByName("DATABASE");
    const oldRecordsSheet = spreadsheet.getSheetByName("RECORDS");
    const oldSummarySheet = spreadsheet.getSheetByName("SUMMARY");
    const oldLogsSheet = spreadsheet.getSheetByName("LOGS");
    const oldPartnerReportsSheet =
        spreadsheet.getSheetByName("PARTNER_REPORTS");

    if (oldDatabaseSheet) spreadsheet.deleteSheet(oldDatabaseSheet);
    if (oldRecordsSheet) spreadsheet.deleteSheet(oldRecordsSheet);
    if (oldSummarySheet) spreadsheet.deleteSheet(oldSummarySheet);
    if (oldLogsSheet) spreadsheet.deleteSheet(oldLogsSheet);
    if (oldPartnerReportsSheet) spreadsheet.deleteSheet(oldPartnerReportsSheet);

    const databaseTemplateSheet =
        spreadsheet.getSheetByName("TEMPLATE_DATABASE");
    const recordsTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_RECORDS");
    const summaryTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_SUMMARY");
    const partnerReportsTemplateSheet = spreadsheet.getSheetByName(
        "TEMPLATE_PARTNER_REPORTS",
    );

    const databaseSheet = databaseTemplateSheet.copyTo(spreadsheet);
    databaseSheet.setName("DATABASE");

    const summarySheet = summaryTemplateSheet.copyTo(spreadsheet);
    summarySheet.setName("SUMMARY");
    summarySheet.showSheet();
    spreadsheet.setActiveSheet(summarySheet);

    const recordsSheet = recordsTemplateSheet.copyTo(spreadsheet);
    recordsSheet.setName("RECORDS");
    recordsSheet.showSheet();

    const partnerReportsSheet = partnerReportsTemplateSheet.copyTo(spreadsheet);
    partnerReportsSheet.setName("PARTNER_REPORTS");

    const logsSheet = spreadsheet.insertSheet("LOGS");
    spreadsheet.setActiveSheet(summarySheet);
}

function createNewSetupSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const oldSetupSheet = spreadsheet.getSheetByName("SETUP");

    if (oldSetupSheet) {
        spreadsheet.deleteSheet(oldSetupSheet);
    }

    const setupTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_SETUP");
    const setupSheet = setupTemplateSheet.copyTo(spreadsheet);
    setupSheet.setName("SETUP");
    setupSheet.showSheet();

    return setupSheet;
}

function hideAllUnusedSheets() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_SETUP");
    const databaseTemplateSheet =
        spreadsheet.getSheetByName("TEMPLATE_DATABASE");
    const recordsTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_RECORDS");
    const summaryTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_SUMMARY");
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const logsSheet = spreadsheet.getSheetByName("LOGS");
    const partnerReportsSheet = spreadsheet.getSheetByName("PARTNER_REPORTS");
    const numOfStudents = getNumOfStudents();

    setupSheet.getRange(3, 3, numOfStudents, 1).clearContent();

    if (setupTemplateSheet) setupTemplateSheet.hideSheet();
    if (databaseTemplateSheet) databaseTemplateSheet.hideSheet();
    if (recordsTemplateSheet) recordsTemplateSheet.hideSheet();
    if (summaryTemplateSheet) summaryTemplateSheet.hideSheet();
    if (setupSheet) setupSheet.hideSheet();
    if (databaseSheet) databaseSheet.hideSheet();
    if (logsSheet) logsSheet.hideSheet();
    if (partnerReportsSheet) partnerReportsSheet.hideSheet();
}
