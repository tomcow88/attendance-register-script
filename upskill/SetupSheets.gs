/**
 * Recreates the working sheets by deleting any existing ones and copying fresh copies
 * from the TEMPLATE_* sheets. Also creates a blank LOGS sheet.
 * If no SETUP sheet exists, one is created from the template first.
 *
 * Working sheets created: DATABASE, SUMMARY, RECORDS, PARTNER_REPORTS, LOGS.
 */
function createSheetsFromTemplates() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Validate all required templates exist before deleting any working sheets.
    // If any are missing, abort so existing data is not permanently lost.
    const requiredTemplates = ["TEMPLATE_DATABASE", "TEMPLATE_SUMMARY", "TEMPLATE_RECORDS", "TEMPLATE_PARTNER_REPORTS"];
    const missingTemplates = requiredTemplates.filter(name => !spreadsheet.getSheetByName(name));
    if (missingTemplates.length > 0) {
        SpreadsheetApp.getUi().alert(`Cannot proceed: the following template sheets are missing:\n${missingTemplates.join("\n")}`);
        return;
    }

    let setupSheet = spreadsheet.getSheetByName("SETUP");
    if (!setupSheet) {
        setupSheet = createNewSetupSheet();
    }

    setupSheet.showSheet();
    spreadsheet.setActiveSheet(setupSheet);

    // Delete any existing working sheets before recreating them to ensure a clean state.
    const oldDatabaseSheet = spreadsheet.getSheetByName("DATABASE");
    const oldRecordsSheet = spreadsheet.getSheetByName("RECORDS");
    const oldSummarySheet = spreadsheet.getSheetByName("SUMMARY");
    const oldLogsSheet = spreadsheet.getSheetByName("LOGS");
    const oldPartnerReportsSheet = spreadsheet.getSheetByName("PARTNER_REPORTS");

    if (oldDatabaseSheet) spreadsheet.deleteSheet(oldDatabaseSheet);
    if (oldRecordsSheet) spreadsheet.deleteSheet(oldRecordsSheet);
    if (oldSummarySheet) spreadsheet.deleteSheet(oldSummarySheet);
    if (oldLogsSheet) spreadsheet.deleteSheet(oldLogsSheet);
    if (oldPartnerReportsSheet) spreadsheet.deleteSheet(oldPartnerReportsSheet);

    const databaseSheet = spreadsheet.getSheetByName("TEMPLATE_DATABASE").copyTo(spreadsheet);
    databaseSheet.setName("DATABASE");

    const summarySheet = spreadsheet.getSheetByName("TEMPLATE_SUMMARY").copyTo(spreadsheet);
    summarySheet.setName("SUMMARY");
    summarySheet.showSheet();
    spreadsheet.setActiveSheet(summarySheet);

    const recordsSheet = spreadsheet.getSheetByName("TEMPLATE_RECORDS").copyTo(spreadsheet);
    recordsSheet.setName("RECORDS");
    recordsSheet.showSheet();

    const partnerReportsSheet = spreadsheet.getSheetByName("TEMPLATE_PARTNER_REPORTS").copyTo(spreadsheet);
    partnerReportsSheet.setName("PARTNER_REPORTS");

    spreadsheet.insertSheet("LOGS");
    spreadsheet.setActiveSheet(summarySheet);
}

/**
 * Deletes the existing SETUP sheet (if any) and creates a fresh one from TEMPLATE_SETUP.
 * Returns the new SETUP sheet object.
 */
function createNewSetupSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const oldSetupSheet = spreadsheet.getSheetByName("SETUP");

    if (oldSetupSheet) spreadsheet.deleteSheet(oldSetupSheet);

    const setupSheet = spreadsheet.getSheetByName("TEMPLATE_SETUP").copyTo(spreadsheet);
    setupSheet.setName("SETUP");
    setupSheet.showSheet();

    return setupSheet;
}

/**
 * Hides all sheets that end users do not need to interact with directly:
 * all TEMPLATE_* sheets, SETUP, DATABASE, LOGS, and PARTNER_REPORTS.
 * Also clears the email column in SETUP so sensitive data is not left visible.
 */
function hideAllUnusedSheets() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const numOfStudents = getNumOfStudents();

    const setupTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_SETUP");
    const databaseTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_DATABASE");
    const recordsTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_RECORDS");
    const summaryTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_SUMMARY");
    const partnerReportsTemplateSheet = spreadsheet.getSheetByName("TEMPLATE_PARTNER_REPORTS");
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const logsSheet = spreadsheet.getSheetByName("LOGS");
    const partnerReportsSheet = spreadsheet.getSheetByName("PARTNER_REPORTS");

    // Clear email column before hiding SETUP to avoid leaving PII in a visible sheet.
    setupSheet.getRange(3, 3, numOfStudents, 1).clearContent();

    if (setupTemplateSheet) setupTemplateSheet.hideSheet();
    if (databaseTemplateSheet) databaseTemplateSheet.hideSheet();
    if (recordsTemplateSheet) recordsTemplateSheet.hideSheet();
    if (summaryTemplateSheet) summaryTemplateSheet.hideSheet();
    if (partnerReportsTemplateSheet) partnerReportsTemplateSheet.hideSheet();
    if (setupSheet) setupSheet.hideSheet();
    if (databaseSheet) databaseSheet.hideSheet();
    if (logsSheet) logsSheet.hideSheet();
    if (partnerReportsSheet) partnerReportsSheet.hideSheet();
}
