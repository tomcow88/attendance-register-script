/**
 * Runs automatically when the spreadsheet is opened.
 * Adds two custom menus to the spreadsheet UI:
 *  - "Setup": one-click sheet setup.
 *  - "Report": generates per-student or per-partner attendance report spreadsheets.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();

    const setupMenu = ui
        .createMenu("Setup")
        .addItem("Setup Sheets", "setupSheets");
    setupMenu.addToUi();

    const reportMenu = ui
        .createMenu("Report")
        .addItem("Generate Student Report", "generateStudentReport")
        .addSeparator()
        .addItem("Generate Partner Report", "generatePartnerReport");
    reportMenu.addToUi();
}
