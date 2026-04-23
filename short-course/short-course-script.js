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
