/**
 * Runs automatically when the spreadsheet is opened.
 * Adds three custom menus to the spreadsheet UI:
 *  - "Setup": manual step-by-step setup and a one-click auto setup, plus database update.
 *  - "Auto Attendance": triggers for checking attendance (today or all).
 *  - "Reports": generates partner-specific attendance report spreadsheets.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();

    // "Manual Setup Steps" is a sub-menu nested inside "Setup", allowing each
    // step to be run individually if the full auto setup is not appropriate.
    const subMenu = ui
        .createMenu("Manual Setup Steps")
        .addItem("Step 1: Create Sheets", "createSheetsFromTemplates")
        .addSeparator()
        .addItem("Step 2: Setup Database", "setupDatabase")
        .addSeparator()
        .addItem("Step 3: Setup Records", "setupRecords")
        .addSeparator()
        .addItem("Step 4: Setup Summary", "setupSummary")
        .addSeparator()
        .addItem("Step 5: Hide Sheets", "hideAllUnusedSheets");

    // Top-level "Setup" menu: create a fresh setup sheet, run manual steps individually,
    // run all steps automatically in sequence, or update the database after initial setup.
    const menu = ui
        .createMenu("Setup")
        .addItem("Create New Setup Sheet", "createNewSetupSheet")
        .addSeparator()
        .addSubMenu(subMenu)
        .addSeparator()
        .addItem("Auto Setup Steps", "setupEverything")
        .addSeparator()
        .addItem("Update Database", "updateDatabase");
    menu.addToUi();

    // "Auto Attendance" menu: check attendance entries for today only, or re-check all dates.
    const autoAttendanceMenu = ui
        .createMenu("Auto Attendance")
        .addItem("Check Today", "checkAttendanceToday")
        .addItem("Check All", "checkAllAttendance");
    autoAttendanceMenu.addToUi();

    // "Reports" menu: generate a spreadsheet report for one named partner, or all partners at once.
    const reportsMenu = ui
        .createMenu("Reports")
        .addItem("Generate Single Partner Report", "generatePartnerReport")
        .addItem(
            "Generate Multiple Partner Reports",
            "autoGeneratePartnerReports",
        );
    reportsMenu.addToUi();
}

/**
 * Runs all five setup steps in sequence:
 * createSheetsFromTemplates → setupDatabase → setupRecords → setupSummary → hideAllUnusedSheets.
 * Alerts the user when complete, reporting the total elapsed time in seconds.
 */
function setupEverything() {
    let startTime = new Date();

    createSheetsFromTemplates();
    setupDatabase();
    setupRecords();
    setupSummary();
    hideAllUnusedSheets();

    let elapsedTime = new Date() - startTime;
    Logger.log("Elapsed time: " + elapsedTime / 1000 + " seconds");
    SpreadsheetApp.getUi().alert(
        "Finished setting up everything\nElapsed time: " +
            elapsedTime / 1000 +
            " seconds",
    );
}

/**
 * Splits learner rows from the SUMMARY sheet into those belonging to the given partner
 * and those that do not.
 *
 * Returns an object with:
 *  - includedLearners: display names of learners kept in the report
 *  - removedLearners: display names of learners excluded from the report
 *  - removedLearnersRows: 1-based SUMMARY sheet row numbers for excluded learners (used to clear rows)
 *  - removedLearnersIndices: 0-based array indices for excluded learners (used as RECORDS block offsets)
 */
function splitLearnersByPartner(partnerName, data) {
    let includedLearners = [];
    let removedLearners = [];
    let removedLearnersRows = [];
    let removedLearnersIndices = [];

    for (let i = 0; i < data.length; i++) {
        let firstName = data[i][0].toString();
        let lastName = data[i][1].toString();
        let partner = data[i][2].toString().toLowerCase();

        if (partner === partnerName.toLowerCase()) {
            includedLearners.push(firstName + " " + lastName);
        } else if (firstName && lastName) {
            removedLearners.push(firstName + " " + lastName);
            removedLearnersRows.push(i + 2); // 1-based row: +1 for header, +1 for 0-to-1 index
            removedLearnersIndices.push(i);  // 0-based index used as offset within each RECORDS block
        }
    }

    return { includedLearners, removedLearners, removedLearnersRows, removedLearnersIndices };
}

/**
 * Creates a new Google Spreadsheet for the given partner, copies the SUMMARY and RECORDS sheets
 * into it, removes all non-partner learner rows, and logs the result to PARTNER_REPORTS.
 *
 * Parameters:
 *  partnerName           - the funding partner name (already lowercased)
 *  spreadsheet           - the source SpreadsheetApp spreadsheet object
 *  summarySheet          - the SUMMARY sheet object
 *  recordsSheet          - the RECORDS sheet object
 *  numOfStudents         - total number of learner rows in SUMMARY
 *  split                 - result of splitLearnersByPartner()
 *  currentWeek           - current week identifier (from SUMMARY column 13)
 *  partnerReportsSheet   - the PARTNER_REPORTS log sheet
 *  startTime             - Date object used to derive today's date for the log entry
 */
function buildPartnerReport(
    partnerName,
    spreadsheet,
    summarySheet,
    recordsSheet,
    numOfStudents,
    split,
    currentWeek,
    partnerReportsSheet,
    startTime,
) {
    const { removedLearnersRows, removedLearnersIndices } = split;

    const newSpreadsheet = SpreadsheetApp.create(
        `Partner Report - ${currentWeek} - ${spreadsheet.getName().split(" - ")[0]}`,
    );

    // Copy SUMMARY and remove non-partner learner rows, then compact upward.
    const newSummarySheet = summarySheet.copyTo(newSpreadsheet);
    newSummarySheet.setName("SUMMARY");

    for (let i = 0; i < removedLearnersRows.length; i++) {
        newSummarySheet.getRange(removedLearnersRows[i], 1, 1, 9).clearContent();
    }

    const maxRow = numOfStudents + 1;
    let currentRow = 2;
    let lastRow = maxRow;

    while (currentRow <= lastRow) {
        let firstName = newSummarySheet.getRange(currentRow, 1, 1, 1).getValue();

        if (firstName) {
            currentRow++;
            continue;
        }

        let nextRow = currentRow + 1;
        let remainingRows = lastRow - currentRow;

        if (remainingRows > 0) {
            let destinationRange = newSummarySheet.getRange(currentRow, 1, remainingRows, 9);
            let copyRange = newSummarySheet.getRange(nextRow, 1, remainingRows, 9);
            copyRange.copyTo(destinationRange);
        }

        let lastRowRange = newSummarySheet.getRange(lastRow, 1, 1, 9);
        lastRowRange.clearContent();
        lastRowRange.clearFormat();
        lastRowRange.clearDataValidations();
        lastRow--;
    }

    // Copy RECORDS and delete rows belonging to non-partner learners.
    // The RECORDS sheet has 16 weekly blocks; each block starts at row 21 + weekIndex * (numOfStudents + 7).
    // Within each block, student rows are at 0-based offsets matching their position in the SUMMARY data array.
    const newRecordsSheet = recordsSheet.copyTo(newSpreadsheet);
    newRecordsSheet.setName("RECORDS");

    let rowsToRemove = [];
    for (let i = 0; i < 16; i++) {
        let baseWeekRow = 21 + i * (numOfStudents + 7);
        for (let j = 0; j < removedLearnersIndices.length; j++) {
            rowsToRemove.push(baseWeekRow + removedLearnersIndices[j]);
        }
    }

    // Delete bottom-up so earlier row indices stay valid as rows are removed.
    for (let i = rowsToRemove.length - 1; i >= 0; i--) {
        newRecordsSheet.deleteRow(rowsToRemove[i]);
    }

    let defaultSheet = newSpreadsheet.getSheetByName("Sheet1");
    if (defaultSheet) newSpreadsheet.deleteSheet(defaultSheet);

    // Append a log entry: [date, partner name, week, URL].
    const newRow = partnerReportsSheet.getLastRow() + 1;
    const todayDate = startTime.toISOString().split("T")[0];
    const formattedTodayDate = formatDate(todayDate);
    partnerReportsSheet
        .getRange(newRow, 1, 1, 4)
        .setValues([[formattedTodayDate, partnerName, currentWeek, newSpreadsheet.getUrl()]]);

    return newSpreadsheet;
}

/**
 * Prompts the user to enter a funding partner name, then generates a new Google Spreadsheet
 * containing only the learners belonging to that partner.
 *
 * Steps:
 *  1. Prompts for a partner name via a UI dialog; cancels if empty or dismissed.
 *  2. Splits learners into included/removed and confirms the split with the user.
 *  3. Builds the filtered report spreadsheet via buildPartnerReport().
 *  4. Logs the new report to PARTNER_REPORTS and navigates the user there.
 */
function generatePartnerReport() {
    let startTime = new Date();
    const ui = SpreadsheetApp.getUi();

    const response = ui.prompt(
        "Generate Partner Report",
        "Please enter the funding partner name:",
        ui.ButtonSet.OK_CANCEL,
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
        ui.alert("Operation canceled.");
        return;
    }

    const partnerName = response.getResponseText().trim();

    if (!partnerName) {
        ui.alert("No partner name entered.");
        return;
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const numOfStudents = getNumOfStudents();
    const data = summarySheet.getRange(2, 1, numOfStudents, 3).getValues();

    const split = splitLearnersByPartner(partnerName, data);

    if (split.includedLearners.length === 0) {
        ui.alert("Partner does not exist.");
        return;
    }

    let message =
        "Learners from this partner:\n" + split.includedLearners.join("\n") + "\n\n";
    message += "Learners to be removed:\n" + split.removedLearners.join("\n");

    const userResponse = ui.alert(
        "Generate Partner Report",
        message,
        ui.ButtonSet.OK_CANCEL,
    );

    if (userResponse !== ui.Button.OK) {
        ui.alert("Operation canceled.");
        return;
    }

    const currentWeek = summarySheet.getRange(2, 13, 1, 1).getValue();

    let partnerReportsSheet = spreadsheet.getSheetByName("PARTNER_REPORTS");
    if (!partnerReportsSheet) {
        const template = spreadsheet.getSheetByName("TEMPLATE_PARTNER_REPORTS");
        partnerReportsSheet = template.copyTo(spreadsheet);
        partnerReportsSheet.setName("PARTNER_REPORTS");
    }

    buildPartnerReport(
        partnerName,
        spreadsheet,
        summarySheet,
        recordsSheet,
        numOfStudents,
        split,
        currentWeek,
        partnerReportsSheet,
        startTime,
    );

    let elapsedTime = new Date() - startTime;

    ui.alert(
        `Partner Report Generated!`,
        `You can view it in the "PARTNER_REPORTS" sheet!\nElapsed Time: ${elapsedTime / 1000} seconds`,
        ui.ButtonSet.OK,
    );

    partnerReportsSheet.showSheet();
    spreadsheet.setActiveSheet(partnerReportsSheet);
}

/**
 * Automatically generates a separate partner report spreadsheet for every unique funding partner
 * found in the SUMMARY sheet, without prompting the user for input.
 *
 * Steps:
 *  1. Reads all learner rows from SUMMARY and collects the unique set of partner names.
 *  2. For each partner, builds a filtered report spreadsheet via buildPartnerReport().
 *  3. Logs each new report to PARTNER_REPORTS.
 *  4. Logs total elapsed time when all reports are complete.
 */
function autoGeneratePartnerReports() {
    let startTime = new Date();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const numOfStudents = getNumOfStudents();
    const data = summarySheet.getRange(2, 1, numOfStudents, 3).getValues();

    const currentWeek = summarySheet.getRange(2, 13, 1, 1).getValue();

    let partnerReportsSheet = spreadsheet.getSheetByName("PARTNER_REPORTS");
    if (!partnerReportsSheet) {
        const template = spreadsheet.getSheetByName("TEMPLATE_PARTNER_REPORTS");
        partnerReportsSheet = template.copyTo(spreadsheet);
        partnerReportsSheet.setName("PARTNER_REPORTS");
    }

    let allPartnerNames = summarySheet
        .getRange(2, 3, numOfStudents, 1)
        .getValues()
        .flat();
    let partnerNames = [
        ...new Set(allPartnerNames.map((item) => item.toLowerCase())),
    ];

    for (let partnerNum = 0; partnerNum < partnerNames.length; partnerNum++) {
        let partnerName = partnerNames[partnerNum];
        const split = splitLearnersByPartner(partnerName, data);
        buildPartnerReport(
            partnerName,
            spreadsheet,
            summarySheet,
            recordsSheet,
            numOfStudents,
            split,
            currentWeek,
            partnerReportsSheet,
            startTime,
        );
        partnerReportsSheet.showSheet();
    }

    let elapsedTime = new Date() - startTime;
    Logger.log(`Finished - ${elapsedTime / 1000} seconds`);
}
