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
        .addItem("Update Database", "updateDatabase")
        .addSeparator()
        .addItem("Add Returning Learner", "addReturningLearner")
        .addSeparator()
        .addItem("Refresh Prior GLH", "refreshPriorGLH");
    menu.addToUi();

    // "Auto Attendance" menu: check attendance entries for today only, or re-check all dates.
    const autoAttendanceMenu = ui
        .createMenu("Auto Attendance")
        .addItem("Check Today", "checkAttendanceToday")
        .addItem("Check All", "checkAllAttendance")
        .addItem("Check Between Dates", "checkAttendanceBetweenDates");
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
        let partner = data[i][2].toString().trim().toLowerCase();

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
        let baseWeekRow = 23 + i * (numOfStudents + 7);
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
        ...new Set(allPartnerNames.map((item) => item.toString().trim().toLowerCase()).filter(Boolean)),
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

/**
 * Inserts a new student row into every weekly RECORDS block for a returning learner.
 * The new row is appended at the end of each block's student section and populated with
 * formulas copied from the row above. Blocks are processed bottom-to-top so that earlier
 * block row positions are not shifted by later insertions.
 *
 * numOfStudents is the count BEFORE the new student is added.
 * fullName is written into column 1 of the new row.
 */
function insertLearnerIntoRecords(recordsSheet, numOfStudents, fullName) {
    const lastCol = recordsSheet.getLastColumn();
    for (let i = 15; i >= 0; i--) {
        const newStudentRow = 23 + numOfStudents + i * (numOfStudents + 7);
        recordsSheet.insertRows(newStudentRow, 1);
        recordsSheet
            .getRange(newStudentRow - 1, 1, 1, lastCol)
            .copyTo(
                recordsSheet.getRange(newStudentRow, 1, 1, lastCol),
                SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
                false
            );
        recordsSheet.getRange(newStudentRow, 1, 1, 1).setValue(fullName);
    }
}

/**
 * Prompts the user for a returning learner's details, looks them up in their prior cohort
 * spreadsheet, and adds them to the end of the current cohort's DATABASE, RECORDS, and SUMMARY.
 *
 * Prior GLH and hackathon attendance are read from the prior cohort's SUMMARY sheet and stored
 * in DATABASE columns 25–28 so that the SUMMARY formulas can include them in the learner's totals.
 *
 * The user is warned if the prior spreadsheet is inaccessible or the learner cannot be matched;
 * they can choose to continue adding without prior data.
 */
function addReturningLearner() {
    const ui = SpreadsheetApp.getUi();

    const nameResponse = ui.prompt(
        "Add Returning Learner",
        "Enter full name (First Last):",
        ui.ButtonSet.OK_CANCEL
    );
    if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
    const fullName = nameResponse.getResponseText().trim();
    if (!fullName || !fullName.includes(" ")) {
        ui.alert("Please enter a full name with first and last name separated by a space.");
        return;
    }

    const partnerResponse = ui.prompt(
        "Add Returning Learner",
        "Enter funding partner name:",
        ui.ButtonSet.OK_CANCEL
    );
    if (partnerResponse.getSelectedButton() !== ui.Button.OK) return;
    const partner = partnerResponse.getResponseText().trim();

    const emailResponse = ui.prompt(
        "Add Returning Learner",
        "Enter learner email address:",
        ui.ButtonSet.OK_CANCEL
    );
    if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
    const email = emailResponse.getResponseText().trim();

    const urlResponse = ui.prompt(
        "Add Returning Learner",
        "Enter prior cohort spreadsheet URL:",
        ui.ButtonSet.OK_CANCEL
    );
    if (urlResponse.getSelectedButton() !== ui.Button.OK) return;
    const priorUrl = urlResponse.getResponseText().trim();

    const meetEmail = getMeetEmail(email);
    let priorGlh = 0;
    let priorHack1 = "";
    let priorHack2 = "";

    try {
        const priorSpreadsheet = SpreadsheetApp.openByUrl(priorUrl);
        const priorDatabase = priorSpreadsheet.getSheetByName("DATABASE");
        const priorSummary = priorSpreadsheet.getSheetByName("SUMMARY");

        if (!priorDatabase || !priorSummary) {
            ui.alert("Could not find DATABASE or SUMMARY in the prior cohort spreadsheet.");
            return;
        }

        // Col 23 = new layout (partner column added); col 22 = old layout. Fall back for prior cohorts.
        const priorNumStudents = priorDatabase.getRange(3, 23, 1, 1).getValue()
            || priorDatabase.getRange(3, 22, 1, 1).getValue();
        let priorStudentDatabaseRow = -1;

        for (let i = 0; i < priorNumStudents; i++) {
            const cellValue = priorDatabase.getRange(3 + i, 5, 1, 1).getValue();
            if (!cellValue) continue;
            try {
                const storedEmails = JSON.parse(cellValue);
                if (storedEmails.includes(meetEmail)) {
                    priorStudentDatabaseRow = 3 + i;
                    break;
                }
            } catch (e) {
                continue;
            }
        }

        if (priorStudentDatabaseRow === -1) {
            const proceed = ui.alert(
                "Learner Not Found",
                `No match for "${email}" in the prior cohort. Continue without prior data?`,
                ui.ButtonSet.YES_NO
            );
            if (proceed !== ui.Button.YES) return;
        } else {
            // SUMMARY row = DATABASE row - 1 (SUMMARY row 2 = DATABASE row 3, etc.)
            const priorSummaryRow = priorStudentDatabaseRow - 1;
            priorGlh = priorSummary.getRange(priorSummaryRow, 6, 1, 1).getValue() || 0;
            const hack1Raw = priorSummary.getRange(priorSummaryRow, 8, 1, 1).getValue();
            const hack2Raw = priorSummary.getRange(priorSummaryRow, 9, 1, 1).getValue();
            priorHack1 = hack1Raw === "Yes" ? "Yes" : "";
            priorHack2 = hack2Raw === "Yes" ? "Yes" : "";
        }
    } catch (e) {
        const proceed = ui.alert(
            "Could Not Access Prior Spreadsheet",
            `Error: ${e.message}\nContinue without prior data?`,
            ui.ButtonSet.YES_NO
        );
        if (proceed !== ui.Button.YES) return;
    }

    const nameParts = fullName.split(" ");
    const firstName = nameParts[0];
    const lastName = nameParts.slice(1).join(" ");

    const confirmMsg =
        `Name: ${fullName}\nPartner: ${partner}\nEmail: ${email}` +
        `\nPrior GLH: ${priorGlh}` +
        `\nPrior Hack1: ${priorHack1 || "N/A"}` +
        `\nPrior Hack2: ${priorHack2 || "N/A"}`;

    const confirmed = ui.alert("Confirm Returning Learner", confirmMsg, ui.ButtonSet.OK_CANCEL);
    if (confirmed !== ui.Button.OK) return;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");

    const numOfStudents = getNumOfStudents();
    const newDatabaseRow = 3 + numOfStudents;

    databaseSheet.getRange(newDatabaseRow, 1, 1, 1).setValue(fullName);
    databaseSheet.getRange(newDatabaseRow, 2, 1, 1).setValue(firstName);
    databaseSheet.getRange(newDatabaseRow, 3, 1, 1).setValue(lastName);
    databaseSheet.getRange(newDatabaseRow, 4, 1, 1).setValue(`["${firstName}"]`);
    databaseSheet.getRange(newDatabaseRow, 5, 1, 1).setValue(`["${meetEmail}"]`);
    databaseSheet.getRange(newDatabaseRow, 26, 1, 1).setValue(priorUrl);
    databaseSheet.getRange(newDatabaseRow, 27, 1, 1).setValue(priorGlh);
    if (priorHack1 === "Yes") databaseSheet.getRange(newDatabaseRow, 28, 1, 1).setValue("Yes");
    if (priorHack2 === "Yes") databaseSheet.getRange(newDatabaseRow, 29, 1, 1).setValue("Yes");

    databaseSheet.getRange(3, 23, 1, 1).setValue(numOfStudents + 1);

    insertLearnerIntoRecords(recordsSheet, numOfStudents, fullName);

    const newSummaryRow = numOfStudents + 2;
    summarySheet
        .getRange(newSummaryRow - 1, 1, 1, 9)
        .copyTo(summarySheet.getRange(newSummaryRow, 1, 1, 9));
    summarySheet.getRange(newSummaryRow, 1, 1, 1).setValue(firstName);
    summarySheet.getRange(newSummaryRow, 2, 1, 1).setValue(lastName);
    summarySheet.getRange(newSummaryRow, 3, 1, 1).setValue(partner);
    summarySheet.getRange(newSummaryRow, 4, 1, 1).setValue("Active");

    ui.alert(`"${fullName}" added successfully as a returning learner.`);
}

/**
 * Re-reads the prior cohort spreadsheet for each returning learner in DATABASE (those with
 * a URL in column 26) and refreshes the prior GLH and hackathon data in columns 27–29.
 *
 * Learners are matched by their stored meet email (DATABASE col 5) against the prior DATABASE.
 * Any inaccessible spreadsheets or unmatched learners are listed in the summary alert.
 */
function refreshPriorGLH() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const numOfStudents = getNumOfStudents();

    let updated = 0;
    let warnings = [];

    for (let i = 0; i < numOfStudents; i++) {
        const databaseRow = 3 + i;
        const priorUrl = databaseSheet.getRange(databaseRow, 26, 1, 1).getValue();
        if (!priorUrl) continue;

        const fullName = databaseSheet.getRange(databaseRow, 1, 1, 1).getValue() || `Row ${databaseRow}`;

        try {
            const priorSpreadsheet = SpreadsheetApp.openByUrl(priorUrl);
            const priorDatabase = priorSpreadsheet.getSheetByName("DATABASE");
            const priorSummary = priorSpreadsheet.getSheetByName("SUMMARY");

            if (!priorDatabase || !priorSummary) {
                warnings.push(`${fullName}: DATABASE or SUMMARY not found in prior spreadsheet.`);
                continue;
            }

            let currentMeetEmails = [];
            try {
                currentMeetEmails = JSON.parse(databaseSheet.getRange(databaseRow, 5, 1, 1).getValue());
            } catch (e) {
                warnings.push(`${fullName}: Could not read stored meet emails.`);
                continue;
            }

            // Col 23 = new layout (partner column added); col 22 = old layout. Fall back for prior cohorts.
            const priorNumStudents = priorDatabase.getRange(3, 23, 1, 1).getValue()
                || priorDatabase.getRange(3, 22, 1, 1).getValue();
            let priorStudentDatabaseRow = -1;

            for (let j = 0; j < priorNumStudents; j++) {
                const cellValue = priorDatabase.getRange(3 + j, 5, 1, 1).getValue();
                if (!cellValue) continue;
                try {
                    const storedEmails = JSON.parse(cellValue);
                    if (currentMeetEmails.some((e) => storedEmails.includes(e))) {
                        priorStudentDatabaseRow = 3 + j;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            if (priorStudentDatabaseRow === -1) {
                warnings.push(`${fullName}: Could not find matching learner in prior spreadsheet.`);
                continue;
            }

            const priorSummaryRow = priorStudentDatabaseRow - 1;
            const priorGlh = priorSummary.getRange(priorSummaryRow, 6, 1, 1).getValue() || 0;
            const hack1Raw = priorSummary.getRange(priorSummaryRow, 8, 1, 1).getValue();
            const hack2Raw = priorSummary.getRange(priorSummaryRow, 9, 1, 1).getValue();

            databaseSheet.getRange(databaseRow, 27, 1, 1).setValue(priorGlh);
            databaseSheet.getRange(databaseRow, 28, 1, 1).setValue(hack1Raw === "Yes" ? "Yes" : "");
            databaseSheet.getRange(databaseRow, 29, 1, 1).setValue(hack2Raw === "Yes" ? "Yes" : "");
            updated++;
        } catch (e) {
            warnings.push(`${fullName}: ${e.message}`);
        }
    }

    let message = `Refresh complete.\nUpdated: ${updated} learner(s).`;
    if (warnings.length > 0) {
        message += "\n\nWarnings:\n" + warnings.join("\n");
    }
    ui.alert("Refresh Prior GLH", message, ui.ButtonSet.OK);
}
