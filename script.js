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

    // Run each setup step in the required order. Each function is defined elsewhere in the project.
    createSheetsFromTemplates(); // Step 1: copy template sheets into the spreadsheet
    setupDatabase();             // Step 2: populate the database sheet
    setupRecords();              // Step 3: build the attendance records structure
    setupSummary();              // Step 4: build the summary sheet
    hideAllUnusedSheets();       // Step 5: hide any sheets not needed by end users

    let endTime = new Date();
    let elapsedTime = endTime - startTime;

    Logger.log("Elapsed time: " + elapsedTime / 1000 + " seconds");
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        "Finished setting up everything\nElapsed time: " +
            elapsedTime / 1000 +
            " seconds",
    );
}

/**
 * Prompts the user to enter a funding partner name, then generates a new Google Spreadsheet
 * containing only the learners belonging to that partner.
 *
 * Steps:
 *  1. Prompts for a partner name via a UI dialog; cancels if the name is empty or dialog is dismissed.
 *  2. Reads all learner rows from the SUMMARY sheet and splits them into included (matching partner)
 *     and removed (non-matching) lists.
 *  3. Confirms the learner split with the user before proceeding.
 *  4. Creates a new spreadsheet, copies the SUMMARY and RECORDS sheets into it, then removes all
 *     rows belonging to non-partner learners and compacts the remaining rows upward.
 *  5. Logs the new report (date, partner name, current week, URL) to the PARTNER_REPORTS sheet,
 *     creating that sheet from its template if it does not yet exist.
 *  6. Alerts the user with a success message and the elapsed time.
 */
function generatePartnerReport() {
    let startTime = new Date();

    const ui = SpreadsheetApp.getUi();

    // Show a prompt dialog so the user can type the partner name they want to report on.
    const response = ui.prompt(
        "Generate Partner Report",
        "Please enter the funding partner name:",
        ui.ButtonSet.OK_CANCEL,
    );

    // If the user clicked Cancel or closed the dialog, abort early.
    if (response.getSelectedButton() !== ui.Button.OK) {
        ui.alert("Operation canceled.");
        return;
    }

    const partnerName = response.getResponseText().trim();

    // Guard against an empty submission (user clicked OK without typing anything).
    if (!partnerName) {
        ui.alert("No partner name entered.");
        return;
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const numOfStudents = getNumOfStudents(); // helper defined elsewhere in the project

    // Read the first three columns of the SUMMARY sheet (first name, last name, partner)
    // starting from row 2 (row 1 is the header row).
    const dataRange = summarySheet.getRange(2, 1, numOfStudents, 3);
    const data = dataRange.getValues();

    // Split learners into two groups: those belonging to the requested partner (kept)
    // and everyone else (to be removed from the report copy).
    let includedLearners = [];
    let removedLearners = [];
    let removedLearnersRows = []; // 1-based sheet row numbers for learners being removed

    for (let i = 0; i < data.length; i++) {
        let firstName = data[i][0].toString();
        let lastName = data[i][1].toString();
        let partner = data[i][2].toString().toLowerCase();

        if (partner === partnerName.toLowerCase()) {
            // This learner belongs to the requested partner — keep them in the report.
            includedLearners.push(firstName + " " + lastName);
        } else if (firstName && lastName) {
            // This learner belongs to a different partner — mark their row for removal.
            // i + 2 converts from 0-based array index to 1-based sheet row (accounting for the header).
            removedLearners.push(firstName + " " + lastName);
            removedLearnersRows.push(i + 2);
        }
    }

    // If no learners matched, the partner name doesn't exist in this spreadsheet.
    if (includedLearners.length === 0) {
        ui.alert("Partner does not exist.");
        return;
    }

    // Show the user a summary of who will be kept and who will be removed before proceeding.
    let message =
        "Learners from this partner:\n" + includedLearners.join("\n") + "\n\n";
    message += "Learners to be removed:\n" + removedLearners.join("\n");

    const userResponse = ui.alert(
        "Generate Partner Report",
        message,
        ui.ButtonSet.OK_CANCEL,
    );

    // If the user clicked Cancel on the confirmation dialog, abort without creating anything.
    if (userResponse !== ui.Button.OK) {
        ui.alert("Operation canceled.");
        return;
    }

    // Read the current week identifier from column 13 of the first data row in SUMMARY.
    // This is used to name the new spreadsheet so it can be identified later.
    const currentWeek = summarySheet.getRange(2, 13, 1, 1).getValue();

    // Create a brand-new Google Spreadsheet to hold this partner's filtered data.
    // The name format is: "Partner Report - <week> - <base spreadsheet name>".
    const newSpreadsheet = SpreadsheetApp.create(
        `Partner Report - ${currentWeek} - ${spreadsheet.getName().split(" - ")[0]}`,
    );

    // Copy the full SUMMARY sheet into the new spreadsheet and rename it.
    const newSummarySheet = summarySheet.copyTo(newSpreadsheet);
    newSummarySheet.setName("SUMMARY");

    // Clear the content of every row that belongs to a non-partner learner.
    // The rows are cleared first, then compacted below (rather than deleted directly)
    // to avoid row-index shifting mid-loop.
    for (let i = 0; i < removedLearnersRows.length; i++) {
        let row = removedLearnersRows[i];
        newSummarySheet.getRange(row, 1, 1, 9).clearContent();
    }

    // Compact the SUMMARY sheet by shifting non-empty rows upward to fill the gaps
    // left by the cleared rows above. Iterates from row 2 (first data row) downward.
    const maxRow = numOfStudents + 1; // last possible data row (inclusive)
    let currentRow = 2;
    let lastRow = maxRow;

    while (currentRow <= lastRow) {
        let firstName = newSummarySheet
            .getRange(currentRow, 1, 1, 1)
            .getValue();

        // Row has data — move on to check the next row.
        if (firstName) {
            currentRow++;
            continue;
        }

        // Row is empty: shift all rows below it up by one to close the gap.
        let nextRow = currentRow + 1;
        let remainingRows = lastRow - currentRow; // number of rows to shift up

        if (remainingRows > 0) {
            let destinationRange = newSummarySheet.getRange(
                currentRow,
                1,
                remainingRows,
                9,
            );
            let copyRange = newSummarySheet.getRange(
                nextRow,
                1,
                remainingRows,
                9,
            );
            copyRange.copyTo(destinationRange);
        }

        // Clear the last row (which is now a duplicate after the shift) and shrink the range.
        let lastRowRange = newSummarySheet.getRange(lastRow, 1, 1, 9);
        lastRowRange.clearContent();
        lastRowRange.clearFormat();
        lastRowRange.clearDataValidations();

        // The effective data range is now one row shorter.
        lastRow--;
    }

    // Copy the full RECORDS sheet into the new spreadsheet and rename it.
    const newRecordsSheet = recordsSheet.copyTo(newSpreadsheet);
    newRecordsSheet.setName("RECORDS");

    // Build the list of RECORDS rows to delete. The RECORDS sheet repeats a block of rows
    // for each of the 16 weeks: each block starts at row 21 + (weekIndex * (numOfStudents + 7)).
    // Within each block, the student rows are offset by their original SUMMARY row numbers.
    // NOTE: the use of removedLearnersRows (absolute SUMMARY row numbers) as offsets here
    // may be a bug — see refactoring notes.
    let rowsToRemove = [];
    for (let i = 0; i < 16; i++) {
        let baseWeekRow = 21 + i * (numOfStudents + 7); // first row of this week's block
        for (let j = 0; j < removedLearnersRows.length; j++) {
            let baseRow = removedLearnersRows[j];
            let row = baseWeekRow + baseRow;
            rowsToRemove.push(row);
        }
    }

    // Delete rows in reverse order so that earlier row indices are not invalidated
    // by the deletion of later rows (deleting from the bottom up keeps indices stable).
    for (let i = rowsToRemove.length - 1; i >= 0; i--) {
        let rowToRemove = rowsToRemove[i];
        newRecordsSheet.deleteRow(rowToRemove);
    }

    // Remove the default "Sheet1" that Google Sheets creates automatically with every new spreadsheet.
    let defaultSheet = newSpreadsheet.getSheetByName("Sheet1");
    if (defaultSheet) newSpreadsheet.deleteSheet(defaultSheet);

    // Ensure the PARTNER_REPORTS log sheet exists in the source spreadsheet.
    // If it doesn't yet exist, create it by copying from the template sheet.
    let partnerReportsSheet = spreadsheet.getSheetByName("PARTNER_REPORTS");

    if (!partnerReportsSheet) {
        const partnerReportsTemplateSheet = spreadsheet.getSheetByName(
            "TEMPLATE_PARTNER_REPORTS",
        );
        partnerReportsSheet = partnerReportsTemplateSheet.copyTo(spreadsheet);
        partnerReportsSheet.setName("PARTNER_REPORTS");
    }

    // Append a new log entry to PARTNER_REPORTS: [date, partner name, week, URL].
    const newRow = partnerReportsSheet.getLastRow() + 1;
    const todayDate = startTime.toISOString().split("T")[0]; // extract YYYY-MM-DD from ISO string
    const formattedTodayDate = formatDate(todayDate);        // helper defined elsewhere in the project
    const newSpreadsheetUrl = newSpreadsheet.getUrl();

    const newPartnerReportValues = [
        [formattedTodayDate, partnerName, currentWeek, newSpreadsheetUrl],
    ];
    partnerReportsSheet
        .getRange(newRow, 1, 1, 4)
        .setValues(newPartnerReportValues);

    let endTime = new Date();
    let elapsedTime = endTime - startTime;

    ui.alert(
        `Partner Report Generated!`,
        `You can view it in the "PARTNER_REPORTS" sheet!\nElapsed Time: ${elapsedTime / 1000} seconds`,
        ui.ButtonSet.OK,
    );

    // Navigate the user to the PARTNER_REPORTS sheet so they can see the new log entry.
    partnerReportsSheet.showSheet();
    spreadsheet.setActiveSheet(partnerReportsSheet);
}

/**
 * Automatically generates a separate partner report spreadsheet for every unique funding partner
 * found in the SUMMARY sheet, without prompting the user for input.
 *
 * Steps:
 *  1. Reads all learner rows from the SUMMARY sheet and collects the unique set of partner names.
 *  2. For each partner, separates learners into included (matching) and removed (non-matching) lists.
 *  3. Creates a new spreadsheet per partner, copies the SUMMARY and RECORDS sheets into it, then
 *     removes all rows belonging to non-partner learners and compacts the remaining rows upward.
 *  4. Logs each new report (date, partner name, current week, URL) to the PARTNER_REPORTS sheet,
 *     creating that sheet from its template if it does not yet exist.
 *  5. Logs total elapsed time to the Apps Script logger when all reports are complete.
 *
 * Note: unlike generatePartnerReport(), this function produces no user-facing confirmation dialogs.
 */
function autoGeneratePartnerReports() {
    let startTime = new Date();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    const numOfStudents = getNumOfStudents(); // helper defined elsewhere in the project

    // Read the first three columns of the SUMMARY sheet (first name, last name, partner)
    // starting from row 2 (row 1 is the header row).
    const dataRange = summarySheet.getRange(2, 1, numOfStudents, 3);
    const data = dataRange.getValues();

    // Read the partner column (column 3) and deduplicate to get the unique set of partner names.
    // getValues() returns a 2D array, so .flat() flattens it to a 1D array of strings.
    // All names are lowercased to avoid generating duplicate reports from casing differences.
    let allPartnerNames = summarySheet
        .getRange(2, 3, numOfStudents, 1)
        .getValues()
        .flat()
        .flat();
    let partnerNames = [
        ...new Set(allPartnerNames.map((item) => item.toLowerCase())),
    ];

    Logger.log(allPartnerNames);
    Logger.log(partnerNames);

    // Iterate over every unique partner and generate a separate report spreadsheet for each.
    for (let parnterNum = 0; parnterNum < partnerNames.length; parnterNum++) {
        let includedLearners = [];
        let removedLearners = [];
        let removedLearnersRows = []; // 1-based sheet row numbers for learners being removed
        let partnerName = partnerNames[parnterNum]; // already lowercased from the Set above

        // Split learners into those belonging to this partner (kept) and all others (removed).
        for (let i = 0; i < data.length; i++) {
            let firstName = data[i][0].toString();
            let lastName = data[i][1].toString();
            let partner = data[i][2].toString().toLowerCase();
            Logger.log(partner);
            Logger.log(partnerName);
            if (partner === partnerName.toLowerCase()) {
                // This learner belongs to the current partner — include them in the report.
                includedLearners.push(firstName + " " + lastName);
            } else if (firstName && lastName) {
                // This learner belongs to a different partner — mark their row for removal.
                // i + 2 converts from 0-based array index to 1-based sheet row (accounting for the header).
                removedLearners.push(firstName + " " + lastName);
                removedLearnersRows.push(i + 2);
            }
        }

        Logger.log(removedLearners);
        Logger.log(includedLearners);

        // Read the current week identifier from column 13 of the first data row in SUMMARY.
        // Used to name the new spreadsheet so it can be identified later.
        // NOTE: this is fetched inside the loop but never changes — a candidate for hoisting out.
        const currentWeek = summarySheet.getRange(2, 13, 1, 1).getValue();

        // Create a brand-new Google Spreadsheet to hold this partner's filtered data.
        // The name format is: "Partner Report - <week> - <base spreadsheet name>".
        const newSpreadsheet = SpreadsheetApp.create(
            `Partner Report - ${currentWeek} - ${spreadsheet.getName().split(" - ")[0]}`,
        );

        // Copy the full SUMMARY sheet into the new spreadsheet and rename it.
        const newSummarySheet = summarySheet.copyTo(newSpreadsheet);
        newSummarySheet.setName("SUMMARY");

        // Clear the content of every row that belongs to a non-partner learner.
        // Rows are cleared first and compacted below (rather than deleted directly)
        // to avoid row-index shifting mid-loop.
        for (let i = 0; i < removedLearnersRows.length; i++) {
            let row = removedLearnersRows[i];
            newSummarySheet.getRange(row, 1, 1, 9).clearContent();
        }

        // Compact the SUMMARY sheet by shifting non-empty rows upward to fill the gaps
        // left by the cleared rows above. Iterates from row 2 (first data row) downward.
        const maxRow = numOfStudents + 1; // last possible data row (inclusive)
        let currentRow = 2;
        let lastRow = maxRow;

        while (currentRow <= lastRow) {
            let firstName = newSummarySheet
                .getRange(currentRow, 1, 1, 1)
                .getValue();

            // Row has data — move on to check the next row.
            if (firstName) {
                currentRow++;
                continue;
            }

            // Row is empty: shift all rows below it up by one to close the gap.
            let nextRow = currentRow + 1;
            let remainingRows = lastRow - currentRow; // number of rows to shift up

            if (remainingRows > 0) {
                let destinationRange = newSummarySheet.getRange(
                    currentRow,
                    1,
                    remainingRows,
                    9,
                );
                let copyRange = newSummarySheet.getRange(
                    nextRow,
                    1,
                    remainingRows,
                    9,
                );
                copyRange.copyTo(destinationRange);
            }

            // Clear the last row (which is now a duplicate after the shift) and shrink the range.
            let lastRowRange = newSummarySheet.getRange(lastRow, 1, 1, 9);
            lastRowRange.clearContent();
            lastRowRange.clearFormat();
            lastRowRange.clearDataValidations();

            // The effective data range is now one row shorter.
            lastRow--;
        }

        // Copy the full RECORDS sheet into the new spreadsheet and rename it.
        const newRecordsSheet = recordsSheet.copyTo(newSpreadsheet);
        newRecordsSheet.setName("RECORDS");

        // Build the list of RECORDS rows to delete. The RECORDS sheet repeats a block of rows
        // for each of the 16 weeks: each block starts at row 21 + (weekIndex * (numOfStudents + 7)).
        // Within each block, the student rows are offset by their original SUMMARY row numbers.
        // NOTE: the use of removedLearnersRows (absolute SUMMARY row numbers) as offsets here
        // may be a bug — see refactoring notes.
        let rowsToRemove = [];
        for (let i = 0; i < 16; i++) {
            let baseWeekRow = 21 + i * (numOfStudents + 7); // first row of this week's block
            for (let j = 0; j < removedLearnersRows.length; j++) {
                let baseRow = removedLearnersRows[j];
                let row = baseWeekRow + baseRow;
                rowsToRemove.push(row);
            }
        }

        // Delete rows in reverse order so that earlier row indices are not invalidated
        // by the deletion of later rows (deleting from the bottom up keeps indices stable).
        for (let i = rowsToRemove.length - 1; i >= 0; i--) {
            let rowToRemove = rowsToRemove[i];
            newRecordsSheet.deleteRow(rowToRemove);
        }

        // Remove the default "Sheet1" that Google Sheets creates automatically with every new spreadsheet.
        let defaultSheet = newSpreadsheet.getSheetByName("Sheet1");
        if (defaultSheet) newSpreadsheet.deleteSheet(defaultSheet);

        // Ensure the PARTNER_REPORTS log sheet exists in the source spreadsheet.
        // If it doesn't yet exist, create it by copying from the template sheet.
        // NOTE: this check runs on every loop iteration but only needs to run once — a candidate
        // for hoisting out of the loop.
        let partnerReportsSheet = spreadsheet.getSheetByName("PARTNER_REPORTS");

        if (!partnerReportsSheet) {
            const partnerReportsTemplateSheet = spreadsheet.getSheetByName(
                "TEMPLATE_PARTNER_REPORTS",
            );
            partnerReportsSheet =
                partnerReportsTemplateSheet.copyTo(spreadsheet);
            partnerReportsSheet.setName("PARTNER_REPORTS");
        }

        // Append a new log entry to PARTNER_REPORTS: [date, partner name, week, URL].
        const newRow = partnerReportsSheet.getLastRow() + 1;
        const todayDate = startTime.toISOString().split("T")[0]; // extract YYYY-MM-DD from ISO string
        const formattedTodayDate = formatDate(todayDate);        // helper defined elsewhere in the project
        const newSpreadsheetUrl = newSpreadsheet.getUrl();

        const newPartnerReportValues = [
            [formattedTodayDate, partnerName, currentWeek, newSpreadsheetUrl],
        ];
        partnerReportsSheet
            .getRange(newRow, 1, 1, 4)
            .setValues(newPartnerReportValues);

        // Make the PARTNER_REPORTS sheet visible in case it was hidden.
        partnerReportsSheet.showSheet();
    }

    let endTime = new Date();
    let elapsedTime = endTime - startTime;
    Logger.log(`Finished - ${elapsedTime / 1000} seconds`);
}
