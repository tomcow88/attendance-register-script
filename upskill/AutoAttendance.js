/**
 * Checks attendance for today's date only.
 *
 * Determines the current week and date from the schedule, runs checkAttendance() for
 * that date, then appends a log entry to the LOGS sheet recording the date, time,
 * and which session types were updated.
 */
function checkAttendanceToday() {
    let startTime = new Date();
    // Pass false to use the real current date. Replace with a YYYY-MM-DD string for testing:
    // let todayDate = '2025-04-28';
    let todayDate = false;
    const currentWeekData = setCurrentWeek(todayDate);
    const currentDate = currentWeekData.todayDate;
    const weekNum = currentWeekData.currentWeekNum;
    const sessionsUpdated = checkAttendance(currentDate, weekNum);

    let endTime = new Date();
    let elapsedTime = endTime - startTime;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName("LOGS");
    logsSheet
        .getRange(logsSheet.getLastRow() + 1, 1, 1, 1)
        .setValue(
            `${currentDate} - ${Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm:ss")} - Sessions Updated: ${JSON.stringify(sessionsUpdated)}`,
        );

    Logger.log("Total Elapsed time: " + elapsedTime / 1000 + " seconds");
}

/**
 * Re-checks attendance for every session day in the schedule, overwriting auto-populated values.
 *
 * Iterates through each week and each day within that week, calling checkAttendance() for
 * each date.
 */
function checkAllAttendance() {
    const ui = SpreadsheetApp.getUi();
    let startTime = new Date();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const scheduleData = JSON.parse(databaseSheet.getRange(3, 23, 1, 1).getValue());
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;

    for (let i = 0; i < weeks; i++) {
        let week = schedule[i];

        for (let j = 0; j < week.length; j++) {
            let date = week[j].date;
            let weekNum = i + 1;
            checkAttendance(date, weekNum, scheduleData);
            Utilities.sleep(500);
        }
    }

    let elapsedTime = new Date() - startTime;
    Logger.log("Total Elapsed time: " + elapsedTime / 1000 + " seconds");
    ui.alert(
        "Finished checking all attendance\nElapsed time: " +
            elapsedTime / 1000 +
            " seconds",
    );
}

/**
 * Checks attendance for a single date within a specific week.
 *
 * For each session type (SU, SD, GS, SME, CC), locates the corresponding Google Drive folder
 * and attendance file for the given date, then calls updateAttendance() to write results to RECORDS.
 *
 * Accepts an optional pre-parsed scheduleData object to avoid re-reading DATABASE when called
 * in a loop (e.g. from checkAllAttendance).
 *
 * Returns a sessionsUpdated object indicating which session types were successfully updated.
 */
function checkAttendance(todayDate, weekNum, scheduleData) {
    let startTime = new Date();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    if (!scheduleData) {
        scheduleData = JSON.parse(databaseSheet.getRange(3, 23, 1, 1).getValue());
    }
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;

    // Session type abbreviations, matched by index to calendarNames and glhs arrays.
    const abreviations = ["SU", "SD", "GS", "SME", "CC"];
    const calendarNames = getCalendarNames();
    const glhs = [0.5, 0.5, 1, 2, 1]; // GLH awarded per session type
    let sessionsUpdated = { SU: false, SD: false, GS: false, SME: false, CC: false };

    for (let i = 0; i < weeks; i++) {
        if (i + 1 != weekNum) continue;
        let week = schedule[i];

        for (let j = 0; j < week.length; j++) {
            let dayInWeek = week[j];
            let date = dayInWeek.date;
            let day = dayInWeek.day;

            if (date != todayDate) continue;

            for (let k = 0; k < abreviations.length; k++) {
                let abreviation = abreviations[k];
                let calendarName = calendarNames[k];
                if (!calendarName) continue;

                let glh = glhs[k];
                let parentFolderId = getReportFolderId(abreviation);
                if (parentFolderId == "null") continue;

                let host = getDeliveryTeam(abreviation);
                let updatedAttendance = updateAttendance(
                    parentFolderId,
                    date,
                    calendarName,
                    day,
                    abreviation,
                    glh,
                    j,
                    i,
                    host,
                );
                if (updatedAttendance == "updated") {
                    sessionsUpdated[abreviation] = true;
                }
            }
        }
    }

    Logger.log(todayDate + " - Elapsed time: " + (new Date() - startTime) / 1000 + " seconds");
    return sessionsUpdated;
}

/**
 * Writes attendance values for a single session to the RECORDS sheet.
 *
 * Steps:
 *  1. Locates the Drive folder matching the session date and calendar name.
 *  2. Finds the "Attendance" file within that folder.
 *  3. Reads the attendee list (names and emails) from the file's "Attendees" sheet.
 *  4. Matches each student against the attendee list by email first, then by first name.
 *  5. Preserves any existing positive numeric values in RECORDS — cells with a value > 0
 *     are assumed to have been manually entered and are not overwritten.
 *     Note: non-numeric manual entries (e.g. text) and manually entered zeros are still overwritten.
 *  6. Writes the final attendance array plus the host signature row to RECORDS.
 *
 * Returns "updated" on success, or false if the folder or file could not be found.
 */
function updateAttendance(
    parentFolderId,
    date,
    calendarName,
    day,
    abreviation,
    glh,
    dayNum,
    weekNum,
    host,
) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const numOfStudents = getNumOfStudents();
    const parentFolder = DriveApp.getFolderById(parentFolderId);

    // Search for a sub-folder whose name contains the session date string.
    const folders = parentFolder.searchFolders(`fullText contains '${date}'`);
    if (!folders.hasNext()) return false;

    let correctFolder;
    while (folders.hasNext()) {
        let folder = folders.next();
        if (folder.getName().includes(calendarName)) {
            correctFolder = folder;
        }
    }
    if (!correctFolder) return false;

    // Find the attendance file within the folder (identified by "Attendance" in the file name).
    const files = correctFolder.getFiles();
    let correctFile;
    while (files.hasNext()) {
        let file = files.next();
        if (file.getName().includes("Attendance")) {
            correctFile = file;
        }
    }
    if (!correctFile) return false;

    const spreadsheetFile = SpreadsheetApp.openById(correctFile.getId());
    const attendeeSheet = spreadsheetFile.getSheetByName("Attendees");
    const firstNames = attendeeSheet.getRange(2, 1, attendeeSheet.getLastRow(), 1).getValues().flat();
    const emails = attendeeSheet.getRange(2, 3, attendeeSheet.getLastRow(), 1).getValues().flat();

    // Use separate working copies so matched attendees can be removed, preventing double-counting.
    let availableFirstNames = [...firstNames];
    let availableEmails = [...emails];

    let attendanceValues = [];
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");
    let lastAttendedRange = summarySheet.getRange(2, 5, numOfStudents, 1);
    let lastAttendedValues = lastAttendedRange.getValues();
    const students = getStudents();

    const [yearSplit, monthSplit, daySplit] = date.split("-");
    const formattedDate = `${daySplit}/${monthSplit}/${yearSplit}`;

    for (let i = 0; i < students.length; i++) {
        let attendanceValue = [];
        let lastAttendedValue = lastAttendedValues[i];
        let student = students[i];

        // Inactive students always receive "X" — no attendance check needed.
        if (student.status != "Active") {
            attendanceValue.push("X");
            attendanceValues.push(attendanceValue);
            continue;
        }

        // Guest speaker, SME, and career coach sessions are marked "-" on project/hackathon days
        // since those sessions do not run during project weeks.
        if (
            isProjectOrHackathonDay(day) &&
            (abreviation == "GS" || abreviation == "SME" || abreviation == "CC")
        ) {
            attendanceValue.push("-");
            attendanceValues.push(attendanceValue);
            continue;
        }

        let meetEmails = student.meetEmails;
        let meetNames = student.meetNames;

        // Try matching by email first (more reliable), then fall back to first name.
        for (let j = 0; j < meetEmails.length; j++) {
            if (attendanceValue.length > 0) continue;
            if (availableEmails.includes(meetEmails[j])) {
                attendanceValue.push(glh);
                attendanceValues.push(attendanceValue);
                lastAttendedValue.pop();
                lastAttendedValue.push(formattedDate);
                let matchedIndex = availableEmails.indexOf(meetEmails[j]);
                availableEmails.splice(matchedIndex, 1);
                availableFirstNames.splice(matchedIndex, 1);
            }
        }

        for (let j = 0; j < meetNames.length; j++) {
            if (attendanceValue.length > 0) continue;
            if (availableFirstNames.includes(meetNames[j])) {
                attendanceValue.push(glh);
                attendanceValues.push(attendanceValue);
                lastAttendedValue.pop();
                lastAttendedValue.push(formattedDate);
                let matchedIndex = availableFirstNames.indexOf(meetNames[j]);
                availableEmails.splice(matchedIndex, 1);
                availableFirstNames.splice(matchedIndex, 1);
            }
        }

        if (attendanceValue.length == 0) {
            attendanceValue.push(0);
            attendanceValues.push(attendanceValue);
        }
    }

    // Append the host/facilitator signature as the final row of the attendance block.
    attendanceValues.push([host]);

    const defaultStartRow = 23;
    const numberOfSourceRows = 7 + numOfStudents;
    const sessionCol = getSessonCol(dayNum + 1, abreviation);
    const sessionStartRow = defaultStartRow + numberOfSourceRows * weekNum;
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const sessionRange = recordsSheet.getRange(sessionStartRow, sessionCol, numOfStudents + 1, 1);

    // Preserve any existing non-empty, non-zero value — these are assumed to be manually entered.
    // Zero is still re-evaluated since it's indistinguishable from a previous auto-check result.
    const currentSessionValues = sessionRange.getValues();
    for (let i = 0; i < currentSessionValues.length; i++) {
        const currentValue = currentSessionValues[i][0];
        if (currentValue !== "" && currentValue !== 0) {
            attendanceValues[i] = currentSessionValues[i];
        }
    }

    sessionRange.setValues(attendanceValues);
    lastAttendedRange.setValues(lastAttendedValues);

    // Apply the Caveat font to the host signature row so it renders as a handwritten signature.
    recordsSheet
        .getRange(sessionStartRow + numOfStudents, sessionCol, 1, 1)
        .setFontFamily("Caveat");

    // Write the GLH value into the schedule header row for this session.
    // On project/hackathon days, also write the project GLH into the PRO column.
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const proGlh = databaseSheet.getRange(8, 9, 1, 1).getValue();
    const proCol = getSessonCol(dayNum + 1, "PRO");
    const scheduleRow = sessionStartRow - 2;

    if (isProjectOrHackathonDay(day)) {
        recordsSheet.getRange(scheduleRow, proCol, 1, 1).setValue(proGlh);
        if (abreviation == "SU" || abreviation == "SD") {
            recordsSheet.getRange(scheduleRow, sessionCol, 1, 1).setValue(glh);
        }
    } else {
        recordsSheet.getRange(scheduleRow, sessionCol, 1, 1).setValue(glh);
    }

    return "updated";
}

/**
 * Returns the Google Drive folder ID for the given session type abbreviation,
 * read from the DATABASE sheet. Returns false if no folder is configured.
 *
 * SU, SD, and GS share the same folder; SME and CC each have their own.
 */
function getReportFolderId(abreviation) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    let reportFolderIdsString = false;
    if (abreviation == "SU" || abreviation == "SD" || abreviation == "GS") {
        reportFolderIdsString = databaseSheet.getRange(3, 17, 1, 1).getValue();
    } else if (abreviation == "SME") {
        reportFolderIdsString = databaseSheet.getRange(4, 17, 1, 1).getValue();
    } else if (abreviation == "CC") {
        reportFolderIdsString = databaseSheet.getRange(5, 17, 1, 1).getValue();
    }

    if (!reportFolderIdsString) return false;

    // Folder IDs are stored as a JSON array; the most recently added ID is used.
    const ids = JSON.parse(reportFolderIdsString);
    return ids[ids.length - 1];
}

/**
 * Returns the delivery team member name (used as the session host/signature) for the given
 * session type, read from the DATABASE sheet. Returns false if not configured.
 */
function getDeliveryTeam(abreviation) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    let signature = false;
    if (abreviation == "SU" || abreviation == "SD" || abreviation == "GS") {
        signature = databaseSheet.getRange(8, 15, 1, 1).getValue();
    } else if (abreviation == "SME") {
        signature = databaseSheet.getRange(9, 15, 1, 1).getValue();
    } else if (abreviation == "CC") {
        signature = databaseSheet.getRange(10, 15, 1, 1).getValue();
    }

    return signature || false;
}

/**
 * Reads the current calendar names for all five session types from DATABASE.
 * Returns an array of [SU, SD, GS, SME, CC] calendar names.
 * Each entry is the last value in a JSON array stored in DATABASE, or false if not configured.
 */
function getCalendarNames() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const namesStrings = databaseSheet.getRange(3, 11, 5, 1).getValues().flat();

    let names = [];
    for (let i = 0; i < namesStrings.length; i++) {
        let nameString = namesStrings[i];
        if (!nameString) {
            names.push(false);
            continue;
        }
        // Calendar names are stored as JSON arrays; the most recently added name is used.
        let nameArr = JSON.parse(nameString);
        names.push(nameArr[nameArr.length - 1]);
    }
    return names;
}

/**
 * Builds and returns an array of student objects, each containing:
 *  - meetNames: array of display name variants used for matching against attendee lists
 *  - meetEmails: array of anonymized emails used for matching
 *  - position: 1-based row index in the student list
 *  - status: "Active", "Withdrawn", etc.
 */
function getStudents() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");

    const numOfStudents = databaseSheet.getRange(3, 22, 1, 1).getValue();
    const meetNamesStrings = databaseSheet.getRange(3, 4, numOfStudents, 1).getValues().flat();
    const meetEmailsStrings = databaseSheet.getRange(3, 5, numOfStudents, 1).getValues().flat();
    const statusOfStudents = summarySheet.getRange(2, 4, numOfStudents, 1).getValues().flat();

    let students = [];
    for (let i = 0; i < statusOfStudents.length; i++) {
        students.push({
            meetNames: JSON.parse(meetNamesStrings[i]),
            meetEmails: JSON.parse(meetEmailsStrings[i]),
            position: i + 1,
            status: statusOfStudents[i],
        });
    }
    return students;
}

/**
 * Returns the RECORDS sheet column number for a given session type and day number within the week.
 * Column arrays are stored as JSON in DATABASE; dayNum is the 1-based index into that array.
 */
function getSessonCol(dayNum, abreviation) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    let sessionColsString = false;
    if (abreviation == "SU") sessionColsString = databaseSheet.getRange(3, 10, 1, 1).getValue();
    else if (abreviation == "SD") sessionColsString = databaseSheet.getRange(4, 10, 1, 1).getValue();
    else if (abreviation == "GS") sessionColsString = databaseSheet.getRange(5, 10, 1, 1).getValue();
    else if (abreviation == "SME") sessionColsString = databaseSheet.getRange(6, 10, 1, 1).getValue();
    else if (abreviation == "CC") sessionColsString = databaseSheet.getRange(7, 10, 1, 1).getValue();
    else if (abreviation == "PRO") sessionColsString = databaseSheet.getRange(8, 10, 1, 1).getValue();

    return JSON.parse(sessionColsString)[dayNum];
}

/**
 * Returns the total number of students for this cohort, read from DATABASE.
 */
function getNumOfStudents() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    return databaseSheet.getRange(3, 22, 1, 1).getValue();
}
