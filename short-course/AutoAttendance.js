function checkAttendanceToday() {
    let startTime = new Date();
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
    const lastRow = logsSheet.getLastRow();
    logsSheet
        .getRange(lastRow + 1, 1, 1, 1)
        .setValue(
            `${currentDate} - ${Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm:ss")} - Sessions Updated: ${JSON.stringify(sessionsUpdated)}`,
        );

    Logger.log("Total Elapsed time: " + elapsedTime / 1000 + " seconds");
}

function checkAllAttendance() {
    const ui = SpreadsheetApp.getUi();
    let startTime = new Date();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const scheduleDataString = databaseSheet.getRange(3, 23, 1, 1).getValue();
    const scheduleData = JSON.parse(scheduleDataString);
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;

    for (let i = 0; i < weeks; i++) {
        if (i > 2) break;
        let week = schedule[i];

        for (let j = 0; j < week.length; j++) {
            let dayInWeek = week[j];

            let date = dayInWeek.date;
            let weekNum = i + 1;

            checkAttendance(date, weekNum);
        }
    }

    let endTime = new Date();

    let elapsedTime = endTime - startTime;

    Logger.log("Total Elapsed time: " + elapsedTime / 1000 + " seconds");
    ui.alert(
        "Finished checking all attendance\nElapsed time: " +
            elapsedTime / 1000 +
            " seconds",
    );
}

function checkAttendance(todayDate, weekNum) {
    let startTime = new Date();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const scheduleDataString = databaseSheet.getRange(3, 23, 1, 1).getValue();
    const scheduleData = JSON.parse(scheduleDataString);
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;
    const abreviations = ["SU", "SD", "GS", "SME", "CC"];
    const calendarNames = getCalendarNames();
    const glhs = [0.5, 0.5, 1, 2, 1];
    let sessionsUpdated = {
        SU: false,
        SD: false,
        GS: false,
        SME: false,
        CC: false,
    };

    for (let i = 0; i < weeks; i++) {
        if (i + 1 != weekNum) continue;
        let week = schedule[i];
        for (let j = 0; j < week.length; j++) {
            let dayInWeek = week[j];
            let date = dayInWeek.date;
            let day = dayInWeek.day;
            if (date != todayDate) {
                continue;
            }
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
                if (!updatedAttendance) {
                    continue;
                }
            }
        }
    }

    let endTime = new Date();

    let elapsedTime = endTime - startTime;

    Logger.log(
        todayDate + " - Elapsed time: " + elapsedTime / 1000 + " seconds",
    );

    return sessionsUpdated;
}

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
    const folders = parentFolder.searchFolders(`fullText contains '${date}'`);
    if (!folders.hasNext()) {
        return false;
    }
    let correctFolder;
    while (folders.hasNext()) {
        let folder = folders.next();
        let folderName = folder.getName();
        if (folderName.includes(calendarName)) {
            correctFolder = folder;
        }
    }
    if (!correctFolder) return false;
    const files = correctFolder.getFiles();
    let correctFile;
    while (files.hasNext()) {
        let file = files.next();
        let fileName = file.getName();
        if (fileName.includes("Attendance")) {
            correctFile = file;
        }
    }
    if (!correctFile) return false;
    const fileId = correctFile.getId();
    const spreadsheetFile = SpreadsheetApp.openById(fileId);
    const attendeeSheet = spreadsheetFile.getSheetByName("Attendees");
    const firstNames = attendeeSheet
        .getRange(2, 1, attendeeSheet.getLastRow(), 1)
        .getValues()
        .flat();
    const emails = attendeeSheet
        .getRange(2, 3, attendeeSheet.getLastRow(), 1)
        .getValues()
        .flat();
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
        let status = student.status;

        if (status != "Active") {
            attendanceValue.push("X");
            attendanceValues.push(attendanceValue);
            continue;
        }

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

        for (let j = 0; j < meetEmails.length; j++) {
            let meetEmail = meetEmails[j];

            if (attendanceValue.length > 0) continue;

            if (availableEmails.includes(meetEmail)) {
                attendanceValue.push(glh);
                attendanceValues.push(attendanceValue);
                lastAttendedValue.pop();
                lastAttendedValue.push(formattedDate);
                let matchedIndex = availableEmails.indexOf(meetEmail);
                availableEmails.splice(matchedIndex, 1);
                availableFirstNames.splice(matchedIndex, 1);
                continue;
            }
        }

        for (let j = 0; j < meetNames.length; j++) {
            let meetName = meetNames[j];

            if (attendanceValue.length > 0) continue;

            if (firstNames.includes(meetName)) {
                attendanceValue.push(glh);
                attendanceValues.push(attendanceValue);
                lastAttendedValue.pop();
                lastAttendedValue.push(formattedDate);
                let matchedIndex = availableFirstNames.indexOf(meetName);
                availableEmails.splice(matchedIndex, 1);
                availableFirstNames.splice(matchedIndex, 1);
                continue;
            }
        }

        if (attendanceValue.length == 0) {
            attendanceValue.push(0);
            attendanceValues.push(attendanceValue);
            continue;
        }
    }

    let hosts = [];
    hosts.push(host);
    attendanceValues.push(hosts);

    const defaultStartRow = 23;
    const numberOfSourceRows = 7 + numOfStudents;
    const sessionCol = getSessonCol(dayNum + 1, abreviation);
    const sessionStartRow = defaultStartRow + numberOfSourceRows * weekNum;
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const sessionRange = recordsSheet.getRange(
        sessionStartRow,
        sessionCol,
        numOfStudents + 1,
        1,
    );

    const currentSessionValues = sessionRange.getValues();
    for (let i = 0; i < currentSessionValues.length; i++) {
        let numberValue = Number(currentSessionValues[i]);
        if (numberValue > 0) {
            attendanceValues[i] = currentSessionValues[i];
        }
    }

    sessionRange.setValues(attendanceValues);
    lastAttendedRange.setValues(lastAttendedValues);
    recordsSheet
        .getRange(sessionStartRow + numOfStudents, sessionCol, 1, 1)
        .setFontFamily("Caveat");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const scheduleDataString = databaseSheet.getRange(3, 23, 1, 1).getValue();
    const scheduleData = JSON.parse(scheduleDataString);

    const scheduleRow = sessionStartRow - 2;
    const proGlh = databaseSheet.getRange(8, 9, 1, 1).getValue();
    const proCol = getSessonCol(dayNum + 1, "PRO");

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

function getReportFolderId(abreviation) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    let reportFolderIdsString = false;
    if (abreviation == "SU" || abreviation == "SD" || abreviation == "GS") {
        reportFolderIdsString = databaseSheet.getRange(3, 17, 1, 1).getValue();
    }
    if (abreviation == "SME") {
        reportFolderIdsString = databaseSheet.getRange(4, 17, 1, 1).getValue();
    }
    if (abreviation == "CC") {
        reportFolderIdsString = databaseSheet.getRange(5, 17, 1, 1).getValue();
    }

    if (!reportFolderIdsString) {
        return false;
    }

    const ids = JSON.parse(reportFolderIdsString);
    return ids[ids.length - 1];
}

function getDeliveryTeam(abreviation) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    let signature = false;
    if (abreviation == "SU" || abreviation == "SD" || abreviation == "GS") {
        signature = databaseSheet.getRange(8, 15, 1, 1).getValue();
    }
    if (abreviation == "SME") {
        signature = databaseSheet.getRange(9, 15, 1, 1).getValue();
    }
    if (abreviation == "CC") {
        signature = databaseSheet.getRange(10, 15, 1, 1).getValue();
    }

    if (!signature) {
        return false;
    }

    return signature;
}

function getCalendarNames() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const namesStrings = databaseSheet.getRange(3, 11, 5, 1).getValues().flat();

    let names = [];
    for (let i = 0; i < namesStrings.length; i++) {
        let nameString = namesStrings[i];
        if (!nameString) {
            names.push(false);
        }
        let nameArr = JSON.parse(nameString);
        let name = nameArr[nameArr.length - 1];
        names.push(name);
    }
    return names;
}

function getStudents() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");

    const numOfStudents = databaseSheet.getRange(3, 22, 1, 1).getValue();
    const meetNamesStrings = databaseSheet
        .getRange(3, 4, numOfStudents, 1)
        .getValues()
        .flat();
    const meetEmailsStrings = databaseSheet
        .getRange(3, 5, numOfStudents, 1)
        .getValues()
        .flat();
    const statusOfStudents = summarySheet
        .getRange(2, 4, numOfStudents, 1)
        .getValues()
        .flat();

    let students = [];

    for (let i = 0; i < statusOfStudents.length; i++) {
        let meetNames = JSON.parse(meetNamesStrings[i]);
        let meetEmails = JSON.parse(meetEmailsStrings[i]);
        let status = statusOfStudents[i];
        let student = {
            meetNames: meetNames,
            meetEmails: meetEmails,
            position: i + 1,
            status: status,
        };
        students.push(student);
    }

    return students;
}

function getSessonCol(dayNum, abreviation) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    let sessionColsString = false;
    if (abreviation == "SU") {
        sessionColsString = databaseSheet.getRange(3, 10, 1, 1).getValue();
    }
    if (abreviation == "SD") {
        sessionColsString = databaseSheet.getRange(4, 10, 1, 1).getValue();
    }
    if (abreviation == "GS") {
        sessionColsString = databaseSheet.getRange(5, 10, 1, 1).getValue();
    }
    if (abreviation == "SME") {
        sessionColsString = databaseSheet.getRange(6, 10, 1, 1).getValue();
    }
    if (abreviation == "CC") {
        sessionColsString = databaseSheet.getRange(7, 10, 1, 1).getValue();
    }
    if (abreviation == "PRO") {
        sessionColsString = databaseSheet.getRange(8, 10, 1, 1).getValue();
    }

    const sessionColsArr = JSON.parse(sessionColsString);
    const col = sessionColsArr[dayNum];
    return col;
}

function getNumOfStudents() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const numOfStudents = databaseSheet.getRange(3, 22, 1, 1).getValue();
    return numOfStudents;
}
