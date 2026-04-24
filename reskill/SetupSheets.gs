function setupSheets() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let setupSheet = spreadsheet.getSheetByName("SETUP");

    if (!setupSheet) {
        const templateSetupSheet = spreadsheet.getSheetByName("TEMPLATE_SETUP");
        const newSetupSheet = templateSetupSheet.copyTo(spreadsheet);
        newSetupSheet.setName("SETUP");
        return;
    }

    const cohortDataRange = setupSheet.getRange(1, 2, 7, 1);

    const numOfStudents = setupSheet.getLastRow() - 13;
    if (numOfStudents < 1) {
        return;
    }

    let numOfWeeksOnly = setupSheet.getRange(8, 2, 1, 1).getValue();
    let numOfSessionsPerWeek = setupSheet.getRange(9, 2, 1, 1).getValue();
    numOfSessionsPerWeek = Number(numOfSessionsPerWeek);
    const glhPerSession = setupSheet.getRange(10, 2, 1, 1).getValue();
    let numOfExtraSessions = setupSheet.getRange(11, 2, 1, 1).getValue();
    numOfExtraSessions = Number(numOfExtraSessions);
    let numOfWeeks = Number(numOfWeeksOnly) + numOfExtraSessions;

    const startDate = setupSheet.getRange(4, 2, 1, 1).getDisplayValue();

    const firstNamesRange = setupSheet.getRange(14, 1, numOfStudents, 1);
    const lastNamesRange = setupSheet.getRange(14, 2, numOfStudents, 1);

    let attendanceSheet = spreadsheet.getSheetByName("ATTENDANCE");
    if (!attendanceSheet) {
        const templateAttendanceSheet = spreadsheet.getSheetByName(
            "TEMPLATE_ATTENDANCE",
        );
        attendanceSheet = templateAttendanceSheet.copyTo(spreadsheet);
        attendanceSheet.setName("ATTENDANCE");
        attendanceSheet.showSheet();
    }
    spreadsheet.setActiveSheet(attendanceSheet);
    spreadsheet.moveActiveSheet(1);

    if (numOfSessionsPerWeek > 1) {
        const defaultAttendanceCol = attendanceSheet.getRange(1, 3, 4, 1);
        attendanceSheet.insertColumnsAfter(3, numOfSessionsPerWeek);

        for (let i = 1; i <= numOfSessionsPerWeek; i++) {
            let colNum = 3 + i;
            defaultAttendanceCol.copyTo(
                attendanceSheet.getRange(1, colNum, 4, 1),
            );
            attendanceSheet.getRange(3, colNum, 1, 1).setValue(`Session ${i}`);
        }

        attendanceSheet.deleteColumn(3);

        let endingLetter = numberToLetter(numOfSessionsPerWeek + 2);
        attendanceSheet
            .getRange(4, 3 + numOfSessionsPerWeek, 1, 1)
            .setValue(`=COUNTIF(C4:${endingLetter}4, TRUE) * SUMMARY!$F$3`);
    }

    const totalAttendanceCols = 5 + numOfSessionsPerWeek;

    const defaultStudentAttendanceRange = attendanceSheet.getRange(
        4,
        1,
        1,
        totalAttendanceCols,
    );
    defaultStudentAttendanceRange.copyTo(
        attendanceSheet.getRange(4, 1, numOfStudents, totalAttendanceCols),
    );
    firstNamesRange.copyTo(attendanceSheet.getRange(4, 1, numOfStudents, 1));
    lastNamesRange.copyTo(attendanceSheet.getRange(4, 2, numOfStudents, 1));
    attendanceSheet.getRange(2, 2, 1, 1).setValue(startDate);

    const rowDiff = 3 + numOfStudents;
    const defaultAttendanceRange = attendanceSheet.getRange(
        2,
        1,
        rowDiff,
        totalAttendanceCols,
    );
    for (let week = 2; week <= numOfWeeks; week++) {
        let row = (week - 1) * rowDiff + 2;
        defaultAttendanceRange.copyTo(
            attendanceSheet.getRange(row, 1, rowDiff, totalAttendanceCols),
        );
        let heading;

        if (week - numOfWeeksOnly > 0) {
            heading = `Extra Session ${week - numOfWeeksOnly}`;
        } else {
            heading = `Week ${week}`;
        }
        attendanceSheet.getRange(row, 1, 1, 1).setValue(heading);
        let date = addWeeks(startDate, week - 1);
        attendanceSheet.getRange(row, 2, 1, 1).setValue(date);
    }

    let summarySheet = spreadsheet.getSheetByName("SUMMARY");
    if (!summarySheet) {
        const templateSummarySheet =
            spreadsheet.getSheetByName("TEMPLATE_SUMMARY");
        summarySheet = templateSummarySheet.copyTo(spreadsheet);
        summarySheet.setName("SUMMARY");
        summarySheet.showSheet();
    }

    cohortDataRange.copyTo(summarySheet.getRange(1, 3, 7, 1));
    summarySheet.getRange(1, 6, 1, 1).setValue(numOfWeeksOnly);
    summarySheet.getRange(2, 6, 1, 1).setValue(numOfSessionsPerWeek);
    summarySheet.getRange(3, 6, 1, 1).setValue(glhPerSession);
    summarySheet.getRange(4, 6, 1, 1).setValue(numOfExtraSessions);
    summarySheet.getRange(5, 6, 1, 1).setValue(numOfStudents);

    let totalGlhLetter = numberToLetter(3 + numOfSessionsPerWeek);
    let totalGlhFormula = `=SUMIFS(ATTENDANCE!$${totalGlhLetter}:$${totalGlhLetter}, ATTENDANCE!$A:$A, $A14, ATTENDANCE!$B:$B, $B14)`;
    summarySheet.getRange(14, 5, 1, 1).setValue(totalGlhFormula);

    const defaultStudentRange = summarySheet.getRange(14, 1, 1, 6);
    defaultStudentRange.copyTo(summarySheet.getRange(14, 1, numOfStudents, 6));
    firstNamesRange.copyTo(summarySheet.getRange(14, 1, numOfStudents, 1));
    lastNamesRange.copyTo(summarySheet.getRange(14, 2, numOfStudents, 1));

    const attendanceCol = attendanceSheet.getRange(
        2,
        3 + numOfSessionsPerWeek,
        attendanceSheet.getLastRow() - 1,
        3,
    );
    const colValues = attendanceCol.getFormulas();
    attendanceCol.setFormulas(colValues);

    spreadsheet.setActiveSheet(summarySheet);
    spreadsheet.moveActiveSheet(1);
}

function addWeeks(dateStr, weeks) {
    // Parse the plain text date (dd/mm/yyyy)
    const [day, month, year] = dateStr.split("/").map(Number);
    const date = new Date(year, month - 1, day);

    // Add weeks (7 days * number of weeks)
    date.setDate(date.getDate() + weeks * 7);

    // Format back to dd/mm/yyyy
    const newDay = String(date.getDate()).padStart(2, "0");
    const newMonth = String(date.getMonth() + 1).padStart(2, "0");
    const newYear = date.getFullYear();

    return `${newDay}/${newMonth}/${newYear}`;
}

function numberToLetter(n) {
    if (n < 1) return "";
    let letters = "";
    while (n > 0) {
        let remainder = (n - 1) % 26;
        letters = String.fromCharCode(65 + remainder) + letters;
        n = Math.floor((n - 1) / 26);
    }
    return letters;
}
