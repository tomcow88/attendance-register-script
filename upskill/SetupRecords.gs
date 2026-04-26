/**
 * Builds the RECORDS sheet structure for the full cohort schedule.
 *
 * Steps:
 *  1. Copies the template name row to create one row per student, populating full names.
 *  2. Copies the weekly block template once per week in the schedule, appending blocks downward.
 *  3. Writes the correct week label, dates, day labels, PRO formulas, and percentage formulas
 *     into each weekly block, and highlights project/hackathon days.
 *  4. Deletes the original template block and calls setCurrentWeek to initialise the week display.
 */
function setupRecords() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const recordsLastCol = recordsSheet.getLastColumn();

    const numOfStudents = databaseSheet.getRange(3, 23, 1, 1).getValue();
    const scheduleData = JSON.parse(databaseSheet.getRange(3, 24, 1, 1).getValue());
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;
    const fullNames = databaseSheet
        .getRange(3, 1, numOfStudents, 1)
        .getValues()
        .flat();

    // Copy the template name row across numOfStudents rows to create the student name block.
    const defaultNameRow = 23;
    const defaultNameRowRange = recordsSheet.getRange(defaultNameRow, 1, 1, recordsLastCol);
    const targetNameRowRange = recordsSheet.getRange(defaultNameRow, 1, numOfStudents, recordsLastCol);
    defaultNameRowRange.copyTo(targetNameRowRange);

    for (let i = 0; i < numOfStudents; i++) {
        recordsSheet.getRange(defaultNameRow + i, 1, 1, 1).setValue(fullNames[i]);
    }

    // Copy the full weekly block template (header rows + student rows) once per week,
    // appending each block immediately below the previous one.
    const sourceStartRow = 17;
    const numberOfSourceRows = 7 + numOfStudents;
    const sourceRange = recordsSheet.getRange(sourceStartRow, 1, numberOfSourceRows, recordsLastCol);

    let targetStartRow = sourceStartRow + numberOfSourceRows;
    for (let i = 0; i < weeks; i++) {
        sourceRange.copyTo(recordsSheet.getRange(targetStartRow, 1, numberOfSourceRows, recordsLastCol));
        targetStartRow += numberOfSourceRows;
    }

    // Layout constants for week label, date, and day rows within each weekly block.
    const defaultWeekRow = 18;
    const defaultWeekCol = 1;
    const defaultDateDayCol = 14;
    const defaultDateDayCols = 10;
    const weekRowDifference = numberOfSourceRows;
    const dateDayColDifference = defaultDateDayCols + 1;

    const percentCol = 12;
    const proCols = JSON.parse(databaseSheet.getRange(8, 11, 1, 1).getValue());
    const suCols = JSON.parse(databaseSheet.getRange(3, 11, 1, 1).getValue());
    const proGlh = databaseSheet.getRange(8, 10, 1, 1).getValue();

    // PRO (project hours) formula: awards project GLH if either the SU or SD column for that
    // student has a positive value or a "-" marker, otherwise 0. Propagates "X" for inactive students.
    const proFormula = `=IF(OR(INDIRECT(ADDRESS(ROW(), COLUMN()-4))="X", INDIRECT(ADDRESS(ROW(), COLUMN()-5))="X"), "X",
    IF(AND(INDIRECT(ADDRESS(ROW(), COLUMN()-4))="-", INDIRECT(ADDRESS(ROW(), COLUMN()-5))="-"), "-",
    IF(OR(
        ISNUMBER(INDIRECT(ADDRESS(ROW(), COLUMN()-4))) * (INDIRECT(ADDRESS(ROW(), COLUMN()-4)) > 0),
        ISNUMBER(INDIRECT(ADDRESS(ROW(), COLUMN()-5))) * (INDIRECT(ADDRESS(ROW(), COLUMN()-5)) > 0)
    ), ${proGlh}, 0)))`;
    const proFormulas = Array.from({ length: numOfStudents }, () => [proFormula]);

    for (let i = 0; i < weeks; i++) {
        let week = schedule[i];
        let weekNum = i + 1;
        let weekRow = defaultWeekRow + weekRowDifference * weekNum;
        let scheduleRow = weekRow + 3;

        recordsSheet
            .getRange(weekRow, defaultWeekCol, 2, 1)
            .setValue(`Week ${weekNum}`);

        // Percentage formula: each student's cumulative GLH for this week divided by the
        // total possible GLH up to week 1 (used to show attendance rate progression).
        let percentFormula = `=IF(AND(ISNUMBER(INDIRECT(ADDRESS(ROW(), COLUMN()-1))), ISNUMBER(INDIRECT(ADDRESS(21 + (${i} * (SUMMARY!K8 + 7)), COLUMN()-1)))), INDIRECT(ADDRESS(ROW(), COLUMN()-1)) / INDIRECT(ADDRESS(21 + (${i} * (SUMMARY!K8 + 7)), COLUMN()-1)), "")`;
        let percentFormulas = Array.from({ length: numOfStudents }, () => [percentFormula]);
        recordsSheet.getRange(scheduleRow + 2, percentCol, numOfStudents, 1).setFormulas(percentFormulas);

        for (let j = 1; j < week.length + 1; j++) {
            let dateObj = week[j - 1];
            let date = dateObj.date;
            let day = dateObj.day;
            let dateDayCol = defaultDateDayCol + dateDayColDifference * (j - 1);

            // Format as DD/MM/YYYY for display in the date row.
            let dateTime = new Date(date);
            const formattedDate = `${String(dateTime.getDate()).padStart(2, "0")}/${String(dateTime.getMonth() + 1).padStart(2, "0")}/${String(dateTime.getFullYear())}`;
            recordsSheet
                .getRange(weekRow, dateDayCol, 1, defaultDateDayCols)
                .setValue(formattedDate);

            const weekDays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
            const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const formattedDay = `Day ${day} - ${weekDays[dateTime.getDay()]} - ${String(dateTime.getDate())} ${months[dateTime.getMonth()]}`;
            recordsSheet
                .getRange(weekRow + 1, dateDayCol, 1, defaultDateDayCols)
                .setValue(formattedDay);

            let proCol = proCols[j];
            let suCol = suCols[j];

            // Project and hackathon days get PRO formulas and an amber background to distinguish them visually.
            if (isProjectOrHackathonDay(day)) {
                recordsSheet
                    .getRange(scheduleRow + 2, proCol, numOfStudents, 1)
                    .setFormulas(proFormulas);
                recordsSheet
                    .getRange(scheduleRow - 3, suCol, 2, 10)
                    .setBackgroundRGB(249, 203, 156);
                let dayString = recordsSheet.getRange(scheduleRow - 2, suCol, 1, 10).getValue();
                recordsSheet
                    .getRange(scheduleRow - 2, suCol, 1, 10)
                    .setValue(dayString + " - Project");
            }
        }
    }

    // Remove the original template block now that all weekly blocks have been generated.
    recordsSheet.deleteRows(defaultWeekRow, numberOfSourceRows);
    setCurrentWeek(false);
}

/**
 * Updates the "current week" and "today's date" display cells in both RECORDS and SUMMARY.
 * Pass a YYYY-MM-DD date string to override today's date (useful for testing), or false to use now.
 * Returns the currentWeekData object from getCurrentWeekData().
 */
function setCurrentWeek(todayDate) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");

    const currentWeekData = getCurrentWeekData(todayDate);

    // Preserve the existing number format when writing the date so cell formatting is not lost.
    const todayDateRange = recordsSheet.getRange(14, 13, 2, 5);
    const todayDateFormat = todayDateRange.getNumberFormat();
    todayDateRange.setValue(currentWeekData.todayDate);
    todayDateRange.setNumberFormat(todayDateFormat);

    recordsSheet.getRange(14, 19, 2, 5).setValue(currentWeekData.currentWeek);
    summarySheet.getRange(2, 11, 1, 1).setValue(currentWeekData.todayDate);
    summarySheet.getRange(2, 12, 1, 1).setValue(currentWeekData.currentDay);
    summarySheet.getRange(2, 13, 1, 1).setValue(currentWeekData.currentWeek);
    summarySheet.getRange(2, 14, 1, 1).setValue(currentWeekData.currentWeekNum);

    return currentWeekData;
}

/**
 * Calculates the current week number, day number, and today's date relative to the cohort schedule.
 * Returns "Not Started" or "Finished" if today falls outside the cohort date range.
 * Pass a YYYY-MM-DD string to override today, or false to use the real current date.
 */
function getCurrentWeekData(todayDate) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const scheduleData = JSON.parse(databaseSheet.getRange(3, 24, 1, 1).getValue());
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;

    const startDateTime = getStartOrEndDate("start", "dateTime");
    const endDateTime = getStartOrEndDate("end", "dateTime");

    let todayDateTime = todayDate ? new Date(todayDate) : new Date();

    // Set to 22:00 so any session earlier in the day is treated as "today" without overlap issues.
    todayDateTime.setHours(22, 0, 0, 0);

    let currentWeekNum = 0;
    let currentDay = 0;
    let currentWeek;

    if (todayDateTime > endDateTime) {
        currentWeek = "Finished";
        currentWeekNum = 16;
        currentDay = 80;
    } else if (todayDateTime < startDateTime) {
        currentWeek = "Not Started";
        currentWeekNum = 1;
        currentDay = 1;
    } else {
        for (let i = 0; i < weeks; i++) {
            currentWeekNum = i + 1;

            // For all weeks except the last, skip ahead if today is on or after the next week's start.
            if (i < weeks - 1) {
                let firstDateTimeOfNextWeek = new Date(schedule[currentWeekNum][0].date);
                firstDateTimeOfNextWeek.setHours(6, 0, 0, 0);
                if (todayDateTime >= firstDateTimeOfNextWeek) continue;
            }

            // Today is within week i (either we're before the next week's start, or this is the last week).
            currentWeek = `Week ${currentWeekNum}`;
            let w = schedule[i];
            for (let j = 0; j < w.length; j++) {
                if (w[j].date == todayDateTime.toISOString().split("T")[0]) {
                    currentDay = w[j].day;
                }
            }
            break;
        }
    }

    return {
        startDateTime: new Date(todayDateTime),
        endDateTime: new Date(todayDateTime),
        currentWeek: currentWeek,
        currentWeekNum: currentWeekNum,
        todayDateTime: todayDateTime,
        todayDate: todayDateTime.toISOString().split("T")[0],
        currentDay: currentDay,
    };
}

/**
 * Reads the cohort start or end date from DATABASE.
 * Pass "start" or "end" for startOrEnd, and "date" (YYYY-MM-DD string) or "dateTime" (Date object)
 * for the return type.
 */
function getStartOrEndDate(startOrEnd, dateOrDateTime) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const dateRow = startOrEnd == "start" ? 3 : 4;
    const dateTime = new Date(databaseSheet.getRange(dateRow, 22, 1, 1).getValue());
    dateTime.setHours(6, 0, 0, 0);

    return dateOrDateTime == "date" ? dateTime.toISOString().split("T")[0] : dateTime;
}

/**
 * Returns true if the given cohort day number falls within a project or hackathon window.
 * Ranges are derived from the shared constants in GenerateSchedule.js.
 */
function isProjectOrHackathonDay(day) {
    return (
        (day >= PROJ1_START_DAY && day <= PROJ1_END_DAY) ||
        (day >= HACK1_START_DAY && day <= HACK1_END_DAY) ||
        (day >= PROJ2_START_DAY && day <= PROJ2_END_DAY) ||
        (day >= HACK2_START_DAY && day <= HACK2_END_DAY)
    );
}
