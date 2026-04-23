function setupRecords() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const recordsLastCol = recordsSheet.getLastColumn();

    const numOfStudents = databaseSheet.getRange(3, 22, 1, 1).getValue();
    const scheduleDataString = databaseSheet.getRange(3, 23, 1, 1).getValue();
    const scheduleData = JSON.parse(scheduleDataString);
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;
    const fullNames = databaseSheet
        .getRange(3, 1, numOfStudents, 1)
        .getValues()
        .flat();

    // Setup Template Weeks

    const defaultNameRow = 23;
    const defaultNameRowRange = recordsSheet.getRange(
        defaultNameRow,
        1,
        1,
        recordsLastCol,
    );

    const targetNameRow = defaultNameRow;
    const numOfTargetNameRows = numOfStudents;
    const targetNameRowRange = recordsSheet.getRange(
        targetNameRow,
        1,
        numOfTargetNameRows,
        recordsLastCol,
    );
    defaultNameRowRange.copyTo(targetNameRowRange);

    // Add Full Names
    for (let i = 0; i < numOfStudents; i++) {
        let nameRow = defaultNameRow + i;
        let fullName = fullNames[i];

        recordsSheet.getRange(nameRow, 1, 1, 1).setValue(fullName);
    }

    const sourceStartRow = 17;
    const numberOfSourceRows = 7 + numOfStudents;
    const sourceRange = recordsSheet.getRange(
        sourceStartRow,
        1,
        numberOfSourceRows,
        recordsLastCol,
    );

    let targetStartRow = sourceStartRow + numberOfSourceRows;
    for (let i = 0; i < weeks; i++) {
        let targetRange = recordsSheet.getRange(
            targetStartRow,
            1,
            numberOfSourceRows,
            recordsLastCol,
        );
        sourceRange.copyTo(targetRange);
        targetStartRow += numberOfSourceRows;
    }

    // Add Correct Dates and Weeks

    const defaultWeekRow = 18;
    const defaultWeekCol = 1;
    const defaultWeekRows = 2;
    const defaultWeekCols = 1;

    const defaultDateDayCol = 14;
    const defaultDateDayRows = 1;
    const defaultDateDayCols = 10;

    const weekRowDifference = numberOfSourceRows;
    const dateRowDifference = 0;
    const dayRowDifference = 1;
    const dateDayColDifference = defaultDateDayCols + 1;

    const defaultDateRow = defaultWeekRow + dateRowDifference;
    const defaultDateFormat = recordsSheet
        .getRange(
            defaultDateRow,
            defaultDateDayCol,
            defaultDateDayRows,
            defaultDateDayCols,
        )
        .getNumberFormat();
    const defaultDayRow = defaultWeekRow + dayRowDifference;
    const defaultDayFormat = recordsSheet
        .getRange(
            defaultDayRow,
            defaultDateDayCol,
            defaultDateDayRows,
            defaultDateDayCols,
        )
        .getNumberFormat();

    const startScheduleRow = 21;
    const proGlh = databaseSheet.getRange(8, 9, 1, 1).getValue();
    const proFormula = `=IF(OR(INDIRECT(ADDRESS(ROW(), COLUMN()-4))="X", INDIRECT(ADDRESS(ROW(), COLUMN()-5))="X"), "X",
    IF(AND(INDIRECT(ADDRESS(ROW(), COLUMN()-4))="-", INDIRECT(ADDRESS(ROW(), COLUMN()-5))="-"), "-", 
    IF(OR(
        ISNUMBER(INDIRECT(ADDRESS(ROW(), COLUMN()-4))) * (INDIRECT(ADDRESS(ROW(), COLUMN()-4)) > 0),
        ISNUMBER(INDIRECT(ADDRESS(ROW(), COLUMN()-5))) * (INDIRECT(ADDRESS(ROW(), COLUMN()-5)) > 0)
    ), ${proGlh}, 0)))`;
    const proFormulas = Array.from({ length: numOfStudents }, () => [
        proFormula,
    ]);
    const percentCol = 12;
    const proCols = JSON.parse(databaseSheet.getRange(8, 10, 1, 1).getValue());
    const suCols = JSON.parse(databaseSheet.getRange(3, 10, 1, 1).getValue());

    for (let i = 0; i < weeks; i++) {
        let week = schedule[i];
        let weekNum = i + 1;
        let weekRow = defaultWeekRow + weekRowDifference * weekNum;
        recordsSheet
            .getRange(weekRow, defaultWeekCol, defaultWeekRows, defaultWeekCols)
            .setValue(`Week ${weekNum}`);

        let scheduleRow = weekRow + 3;
        let percentFormula = `=IF(AND(ISNUMBER(INDIRECT(ADDRESS(ROW(), COLUMN()-1))), ISNUMBER(INDIRECT(ADDRESS(21 + (${i} * (SUMMARY!K8 + 7)), COLUMN()-1)))), INDIRECT(ADDRESS(ROW(), COLUMN()-1)) / INDIRECT(ADDRESS(21 + (${i} * (SUMMARY!K8 + 7)), COLUMN()-1)), "")`;
        let percentFormulas = Array.from({ length: numOfStudents }, () => [
            percentFormula,
        ]);
        recordsSheet
            .getRange(scheduleRow + 2, percentCol, numOfStudents, 1)
            .setFormulas(percentFormulas);

        let dateRow = weekRow + dateRowDifference;
        let dayRow = weekRow + dayRowDifference;
        for (let j = 1; j < week.length + 1; j++) {
            let dateObj = week[j - 1];
            let date = dateObj.date;
            let day = dateObj.day;
            let dateDayCol = defaultDateDayCol + dateDayColDifference * (j - 1);

            let dateTime = new Date(date);
            const formattedDate = `${String(dateTime.getDate()).padStart(2, "0")}/${String(dateTime.getMonth() + 1).padStart(2, "0")}/${String(dateTime.getFullYear())}`;
            let dateRange = recordsSheet.getRange(
                dateRow,
                dateDayCol,
                defaultDateDayRows,
                defaultDateDayCols,
            );
            dateRange.setValue(formattedDate);

            const weekDays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
            const months = [
                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "May",
                "Jun",
                "Jul",
                "Aug",
                "Sep",
                "Oct",
                "Nov",
                "Dec",
            ];
            const formattedDay = `Day ${day} - ${weekDays[dateTime.getDay()]} - ${String(dateTime.getDate())} ${months[dateTime.getMonth()]}`;
            let dayRange = recordsSheet.getRange(
                dayRow,
                dateDayCol,
                defaultDateDayRows,
                defaultDateDayCols,
            );
            dayRange.setValue(formattedDay);

            let proCol = proCols[j];
            let suCol = suCols[j];

            if (isProjectOrHackathonDay(day)) {
                recordsSheet
                    .getRange(scheduleRow + 2, proCol, numOfStudents, 1)
                    .setFormulas(proFormulas);
                recordsSheet
                    .getRange(scheduleRow - 3, suCol, 2, 10)
                    .setBackgroundRGB(249, 203, 156);
                let dayString = recordsSheet
                    .getRange(scheduleRow - 2, suCol, 1, 10)
                    .getValue();
                let newDayString = dayString + " - Project";
                recordsSheet
                    .getRange(scheduleRow - 2, suCol, 1, 10)
                    .setValue(newDayString);
            }
        }
    }

    recordsSheet.deleteRows(defaultWeekRow, numberOfSourceRows);
    setCurrentWeek(false);
}

function setCurrentWeek(todayDate) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const recordsSheet = spreadsheet.getSheetByName("RECORDS");
    const summarySheet = spreadsheet.getSheetByName("SUMMARY");

    const currentWeekData = getCurrentWeekData(todayDate);
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

function getCurrentWeekData(todayDate) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const scheduleDataString = databaseSheet.getRange(3, 23, 1, 1).getValue();
    const scheduleData = JSON.parse(scheduleDataString);
    const weeks = scheduleData.weeks;
    const schedule = scheduleData.schedule;

    const startDateTime = getStartOrEndDate("start", "dateTime");
    const endDateTime = getStartOrEndDate("end", "dateTime");

    let todayDateTime;

    if (todayDate) {
        todayDateTime = new Date(todayDate);
    } else {
        todayDateTime = new Date();
    }

    todayDateTime.setHours(22, 0, 0, 0);

    let startOfWeekDateTime = new Date(todayDateTime);
    let endOfWeekDateTime = new Date(todayDateTime);

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
            let nextWeek = schedule[currentWeekNum];
            if (i == weeks - 1) {
                nextWeek = schedule[i];
            }
            let firstDateOfWeek = nextWeek[0].date;
            let firstDateTimeOfWeek = new Date(firstDateOfWeek);
            firstDateTimeOfWeek.setHours(6, 0, 0, 0);
            if (todayDateTime < firstDateTimeOfWeek) {
                currentWeek = `Week ${currentWeekNum}`;
                let w = schedule[i];
                for (let j = 0; j < w.length; j++) {
                    let d = w[j];
                    if (d.date == todayDateTime.toISOString().split("T")[0]) {
                        currentDay = d.day;
                    }
                }
                break;
            }
        }
    }

    let todayDateString = todayDateTime.toISOString().split("T")[0];

    const currentWeekData = {
        startDateTime: startOfWeekDateTime,
        endDateTime: endOfWeekDateTime,
        currentWeek: currentWeek,
        currentWeekNum: currentWeekNum,
        todayDateTime: todayDateTime,
        todayDate: todayDateString,
        currentDay: currentDay,
    };

    return currentWeekData;
}

function getStartOrEndDate(startOrEnd, dateOrDateTime) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    let dateRow;

    if (startOrEnd == "start") {
        dateRow = 3;
    } else {
        dateRow = 4;
    }

    const dateString = databaseSheet.getRange(dateRow, 21, 1, 1).getValue();
    const dateTime = new Date(dateString);
    dateTime.setHours(6, 0, 0, 0);

    if (dateOrDateTime == "date") {
        const date = dateTime.toISOString().split("T")[0];
        return date;
    } else {
        return dateTime;
    }
}

function isProjectOrHackathonDay(day) {
    const projectHackathonDays = [
        25, 26, 27, 28, 38, 39, 40, 41, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70,
        71, 72, 73, 74, 75, 76, 77, 78, 79,
    ];

    return projectHackathonDays.includes(day);
}
