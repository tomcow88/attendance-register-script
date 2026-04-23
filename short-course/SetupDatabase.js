function setupDatabase() {
    setupStudentData();
    setupSessionsData();
    setupDeliveryTeamData();
    setupCohortData();
}

function setupStudentData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const numOfStudents =
        setupSheet.getRange("A:A").getValues().filter(String).length - 2;
    databaseSheet.getRange(3, 22, 1, 1).setValue(numOfStudents);

    let studentFirstNames = setupSheet
        .getRange(3, 1, numOfStudents, 1)
        .getValues()
        .flat();
    let studentLastNames = setupSheet
        .getRange(3, 2, numOfStudents, 1)
        .getValues()
        .flat();
    let studentEmails = setupSheet
        .getRange(3, 3, numOfStudents, 1)
        .getValues()
        .flat();

    for (let i = 0; i < numOfStudents; i++) {
        let firstName = studentFirstNames[i];
        let lastName = studentLastNames[i];
        let email = studentEmails[i];

        if (!firstName) {
            return;
        }

        if (!lastName) {
            return;
        }

        if (!email) {
            return;
        }

        let fullName = firstName + " " + lastName;
        let meetName = firstName;
        let meetEmail = getMeetEmail(email);

        let studentRow = 3 + i;

        let fullNameRange = databaseSheet.getRange(studentRow, 1, 1, 1);
        let firstNameRange = databaseSheet.getRange(studentRow, 2, 1, 1);
        let lastNameRange = databaseSheet.getRange(studentRow, 3, 1, 1);
        let meetNameRange = databaseSheet.getRange(studentRow, 4, 1, 1);
        let meetEmailRange = databaseSheet.getRange(studentRow, 5, 1, 1);

        fullNameRange.setValue(fullName);
        firstNameRange.setValue(firstName);
        lastNameRange.setValue(lastName);
        meetNameRange.setValue(`["${meetName}"]`);
        meetEmailRange.setValue(`["${meetEmail}"]`);
    }
}

function getMeetEmail(email) {
    let [localPart, domainPart] = email.split("@");

    let localPartSplit = localPart.split("");
    for (let i = 4; i < localPartSplit.length; i++) {
        localPartSplit[i] = "*";
    }
    let newLocalPart = localPartSplit.join("");

    let [domainName, extension] = domainPart.split(".");
    let newDomainName = "***";

    let newDomainPart = newDomainName + "." + extension;

    let meetEmail = newLocalPart + "@" + newDomainPart;

    return meetEmail;
}

function setupSessionsData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const recordsSheet = spreadsheet.getSheetByName("Records");
    const recordsLastCol = recordsSheet.getLastColumn();

    const setupCalSU = setupSheet.getRange(3, 6, 1, 1).getValue();
    const setupCalSD = setupSheet.getRange(4, 6, 1, 1).getValue();
    const setupCalGS = setupSheet.getRange(5, 6, 1, 1).getValue();
    const setupCalSME = setupSheet.getRange(6, 6, 1, 1).getValue();
    const setupCalCC = setupSheet.getRange(7, 6, 1, 1).getValue();

    databaseSheet.getRange(3, 11, 1, 1).setValue(`["${setupCalSU}"]`);
    databaseSheet.getRange(4, 11, 1, 1).setValue(`["${setupCalSD}"]`);
    databaseSheet.getRange(5, 11, 1, 1).setValue(`["${setupCalGS}"]`);
    databaseSheet.getRange(6, 11, 1, 1).setValue(`["${setupCalSME}"]`);
    databaseSheet.getRange(7, 11, 1, 1).setValue(`["${setupCalCC}"]`);

    const recordsSessionsRow = 20;
    const recordsSessionsValues = recordsSheet
        .getRange(recordsSessionsRow, 1, 1, recordsLastCol)
        .getValues()
        .flat();
    let suCols = [];
    let sdCols = [];
    let gsCols = [];
    let smeCols = [];
    let ccCols = [];
    let proCols = [];
    for (let i = 0; i < recordsLastCol; i++) {
        let value = recordsSessionsValues[i];

        if (value == "SU") {
            suCols.push(String(i + 1));
        } else if (value == "SD") {
            sdCols.push(String(i + 1));
        } else if (value == "GS") {
            gsCols.push(String(i + 1));
        } else if (value == "SME") {
            smeCols.push(String(i + 1));
        } else if (value == "CC") {
            ccCols.push(String(i + 1));
        } else if (value == "PRO") {
            proCols.push(String(i + 1));
        }
    }

    databaseSheet.getRange(3, 10, 1, 1).setValue(JSON.stringify(suCols));
    databaseSheet.getRange(4, 10, 1, 1).setValue(JSON.stringify(sdCols));
    databaseSheet.getRange(5, 10, 1, 1).setValue(JSON.stringify(gsCols));
    databaseSheet.getRange(6, 10, 1, 1).setValue(JSON.stringify(smeCols));
    databaseSheet.getRange(7, 10, 1, 1).setValue(JSON.stringify(ccCols));
    databaseSheet.getRange(8, 10, 1, 1).setValue(JSON.stringify(proCols));
}

function setupDeliveryTeamData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const setupFacName = setupSheet.getRange(3, 9, 1, 1).getValue();
    const setupFacEmail = setupSheet.getRange(3, 10, 1, 1).getValue();
    const setupFacUrl = setupSheet.getRange(3, 11, 1, 1).getValue();
    const facFolderId = getFolderIdFromUrl(setupFacUrl);

    const setupSmeName = setupSheet.getRange(4, 9, 1, 1).getValue();
    const setupSmeEmail = setupSheet.getRange(4, 10, 1, 1).getValue();
    const setupSmeUrl = setupSheet.getRange(4, 11, 1, 1).getValue();
    const smeFolderId = getFolderIdFromUrl(setupSmeUrl);

    const setupCcName = setupSheet.getRange(5, 9, 1, 1).getValue();
    const setupCcEmail = setupSheet.getRange(5, 10, 1, 1).getValue();
    const setupCcUrl = setupSheet.getRange(5, 11, 1, 1).getValue();
    const ccFolderId = getFolderIdFromUrl(setupCcUrl);

    databaseSheet.getRange(3, 15, 1, 1).setValue(`["${setupFacName}"]`);
    databaseSheet.getRange(3, 16, 1, 1).setValue(`["${setupFacEmail}"]`);
    databaseSheet.getRange(3, 17, 1, 1).setValue(`["${facFolderId}"]`);

    databaseSheet.getRange(4, 15, 1, 1).setValue(`["${setupSmeName}"]`);
    databaseSheet.getRange(4, 16, 1, 1).setValue(`["${setupSmeEmail}"]`);
    databaseSheet.getRange(4, 17, 1, 1).setValue(`["${smeFolderId}"]`);

    databaseSheet.getRange(5, 15, 1, 1).setValue(`["${setupCcName}"]`);
    databaseSheet.getRange(5, 16, 1, 1).setValue(`["${setupCcEmail}"]`);
    databaseSheet.getRange(5, 17, 1, 1).setValue(`["${ccFolderId}"]`);

    databaseSheet.getRange(8, 15, 1, 1).setValue(setupFacName);
    databaseSheet.getRange(9, 15, 1, 1).setValue(setupSmeName);
    databaseSheet.getRange(10, 15, 1, 1).setValue(setupCcName);
}

function getFolderIdFromUrl(url) {
    let folderIdMatch = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    let folderId = folderIdMatch ? folderIdMatch[1] : null;

    return folderId;
}

function setupCohortData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const partner = setupSheet.getRange(3, 13, 1, 1).getValue();
    const location = setupSheet.getRange(3, 14, 1, 1).getValue();
    const startDate = setupSheet.getRange(3, 15, 1, 1).getValue();

    if (!(partner && location && startDate)) {
        return;
    }

    databaseSheet.getRange(3, 19, 1, 1).setValue(partner);
    databaseSheet.getRange(3, 20, 1, 1).setValue(location);
    databaseSheet.getRange(3, 21, 1, 1).setValue(startDate);

    const holidays = getPublicHolidays(location);
    const scheduleData = generateSchedule(startDate, holidays);

    const scheduleDataString = JSON.stringify(scheduleData);
    databaseSheet.getRange(3, 23, 1, 1).setValue(scheduleDataString);

    const schedule = scheduleData.schedule;
    const weeks = scheduleData.weeks;
    const lastWeek = schedule[weeks - 1];
    const lastDay = lastWeek[lastWeek.length - 1].date;

    databaseSheet.getRange(4, 21, 1, 1).setValue(lastDay);
    databaseSheet.getRange(7, 21, 1, 1).setValue(scheduleData.proj1StartDate);
    databaseSheet.getRange(8, 21, 1, 1).setValue(scheduleData.proj1EndDate);
    databaseSheet.getRange(11, 21, 1, 1).setValue(scheduleData.hack1StartDate);
    databaseSheet.getRange(12, 21, 1, 1).setValue(scheduleData.hack1EndDate);
    databaseSheet.getRange(15, 21, 1, 1).setValue(scheduleData.proj2StartDate);
    databaseSheet.getRange(16, 21, 1, 1).setValue(scheduleData.proj2EndDate);
    databaseSheet.getRange(19, 21, 1, 1).setValue(scheduleData.hack2StartDate);
    databaseSheet.getRange(20, 21, 1, 1).setValue(scheduleData.hack2EndDate);
}

function updateDatabase() {
    updateStudentData();
}

function updateStudentData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const numOfStudents =
        setupSheet.getRange("A:A").getValues().filter(String).length - 2;

    let studentEmails = setupSheet
        .getRange(3, 3, numOfStudents, 1)
        .getValues()
        .flat();

    for (let i = 0; i < numOfStudents; i++) {
        let email = studentEmails[i];

        if (!email) {
            continue;
        }

        let meetEmail = getMeetEmail(email);

        let studentRow = 3 + i;

        let meetEmailRange = databaseSheet.getRange(studentRow, 5, 1, 1);

        let meetEmailArr = JSON.parse(meetEmailRange.getValue());

        meetEmailArr.push(meetEmail);

        meetEmailRange.setValue(JSON.stringify(meetEmailArr));
    }

    setupSheet.getRange(3, 3, numOfStudents, 1).clearContent();
}
