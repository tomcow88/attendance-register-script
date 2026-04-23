/**
 * Orchestrates all four database setup sub-functions in the required order:
 * setupStudentData → setupSessionsData → setupDeliveryTeamData → setupCohortData.
 */
function setupDatabase() {
    setupStudentData();
    setupSessionsData();
    setupDeliveryTeamData();
    setupCohortData();
}

/**
 * Reads student names and emails from the SETUP sheet and writes them to DATABASE.
 * Also generates an anonymized meet email (used for matching against Google Meet attendee lists)
 * and stores the student count.
 */
function setupStudentData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    // Count data rows in column A (excluding the two header rows) to get the number of students.
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

        if (!firstName) return;
        if (!lastName) return;
        if (!email) return;

        let fullName = firstName + " " + lastName;
        let meetName = firstName;
        let meetEmail = getMeetEmail(email);

        let studentRow = 3 + i;

        databaseSheet.getRange(studentRow, 1, 1, 1).setValue(fullName);
        databaseSheet.getRange(studentRow, 2, 1, 1).setValue(firstName);
        databaseSheet.getRange(studentRow, 3, 1, 1).setValue(lastName);
        // Meet names and emails are stored as JSON arrays to support multiple values per learner.
        databaseSheet.getRange(studentRow, 4, 1, 1).setValue(`["${meetName}"]`);
        databaseSheet.getRange(studentRow, 5, 1, 1).setValue(`["${meetEmail}"]`);
    }
}

/**
 * Anonymizes an email address for safe storage and matching against Google Meet attendee names.
 * Masks characters after the first 4 in the local part and replaces the domain name with "***".
 * Example: "john.doe@example.com" → "john***@***.com"
 */
function getMeetEmail(email) {
    let [localPart, domainPart] = email.split("@");

    let localPartSplit = localPart.split("");
    for (let i = 4; i < localPartSplit.length; i++) {
        localPartSplit[i] = "*";
    }
    let newLocalPart = localPartSplit.join("");

    let [domainName, extension] = domainPart.split(".");
    let newDomainPart = "***" + "." + extension;

    return newLocalPart + "@" + newDomainPart;
}

/**
 * Reads calendar names from SETUP and session column mappings from RECORDS row 20,
 * then writes them to DATABASE as JSON arrays.
 * Session types: SU (Stand Up), SD (Stand Down), GS (Guest Speaker), SME, CC (Career Coach), PRO (Project).
 */
function setupSessionsData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");
    const recordsSheet = spreadsheet.getSheetByName("Records");
    const recordsLastCol = recordsSheet.getLastColumn();

    // Read calendar names for each session type from SETUP.
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

    // Row 20 of RECORDS contains session type labels (e.g. "SU", "SD") in each column.
    // Scan every column to build an ordered list of column numbers per session type.
    const recordsSessionsRow = 20;
    const recordsSessionsValues = recordsSheet
        .getRange(recordsSessionsRow, 1, 1, recordsLastCol)
        .getValues()
        .flat();
    let suCols = [], sdCols = [], gsCols = [], smeCols = [], ccCols = [], proCols = [];

    for (let i = 0; i < recordsLastCol; i++) {
        let value = recordsSessionsValues[i];
        if (value == "SU") suCols.push(String(i + 1));
        else if (value == "SD") sdCols.push(String(i + 1));
        else if (value == "GS") gsCols.push(String(i + 1));
        else if (value == "SME") smeCols.push(String(i + 1));
        else if (value == "CC") ccCols.push(String(i + 1));
        else if (value == "PRO") proCols.push(String(i + 1));
    }

    databaseSheet.getRange(3, 10, 1, 1).setValue(JSON.stringify(suCols));
    databaseSheet.getRange(4, 10, 1, 1).setValue(JSON.stringify(sdCols));
    databaseSheet.getRange(5, 10, 1, 1).setValue(JSON.stringify(gsCols));
    databaseSheet.getRange(6, 10, 1, 1).setValue(JSON.stringify(smeCols));
    databaseSheet.getRange(7, 10, 1, 1).setValue(JSON.stringify(ccCols));
    databaseSheet.getRange(8, 10, 1, 1).setValue(JSON.stringify(proCols));
}

/**
 * Reads delivery team member details (name, email, Drive folder URL) from SETUP
 * and writes them to DATABASE. The folder URL is converted to a folder ID for later use
 * when locating session attendance files.
 */
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

    // Names and emails are stored as JSON arrays to support future multi-value expansion.
    databaseSheet.getRange(3, 15, 1, 1).setValue(`["${setupFacName}"]`);
    databaseSheet.getRange(3, 16, 1, 1).setValue(`["${setupFacEmail}"]`);
    databaseSheet.getRange(3, 17, 1, 1).setValue(`["${facFolderId}"]`);

    databaseSheet.getRange(4, 15, 1, 1).setValue(`["${setupSmeName}"]`);
    databaseSheet.getRange(4, 16, 1, 1).setValue(`["${setupSmeEmail}"]`);
    databaseSheet.getRange(4, 17, 1, 1).setValue(`["${smeFolderId}"]`);

    databaseSheet.getRange(5, 15, 1, 1).setValue(`["${setupCcName}"]`);
    databaseSheet.getRange(5, 16, 1, 1).setValue(`["${setupCcEmail}"]`);
    databaseSheet.getRange(5, 17, 1, 1).setValue(`["${ccFolderId}"]`);

    // Plain name values are also stored separately for use as signatures in RECORDS.
    databaseSheet.getRange(8, 15, 1, 1).setValue(setupFacName);
    databaseSheet.getRange(9, 15, 1, 1).setValue(setupSmeName);
    databaseSheet.getRange(10, 15, 1, 1).setValue(setupCcName);
}

/**
 * Extracts a Google Drive folder ID from a full Drive folder URL.
 * Returns null if the URL does not contain a recognisable folder ID pattern.
 */
function getFolderIdFromUrl(url) {
    let folderIdMatch = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    return folderIdMatch ? folderIdMatch[1] : null;
}

/**
 * Reads cohort metadata (funding partner, location, start date) from SETUP, generates the
 * full 80-day schedule using getPublicHolidays and generateSchedule, and writes all of it
 * to DATABASE — including the serialised schedule JSON and computed project/hackathon dates.
 */
function setupCohortData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = spreadsheet.getSheetByName("SETUP");
    const databaseSheet = spreadsheet.getSheetByName("DATABASE");

    const partner = setupSheet.getRange(3, 13, 1, 1).getValue();
    const location = setupSheet.getRange(3, 14, 1, 1).getValue();
    const startDate = setupSheet.getRange(3, 15, 1, 1).getValue();

    if (!(partner && location && startDate)) return;

    databaseSheet.getRange(3, 19, 1, 1).setValue(partner);
    databaseSheet.getRange(3, 20, 1, 1).setValue(location);
    databaseSheet.getRange(3, 21, 1, 1).setValue(startDate);

    const holidays = getPublicHolidays(location);
    const scheduleData = generateSchedule(startDate, holidays);

    // Serialise the full schedule as JSON so it can be read back by other functions at runtime.
    databaseSheet.getRange(3, 23, 1, 1).setValue(JSON.stringify(scheduleData));

    const schedule = scheduleData.schedule;
    const weeks = scheduleData.weeks;
    const lastDay = schedule[weeks - 1][schedule[weeks - 1].length - 1].date;

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

/**
 * Updates student data after initial setup.
 * Called when new email addresses need to be added to existing learner records
 * (e.g. a learner joined a session using a different email address).
 */
function updateDatabase() {
    updateStudentData();
}

/**
 * Reads new email addresses from the SETUP sheet and appends them to the existing
 * meet email arrays stored in DATABASE. Clears the email column in SETUP after processing
 * so it is ready for the next update batch.
 *
 * This is the primary mechanism for handling learners who attend sessions under
 * multiple different email addresses.
 */
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
        if (!email) continue;

        let meetEmail = getMeetEmail(email);
        let studentRow = 3 + i;
        let meetEmailRange = databaseSheet.getRange(studentRow, 5, 1, 1);

        // Append the new email to the existing JSON array rather than overwriting it.
        let meetEmailArr = JSON.parse(meetEmailRange.getValue());
        meetEmailArr.push(meetEmail);
        meetEmailRange.setValue(JSON.stringify(meetEmailArr));
    }

    // Clear the email column so it doesn't get re-processed on the next update.
    setupSheet.getRange(3, 3, numOfStudents, 1).clearContent();
}
