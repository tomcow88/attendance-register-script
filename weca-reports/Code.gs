/**
 * Adds a "Custom Menu" to the spreadsheet UI with a single "Generate Reports" item.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Custom Menu")
        .addItem("Generate Reports", "showPrompt")
        .addToUi();
}

/**
 * Prompts the user for a Google Doc template URL and passes it to generateReports.
 * The Doc must contain placeholder strings matching those replaced in generateReports.
 */
function showPrompt() {
    const ui = SpreadsheetApp.getUi();
    const responseDocUrl = ui.prompt("Enter the Report Template Google Doc URL");
    const docUrl = responseDocUrl.getResponseText();

    if (docUrl) {
        generateReports(docUrl);
    } else {
        ui.alert("Please enter the Report Template Google Doc URL.");
    }
}

/**
 * Generates one Training Completion Declaration Doc per learner row in the active sheet.
 *
 * Expected spreadsheet columns (0-indexed, row 1 onwards — row 0 is the header):
 *   0  firstName
 *   1  lastName
 *   2  totalGlh
 *   3  pctGlh
 *   4  employabilityGlh
 *   5  careersGlh
 *   6  guidedBehavioursGlh
 *   7  pastoralHours
 *
 * Each row produces a copy of the template Doc with the above values substituted
 * into the corresponding {{ Placeholder }} strings.
 */
function generateReports(docUrl) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const docId = extractDocId(docUrl);
    if (!docId) {
        SpreadsheetApp.getUi().alert("Could not extract a document ID from the URL provided. Please check the URL and try again.");
        return;
    }

    const templateFile = DriveApp.getFileById(docId);
    const failures = [];

    for (let i = 1; i < data.length; i++) {
        const [firstName, lastName, totalGlh, pctGlh, employabilityGlh, careersGlh, guidedBehavioursGlh, pastoralHours] = data[i];
        if (!firstName || typeof totalGlh !== "number") continue;
        const fullName = `${firstName} ${lastName}`;
        const formattedPct = typeof pctGlh === "number" ? `${(pctGlh * 100).toFixed(2)}%` : pctGlh;

        try {
            const newDoc = templateFile.makeCopy(`${fullName} | Training Completion Declaration`);
            const newDocFile = DocumentApp.openById(newDoc.getId());
            const newDocBody = newDocFile.getBody();

            newDocBody.replaceText("{{ Learner First Name }}", firstName);
            newDocBody.replaceText("{{ Learner Last Name }}", lastName);
            newDocBody.replaceText("{{ Total GLH }}", totalGlh);
            newDocBody.replaceText("{{ % GLH }}", formattedPct);
            newDocBody.replaceText("{{ Total Employability GLH }}", employabilityGlh);
            newDocBody.replaceText("{{ Total Careers GLH }}", careersGlh);
            newDocBody.replaceText("{{ Guided Behaviours & Workplace Skills GLH }}", guidedBehavioursGlh);
            newDocBody.replaceText("{{ Pastoral Support Hours }}", pastoralHours || 0);

            newDocFile.saveAndClose();
        } catch (e) {
            failures.push(fullName);
            Logger.log(`Failed to generate report for ${fullName}: ${e.message}`);
        }
    }

    if (failures.length > 0) {
        SpreadsheetApp.getUi().alert("Reports generated with errors. The following learners failed:\n\n" + failures.join("\n"));
    } else {
        SpreadsheetApp.getUi().alert("Reports generated successfully.");
    }
}

/**
 * Extracts the file ID from a Google Drive or Docs URL.
 * Matches the first segment of 25+ word characters or hyphens, which is the
 * format Google uses for all Drive file IDs.
 */
function extractDocId(url) {
    const docId = url.match(/[-\w]{25,}/);
    return docId ? docId[0] : null;
}
