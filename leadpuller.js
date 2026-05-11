const SOURCE_SHEET_ID = "1dJrQYJ6MtF5ixrmjXik7dHJsXIiFC64Eu4vMoy3-kls"; // BD DME 2026

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("Menu")
        .addItem("Pull Leads", "pullLeads")
        .addToUi();
}

function pullLeads() {
    const sourceSS = SpreadsheetApp.openById(SOURCE_SHEET_ID);
    const destSS = SpreadsheetApp.openById("1ATtLyHfkoVPdzLL5kaEaeDusKL7O7JHDICWZ6RDj0z4");

    const tabsToCopy = ["MAIN", "LABS"];

    tabsToCopy.forEach(tabName => {
        const sourceSheet = sourceSS.getSheetByName(tabName);
        if (!sourceSheet) {
            Logger.log(`Tab "${tabName}" not found in source. Skipping.`);
            return;
        }

        const destSheet = destSS.getSheetByName(tabName);
        if (!destSheet) {
            Logger.log(`Tab "${tabName}" not found in this sheet. Skipping.`);
            return;
        }

        const data = sourceSheet.getDataRange().getValues();
        if (data.length === 0) return;

        const header = data[0];
        const benRows = data.slice(1).filter(row => {
            const colD = String(row[3]).trim();
            return colD.toLowerCase().startsWith("ben");
        });

        destSheet.clearContents();
        const output = [header, ...benRows];
        destSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

        Logger.log(`"${tabName}": pulled ${benRows.length} leads.`);
    });
}