// API Keys for Phone Verification
// Replace these with your actual API keys.
const PHONE_API_KEYS = {
    APITIER: "YOUR_APITIER_KEY",
    IPQS: "YOUR_IPQS_KEY"
};

/**
 * Custom function to verify a phone number using a stacked API approach.
 * Falls over to the next API if quota is exhausted or an error occurs.
 * 
 * @param {string|number} phoneNumber The phone number to verify.
 * @returns {string} Returns "Dis/Wn" if invalid, or empty string if valid or unknown.
 * @customfunction
 */
function verifyPhoneStacked(phoneNumber) {
    if (!phoneNumber) return "";
    
    const cleanPhone = String(phoneNumber).replace(/\D/g, '');
    if (cleanPhone.length < 10) return "Dis/Wn";
    
    // API 1: Antideo (Free Tier, No Key Required)
    try {
        const url = `https://api.antideo.com/phone/${cleanPhone}`;
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
            const result = JSON.parse(response.getContentText());
            if (result.valid === false) return "Dis/Wn";
            return ""; 
        }
    } catch (e) {
        // Fallback
    }
    
    // API 2: APITier
    if (PHONE_API_KEYS.APITIER && PHONE_API_KEYS.APITIER !== "YOUR_APITIER_KEY") {
        try {
            const url = `https://phone.apitier.com/v1/validate?number=${cleanPhone}&x-api-key=${PHONE_API_KEYS.APITIER}`;
            const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
            if (response.getResponseCode() === 200) {
                const result = JSON.parse(response.getContentText());
                if (result.valid === false) return "Dis/Wn";
                return "";
            }
        } catch (e) {
            // Fallback
        }
    }
    
    // API 3: IP Quality Score
    if (PHONE_API_KEYS.IPQS && PHONE_API_KEYS.IPQS !== "YOUR_IPQS_KEY") {
        try {
            const url = `https://www.ipqualityscore.com/api/json/phone/${PHONE_API_KEYS.IPQS}/${cleanPhone}`;
            const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
            if (response.getResponseCode() === 200) {
                const result = JSON.parse(response.getContentText());
                if (result.success === true && (result.valid === false || result.active === false)) {
                    return "Dis/Wn";
                }
                return "";
            }
        } catch (e) {
            // Fallback
        }
    }
    
    return "";
}

/**
 * Menu entry point to verify checked rows in Column S.
 */
function verifySelectedPhones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const allowedTabs = ["MAIN", "LABS", "NEWDME"];
  
  if (!allowedTabs.includes(sheet.getName())) {
    ss.toast("This function can only be run on MAIN, LABS, or NEWDME tabs.", "Blocked");
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    ss.toast("Another operation is running. Please try again shortly.", "Locked");
    return;
  }
  
  try {
    ss.toast("Loading verification data...", "Verification");
    
    // 1. Load existing Disconnected numbers into memory Set
    const disconnectedPhones = new Set();
    const disconnectedSheet = ss.getSheetByName("Disconnected");
    if (disconnectedSheet) {
      const deacLastRow = disconnectedSheet.getLastRow();
      if (deacLastRow >= 2) {
        // Reads from Col H to Col L (5 columns) starting at row 2
        const deacData = disconnectedSheet.getRange(2, COL_SEARCH_START, deacLastRow - 1, COL_LOOKUP_WIDTH_MAX).getValues();
        deacData.forEach(row => {
          const p1 = row[COL_LOOKUP_PH1_OFFSET];
          const p2 = row[COL_LOOKUP_PH2_OFFSET];
          if (p1) disconnectedPhones.add(String(p1).replace(/\D/g, ''));
          if (p2) disconnectedPhones.add(String(p2).replace(/\D/g, ''));
        });
      }
    }
    
    // 2. Load/Create Verified Active cache sheet
    let verifiedActivePhones = new Set();
    let verifiedActiveSheet = ss.getSheetByName("Verified_Valid");
    if (!verifiedActiveSheet) {
      verifiedActiveSheet = ss.insertSheet("Verified_Valid");
      verifiedActiveSheet.appendRow(["Phone Number", "Verification Date"]);
    } else {
      const activeLastRow = verifiedActiveSheet.getLastRow();
      if (activeLastRow >= 2) {
        const activeData = verifiedActiveSheet.getRange(2, 1, activeLastRow - 1, 1).getValues();
        activeData.forEach(row => {
          if (row[0]) verifiedActivePhones.add(String(row[0]).replace(/\D/g, ''));
        });
      }
    }
    
    // 3. Scan Sheet for Checkboxes
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let apiCallsCount = 0;
    const maxApiCalls = 30; // Safety guardrail
    
    const newlyDisconnected = [];
    const newlyVerifiedValid = [];
    
    for (let r = 1; r < values.length; r++) {
      const rowNum = r + 1;
      const isChecked = values[r][COL_VERIFY_CHECKBOX - 1]; // Column S
      
      if (isChecked === true) {
        const rawPhone1 = values[r][COL_PHONE1 - 1]; // Column J
        const rawPhone2 = values[r][COL_PHONE2 - 1]; // Column L
        
        const cleanPhone1 = String(rawPhone1 || '').replace(/\D/g, '');
        const cleanPhone2 = String(rawPhone2 || '').replace(/\D/g, '');
        
        let p1Result = "";
        let p2Result = "";
        let rowUpdated = false;
        
        // --- Process Phone 1 ---
        if (cleanPhone1.length >= 10) {
          if (disconnectedPhones.has(cleanPhone1)) {
            p1Result = "Dis/Wn";
            rowUpdated = true;
          } else if (verifiedActivePhones.has(cleanPhone1)) {
            p1Result = ""; // Already validated as active
          } else if (apiCallsCount < maxApiCalls) {
            const apiRes = verifyPhoneStacked(cleanPhone1);
            apiCallsCount++;
            if (apiRes === "Dis/Wn") {
              p1Result = "Dis/Wn";
              disconnectedPhones.add(cleanPhone1);
              newlyDisconnected.push([`API_Verified_${new Date().toLocaleDateString()}`, "", cleanPhone1, "", cleanPhone1]); 
              rowUpdated = true;
            } else {
              verifiedActivePhones.add(cleanPhone1);
              newlyVerifiedValid.push([cleanPhone1, new Date().toISOString()]);
            }
          }
        }
        
        // --- Process Phone 2 ---
        if (cleanPhone2.length >= 10) {
          if (disconnectedPhones.has(cleanPhone2)) {
            p2Result = "Dis/Wn";
            rowUpdated = true;
          } else if (verifiedActivePhones.has(cleanPhone2)) {
            p2Result = ""; // Already validated as active
          } else if (apiCallsCount < maxApiCalls) {
            const apiRes = verifyPhoneStacked(cleanPhone2);
            apiCallsCount++;
            if (apiRes === "Dis/Wn") {
              p2Result = "Dis/Wn";
              disconnectedPhones.add(cleanPhone2);
              newlyDisconnected.push([`API_Verified_${new Date().toLocaleDateString()}`, "", cleanPhone2, "", cleanPhone2]);
              rowUpdated = true;
            } else {
              verifiedActivePhones.add(cleanPhone2);
              newlyVerifiedValid.push([cleanPhone2, new Date().toISOString()]);
            }
          }
        }
        
        // Write changes to current row
        if (rowUpdated) {
          if (p1Result === "Dis/Wn") sheet.getRange(rowNum, COL_PHONE1_MATCH).setValue("Dis/Wn");
          if (p2Result === "Dis/Wn") sheet.getRange(rowNum, COL_PHONE2_MATCH).setValue("Dis/Wn");
        }
        
        // Uncheck the checkbox
        sheet.getRange(rowNum, COL_VERIFY_CHECKBOX).setValue(false);
      }
    }
    
    // 4. Bulk Write changes to Lookup Sheets
    if (newlyDisconnected.length > 0 && disconnectedSheet) {
      const deacLast = disconnectedSheet.getLastRow();
      disconnectedSheet.getRange(deacLast + 1, COL_SEARCH_START, newlyDisconnected.length, 5).setValues(newlyDisconnected);
    }
    
    if (newlyVerifiedValid.length > 0 && verifiedActiveSheet) {
      const activeLast = verifiedActiveSheet.getLastRow();
      verifiedActiveSheet.getRange(activeLast + 1, 1, newlyVerifiedValid.length, 2).setValues(newlyVerifiedValid);
    }
    
    SpreadsheetApp.flush();
    ss.toast(`Successfully verified numbers (${apiCallsCount} API queries run).`, "Done");
  } finally {
    lock.releaseLock();
  }
}
