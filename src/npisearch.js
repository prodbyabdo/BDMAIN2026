/**
 * Fetches NPI data for a range or single cell using batch processing.
 * @param {Array<Array<string>>|string} range The range containing NPI numbers.
 * @param {string} fields Comma-separated fields (company, person, position, phone, state, lastupdate, taxonomies_group).
 * @return {Array<Array<string>>} The requested data.
 * @customfunction
 */
function dbgetnpi(range, fields = "company") {
  if (!range) return [["No Input"]];

  const grid = Array.isArray(range) 
    ? (Array.isArray(range[0]) ? range : [range]) 
    : [[range]];

  const requestedFields = fields.toLowerCase().split(",").map(f => f.trim());
  
  const npiToFetch = [...new Set(grid.flat().map(cell => {
    if (!cell) return null;
    const match = cell.toString().match(/(\d{10})/);
    return match ? match[1] : null;
  }).filter(Boolean))];

  const resultsMap = {};

  if (npiToFetch.length > 0) {
    const requests = npiToFetch.map(npi => ({
      url: `https://npiregistry.cms.hhs.gov/api/?number=${npi}&version=2.1`,
      method: "get",
      muteHttpExceptions: true
    }));

    try {
      const responses = UrlFetchApp.fetchAll(requests);
      
      responses.forEach((res, i) => {
        const npi = npiToFetch[i];
        
        if (res.getResponseCode() !== 200) {
          resultsMap[npi] = requestedFields.map(() => "Error");
          return;
        }

        const data = JSON.parse(res.getContentText());
        const result = data.results?.[0];

        if (!result) {
          resultsMap[npi] = requestedFields.map(() => "Not Found");
          return;
        }

        const { basic = {}, addresses = [], taxonomies = [] } = result;
        const addr = addresses[0] || {};

        // ---- NEW: join all taxonomy descriptions with newline ----
        const allTaxonomies = taxonomies
          .map(t => t.desc)
          .filter(Boolean)
          .join("\n") || "N/A";

        const fieldValues = {
          company: basic.organization_name || "N/A",
          person: [
            basic.first_name || basic.authorized_official_first_name,
            basic.last_name || basic.authorized_official_last_name
          ].filter(Boolean).join(" ") || "N/A",
          position: basic.authorized_official_title_or_position || "N/A",
          phone : basic.authorized_official_telephone_number || "N/A",
          companyphone: addr.telephone_number || "N/A",
          state: addr.state || "N/A",
          city: addr.city || "N/A",
          lastupdate: basic.last_updated || "N/A",
          taxonomies_group: taxonomies[0]?.desc || "N/A",   // kept for backward compatibility
          taxonomies_all: allTaxonomies,                   // <-- NEW field
          enum: basic.enumeration_date || "N/A"
        };

        resultsMap[npi] = requestedFields.map(f => fieldValues[f] || "N/A");
      });
    } catch (e) {
      return grid.map(row => requestedFields.map(() => "Fetch Error"));
    }
  }

  return grid.map(row => {
    return row.flatMap(cell => {
      if (!cell) return requestedFields.map(() => "");

      const match = cell.toString().match(/(\d{10})/);
      if (!match) return requestedFields.map(() => "Invalid NPI");

      return resultsMap[match[1]] || requestedFields.map(() => "No Data");
    });
  });
}
function autofillNpiData() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName("NEWDME");

  if (!sheet) {
    ss.toast('Sheet "MAIN" not found.', 'Error', 5);
    return;
  }

  const BATCH_SIZE  = 50;
  const NPI_COL     = 15;   // O
  const UPDATE_COL  = 7;    // G  — UpdateDate (guard)
  const WRITE_START = 7;    // G  — first column we write into
  const WRITE_WIDTH = 8;    // <-- now 8 columns: G through N

  // Field order matches write columns G → N exactly
  const FIELD_ORDER = [
    "lastupdate",     // G
    "company",        // H
    "state",          // I
    "companyphone",   // J
    "person",         // K
    "phone",          // L
    "position",       // M
    "taxonomies_all"  // N   <-- NEW: all taxonomy descriptions
  ];
  const FIELDS_STR = FIELD_ORDER.join(",");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast('No data rows found.', 'Done', 3);
    return;
  }

  const npiValues    = sheet.getRange(2, NPI_COL,    lastRow - 1, 1).getValues();
  const updateValues = sheet.getRange(2, UPDATE_COL, lastRow - 1, 1).getValues();

  const workList = [];

  const extractNpi = (cell) => {
    if (!cell) return null;
    let str = String(cell).trim();
    if (str.includes('.')) str = str.split('.')[0];
    const cleaned = str.replace(/\D/g, '');
    return cleaned.length === 10 ? cleaned : null;
  };

  for (let i = 0; i < npiValues.length; i++) {
    const npi        = extractNpi(npiValues[i][0]);
    const hasUpdate  = String(updateValues[i][0] || "").trim() !== "";
    if (npi && !hasUpdate) {
      workList.push({ sheetRow: i + 2, npi });
    }
  }

  if (workList.length === 0) {
    ss.toast('All NPI rows already have UpdateDate filled.', 'Done', 5);
    return;
  }

  ss.toast(`Found ${workList.length} rows to fill. Starting...`, 'NPI Autofill', 4);

  let filled  = 0;
  let errors  = 0;
  const totalBatches = Math.ceil(workList.length / BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchStart = batchIndex * BATCH_SIZE;
    const batch      = workList.slice(batchStart, batchStart + BATCH_SIZE);

    ss.toast(
      `Batch ${batchIndex + 1}/${totalBatches} — fetching ${batch.length} NPIs...`,
      'NPI Autofill', 5
    );

    const npiGrid = batch.map(item => [item.npi]);

    let results;
    try {
      results = cbgetnpi(npiGrid, FIELDS_STR);
    } catch (err) {
      console.error(`Batch ${batchIndex + 1} fetch failed: ${err.message}`);
      errors += batch.length;
      continue;
    }

    for (let i = 0; i < batch.length; i++) {
      const rowData = results[i];

      if (!rowData || rowData.length === 0) {
        errors++;
        continue;
      }

      const firstVal = String(rowData[0] || "").trim();
      if (firstVal === "Fetch Error" || firstVal === "Not Found" || firstVal === "Error") {
        console.warn(`Row ${batch[i].sheetRow} NPI ${batch[i].npi}: ${firstVal} — skipping write.`);
        errors++;
        continue;
      }

      sheet
        .getRange(batch[i].sheetRow, WRITE_START, 1, WRITE_WIDTH)
        .setValues([rowData]);

      filled++;
    }

    SpreadsheetApp.flush();
  }

  ss.toast(
    `Done. Filled: ${filled} rows. Skipped/errors: ${errors}.`,
    'NPI Autofill Complete', 8
  );
}