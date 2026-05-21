/**
 * Fetches NPI data for a range or single cell using batch processing.
 * @param {Array<Array<string>>|string} range The range containing NPI numbers.
 * @param {string} fields Comma-separated fields (company, person, position, phone, state, lastupdate, taxonomies_group).
 * @return {Array<Array<string>>} The requested data.
 * @customfunction
 */
function bgetnpi(range, fields = "company") {
  if (!range) return [["No Input"]];

  // Normalize input: handles single cells, 1D arrays, and 2D arrays
  const grid = Array.isArray(range)
    ? (Array.isArray(range[0]) ? range : [range])
    : [[range]];

  const requestedFields = fields.toLowerCase().split(",").map(f => f.trim());

  // Extract unique NPIs
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

        const fieldValues = {
          company: basic.organization_name || "N/A",
          person: [
            basic.first_name || basic.authorized_official_first_name,
            basic.last_name || basic.authorized_official_last_name
          ].filter(Boolean).join(" ") || "N/A",
          position: basic.authorized_official_title_or_position || "N/A",
          phone: basic.authorized_official_telephone_number || "N/A",
          companyphone: addr.telephone_number || "N/A",
          state: addr.state || "N/A",
          lastupdate: basic.last_updated || "N/A",
          taxonomies_group: taxonomies[0]?.desc || "N/A",
          enum: basic.enumeration_date || "N/A"
        };

        resultsMap[npi] = requestedFields.map(f => fieldValues[f] || "N/A");
      });
    } catch (e) {
      return grid.map(row => requestedFields.map(() => "Fetch Error"));
    }
  }

  // Map results back to the exact grid structure required by Sheets
  return grid.map(row => {
    return row.flatMap(cell => {
      if (!cell) return requestedFields.map(() => "");

      const match = cell.toString().match(/(\d{10})/);
      if (!match) return requestedFields.map(() => "Invalid NPI");

      return resultsMap[match[1]] || requestedFields.map(() => "No Data");
    });
  });
}