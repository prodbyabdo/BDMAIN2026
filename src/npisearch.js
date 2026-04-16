/** * Gets NPI data. Works with raw 10-digit NPIs or URLs.
 * * @customfunction  */
function getnpi(input, fields = "company,person,position,phone") {
    if (!input) return "";
    // Extract 10-digit NPI from number or URL string
    const match = input.toString().match(/(\d{10})/);
    if (!match) return "Invalid NPI";
    const npi = match[1];
    
    const cache = CacheService.getScriptCache();
    const cacheKey = `npi_${npi}_${fields}`;
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const apiUrl = `https://npiregistry.cms.hhs.gov/api/?number=${npi}&version=2.1`;
    try {
        const response = UrlFetchApp.fetch(apiUrl);
        const data = JSON.parse(response.getContentText());
        const result = data.results?.[0];
        if (!result) return "Not Found";
        const { basic = {}, addresses = [], taxonomies = [] } = result;
        const addr = addresses[0] || {};
        // Map all potential data points
        const fieldValues = {
            'company': basic.organization_name || "N/A",
            'person': [basic.first_name || basic.authorized_official_first_name, basic.last_name || basic.authorized_official_last_name].filter(Boolean).join(" ") || "N/A",
            'position': basic.authorized_official_title_or_position || "N/A",
            'phone': basic.authorized_official_telephone_number || addr.telephone_number || "N/A",
            'state': addr.state || "N/A",
            'lastupdate': basic.last_updated || "N/A",
            'taxonomies_group': taxonomies[0]?.desc || "N/A"
        };
        // Process requested fields
        const requested = fields.toLowerCase().split(",").map(f => f.trim());
        const output = requested.map(f => fieldValues[f] || "");
        // Return single string if 1 field, or a horizontal array if multiple
        const finalResult = output.length === 1 ? output[0] : [output];
        cache.put(cacheKey, JSON.stringify(finalResult), 21600); // 6 hours
        return finalResult;

    } catch (err) {
        console.error(err);
        return "Error";
    }
}
// Shortcuts for bundles
const gettt4 = (input) => getnpi(input, "company,person,position,phone");
const gettt6 = (input) => getnpi(input, "company,person,position,phone,lastupdate,taxonomies_group");