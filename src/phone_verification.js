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
