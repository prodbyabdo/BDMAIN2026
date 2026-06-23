/**
 * Determines if a lead is a qualified DMEDesk prospect.
 * 
 * @param {string} input The lead description or text to analyze.
 * @return {string} "Yes" if it is a qualified DME lead, "No" otherwise, plus a short reason.
 * @customfunction
 */
function IS_DME_LEAD(input) {
  if (!input) return "No input";

  const url = "https://hermes.ai.unturf.com/v1/chat/completions";
  const apiKey = "dummy-api-key"; 

  // Detailed context injected here
  const systemPrompt = `You are a Lead Qualification Specialist for DMEDesk. 
  Your goal is to identify if a business is a high-value DME provider.
  
  CRITERIA FOR 'YES':
  - Specializes in: Orthopedic bracing, Continuous Glucose Monitors (CGM), or Lymphedema products.
  - Capability: Must be able to drop-ship equipment and manage patient intake.
  - Requirement: Focuses on Medicare compliance/billing (PTAN), or private insurances and PPO/HMO Plans.
  - Business Type: Medical suppliers, orthotic/prosthetic providers.
  
  CRITERIA FOR 'NO':
  - Retail-only stores (e.g., general pharmacies, drug stores, non-medical retail).
  - General hospitals or clinics without DME departments.
  - Entities that do not handle equipment/prescriptions.

  Respond in this format: "Yes: [Reason]" or "No: [Reason]". Be brief.`;

  const payload = {
    model: "adamo1139/Hermes-3-Llama-3.1-8B-FP8-Dynamic",
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: input }
    ],
    temperature: 0.1 // Low temperature for more consistent, logical classification
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return json.choices[0].message.content.trim();
  } catch (e) {
    return "Error: " + e.toString();
  }
}