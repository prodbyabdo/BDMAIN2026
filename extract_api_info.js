const fs = require('fs');
const ipqsHtml = fs.readFileSync('docs/apiphone/Phone Number Validation API Documentation _ IPQS.html', 'utf8');

function extractText(html) {
    let text = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
    text = text.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
    text = text.replace(/<[^>]+>/g, ' ');
    text = text.replace(/\s+/g, ' ');
    return text;
}

const ipqsText = extractText(ipqsHtml);
console.log("--- IPQS PHONE ENDPOINTS ---");
const ipqsMatches = ipqsText.match(/https:\/\/[^\s]+/gi);
if (ipqsMatches) {
    console.log([...new Set(ipqsMatches)].filter(m => m.includes('ipqualityscore') && m.includes('phone')));
}

console.log("\n--- JSON EXAMPLES ---");
console.log(ipqsText.match(/\{[^}]+"valid"[^}]+\}/gi)?.slice(0,2));
