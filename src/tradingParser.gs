// Trading parser module (Phase 1-2 split only).
// TODO(v2.2): Harden Gemini response parsing and validation paths.
function callGeminiVision(imageUrl, prompt) {
  const imgBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  const base64Img = Utilities.base64Encode(imgBlob.getBytes());
  const payload = { "contents": [{ "parts": [{ "text": prompt }, { "inline_data": { "mime_type": "image/jpeg", "data": base64Img } }] }] };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
 
  const response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + GEMINI_API_KEY, options);
  const rawResponse = response.getContentText();
  const parsed = JSON.parse(rawResponse);

  if (!parsed.candidates || !parsed.candidates[0]) {
    throw new Error("Gemini Vision response missing candidates.");
  }

  if (!parsed.candidates[0].content) {
    throw new Error("Gemini Vision response missing content.");
  }

  if (!parsed.candidates[0].content.parts || !parsed.candidates[0].content.parts[0]) {
    throw new Error("Gemini Vision response missing content parts.");
  }

  if (!parsed.candidates[0].content.parts[0].text) {
    throw new Error("Gemini Vision response missing text output.");
  }

  const resText = parsed.candidates[0].content.parts[0].text;
 
  const start = resText.indexOf('[');
  const end = resText.lastIndexOf(']') + 1;
  if (start !== -1 && end !== -1) return JSON.parse(resText.substring(start, end));
  throw new Error("Format JSON tidak valid.");
}
