const API_KEY = "AIzaSyD0o7cZiVMiCZb5ZBLwSJljCoiM1irFcsg";
const SHEET_NAME = "Sheet1";

function processNewRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  for (let row = 2; row <= lastRow; row++) {
    const inputData = sheet.getRange(row, 2).getValue();
    const status = sheet.getRange(row, 5).getValue();

    if (!inputData || status === "DONE" || status === "PROCESSING") continue;

    try {
      sheet.getRange(row, 5).setValue("PROCESSING");
      SpreadsheetApp.flush();

      const result = callGemini(inputData);

      sheet.getRange(row, 3).setValue(result.summary);
      sheet.getRange(row, 4).setValue(result.actions);
      sheet.getRange(row, 5).setValue("DONE");
      sheet.getRange(row, 6).setValue(new Date().toISOString());

    } catch (e) {
      sheet.getRange(row, 5).setValue("ERROR");
      sheet.getRange(row, 6).setValue(e.message);
    }
  }
}

function callGemini(inputText) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + API_KEY;

  const payload = {
    contents: [{
      parts: [{
        text: "You are a business analyst. Given raw input data, return a JSON object with exactly two keys: 'summary' (2-3 sentence summary) and 'actions' (comma-separated action items). Return only valid JSON, nothing else.\n\nInput: " + inputText
      }]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error("Gemini API error: HTTP " + responseCode + " - " + response.getContentText());
  }

  const json = JSON.parse(response.getContentText());
  const content = json.candidates[0].content.parts[0].text.trim()
    .replace(/```json/g, "").replace(/```/g, "").trim();

  try {
    return JSON.parse(content);
  } catch (e) {
    return {
      summary: content,
      actions: "Could not parse actions"
    };
  }
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  if (range.getColumn() === 2 && sheet.getName() === SHEET_NAME) {
    processNewRows();
  }
}
