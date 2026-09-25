const INITIAL_SHEET_NAME = "InitialLow";
const FINAL_SHEET_NAME = "FinalLow";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse(e.postData.contents);
    const sheetName = payload.surveyType === "initial"
      ? INITIAL_SHEET_NAME
      : payload.surveyType === "final"
        ? FINAL_SHEET_NAME
        : null;

    if (!sheetName) throw new Error("Unknown survey type.");
    if (!payload.responses || !payload.responses.prolificId) {
      throw new Error("A Prolific ID is required.");
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
    const row = Object.assign({ serverTimestamp: new Date() }, payload.responses);
    appendObjectRow_(sheet, row);

    return jsonResponse_({ status: "success" });
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function appendObjectRow_(sheet, record) {
  let headers = [];

  if (sheet.getLastRow() > 0) {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  Object.keys(record).forEach(function(key) {
    if (headers.indexOf(key) === -1) headers.push(key);
  });

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else if (headers.length > sheet.getLastColumn()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const values = headers.map(function(header) {
    return safeCell_(record[header]);
  });
  sheet.appendRow(values);
}

function safeCell_(value) {
  if (typeof value === "string" && /^[=+\-@]/.test(value)) return "'" + value;
  return value === undefined || value === null ? "" : value;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
