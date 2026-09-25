const SHEET_NAME = "GameLow";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

    const parsed = JSON.parse(e.postData.contents);
    const dataArray = Array.isArray(parsed) ? parsed : [parsed];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Prolific ID",
        "SessionID",
        "Round",
        "Player Captured",
        "AI Captured"
      ]);
    } else if (sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf("Prolific ID") === -1) {
      sheet.insertColumnAfter(1);
      sheet.getRange(1, 2).setValue("Prolific ID");
    }

    dataArray.forEach(function(data) {
      sheet.appendRow([
        safeCell_(data.timestamp),
        safeCell_(data.prolificId),
        safeCell_(data.sessionID),
        safeCell_(data.round),
        safeCell_(data.playerCaptured),
        safeCell_(data.aiCaptured)
      ]);
    });

    return jsonResponse_({ status: "success" });
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.message });
  } finally {
    lock.releaseLock();
  }
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
