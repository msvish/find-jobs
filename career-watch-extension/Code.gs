// Google Apps Script for your Career Watch sheet.
// Setup: open the sheet > Extensions > Apps Script > paste this > Deploy > New deployment
//   Type: Web app | Execute as: Me | Who has access: Anyone
// Copy the /exec URL into the extension's config.js.

const TOKEN = "change-me";  // pick any secret; put the same value in config.js
const SHEET_GID = 0;        // same tab the watcher reads

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) return reply({ ok: false, error: "Wrong token. Check config.js." });

    const name = String(data.company || "").trim();
    const url = String(data.url || "").trim();
    if (!name || !/^https?:\/\//i.test(url)) {
      return reply({ ok: false, error: "Enter a company name and a full https:// link." });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()
      .find(s => s.getSheetId() === SHEET_GID);
    const last = sheet.getLastRow();
    const existing = last > 1 ? sheet.getRange(2, 2, last - 1, 1).getValues().flat() : [];
    if (existing.includes(url)) return reply({ ok: false, error: "That careers page is already in the sheet." });

    sheet.appendRow([name, url]);  // CompanyName | CareersPage
    return reply({ ok: true, row: sheet.getLastRow() });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
