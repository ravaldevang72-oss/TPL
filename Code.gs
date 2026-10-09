// TPL Season 2 – Player Registration backend (Google Apps Script)
// Saves each player as a row in the active Google Sheet and the photo in Google Drive.
const SHEET_NAME  = "Players";
const FOLDER_NAME = "TPL Player Photos";

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || "").trim().slice(0, 60);
    const mobile = String(d.mobile || "").trim();
    const village = String(d.village || "").trim().slice(0, 60);
    const category = String(d.category || "");

    if (!name || !village) return out({ ok: false, error: "Name and village are required." });
    if (!/^[6-9]\d{9}$/.test(mobile)) return out({ ok: false, error: "Invalid mobile number." });
    if (["Batsman", "Bowler", "All Rounder"].indexOf(category) < 0) return out({ ok: false, error: "Invalid category." });
    if (!d.photo || d.photo.indexOf("base64,") < 0) return out({ ok: false, error: "Photo is required." });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(["Timestamp", "Player Name", "Mobile", "Village", "Category", "Photo Link"]);
      sh.getRange("C:C").setNumberFormat("@");
      sh.setFrozenRows(1);
    }

    // Block duplicate mobile numbers
    if (sh.getLastRow() > 1) {
      const mobiles = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().flat().map(String);
      if (mobiles.indexOf(mobile) >= 0) return out({ ok: false, error: "This mobile number is already registered." });
    }

    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
    const blob = Utilities.newBlob(Utilities.base64Decode(d.photo.split("base64,")[1]), "image/jpeg",
                                   name.replace(/[^\w ]/g, "") + "_" + mobile + ".jpg");
    const file = folder.createFile(blob);

    sh.appendRow([new Date(), name, mobile, village, category, file.getUrl()]);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: "Server error: " + err.message });
  } finally {
    lock.releaseLock();
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
