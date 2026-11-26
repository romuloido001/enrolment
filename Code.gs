/**
 * Calinan National High School - Enrolment Portal backend
 * (c) 2026 ridojr. All rights reserved.
 *
 * SETUP
 * 1. Create a Google Sheet (e.g. "CNHS Enrolment"). Open Extensions > Apps Script.
 * 2. Paste this file as Code.gs and save.
 * 3. Deploy > New deployment > Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Authorize, copy the Web App URL, paste it into SCRIPT_URL in index.html.
 * Uploaded PSA Birth Certificates and Report Cards go to a Drive folder
 * named "CNHS Enrolment Uploads"; links are written to the sheet.
 * Re-deploy (new version) after any edit to this file.
 */

var SHEET_NAME = "Enrolments";
var FOLDER_NAME = "CNHS Enrolment Uploads";
var HEADERS = ["Timestamp","Reference","School Year","Grade","Strand","Learner Type","LRN",
  "Last Name","First Name","Middle Name","Ext","Sex","Birthdate","Previous School",
  "Guardian","Mobile","Email","Address","PSA Birth Certificate","Report Card","Status","Copyright"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    var folder = getFolder_();
    var sub = folder.createFolder(d.reference + " - " + clean_(d.lastName) + ", " + clean_(d.firstName));

    var psa = saveFile_(sub, "PSA Birth Certificate", d.files.psa);
    var card = saveFile_(sub, "Report Card", d.files.card);

    sheet.appendRow([
      new Date(), d.reference, d.schoolYear, "Grade " + d.grade, d.strand || "", d.learnerType,
      "'" + (d.lrn || ""), clean_(d.lastName), clean_(d.firstName), clean_(d.middleName), clean_(d.ext),
      d.sex, d.birthdate, clean_(d.prevSchool), clean_(d.guardian), "'" + clean_(d.phone),
      clean_(d.email), clean_(d.address), psa.getUrl(), card.getUrl(), "Pending", "ridojr"
    ]);
    return json_({ ok: true, reference: d.reference });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json_({ ok: true, service: "CNHS Enrolment Portal by ridojr" }); }

function saveFile_(folder, label, f) {
  var ext = (f.name.match(/\.[A-Za-z0-9]+$/) || [""])[0];
  var blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.type, label + ext);
  return folder.createFile(blob);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#0f5c36").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  return sh;
}

function getFolder_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

// Stops spreadsheet formula injection from typed text.
function clean_(v) {
  v = String(v == null ? "" : v).trim();
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
