/**
 * Calinan National High School - Enrolment Portal backend
 * (c) 2026 ridojr. All rights reserved.
 *
 * SETUP
 * 1. Open your Google Sheet > Extensions > Apps Script. Paste this file as Code.gs and save.
 * 2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 3. Paste the Web App URL into SCRIPT_URL in index.html.
 * After editing this file: Deploy > Manage deployments > Edit > New version > Deploy.
 *
 * Each LRN can enrol only once. A repeated LRN is rejected before anything is saved.
 */

var SHEET_NAME = "Enrolments";
var FOLDER_NAME = "CNHS Enrolment Uploads";
var LRN_COL = 7, PHONE_COL = 16;
var HEADERS = ["Timestamp","Reference","School Year","Grade","Strand","Learner Type","LRN",
  "Last Name","First Name","Middle Name","Ext","Sex","Birthdate","Previous School",
  "Guardian","Mobile","Email","Address","PSA Birth Certificate","Report Card","Status","Copyright"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    var lrn = lrnKey_(d.lrn);

    if (lrn && lrnExists_(sheet, lrn)) {
      return json_({ ok: false, code: "DUPLICATE_LRN",
        error: "This LRN has already been used for an enrolment." });
    }

    var folder = getFolder_();
    var sub = folder.createFolder(d.reference + " - " + clean_(d.lastName) + ", " + clean_(d.firstName));
    var psa = saveFile_(sub, "PSA Birth Certificate", d.files.psa);
    var card = saveFile_(sub, "Report Card", d.files.card);

    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, LRN_COL).setNumberFormat("@");
    sheet.getRange(row, PHONE_COL).setNumberFormat("@");
    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
      new Date(), d.reference, d.schoolYear, "Grade " + d.grade, d.strand || "", d.learnerType,
      lrn, clean_(d.lastName), clean_(d.firstName), clean_(d.middleName), clean_(d.ext),
      d.sex, d.birthdate, clean_(d.prevSchool), clean_(d.guardian), clean_(d.phone),
      clean_(d.email), clean_(d.address), psa.getUrl(), card.getUrl(), "Pending", "ridojr"
    ]]);
    return json_({ ok: true, reference: d.reference });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// GET ?lrn=123456789012 tells the page early whether the LRN is already used.
function doGet(e) {
  var lrn = lrnKey_(e && e.parameter && e.parameter.lrn);
  if (lrn) return json_({ ok: true, version: 2, exists: lrnExists_(getSheet_(), lrn) });
  return json_({ ok: true, version: 2, service: "CNHS Enrolment Portal by ridojr" });
}

function lrnExists_(sheet, lrn) {
  var n = sheet.getLastRow() - 1;
  if (n < 1) return false;
  var vals = sheet.getRange(2, LRN_COL, n, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (lrnKey_(vals[i][0]) === lrn) return true;
  }
  return false;
}

// Digits only, padded to 12 so a leading zero lost by the sheet still matches.
function lrnKey_(v) {
  var s = digits_(v);
  while (s && s.length < 12) s = "0" + s;
  return s;
}

function digits_(v) { return String(v == null ? "" : v).replace(/\D/g, ""); }

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
