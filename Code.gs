
// Rows are written to this tab (bottom of the spreadsheet). Not the Google Form "Form Responses" sheet.
const SHEET_NAME = "Registrations";
const PAYMENT_PROOFS_FOLDER_NAME = "Great Skill Circuit – Payment proofs";

const STATIONS = [
  "Mystery Lab",
  "Writing Den",
  "Creative Studio",
  "Byte Zone",
  "Innovation Garage",
];
const STATION_MAX = 20;

const HEADERS = [
  "Timestamp",
  "Full Name",
  "Phone Number",
  "Email",
  "Age",
  "Gender",
  "How heard about event",
  "Referred by (name)",
  "Station Preferences",
  "Payment proof file name",
  "Payment proof link",
  "Assigned Station",
  "Terms accepted",
];

function doGet() {
  return ContentService.createTextOutput(
    "Web app is live. Submissions are saved to the sheet tab named: " + SHEET_NAME
  ).setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    Logger.log("doPost received");
    const sheet = getOrCreateSheet_();
    ensureHeaders_(sheet);

    const fullName = clean_(e.parameter.fullName);
    const phoneNumber = clean_(e.parameter.phoneNumber);
    const email = clean_(e.parameter.email);
    const age = clean_(e.parameter.age);
    const gender = clean_(e.parameter.gender);
    const heardHow = clean_(e.parameter.heardHow);
    const heardPersonName = clean_(e.parameter.heardPersonName);
    const paymentProofName = clean_(e.parameter.paymentProofName);
    const paymentProofBase64 = clean_(e.parameter.paymentProofBase64);
    const paymentProofMime = clean_(e.parameter.paymentProofMime) || "image/png";
    const termsAccepted = clean_(e.parameter.termsAccepted);

    const stationPrefs = normalizePrefs_(getStationPrefsFromEvent_(e));

    if (!fullName || !phoneNumber || !email || !age || !gender || stationPrefs.length === 0) {
      return json_({ ok: false, message: "Missing required fields." });
    }
    if (!heardHow) {
      return json_({ ok: false, message: "Please tell us how you heard about the event." });
    }
    if (heardHow === "Through a person" && !heardPersonName) {
      return json_({ ok: false, message: "Please enter the name of the person who told you." });
    }
    if (termsAccepted !== "yes") {
      return json_({ ok: false, message: "Please accept the event terms to continue." });
    }

    var paymentProofLink = "";
    if (paymentProofBase64 && paymentProofName) {
      try {
        paymentProofLink = savePaymentProofToDrive_(paymentProofBase64, paymentProofMime, paymentProofName, fullName);
      } catch (driveErr) {
        Logger.log("Drive upload error: " + driveErr);
        return json_({ ok: false, message: "Could not save payment proof. Try a smaller image or authorize Drive for this script." });
      }
    } else {
      return json_({ ok: false, message: "Payment screenshot is required." });
    }

    const assignedStation = allocateStation_(sheet, stationPrefs);

    Logger.log("Appending row for: " + fullName);
    // Plain URL — Google Sheets turns it into a clickable link automatically.
    sheet.appendRow([
      new Date(),
      fullName,
      phoneNumber,
      email,
      age,
      gender,
      heardHow,
      heardHow === "Through a person" ? heardPersonName : "",
      stationPrefs.join(", "),
      paymentProofName,
      paymentProofLink,
      assignedStation,
      "Yes",
    ]);

    return json_({ ok: true, assignedStation: assignedStation });
  } catch (err) {
    return json_({ ok: false, message: String(err) });
  }
}

function savePaymentProofToDrive_(base64, mime, fileName, registrantName) {
  const bytes = Utilities.base64Decode(base64);
  const safeName =
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm") +
    "_" +
    String(registrantName || "registrant")
      .replace(/[^\w\s-]/g, "")
      .slice(0, 40) +
    "_" +
    (fileName || "proof.png");
  const blob = Utilities.newBlob(bytes, mime, safeName);
  const folder = getOrCreateProofsFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateProofsFolder_() {
  const it = DriveApp.getFoldersByName(PAYMENT_PROOFS_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PAYMENT_PROOFS_FOLDER_NAME);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }
  const existing = sheet
    .getRange(1, 1, 1, HEADERS.length)
    .getValues()[0]
    .map((h) => String(h || "").trim());
  const isEmpty = existing.every((h) => !h);
  if (isEmpty) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

function allocateStation_(sheet, prefs) {
  const assignedCol = HEADERS.indexOf("Assigned Station") + 1;
  const lastRow = sheet.getLastRow();
  const counts = {};
  STATIONS.forEach((s) => (counts[s] = 0));

  if (lastRow >= 2) {
    const assignedValues = sheet
      .getRange(2, assignedCol, lastRow, 1)
      .getValues()
      .flat()
      .map((v) => String(v).trim())
      .filter((v) => v);
    assignedValues.forEach((v) => {
      if (counts[v] !== undefined) counts[v] += 1;
    });
  }

  for (var i = 0; i < prefs.length; i++) {
    var pref = prefs[i];
    if (counts[pref] !== undefined && counts[pref] < STATION_MAX) return pref;
  }
  return "WAITLIST";
}

function normalizePrefs_(raw) {
  if (!raw) return [];
  const list = Array.isArray(raw) ? raw : [raw];
  const out = [];
  list.forEach((item) => {
    const s = String(item || "").trim();
    if (!s) return;
    s.split(",").forEach((part) => {
      const p = part.trim();
      if (p) out.push(p);
    });
  });
  return out.slice(0, 3);
}

function clean_(v) {
  return String(v || "").trim();
}

function getStationPrefsFromEvent_(e) {
  if (!e) return [];
  if (e.parameters && e.parameters.stationPrefs !== undefined) {
    const v = e.parameters.stationPrefs;
    return Array.isArray(v) ? v : [v];
  }
  if (e.parameter && e.parameter.stationPrefs !== undefined) {
    return [e.parameter.stationPrefs];
  }
  return [];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
