
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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Registration form")
    .addItem("Fix Registrations column layout", "fixRegistrationSheetLayout")
    .addToUi();
}

/** Run once from the sheet menu if columns still look wrong after updating the script. */
function fixRegistrationSheetLayout() {
  const sheet = getOrCreateSheet_();
  ensureHeaders_(sheet);
  SpreadsheetApp.getUi().alert(
    "Registrations tab checked. If the first row was the old 9-column header, data rows were realigned to match the form (13 columns)."
  );
}

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

    // Payment proof: try Drive upload. Always append a row so registrations appear even if upload fails
    // (large/truncated POST payloads used to skip appendRow entirely).
    var paymentProofLink = "";
    if (paymentProofBase64 && paymentProofName) {
      try {
        paymentProofLink = savePaymentProofToDrive_(
          paymentProofBase64,
          paymentProofMime,
          paymentProofName,
          fullName
        );
      } catch (driveErr) {
        Logger.log("Drive upload error: " + driveErr);
        paymentProofLink =
          "Drive upload failed — authorize Drive for this script or use a smaller image. Registrant: " +
          email;
      }
    } else {
      paymentProofLink =
        "No file received — payload may be too large; ask registrant to resend a smaller screenshot.";
    }

    const assignedStation = allocateStation_(sheet, stationPrefs);

    Logger.log("Appending row for: " + fullName);
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
      paymentProofName || "(none)",
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

/** Old layout (9 cols): prefs were in G, no "how heard" / proof link / terms columns. */
const LEGACY_HEADER_G_STATION_PREFS = "Station Preferences";

/**
 * True when row values match the OLD 9-column pattern: proof filename in H, nothing in J+,
 * and not the new pattern (heard-how in G is usually short; prefs span multiple stations).
 */
function looksLegacyNineColDataRow_(row) {
  const proofName = String(row[7] || "").trim(); // col H
  const colJ = String(row[9] || "").trim(); // col J
  if (colJ) return false;
  if (!proofName) return false;
  if (!/\.(png|jpg|jpeg|webp)$/i.test(proofName) && proofName.indexOf("Screenshot") === -1)
    return false;
  return true;
}

/**
 * Fixes sheets that still show the 9-column header while doPost writes 13 values — that made
 * new rows look "shifted". Rebuilds header + moves legacy rows into the correct columns.
 * Rows already in 13-column format (e.g. J has proof filename) are left as-is.
 */
function normalizeLegacyRegistrationSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const g1 = String(sheet.getRange(1, 7).getValue() || "").trim();
  if (g1 !== LEGACY_HEADER_G_STATION_PREFS) return;

  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const out = [];
  out.push(HEADERS.slice());

  for (var r = 1; r < data.length; r++) {
    var row = data[r].slice();
    while (row.length < HEADERS.length) row.push("");

    if (looksLegacyNineColDataRow_(row)) {
      out.push([
        row[0],
        row[1],
        row[2],
        row[3],
        row[4],
        row[5],
        "",
        "",
        row[6],
        row[7],
        "",
        row[8],
        "Yes",
      ]);
      continue;
    }

    // Already 13-wide (or new submission): keep A–M as sent by the web app
    out.push(row.slice(0, HEADERS.length));
  }

  sheet.clearContents();
  sheet.getRange(1, 1, out.length, HEADERS.length).setValues(out);
}

function headersRowMatches_(existing) {
  for (var i = 0; i < HEADERS.length; i++) {
    if (String(existing[i] || "").trim() !== HEADERS[i]) return false;
  }
  return true;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  const g1 = String(sheet.getRange(1, 7).getValue() || "").trim();
  if (g1 === LEGACY_HEADER_G_STATION_PREFS) {
    normalizeLegacyRegistrationSheet_(sheet);
    return;
  }

  const width = Math.max(sheet.getLastColumn(), HEADERS.length);
  const existing = sheet
    .getRange(1, 1, 1, width)
    .getValues()[0]
    .map(function (h) {
      return String(h || "").trim();
    });
  const isEmpty = existing.every(function (h) {
    return !h;
  });
  if (isEmpty) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  var headSlice = existing.slice(0, HEADERS.length);
  if (!headersRowMatches_(headSlice)) {
    // Refresh labels only; do not clear data (migration handled legacy case above).
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
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
