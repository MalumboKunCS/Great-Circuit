
// Rows are written to this tab (bottom tabs in the spreadsheet). Not the Google Form "Form Responses" sheet.
const SHEET_NAME = "Registrations";
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
  "Station Preferences",
  "Payment Proof Name",
  "Assigned Station",
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
    const paymentProofName = clean_(e.parameter.paymentProofName);
    const stationPrefs = normalizePrefs_(getStationPrefsFromEvent_(e));

    if (!fullName || !phoneNumber || !email || !age || !gender || stationPrefs.length === 0) {
      return json_({ ok: false, message: "Missing required fields." });
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
      stationPrefs.join(", "),
      paymentProofName,
      assignedStation,
    ]);

    return json_({ ok: true, assignedStation: assignedStation });
  } catch (err) {
    return json_({ ok: false, message: String(err) });
  }
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0].map((h) => String(h || "").trim());
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
      .getRange(2, assignedCol, lastRow - 1, 1)
      .getValues()
      .flat()
      .map((v) => String(v).trim())
      .filter((v) => v);
    assignedValues.forEach((v) => {
      if (counts[v] !== undefined) counts[v] += 1;
    });
  }

  for (const pref of prefs) {
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

/** Multiple hidden inputs named stationPrefs → e.parameters.stationPrefs is an array. */
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

