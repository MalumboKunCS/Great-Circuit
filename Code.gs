
// Rows are written to this tab (bottom of the spreadsheet). Not the Google Form "Form Responses" sheet.
const SHEET_NAME = "Registrations";

/** Google Form file-upload folders (exact names as in Drive). Proofs go here — not created by this script. */
const PAYMENT_PROOFS_PARENT_FOLDER_NAME = "The Great Circuit (File responses)";
const PAYMENT_PROOFS_SUBFOLDER_NAME = "Payment Screenshot (File responses)";
/**
 * Optional: paste ONLY the folder ID (about 25–45 chars), OR the full browser URL.
 * Wrong: pasting a long string or two IDs — getFolderById will fail.
 * Example URL: https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz
 *              → ID is only: 1AbCdEfGhIjKlMnOpQrStUvWxYz
 */
const PAYMENT_PROOFS_FOLDER_ID = "1rnqd-HDnFykqHnoPF43FW4qXXT3kqwZj";

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
    .addItem("Test Drive upload (tiny image)", "testPaymentProofUploadToDrive")
    .addItem("Log folder IDs (fast — View → Logs)", "listPaymentProofFolderIds_")
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
        paymentProofLink = formatDriveFailureForSheet_(driveErr, email);
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

/** Short text for the sheet so you can see the real failure (truncated). */
function formatDriveFailureForSheet_(driveErr, email) {
  var msg = String(driveErr || "Unknown error").replace(/\s+/g, " ").trim();
  if (msg.length > 160) msg = msg.substring(0, 157) + "...";
  return (
    "Drive upload failed — " +
    msg +
    " | Registrant: " +
    email +
    " | Check: Deploy Web app as Execute as: Me + Apps Script → Executions."
  );
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
  const folder = getPaymentProofsTargetFolder_();
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {
    // File exists; some Workspace orgs block "anyone with the link". Owner can still open in Drive.
    Logger.log("setSharing skipped or failed (file created): " + shareErr);
  }
  return file.getUrl();
}

/**
 * Run this from Apps Script (▶ Run) after saving. If it completes and shows a URL in the dialog,
 * Drive is authorized for this project. If it fails, read the error — same as failed form uploads.
 */
function testPaymentProofUploadToDrive() {
  const tinyPngBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
  const url = savePaymentProofToDrive_(tinyPngBase64, "image/png", "test-proof.png", "DriveTest");
  SpreadsheetApp.getUi().alert("Drive upload OK. Test file URL (open in browser):\n\n" + url);
}

function normalizeFolderLabel_(s) {
  return String(s || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function folderNamesMatch_(a, b) {
  return normalizeFolderLabel_(a) === normalizeFolderLabel_(b);
}

/** Direct child only — case-insensitive name. */
function findDirectChildFolderInsensitive_(parentFolder, childName) {
  const want = normalizeFolderLabel_(childName);
  const it = parentFolder.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (normalizeFolderLabel_(f.getName()) === want) return f;
  }
  return null;
}

/**
 * Accepts a bare ID or a full Drive URL; returns the folder resource id only.
 * Google folder/file IDs are typically 25–45 chars (letters, digits, _ -).
 */
function extractDriveFolderId_(raw) {
  var s = String(raw || "")
    .trim()
    .replace(/^["']|["']$/g, "");
  if (!s) return "";
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1].split(/[?#]/)[0];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]+$/.test(s)) return s;
  return "";
}

function isLikelyDriveFolderId_(id) {
  return typeof id === "string" && /^[a-zA-Z0-9_-]{20,60}$/.test(id);
}

/**
 * Uses Form folders under My Drive.
 * IMPORTANT: DriveApp.getFoldersByName() searches your whole Drive and can take many minutes on
 * large accounts — we avoid it. Prefer PAYMENT_PROOFS_FOLDER_ID (copy from the folder URL).
 */
function getPaymentProofsTargetFolder_() {
  var rawId = PAYMENT_PROOFS_FOLDER_ID && String(PAYMENT_PROOFS_FOLDER_ID).trim();
  if (rawId) {
    var folderId = extractDriveFolderId_(rawId);
    if (!folderId) {
      throw new Error(
        "PAYMENT_PROOFS_FOLDER_ID is not valid. Paste only the ID (letters/numbers after /folders/) or the full folder URL."
      );
    }
    if (!isLikelyDriveFolderId_(folderId)) {
      throw new Error(
        'Folder ID looks wrong (length ' +
          folderId.length +
          '). Open the folder in Drive, copy the link, and use only the part after /folders/ — usually about 33 characters, not a long or doubled string.'
      );
    }
    try {
      return DriveApp.getFolderById(folderId);
    } catch (idErr) {
      throw new Error(
        "Cannot open folder id " +
          folderId +
          ". Check the ID, that this account owns the folder, and Drive access is authorized. " +
          idErr
      );
    }
  }

  const parentName = PAYMENT_PROOFS_PARENT_FOLDER_NAME;
  const childName = PAYMENT_PROOFS_SUBFOLDER_NAME;
  var root = DriveApp.getRootFolder();

  // 1) Fast path only: under My Drive root (case-insensitive), then direct child
  var top = root.getFolders();
  while (top.hasNext()) {
    var folder = top.next();
    if (!folderNamesMatch_(folder.getName(), parentName)) continue;
    var found = findDirectChildFolderInsensitive_(folder, childName);
    if (found) return found;
  }

  // 2) One level deeper: e.g. Shared/shortcut wrappers — still bounded, no whole-Drive scan
  top = root.getFolders();
  while (top.hasNext()) {
    var level1 = top.next();
    var inner = level1.getFolders();
    while (inner.hasNext()) {
      var maybeParent = inner.next();
      if (!folderNamesMatch_(maybeParent.getName(), parentName)) continue;
      found = findDirectChildFolderInsensitive_(maybeParent, childName);
      if (found) return found;
    }
  }

  throw new Error(
    'Could not find "' +
      childName +
      '" under "' +
      parentName +
      '" within the first two levels of My Drive. Open that inner folder in Drive, copy the ID from the URL (folders/XXXX), and set PAYMENT_PROOFS_FOLDER_ID in Code.gs. Run listPaymentProofFolderIds_ for top-level names only.'
  );
}

/**
 * Fast: logs only top-level My Drive folders + immediate children of folders whose names look
 * Form-related. Does NOT walk your whole Drive (that can run for many minutes).
 * View → Logs after running.
 */
function listPaymentProofFolderIds_() {
  var root = DriveApp.getRootFolder();
  Logger.log("=== Top-level folders in My Drive (id = copy into PAYMENT_PROOFS_FOLDER_ID) ===");
  var it = root.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    Logger.log('"' + f.getName() + '"  id=' + f.getId());
  }
  Logger.log("=== Immediate children of folders matching Payment / Circuit / Great / Screenshot ===");
  it = root.getFolders();
  var maxParents = 80;
  while (it.hasNext() && maxParents-- > 0) {
    var parent = it.next();
    var pn = parent.getName();
    if (!/payment|circuit|great|screenshot|file responses/i.test(pn)) continue;
    Logger.log('--- under "' + pn + '" ---');
    var sub = parent.getFolders();
    var maxKids = 40;
    while (sub.hasNext() && maxKids-- > 0) {
      var c = sub.next();
      Logger.log('  "' + c.getName() + '"  id=' + c.getId());
    }
  }
  Logger.log("Done. For uploads, paste the inner folder id into PAYMENT_PROOFS_FOLDER_ID.");
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
