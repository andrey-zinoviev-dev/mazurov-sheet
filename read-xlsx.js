require("dotenv").config({ quiet: true });

const path = require("path");
const { google } = require("googleapis");
const XLSX = require("xlsx");

const CREDENTIALS_PATH = path.join(__dirname, "google-credentials.json");
const SCOPE = "https://www.googleapis.com/auth/drive.readonly";

const DRIVE_FILE_ID = process.env.DRIVE_FILE_ID;
const XLSX_SHEET_NAME = process.env.XLSX_SHEET_NAME;

const MIME_GOOGLE_SHEET = "application/vnd.google-apps.spreadsheet";
const MIME_XLSX =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function isHttpStatus(err, status) {
  const s = err?.code;
  if (s === status || s === String(status)) return true;
  return err?.response?.status === status;
}

function isEnoent(err) {
  for (const e of walkErrorChain(err)) {
    if (e?.code === "ENOENT") return true;
  }
  return false;
}

function* walkErrorChain(err) {
  const seen = new Set();
  let current = err;
  while (current && !seen.has(current)) {
    seen.add(current);
    yield current;
    current = current.cause;
  }
}

function assertConfig() {
  if (!DRIVE_FILE_ID) {
    console.error(
      "DRIVE_FILE_ID is missing in .env (id from drive.google.com/file/d/ID/ or spreadsheets/d/ID/)."
    );
    process.exitCode = 1;
    return false;
  }
  return true;
}

async function getDriveClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS_PATH,
    scopes: [SCOPE],
  });
  const authClient = await auth.getClient();
  return google.drive({ version: "v3", auth: authClient });
}

async function downloadXlsxBuffer(drive) {
  const meta = await drive.files.get({
    fileId: DRIVE_FILE_ID,
    fields: "mimeType",
  });
  const mime = meta.data.mimeType;

  if (mime === MIME_GOOGLE_SHEET) {
    const res = await drive.files.export(
      { fileId: DRIVE_FILE_ID, mimeType: MIME_XLSX },
      { responseType: "arraybuffer" }
    );
    return Buffer.from(res.data);
  }

  const res = await drive.files.get(
    { fileId: DRIVE_FILE_ID, alt: "media" },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(res.data);
}

function bufferToRows(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheetName =
    XLSX_SHEET_NAME && workbook.SheetNames.includes(XLSX_SHEET_NAME)
      ? XLSX_SHEET_NAME
      : workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("Workbook has no sheets.");
  }
  if (XLSX_SHEET_NAME && !workbook.SheetNames.includes(XLSX_SHEET_NAME)) {
    throw new Error(
      `Sheet "${XLSX_SHEET_NAME}" not found. Available: ${workbook.SheetNames.join(", ")}`
    );
  }
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
}

function printResults(values) {
  const rows = Array.isArray(values) ? values : [];
  console.log(`Row count: ${rows.length}`);
  if (rows.length) {
    console.table(rows.slice(0, 5));
  }
}

function handleError(err) {
  if (isEnoent(err)) {
    console.error(
      "File ./google-credentials.json is missing. Add your Service Account JSON key and try again."
    );
    return;
  }
  if (isHttpStatus(err, 403)) {
    console.error(
      "Drive API returned 403. In Google Cloud: APIs and Services, Library, enable Google Drive API for the same project as this service account key. Enabling only Google Sheets API is not enough for this script. Also share the file with the service account (Viewer), or add it to a Shared drive the account can access."
    );
    return;
  }
  if (isHttpStatus(err, 404)) {
    console.error(
      "Drive API returned 404. Check DRIVE_FILE_ID in .env and that the file exists and is shared to the service account."
    );
    return;
  }
  console.error(err);
}

async function main() {
  if (!assertConfig()) return;
  const drive = await getDriveClient();
  const buffer = await downloadXlsxBuffer(drive);
  const values = bufferToRows(buffer);
  printResults(values);
}

main().catch((err) => {
  handleError(err);
  process.exitCode = 1;
});
