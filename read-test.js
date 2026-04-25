require("dotenv").config({ quiet: true });

const path = require("path");
const { google } = require("googleapis");

const CREDENTIALS_PATH = path.join(__dirname, "google-credentials.json");
const SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly";

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_RANGE = process.env.SHEET_RANGE;

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
  if (!SPREADSHEET_ID) {
    console.error("SPREADSHEET_ID is missing in .env");
    process.exitCode = 1;
    return false;
  }
  if (!SHEET_RANGE) {
    console.error("SHEET_RANGE is missing in .env");
    process.exitCode = 1;
    return false;
  }
  return true;
}

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS_PATH,
    scopes: [SCOPE],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
}

async function fetchValues(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: SHEET_RANGE,
  });
  return res.data.values;
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
      "The Sheets API returned 403 (Forbidden). Grant the Service Account email 'Viewer' access to this Google Sheet, then retry."
    );
    return;
  }
  if (isHttpStatus(err, 404)) {
    console.error(
      "The Sheets API returned 404. Check that SPREADSHEET_ID in .env matches a real spreadsheet and that the ID is not mistyped."
    );
    return;
  }
  console.error(err);
}

async function main() {
  if (!assertConfig()) return;
  const sheets = await getSheetsClient();
  const values = await fetchValues(sheets);
  printResults(values);
}

main().catch((err) => {
  handleError(err);
  process.exitCode = 1;
});
