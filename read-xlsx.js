require("dotenv").config({ quiet: true });

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const mysql = require("mysql2/promise");
const XLSX = require("xlsx");

const CREDENTIALS_PATH = path.join(__dirname, "google-credentials.json");
const SCOPE = "https://www.googleapis.com/auth/drive.readonly";

const DRIVE_FILE_ID = process.env.DRIVE_FILE_ID;
const XLSX_SHEET_NAME = process.env.XLSX_SHEET_NAME;
const DB_HOST = process.env.DB_HOST;
const DB_USER = process.env.DB_USER;
const DB_PASSWORD = process.env.DB_PASSWORD;
const DB_NAME = process.env.DB_NAME;
const DB_PORT = parseInt(process.env.DB_PORT || "3306", 10);
const DB_INSERT = String(process.env.DB_INSERT || "1") !== "0";
const DB_TRUNCATE = String(process.env.DB_TRUNCATE || "") === "1";
const DRIVE_STATE_FILE =
  process.env.DRIVE_STATE_FILE ||
  path.join(__dirname, ".drive-import-state.json");
const FORCE_DOWNLOAD =
  String(process.env.FORCE_DOWNLOAD || process.env.FORCE_REFRESH || "") ===
  "1";

// 0-based column indices. Default: A = name, B = qty (e.g. "5 шт"), C = price (e.g. "11 000 ₽")
const COL_NAME = intEnv("COL_NAME", 0);
const COL_STOCK = intEnv("COL_STOCK", 1);
const COL_PRICE = intEnv("COL_PRICE", 2);
// Optional; if the row has no such column, image URL is derived from the assigned id
const COL_IMAGE = intEnv("COL_IMAGE", 3);
const COL_KIT_ORDER = intEnv("COL_KIT_ORDER", 4);
// Rows in column A with only text (no qty/price) and text longer than this are skipped (preface/duplicate)
const CATEGORY_MAX_LEN = intEnv("CATEGORY_MAX_LEN", 48);
const IMAGE_BASE = (
  process.env.IMAGE_BASE_URL || "https://mazurov-rental.ru/images"
).replace(/\/$/, "");
const FALLBACK_CATEGORY = process.env.FALLBACK_CATEGORY || "uncategorized";

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

function intEnv(name, defaultValue) {
  const raw = process.env[name];
  if (raw === undefined || raw === "") return defaultValue;
  const n = parseInt(String(raw), 10);
  return Number.isFinite(n) && n >= 0 ? n : defaultValue;
}

function asTrimmedString(v) {
  if (v == null) return "";
  return String(v).trim();
}

const CATEGORY_SLUG_MAP = {
  камеры: "camera",
  камера: "camera",
  оптика: "optics",
  свет: "lighting",
  свето: "lighting",
  "звук": "audio",
  штатив: "tripod",
  штативы: "tripod",
  аксессуар: "accessories",
  аксессуары: "accessories",
};

const HEADER_PATTERNS = {
  name: (c) =>
    /^(name|название|наименование|модель|состав|товар|позиция)$/i.test(c) ||
    /^модел/i.test(c),
  price: (c) => /^(price|цена|стоимость)$/i.test(c) || /^цена/i.test(c),
  stock: (c) =>
    /^(stock|остаток|кол-?во|кол\.)$/i.test(c) || /^кол-?в/i.test(c) || /^штук/i.test(c),
  image: (c) =>
    /^(image|url|photo|фото|картинк|изображен)/i.test(c) || /^img$/i.test(c),
  kitOrder: (c) => /^(kit_?order|kit|order|порядок|сорт)$/i.test(c),
};

/** A-only rows that are not section categories (sheet title, etc.) */
const SKIP_CATEGORY_A_PATTERNS = [/^\s*список(\s+оборудования)?\s*$/i, /^\s*list(\s+of)?\s*$/i];

function colFromHeaderRow(headerRow, key) {
  for (let i = 0; i < headerRow.length; i++) {
    const c = asTrimmedString(headerRow[i]);
    if (!c) continue;
    const fn = HEADER_PATTERNS[key];
    if (fn && fn(c)) return i;
  }
  return -1;
}

function findHeaderRowInfo(rows) {
  const cap = Math.min(rows.length, 30);
  for (let r = 0; r < cap; r++) {
    const row = rows[r] || [];
    if (!row.length) continue;
    const iName = colFromHeaderRow(row, "name");
    const iPrice = colFromHeaderRow(row, "price");
    const iStock = colFromHeaderRow(row, "stock");
    if (iName >= 0 && (iPrice >= 0 || iStock >= 0)) {
      return { headerIndex: r, headerRow: row };
    }
  }
  return { headerIndex: -1, headerRow: null };
}

function buildColumnMap(headerRow) {
  if (!headerRow) {
    return {
      name: COL_NAME,
      price: COL_PRICE,
      stock: COL_STOCK,
      image: COL_IMAGE,
      kitOrder: COL_KIT_ORDER,
    };
  }
  const hName = colFromHeaderRow(headerRow, "name");
  const hPrice = colFromHeaderRow(headerRow, "price");
  const hStock = colFromHeaderRow(headerRow, "stock");
  const hImage = colFromHeaderRow(headerRow, "image");
  const hKit = colFromHeaderRow(headerRow, "kitOrder");
  return {
    name: hName >= 0 ? hName : COL_NAME,
    price: hPrice >= 0 ? hPrice : COL_PRICE,
    stock: hStock >= 0 ? hStock : COL_STOCK,
    image: hImage >= 0 ? hImage : COL_IMAGE,
    kitOrder: hKit >= 0 ? hKit : COL_KIT_ORDER,
  };
}

// Column A: section titles like "Камеры" (short, no qty/price in B/C) set category; preface with long A and empty B/C is skipped
function isEmptyRow(row) {
  if (!row || !row.length) return true;
  return row.every((c) => c === "" || c == null);
}

function cellAt(row, idx) {
  if (idx < 0 || !row || idx >= row.length) return "";
  const v = row[idx];
  if (v === null || v === undefined) return "";
  return v;
}

/** "5 шт", "1 шт", or a plain integer in the qty column */
function parseQtyRu(v) {
  const s = asTrimmedString(v);
  if (!s) return null;
  if (/шт|штук|pcs?|pc\b|×/i.test(s)) {
    const m = s.match(/(\d+)/);
    return m ? parseInt(m[1], 10) : null;
  }
  const compact = s.replace(/\s/g, "").replace(",", ".");
  if (/^\d+(\.\d+)?$/.test(compact)) {
    const n = Number(compact);
    return Number.isFinite(n) && n >= 0 ? n : null;
  }
  return null;
}

/** "11 000 ₽", "6 000" — digits are concatenated; requires ₽/rub hint or 2+ digit amount */
function parsePriceRu(v) {
  const s = asTrimmedString(v);
  if (!s) return null;
  const onlyDigits = s.replace(/[^\d]/g, "");
  if (!onlyDigits) return null;
  const n = parseInt(onlyDigits, 10);
  if (!Number.isFinite(n) || n <= 0) return null;
  const hasMoneyHint =
    /₽|руб|р\.?|\$|€|eur|usd/i.test(s) || onlyDigits.length >= 2;
  if (!hasMoneyHint) return null;
  return n;
}

function shouldSkipCategoryLine(colA) {
  const t = asTrimmedString(colA);
  if (!t) return true;
  return SKIP_CATEGORY_A_PATTERNS.some((re) => re.test(t));
}

function isProductRow(row, cols) {
  const name = asTrimmedString(cellAt(row, cols.name)) || asTrimmedString(row[0]);
  if (!name) return false;
  const q = parseQtyRu(cellAt(row, cols.stock));
  const p = parsePriceRu(cellAt(row, cols.price));
  return q != null || p != null;
}

function parseNullableNumber(v) {
  const s = asTrimmedString(v);
  if (!s) return null;
  const n = Number(s.replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function parseNullableInt(v) {
  const s = asTrimmedString(v);
  if (!s) return null;
  const n = parseInt(s.replace(/\s/g, ""), 10);
  return Number.isFinite(n) ? n : null;
}

function mapCategoryLabel(colAText) {
  const t = asTrimmedString(colAText);
  if (!t) return null;
  const k = t.toLowerCase();
  if (CATEGORY_SLUG_MAP[k] !== undefined) return CATEGORY_SLUG_MAP[k];
  return t
    .toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[^a-z0-9_а-яё-]/gi, "");
}

function safeCategory(category) {
  const c = asTrimmedString(category);
  return c || FALLBACK_CATEGORY;
}

function buildImageUrl(row, cols, id) {
  const cell = asTrimmedString(cellAt(row, cols.image));
  if (cell) return cell;
  return `${IMAGE_BASE}/${id}.jpg`;
}

/**
 * @param {unknown[][]} rows from sheet_to_json with header: 1
 * @param {{ startId?: number }} [options]
 * @returns {{ items: Array<{ id: number, name: string, category: string | null, price: number | null, stock: number | null, image: string, kit_order: number | null }>, nextId: number }}
 */
function rowsToInventoryItems(rows, options) {
  const startId = options?.startId != null ? options.startId : 1;
  const { headerIndex, headerRow } = findHeaderRowInfo(rows);
  const cols = buildColumnMap(headerRow);
  const startRow = headerIndex >= 0 ? headerIndex + 1 : 0;
  let currentCategory = null;
  const out = [];
  let nextId = startId;

  const applyNonProductRow = (row) => {
    const colA = asTrimmedString(row[0]);
    if (isProductRow(row, cols)) return;
    if (!colA || colA.length > CATEGORY_MAX_LEN) return;
    if (shouldSkipCategoryLine(colA)) return;
    currentCategory = mapCategoryLabel(colA);
  };

  if (headerIndex > 0) {
    for (let r = 0; r < headerIndex; r++) {
      const row = rows[r] || [];
      if (isEmptyRow(row)) continue;
      applyNonProductRow(row);
    }
  }

  for (let r = startRow; r < rows.length; r++) {
    const row = rows[r] || [];
    if (isEmptyRow(row)) continue;

    if (isProductRow(row, cols)) {
      const colA = asTrimmedString(row[0]);
      const name = asTrimmedString(cellAt(row, cols.name)) || colA;
      const id = nextId++;
      out.push({
        id,
        name,
        category: safeCategory(currentCategory),
        price: parsePriceRu(cellAt(row, cols.price)),
        stock: parseQtyRu(cellAt(row, cols.stock)),
        image: buildImageUrl(row, cols, id),
        kit_order: parseNullableInt(cellAt(row, cols.kitOrder)),
      });
      continue;
    }

    applyNonProductRow(row);
  }
  return { items: out, nextId };
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

function hasDbConfig() {
  return Boolean(DB_HOST && DB_USER && DB_NAME);
}

function assertDbConfig() {
  const missing = [];
  if (!DB_HOST) missing.push("DB_HOST");
  if (!DB_USER) missing.push("DB_USER");
  if (!DB_PASSWORD) missing.push("DB_PASSWORD");
  if (!DB_NAME) missing.push("DB_NAME");
  if (missing.length) {
    console.error(
      `MySQL config is incomplete. Missing: ${missing.join(", ")}. Set DB_INSERT=0 to skip DB write.`
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

async function downloadXlsxBuffer(drive, mimeType) {
  const mime = mimeType;

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

function getSheetNamesToProcess(workbook) {
  if (XLSX_SHEET_NAME) {
    if (!workbook.SheetNames.includes(XLSX_SHEET_NAME)) {
      throw new Error(
        `Sheet "${XLSX_SHEET_NAME}" not found. Available: ${workbook.SheetNames.join(", ")}`
      );
    }
    return [XLSX_SHEET_NAME];
  }
  if (!workbook.SheetNames.length) {
    throw new Error("Workbook has no sheets.");
  }
  return workbook.SheetNames.slice();
}

/**
 * Parse all selected sheets in order; ids are one global sequence (for one MySQL table).
 * @param {Buffer} buffer
 */
function bufferToMergedInventory(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const names = getSheetNamesToProcess(workbook);
  const allItems = [];
  let rowTotal = 0;
  let nextId = 1;
  for (const sheetName of names) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    rowTotal += rows.length;
    const { items, nextId: after } = rowsToInventoryItems(rows, {
      startId: nextId,
    });
    nextId = after;
    allItems.push(...items);
  }
  return { items: allItems, sheetNames: names, rowTotal };
}

/** Single sheet (first or XLSX_SHEET_NAME); for tests or custom tools. */
function bufferToRows(buffer, sheetNameOverride) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheetName =
    sheetNameOverride ||
    (XLSX_SHEET_NAME && workbook.SheetNames.includes(XLSX_SHEET_NAME)
      ? XLSX_SHEET_NAME
      : workbook.SheetNames[0]);
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

function loadStoredModifiedTime() {
  try {
    const raw = fs.readFileSync(DRIVE_STATE_FILE, "utf8");
    const data = JSON.parse(raw);
    return typeof data.modifiedTime === "string" ? data.modifiedTime : null;
  } catch (e) {
    if (isEnoent(e)) return null;
    console.error(
      `Could not read ${DRIVE_STATE_FILE}; will re-import on next run. (${e.message})`
    );
    return null;
  }
}

function saveStoredModifiedTime(modifiedTime) {
  const payload = JSON.stringify(
    { modifiedTime, updatedAt: new Date().toISOString() },
    null,
    0
  );
  fs.writeFileSync(DRIVE_STATE_FILE, payload, "utf8");
}

async function getDriveFileMeta(drive) {
  const res = await drive.files.get({
    fileId: DRIVE_FILE_ID,
    fields: "mimeType,modifiedTime",
  });
  return {
    mimeType: res.data.mimeType || "",
    modifiedTime: res.data.modifiedTime || "",
  };
}

function printMergedInventoryJson(result) {
  const { items, sheetNames, rowTotal } = result;
  console.error(
    `Sheets (${sheetNames.length}): ${sheetNames.join(", ")} — ${rowTotal} sheet row(s) scanned → ${items.length} inventory item(s) (JSON to stdout)`
  );
  console.log(JSON.stringify(items, null, 2));
}

async function insertItemsToDb(items) {
  const conn = await mysql.createConnection({
    host: DB_HOST,
    user: DB_USER,
    password: DB_PASSWORD,
    database: DB_NAME,
    port: Number.isFinite(DB_PORT) ? DB_PORT : 3306,
  });

  try {
    await conn.beginTransaction();
    if (DB_TRUNCATE) {
      await conn.execute("TRUNCATE TABLE products");
    }
    const sql = `
      INSERT INTO products (name, category, price, stock, image, kit_order)
      VALUES (?, ?, ?, ?, ?, ?)
    `;
    for (const it of items) {
      await conn.execute(sql, [
        it.name ?? null,
        safeCategory(it.category),
        it.price ?? null,
        it.stock ?? 0,
        it.image ?? null,
        it.kit_order ?? null,
      ]);
    }
    await conn.commit();
    console.error(
      `MySQL insert completed: ${items.length} row(s) into ${DB_NAME}.products`
    );
  } catch (e) {
    await conn.rollback();
    throw e;
  } finally {
    await conn.end();
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
  if (DB_INSERT && !assertDbConfig()) return;
  const drive = await getDriveClient();
  const meta = await getDriveFileMeta(drive);
  const stored = loadStoredModifiedTime();
  if (!FORCE_DOWNLOAD && stored && stored === meta.modifiedTime) {
    console.error(
      "Drive file unchanged (modifiedTime matches local state). Skipping download and import. Set FORCE_DOWNLOAD=1 to re-import."
    );
    return;
  }
  const buffer = await downloadXlsxBuffer(drive, meta.mimeType);
  const result = bufferToMergedInventory(buffer);
  if (DB_INSERT && hasDbConfig()) {
    await insertItemsToDb(result.items);
  }
  printMergedInventoryJson(result);
  saveStoredModifiedTime(meta.modifiedTime);
}

if (require.main === module) {
  main().catch((err) => {
    handleError(err);
    process.exitCode = 1;
  });
} else {
  module.exports = {
    bufferToRows,
    bufferToMergedInventory,
    rowsToInventoryItems,
  };
}
