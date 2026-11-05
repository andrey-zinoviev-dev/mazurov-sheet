require("dotenv").config();
const { google } = require("googleapis");
const mysql = require("mysql2/promise");
const path = require("path");
const fs = require("fs");

/**
 * Sync Google Sheet → MySQL: categories + products.
 *
 * Секции (подкатегории) определяются по оформлению колонки A:
 * bold + серый фон + fontSize 20 → upsert в categories (parent_id = id листа-родителя).
 * Товары: category_id → categories.id, price DECIMAL, item_order.
 *
 * last_sync.txt: синхронизация только если таблица в Drive изменилась.
 */

const UNCATEGORIZED_SLUG = "uncategorized";
const UNCATEGORIZED_NAME = "Без категории";

/** Имя листа → slug корневой категории (parent_id = NULL). */
const SHEET_TO_PARENT = Object.freeze({
    "Камеры и трансляция": "camera_broadcast",
    Звук: "audio",
    Свет: "light",
    Прочее: "other",
});

const SHEETS_CONFIG = Object.keys(SHEET_TO_PARENT);

/**
 * Первая строка листа в диапазоне чтения (колонка A–F).
 * Задайте SHEET_DATA_FIRST_ROW в .env, если нужно отрезать служебные строки сверху.
 */
const _firstRowParsed = parseInt(process.env.SHEET_DATA_FIRST_ROW ?? "1", 10);
// const SHEET_DATA_FIRST_ROW = process.env.SHEET_DATA_FIRST_ROW ? process.env.SHEET_DATA_FIRST_ROW : Number.isFinite(_firstRowParsed) && _firstRowParsed >= 1 ? _firstRowParsed : 1;
const SHEET_DATA_FIRST_ROW = Number(process.env.SHEET_DATA_FIRST_ROW) || 3;
const LAST_SYNC_FILE = path.join(__dirname, "last_sync.txt");

const CYR_TO_LAT = {
    а: "a",
    б: "b",
    в: "v",
    г: "g",
    д: "d",
    е: "e",
    ё: "e",
    ж: "zh",
    з: "z",
    и: "i",
    й: "y",
    к: "k",
    л: "l",
    м: "m",
    н: "n",
    о: "o",
    п: "p",
    р: "r",
    с: "s",
    т: "t",
    у: "u",
    ф: "f",
    х: "h",
    ц: "ts",
    ч: "ch",
    ш: "sh",
    щ: "sch",
    ъ: "",
    ы: "y",
    ь: "",
    э: "e",
    ю: "yu",
    я: "ya",
};

function normalizeCategoryKey(s) {
    return String(s || "")
        .replace(/\uFEFF/g, "")
        .replace(/[\u200B-\u200D\u2060]/g, "")
        .normalize("NFKC")
        .replace(/\u00a0/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function slugifyCategoryName(name) {
    const normalized = normalizeCategoryKey(name).toLowerCase();
    let out = "";
    for (const ch of normalized) {
        if (CYR_TO_LAT[ch] !== undefined) {
            out += CYR_TO_LAT[ch];
        } else if (/[a-z0-9]/.test(ch)) {
            out += ch;
        } else {
            out += "-";
        }
    }
    return out
        .replace(/-+/g, "-")
        .replace(/^-|-$/g, "")
        .slice(0, 64);
}

function escapeSheetTitleForRange(title) {
    return `'${String(title).replace(/'/g, "''")}'`;
}

/** Серый фон секций в таблице ≈ rgb(208,206,206); допускаем соседние оттенки. */
function isSectionGrayColor(color) {
    if (!color) return false;
    const r = Math.round((color.red ?? 0) * 255);
    const g = Math.round((color.green ?? 0) * 255);
    const b = Math.round((color.blue ?? 0) * 255);
    return r >= 200 && r <= 220 && g >= 200 && g <= 220 && b >= 200 && b <= 220;
}

function pickCellBgColor(fmt) {
    if (!fmt) return null;
    if (fmt.backgroundColor) return fmt.backgroundColor;
    if (fmt.backgroundColorStyle?.rgbColor) return fmt.backgroundColorStyle.rgbColor;
    return null;
}

/**
 * Оформление колонки A: секция = bold + серый фон + кегль 20.
 * @returns {Array<{ bold: boolean, gray: boolean, fontSize: number|null, isSectionStyle: boolean }>}
 */
async function fetchColumnASectionFormats(sheets, spreadsheetId, sheetName, startRow, rowCount) {
    const empty = { bold: false, gray: false, fontSize: null, isSectionStyle: false };
    const out = Array.from({ length: rowCount }, () => ({ ...empty }));
    if (rowCount <= 0) return out;

    const endRow = startRow + rowCount - 1;
    const range = `${escapeSheetTitleForRange(sheetName)}!A${startRow}:A${endRow}`;
    try {
        const res = await sheets.spreadsheets.get({
            spreadsheetId,
            ranges: [range],
            includeGridData: true,
            fields: [
                "sheets.data.rowData.values.userEnteredFormat",
                "sheets.data.rowData.values.effectiveFormat",
                "sheets.data.rowData.values.textFormatRuns",
            ].join(","),
        });
        const grid = res.data.sheets?.[0]?.data?.[0];
        const rowData = grid?.rowData || [];
        if (rowData.length !== rowCount) {
            console.warn(
                `  ⚠ ${sheetName}: values=${rowCount}, rowData=${rowData.length} (риск сдвига индексов)`
            );
        }
        for (let i = 0; i < rowCount; i++) {
            const cell = rowData[i]?.values?.[0];
            const ue = cell?.userEnteredFormat;
            const ef = cell?.effectiveFormat;
            const bold =
                ue?.textFormat?.bold === true ||
                ef?.textFormat?.bold === true ||
                (cell?.textFormatRuns || []).some((r) => r.format?.bold === true);
            const gray =
                isSectionGrayColor(pickCellBgColor(ef)) ||
                isSectionGrayColor(pickCellBgColor(ue));
            const fontSize = ef?.textFormat?.fontSize ?? ue?.textFormat?.fontSize ?? null;
            const isSectionStyle = bold && gray && fontSize === 20;
            out[i] = { bold, gray, fontSize, isSectionStyle };
        }
    } catch (e) {
        console.warn(
            `Не удалось прочитать оформление колонки A (${sheetName}): ${e.message}.`
        );
    }
    return out;
}

function isKomplektHeading(normalizedA) {
    return /^комплект/i.test(normalizedA);
}

function isSkipHeading(normalizedA) {
    return normalizedA === "Список оборудования";
}

/**
 * Секция = текст в A + bold + серый фон + fontSize 20.
 * @returns {{ kind: 'skip' } | { kind: 'komplekt', name: string } | { kind: 'section', name: string } | null}
 */
function resolveSectionHeader(cellAText, format) {
    const a = normalizeCategoryKey(cellAText);
    if (!a) return null;
    if (isSkipHeading(a)) return { kind: "skip" };
    if (!format?.isSectionStyle) return null;

    if (isKomplektHeading(a)) {
        return { kind: "komplekt", name: a };
    }
    return { kind: "section", name: a };
}

/**
 * Секции (дети): UNIQUE(parent_id, slug) — ON DUPLICATE KEY надёжен.
 * parent_id здесь всегда NOT NULL.
 */
async function upsertCategory(connection, { parentId, slug, name, sortOrder }) {
    const [result] = await connection.execute(
        `INSERT INTO categories (parent_id, slug, name, sort_order, is_active)
         VALUES (?, ?, ?, ?, 1)
         ON DUPLICATE KEY UPDATE
           name = VALUES(name),
           sort_order = VALUES(sort_order),
           is_active = 1,
           id = LAST_INSERT_ID(id)`,
        [parentId, slug, name, sortOrder]
    );
    return result.insertId;
}

/**
 * Корни (parent_id IS NULL): UNIQUE(parent_id, slug) в MySQL не ловит дублики NULL,
 * поэтому явный SELECT → UPDATE / INSERT.
 * Покрывает и уже существующие корни, и новые (добавленные в SHEET_TO_PARENT).
 */
async function upsertRootCategory(connection, { slug, name, sortOrder }) {
    const [rows] = await connection.execute(
        `SELECT id FROM categories WHERE parent_id IS NULL AND slug = ? LIMIT 1`,
        [slug]
    );
    if (rows.length > 0) {
        const id = rows[0].id;
        await connection.execute(
            `UPDATE categories
             SET name = ?, sort_order = ?, is_active = 1
             WHERE id = ?`,
            [name, sortOrder, id]
        );
        return id;
    }
    const [result] = await connection.execute(
        `INSERT INTO categories (parent_id, slug, name, sort_order, is_active)
         VALUES (NULL, ?, ?, ?, 1)`,
        [slug, name, sortOrder]
    );
    return result.insertId;
}

/**
 * Корневые категории листов + uncategorized (для товаров без секции / комплектов).
 * @returns {{ parentIdsBySlug: Record<string, number>, uncategorizedId: number }}
 */
async function ensureRootCategories(connection) {
    const parentIdsBySlug = {};
    let sortOrder = 0;
    for (const sheetName of SHEETS_CONFIG) {
        const slug = SHEET_TO_PARENT[sheetName];
        sortOrder += 1;
        parentIdsBySlug[slug] = await upsertRootCategory(connection, {
            slug,
            name: sheetName,
            sortOrder,
        });
        console.log(`  root cat: ${sheetName} → ${slug} (id=${parentIdsBySlug[slug]})`);
    }

    const uncategorizedId = await upsertRootCategory(connection, {
        slug: UNCATEGORIZED_SLUG,
        name: UNCATEGORIZED_NAME,
        sortOrder: sortOrder + 1,
    });
    console.log(`  root cat: ${UNCATEGORIZED_NAME} → ${UNCATEGORIZED_SLUG} (id=${uncategorizedId})`);

    return { parentIdsBySlug, uncategorizedId };
}

async function upsertProduct(connection, { id, name, price, stock, categoryId, itemOrder }) {
    await connection.execute(
        `INSERT INTO products (id, name, category_id, price, stock, item_order, is_active)
         VALUES (?, ?, ?, ?, ?, ?, 1)
         ON DUPLICATE KEY UPDATE
           name = VALUES(name),
           category_id = VALUES(category_id),
           price = VALUES(price),
           stock = VALUES(stock),
           item_order = VALUES(item_order),
           is_active = 1`,
        [id, name, categoryId, price, stock, itemOrder]
    );
}

function parseProductRow(row) {
    const rawName = row[0];
    const rawStock = row[1];
    const rawPrice = row[2];
    const rawSku = row[4];
    const id = parseInt(rawSku, 10);
    const name = rawName ? String(rawName).trim() : "";
    if (Number.isNaN(id) || id <= 0 || !name) return null;

    const stock = parseInt(String(rawStock || "0").replace(/\D/g, ""), 10) || 0;
    const priceSheet =
        parseFloat(
            String(rawPrice || "0")
                .replace(/[^\d,.-]/g, "")
                .replace(",", ".")
        ) || 0;
    // В таблице — полная цена, в каталог пишем половину (DECIMAL)
    const price = Math.round((priceSheet / 2) * 100) / 100;
    return { id, name, stock, price };
}

/** Нужна ли синхронизация: сравниваем modifiedTime таблицы в Drive с last_sync.txt. */
async function spreadsheetNeedsSyncFromDrive(drive, fileId) {
    try {
        const response = await drive.files.get({
            fileId: fileId,
            fields: "modifiedTime",
        });

        const cloudTime = new Date(response.data.modifiedTime).getTime();

        if (!fs.existsSync(LAST_SYNC_FILE)) {
            return true;
        }

        const localTime = parseInt(fs.readFileSync(LAST_SYNC_FILE, "utf8"), 10);
        if (Number.isNaN(localTime)) {
            return true;
        }

        return cloudTime > localTime;
    } catch (error) {
        console.error("Ошибка при проверке даты изменения:", error.message);
        return true;
    }
}

/** Записать штамп только после успешной записи в MySQL. */
async function writeLastSyncStamp(drive, fileId) {
    try {
        const response = await drive.files.get({
            fileId: fileId,
            fields: "modifiedTime",
        });
        const t = new Date(response.data.modifiedTime).getTime();
        fs.writeFileSync(LAST_SYNC_FILE, String(t));
    } catch (e) {
        console.warn(`Не удалось записать ${LAST_SYNC_FILE} после синхронизации:`, e.message);
    }
}

async function runSync() {
    let connection;
    try {
        const spreadsheetId = process.env.SPREADSHEET_ID;
        if (!spreadsheetId) {
            throw new Error("SPREADSHEET_ID не задан в .env");
        }

        connection = await mysql.createConnection({
            host: process.env.DB_HOST,
            user: process.env.DB_USER,
            password: process.env.DB_PASSWORD,
            database: process.env.DB_NAME,
        });

        const auth = new google.auth.GoogleAuth({
            keyFile: path.join(__dirname, "google-credentials.json"),
            scopes: [
                "https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.metadata.readonly",
            ],
        });

        const sheets = google.sheets({ version: "v4", auth });
        const drive = google.drive({ version: "v3", auth });

        const needUpdate = await spreadsheetNeedsSyncFromDrive(drive, spreadsheetId);

        if (!needUpdate) {
            console.log("--- Изменений в Google Sheet не обнаружено. Синхронизация пропущена. ---");
            return;
        }

        console.log("--- Обнаружены изменения. Начинаю синхронизацию... ---");

        console.log("--- Корневые категории ---");
        const { parentIdsBySlug, uncategorizedId } = await ensureRootCategories(connection);

        let totalCategories = 0;
        let itemOrder = 0;
        const allProcessedSkus = [];

        for (const sheetName of SHEETS_CONFIG) {
            const parentSlug = SHEET_TO_PARENT[sheetName];
            const parentId = parentIdsBySlug[parentSlug];
            console.log(`--- Лист "${sheetName}" → parent=${parentSlug} (id=${parentId}) ---`);

            const response = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `${sheetName}!A${SHEET_DATA_FIRST_ROW}:${process.env.SHEET_END_COLUMN}`,
            });

            const rows = response.data.values;
            if (!rows?.length) {
                console.log("  (пусто)");
                continue;
            }

            const formatsA = await fetchColumnASectionFormats(
                sheets,
                spreadsheetId,
                sheetName,
                SHEET_DATA_FIRST_ROW,
                rows.length
            );

            let sortOrder = 0;
            let currentCategoryId = uncategorizedId;
            let productsInSheet = 0;

            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const header = resolveSectionHeader(row[0], formatsA[i]);

                if (header?.kind === "skip") continue;

                if (header?.kind === "komplekt") {
                    currentCategoryId = uncategorizedId;
                    console.log(`  section komplekt → uncategorized: "${header.name}"`);
                    continue;
                }

                if (header?.kind === "section") {
                    const slug = slugifyCategoryName(header.name);
                    if (!slug) {
                        console.warn(`  пустой slug для "${header.name}" — пропуск`);
                        continue;
                    }
                    sortOrder += 1;
                    currentCategoryId = await upsertCategory(connection, {
                        parentId,
                        slug,
                        name: header.name,
                        sortOrder,
                    });
                    totalCategories += 1;
                    console.log(
                        `  + cat ${sortOrder}. ${header.name} → ${slug} (id=${currentCategoryId})`
                    );
                    continue;
                }

                const product = parseProductRow(row);
                if (!product) continue;

                let categoryId = currentCategoryId;
                if (/^комплект/i.test(normalizeCategoryKey(product.name))) {
                    categoryId = uncategorizedId;
                }

                itemOrder += 1;
                await upsertProduct(connection, {
                    id: product.id,
                    name: product.name,
                    price: product.price,
                    stock: product.stock,
                    categoryId,
                    itemOrder,
                });
                allProcessedSkus.push(product.id);
                productsInSheet += 1;
            }

            console.log(`  products upserted on sheet: ${productsInSheet}`);
        }

        if (allProcessedSkus.length > 0) {
            const [result] = await connection.execute(
                `UPDATE products SET is_active = 0
                 WHERE is_active = 1 AND id NOT IN (${allProcessedSkus.join(",")})`
            );
            console.log(`--- Синхронизация завершена. Деактивировано товаров: ${result.affectedRows} ---`);
        }

        console.log(`--- Categories upserted (passes): ${totalCategories} ---`);
        console.log(`--- Products touched: ${allProcessedSkus.length} ---`);

        await writeLastSyncStamp(drive, spreadsheetId);
    } catch (err) {
        console.error("Критическая ошибка:", err.message);
        process.exitCode = 1;
    } finally {
        if (connection) await connection.end();
    }
}

runSync();
