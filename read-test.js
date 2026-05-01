require('dotenv').config();
const { google } = require('googleapis');
const mysql = require('mysql2/promise');
const path = require('path');

const ALLOWED_CATEGORIES = [
    "camera", "lens", "adapter", "filter", "tripod", "stabilizer", "focus", "operator", "rig",
    "audio_mixer", "stagebox", "ethercon", "audio_recorder", "wireless_mic", "iem", "dpa_mic",
    "dynamic_mic", "condenser_mic", "shotgun_mic", "mic_mount", "di_box", "audio_interface",
    "audio_monitor", "speaker", "speaker_kit", "speaker_stand", "xlr_cable", "audio_power",
    "audio_accessory", "comms", "wireless_kit", "light_fixture", "light_modifier", "light_stand",
    "grip", "frame", "flag", "sandbag", "light_power", "light_accessory", "ring_light", "light_kit",
    "storage", "reader", "case", "bag", "computer", "broadcast", "network", "video_monitor",
    "recorder", "converter", "capture", "cable", "power", "battery", "tool", "furniture",
    "vehicle", "work_light", "cart", "cable_channel", "power_tool", "misc", "kit"
];

const SHEETS_CONFIG = ["Камеры и трансляция", "Звук", "Свет", "Прочее"];

// Логика разделения текста и изображения
function processContent(value) {
    if (!value) return { note: null, image: null };
    const str = String(value).trim();
    const isImageUrl = /^http(s)?:\/\/.*\.(jpg|jpeg|png|webp|gif)/i.test(str);
    return isImageUrl ? { note: null, image: str } : { note: str, image: null };
}

async function runSync() {
    let connection;
    try {
        connection = await mysql.createConnection({
            host: process.env.DB_HOST,
            user: process.env.DB_USER,
            password: process.env.DB_PASSWORD,
            database: process.env.DB_NAME
        });

        const auth = new google.auth.GoogleAuth({
            keyFile: path.join(__dirname, "google-credentials.json"),
            scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
        });

        const sheets = google.sheets({ version: 'v4', auth });
        const allProcessedSkus = [];

        for (const sheetName of SHEETS_CONFIG) {
            console.log(`--- Обработка листа: ${sheetName} ---`);
            
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: process.env.SPREADSHEET_ID,
                range: `${sheetName}!A4:F`, 
            });

            const rows = response.data.values;
            if (!rows) continue;

            for (const row of rows) {
                const title = row[0];        // A
                const stock = parseInt(row[1]) || 0; // B
                // Чистка цены от пробелов и замена запятой на точку
                const priceRaw = String(row[2] || "0").replace(/\s/g, '').replace(',', '.');
                const price = parseFloat(priceRaw) || 0;
                const rawNoteImage = row[3]; // D
                const sku = parseInt(row[4]); // E
                let category = String(row[5] || 'misc').trim().toLowerCase(); // F

                // КРИТИЧЕСКИЙ ФИЛЬТР: Игнорируем всё, что не имеет валидного SKU (числа)
                if (isNaN(sku) || !title || title.trim() === "") continue;

                if (!ALLOWED_CATEGORIES.includes(category)) {
                    category = 'misc';
                }

                const { note, image } = processContent(rawNoteImage);
                allProcessedSkus.push(sku);

                await connection.execute(
                    `INSERT INTO products (sku, title, price, stock, note, image_url, category, is_active) 
                     VALUES (?, ?, ?, ?, ?, ?, ?, 1) 
                     ON DUPLICATE KEY UPDATE 
                     title = VALUES(title), 
                     price = VALUES(price), 
                     stock = VALUES(stock), 
                     note = VALUES(note), 
                     image_url = VALUES(image_url), 
                     category = VALUES(category), 
                     is_active = 1`,
                    [sku, title, price, stock, note, image, category]
                );
            }
        }

        // Мягкое удаление: если товара нет в таблице, деактивируем его в БД
        if (allProcessedSkus.length > 0) {
            const [delResult] = await connection.query(
                'UPDATE products SET is_active = 0 WHERE sku NOT IN (?)',
                [allProcessedSkus]
            );
            console.log(`Деактивировано строк: ${delResult.changedRows}`);
        }

        console.log(`Успешно. Всего SKU в базе: ${allProcessedSkus.length}`);

    } catch (err) {
        console.error('Критическая ошибка:', err.message);
    } finally {
        if (connection) await connection.end();
    }
}

runSync();




// require("dotenv").config({ quiet: true });

// const path = require("path");
// const { google } = require("googleapis");

// const CREDENTIALS_PATH = path.join(__dirname, "google-credentials.json");
// const SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly";

// const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
// const SHEET_RANGE = process.env.SHEET_RANGE;

// function isHttpStatus(err, status) {
//   const s = err?.code;
//   if (s === status || s === String(status)) return true;
//   return err?.response?.status === status;
// }

// function isEnoent(err) {
//   for (const e of walkErrorChain(err)) {
//     if (e?.code === "ENOENT") return true;
//   }
//   return false;
// }

// function* walkErrorChain(err) {
//   const seen = new Set();
//   let current = err;
//   while (current && !seen.has(current)) {
//     seen.add(current);
//     yield current;
//     current = current.cause;
//   }
// }

// function assertConfig() {
//   if (!SPREADSHEET_ID) {
//     console.error("SPREADSHEET_ID is missing in .env");
//     process.exitCode = 1;
//     return false;
//   }
//   if (!SHEET_RANGE) {
//     console.error("SHEET_RANGE is missing in .env");
//     process.exitCode = 1;
//     return false;
//   }
//   return true;
// }

// async function getSheetsClient() {
//   const auth = new google.auth.GoogleAuth({
//     keyFile: CREDENTIALS_PATH,
//     scopes: [SCOPE],
//   });
//   const authClient = await auth.getClient();
//   return google.sheets({ version: "v4", auth: authClient });
// }

// async function fetchValues(sheets) {
//   const res = await sheets.spreadsheets.values.get({
//     spreadsheetId: SPREADSHEET_ID,
//     range: SHEET_RANGE,
//   });
//   return res.data.values;
// }

// function printResults(values) {
//   const rows = Array.isArray(values) ? values : [];
//   console.log(`Row count: ${rows.length}`);
//   if (rows.length) {
//     console.table(rows.slice(0, 5));
//   }
// }

// function handleError(err) {
//   if (isEnoent(err)) {
//     console.error(
//       "File ./google-credentials.json is missing. Add your Service Account JSON key and try again."
//     );
//     return;
//   }
//   if (isHttpStatus(err, 403)) {
//     console.error(
//       "The Sheets API returned 403 (Forbidden). Grant the Service Account email 'Viewer' access to this Google Sheet, then retry."
//     );
//     return;
//   }
//   if (isHttpStatus(err, 404)) {
//     console.error(
//       "The Sheets API returned 404. Check that SPREADSHEET_ID in .env matches a real spreadsheet and that the ID is not mistyped."
//     );
//     return;
//   }
//   console.error(err);
// }

// async function main() {
//   if (!assertConfig()) return;
//   const sheets = await getSheetsClient();
//   const values = await fetchValues(sheets);
//   printResults(values);
// }

// main().catch((err) => {
//   handleError(err);
//   process.exitCode = 1;
// });
