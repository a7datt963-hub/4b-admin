/**
 * سيرفر REST للتعامل مع Google Sheets
 * يعتمد على حساب خدمة Google وورقة تحتوي على الأعمدة:
 * personalNumber, name, email, password, phone, balance, LoginNumber, VIP
 */

const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
app.use(express.json());
app.use(cors());

// إعدادات من البيئة
const SHEET_ID = process.env.SHEET_ID || "162XC1_wqWJbEtvjxTFVyxIwA8WWqDEEefFFr_ceQztM";
const SHEET_NAME = process.env.SHEET_NAME || "profiles";
const GOOGLE_SA_KEY_JSON = process.env.GOOGLE_SA_KEY_JSON || "";

if (!GOOGLE_SA_KEY_JSON) {
  console.warn("تحذير: لم يتم ضبط GOOGLE_SA_KEY_JSON في المتغيرات البيئية.");
}

// تحويل النص مع محارف الهروب إلى JSON صالح
function parseServiceAccount(jsonStr) {
  try {
    // أولاً نحول النص إلى كائن
    const obj = JSON.parse(jsonStr);

    // المفتاح الخاص قد يحتوي على \n كنص، نحوله إلى أسطر حقيقية
    if (obj.private_key) {
      obj.private_key = obj.private_key.replace(/\\n/g, "\n");
    }

    return obj;
  } catch (e) {
    throw new Error("JSON الخدمة غير صالح: " + e.message);
  }
}

// إنشاء عميل Google Sheets
function getSheetsClient() {
  const key = parseServiceAccount(GOOGLE_SA_KEY_JSON);
  const auth = new google.auth.JWT(
    key.client_email,
    undefined,
    key.private_key,
    ["https://www.googleapis.com/auth/spreadsheets"]
  );
  return google.sheets({ version: "v4", auth });
}

// قراءة كل البيانات
async function readAllRows() {
  const sheets = getSheetsClient();
  const range = `${SHEET_NAME}!A:Z`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range,
  });
  const rows = res.data.values || [];
  if (rows.length === 0) return { headers: [], items: [] };

  const headers = rows[0];
  const items = rows.slice(1).map((row, idx) => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] !== undefined ? row[i] : "";
    });
    obj._rowNumber = idx + 2;
    return obj;
  });
  return { headers, items };
}

// البحث عن صف حسب personalNumber
async function findByPersonalNumber(pn) {
  const { items } = await readAllRows();
  return items.find((it) => String(it.personalNumber).trim() === String(pn).trim()) || null;
}

// تحديث خلية واحدة
async function updateCell(rowNumber, columnName, value) {
  const sheets = getSheetsClient();
  const { headers } = await readAllRows();
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) throw new Error(`العمود ${columnName} غير موجود`);
  const colLetter = toColumnLetter(colIndex + 1);
  const range = `${SHEET_NAME}!${colLetter}${rowNumber}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[value]] },
  });
}

// تحويل رقم عمود إلى حرف
function toColumnLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// API: جلب جميع المستخدمين
app.get("/users", async (req, res) => {
  try {
    const { items } = await readAllRows();
    res.json({ users: items });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: جلب مستخدم واحد
app.get("/user/:personalNumber", async (req, res) => {
  try {
    const found = await findByPersonalNumber(req.params.personalNumber);
    if (!found) return res.status(404).json({ error: "Not found" });
    res.json(found);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: تعديل LoginNumber
app.post("/user/:personalNumber/block", async (req, res) => {
  try {
    const found = await findByPersonalNumber(req.params.personalNumber);
    if (!found) return res.status(404).json({ error: "Not found" });
    await updateCell(found._rowNumber, "LoginNumber", String(req.body.value));
    const updated = await findByPersonalNumber(req.params.personalNumber);
    res.json(updated);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: تعديل VIP
app.post("/user/:personalNumber/vip", async (req, res) => {
  try {
    const found = await findByPersonalNumber(req.params.personalNumber);
    if (!found) return res.status(404).json({ error: "Not found" });
    await updateCell(found._rowNumber, "VIP", String(req.body.value));
    const updated = await findByPersonalNumber(req.params.personalNumber);
    res.json(updated);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
// صحّة السيرفر
app.get("/", (req, res) => {
  res.json({ ok: true, service: "fourb-admin" });
});

// التشغيل
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on :${PORT}`);
});
