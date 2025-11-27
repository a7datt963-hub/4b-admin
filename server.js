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
const SHEET_ID = process.env.SHEET_ID || "162XC1wqWJbEtvjxTFVyxIwA8WWqDEEefFFr_ceQztM";
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1"; // عدّل الاسم حسب شيتك
const GOOGLE_SA_KEY_JSON =
  process.env.GOOGLE_SA_KEY_JSON ||
  process.env.GOOGLESAKEYJSON ||
  "";

if (!GOOGLE_SA_KEY_JSON) {
  console.warn("تحذير: لم يتم ضبط GOOGLE_SA_KEY_JSON في المتغيرات البيئية.");
}

// التعامل مع اختلاف أسماء المفاتيح داخل JSON
function normalizeServiceAccount(jsonStr) {
  let obj;
  try {
    obj = JSON.parse(jsonStr);
  } catch (e) {
    throw new Error("JSON الخدمة غير صالح");
  }
  // تطبيع المفاتيح إلى النموذج القياسي
  const normalized = {
    type: obj.type === "service_account" || obj.type === "serviceaccount" ? "service_account" : "service_account",
    project_id: obj.project_id || obj.projectid,
    private_key_id: obj.private_key_id || obj.privatekeyid,
    private_key: obj.private_key || obj.privatekey,
    client_email: obj.client_email || obj.clientemail,
    client_id: obj.client_id || obj.clientid,
    auth_uri: obj.auth_uri || obj.authuri,
    token_uri: obj.token_uri || obj.tokenuri,
    auth_provider_x509_cert_url:
      obj.auth_provider_x509_cert_url || obj.authproviderx509certurl,
    client_x509_cert_url:
      obj.client_x509_cert_url || obj.clientx509certurl,
    universe_domain: obj.universe_domain || obj.universedomain || "googleapis.com",
  };
  return normalized;
}

// إنشاء عميل Google
function getSheetsClient() {
  const key = normalizeServiceAccount(GOOGLE_SA_KEY_JSON);
  const auth = new google.auth.JWT(
    key.client_email,
    undefined,
    key.private_key,
    ["https://www.googleapis.com/auth/spreadsheets"],
    undefined
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
    obj._rowNumber = idx + 2; // رقم الصف الفعلي (مع الهيدر)
    return obj;
  });
  return { headers, items };
}

// البحث عن صف حسب personalNumber
async function findByPersonalNumber(pn) {
  const { items } = await readAllRows();
  const found = items.find((it) => String(it.personalNumber).trim() === String(pn).trim());
  if (!found) return null;
  return found;
}

// تحديث خلية واحدة في صف محدد
async function updateCell(rowNumber, columnName, value) {
  const sheets = getSheetsClient();
  // جلب الهيدر لتحديد رقم العمود
  const { headers } = await readAllRows();
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) throw new Error(`العمود ${columnName} غير موجود`);
  const colLetter = toColumnLetter(colIndex + 1);
  const range = `${SHEET_NAME}!${colLetter}${rowNumber}:${colLetter}${rowNumber}`;
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
    // تحجيم البيانات للعرض الأساسي (الاسم + رقم الدخول)
    const users = items.map((it) => ({
      personalNumber: it.personalNumber || "",
      name: it.name || "",
      LoginNumber: it.LoginNumber || "",
    }));
    res.json({ users });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: جلب مستخدم واحد
app.get("/user/:personalNumber", async (req, res) => {
  try {
    const pn = req.params.personalNumber;
    const found = await findByPersonalNumber(pn);
    if (!found) return res.status(404).json({ error: "Not found" });
    res.json(found);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: حظر دائم/مؤقت عبر تعديل LoginNumber
app.post("/user/:personalNumber/block", async (req, res) => {
  try {
    const pn = req.params.personalNumber;
    const { value } = req.body;
    if (value === undefined) return res.status(400).json({ error: "value مطلوب" });

    const found = await findByPersonalNumber(pn);
    if (!found) return res.status(404).json({ error: "Not found" });

    await updateCell(found._rowNumber, "LoginNumber", String(value));
    const updated = await findByPersonalNumber(pn);
    res.json(updated);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// API: ترقية VIP
app.post("/user/:personalNumber/vip", async (req, res) => {
  try {
    const pn = req.params.personalNumber;
    const { value } = req.body;
    if (value === undefined) return res.status(400).json({ error: "value مطلوب" });

    const found = await findByPersonalNumber(pn);
    if (!found) return res.status(404).json({ error: "Not found" });

    await updateCell(found._rowNumber, "VIP", String(value));
    const updated = await findByPersonalNumber(pn);
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
