// server.js
require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(cors({ origin: true }));
app.use(bodyParser.json());

// =====================
//   قراءة متغيرات البيئة
// =====================
const SHEET_ID = process.env.SHEET_ID;
let GOOGLE_SA_KEY_JSON = process.env.GOOGLE_SA_KEY_JSON;

if (!SHEET_ID) console.warn("⚠ SHEET_ID is missing");
if (!GOOGLE_SA_KEY_JSON) console.warn("⚠ GOOGLE_SA_KEY_JSON is missing");

// =====================
//   إصلاح JSON المعطوب
// =====================
function fixServiceAccount(jsonString) {
  try {
    // نزيل علامات البداية والنهاية المكسورة
    let clean = jsonString.trim();

    // Render يخزن \n و \\n → نرجعهم لسطر جديد
    clean = clean.replace(/\\\\n/g, "\n");
    clean = clean.replace(/\\n/g, "\n");

    // الآن نحوله JSON فعلي
    return JSON.parse(clean);
  } catch (e) {
    console.error("❌ Error parsing GOOGLE_SA_KEY_JSON");
    console.error(e);
    return null;
  }
}

const GOOGLE_CREDENTIALS = fixServiceAccount(GOOGLE_SA_KEY_JSON);

if (!GOOGLE_CREDENTIALS || !GOOGLE_CREDENTIALS.client_email) {
  console.error("❌ Invalid GOOGLE_SA_KEY_JSON – missing client_email");
}

// =====================
//   Google Auth
// =====================
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];

const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: GOOGLE_CREDENTIALS.client_email,
    private_key: GOOGLE_CREDENTIALS.private_key,
  },
  scopes: SCOPES,
});

async function getSheetsClient() {
  const client = await auth.getClient();
  return google.sheets({ version: "v4", auth: client });
}

// تحويل رقم عمود → حرف
function colToLetter(col) {
  let s = "";
  while (col >= 0) {
    s = String.fromCharCode((col % 26) + 65) + s;
    col = Math.floor(col / 26) - 1;
  }
  return s;
}

// =====================
//       GET USERS
// =====================
app.get("/users", async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Sheet1!A:Z",
    });

    const rows = r.data.values || [];
    if (rows.length === 0) return res.json({ users: [] });

    const headers = rows[0];
    const users = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const obj = { __row: i + 1 };
      headers.forEach((h, idx) => (obj[h] = row[idx] || ""));
      users.push(obj);
    }

    res.json({ users, headers });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Failed to load users", details: e.message });
  }
});

// =====================
//       SEARCH USER
// =====================
app.get("/search", async (req, res) => {
  const personal = (req.query.personalNumber || "").trim();
  if (!personal) return res.status(400).json({ error: "personalNumber required" });

  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Sheet1!A:Z",
    });

    const rows = r.data.values || [];
    const headers = rows[0];

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || "") === personal) {
        const obj = { __row: i + 1 };
        headers.forEach((h, idx) => (obj[h] = rows[i][idx] || ""));
        return res.json({ found: true, user: obj });
      }
    }

    res.json({ found: false });
  } catch (e) {
    res.status(500).json({ error: "Search failed", details: e.message });
  }
});

// =====================
//   UPDATE CELL HELPERS
// =====================
async function updateCellByHeader(header, row, value) {
  const sheets = await getSheetsClient();
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "Sheet1!1:1",
  });

  const headers = r.data.values[0] || [];
  const colIndex = headers.findIndex((h) => h.trim() === header);
  const colLetter = colToLetter(colIndex);
  const range = `Sheet1!${colLetter}${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[value]] },
  });
}

// =====================
//   SET LOGIN NUMBER
// =====================
app.post("/action/set-login", async (req, res) => {
  const { personalNumber, loginValue } = req.body;
  if (!personalNumber || loginValue === undefined)
    return res.status(400).json({ error: "Missing fields" });

  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Sheet1!A:Z",
    });

    const rows = r.data.values;

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] == personalNumber) {
        const row = i + 1;
        await updateCellByHeader("LoginNumber", row, loginValue);
        return res.json({ ok: true, row });
      }
    }

    res.status(404).json({ error: "personalNumber not found" });
  } catch (e) {
    res.status(500).json({ error: "Update failed", details: e.message });
  }
});

// =====================
//   SET VIP STATUS
// =====================
app.post("/action/set-vip", async (req, res) => {
  const { personalNumber, vipValue } = req.body;
  if (!personalNumber)
    return res.status(400).json({ error: "Missing personalNumber" });

  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Sheet1!A:Z",
    });

    const rows = r.data.values;

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] == personalNumber) {
        const row = i + 1;
        await updateCellByHeader("VIP", row, vipValue || "vip");
        return res.json({ ok: true, row });
      }
    }

    res.status(404).json({ error: "personalNumber not found" });
  } catch (e) {
    res.status(500).json({ error: "Update failed", details: e.message });
  }
});

// =====================
//   START SERVER
// =====================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
