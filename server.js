// server.js
require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(cors({ origin: true }));
app.use(bodyParser.json());

// =========================
//  متغيرات البيئة
// =========================
const SHEET_ID = process.env.SHEET_ID || '';
let GOOGLE_SA_KEY_JSON = process.env.GOOGLE_SA_KEY_JSON || '';

if (!SHEET_ID) {
  console.warn('Warning: SHEET_ID is not set in environment variables.');
}
if (!GOOGLE_SA_KEY_JSON) {
  console.warn('Warning: GOOGLE_SA_KEY_JSON is not set in environment variables.');
}

// =========================
// إصلاح GOOGLE_SA_KEY_JSON
// Render يضيف \\n ويكسر التنسيق
// =========================
try {
  if (typeof GOOGLE_SA_KEY_JSON === 'string') {
    // تحويل \n النصية إلى أسطر حقيقية
    GOOGLE_SA_KEY_JSON = GOOGLE_SA_KEY_JSON.replace(/\\n/g, '\n');

    // ثم نعمل Parse
    GOOGLE_SA_KEY_JSON = JSON.parse(GOOGLE_SA_KEY_JSON);
  }
} catch (err) {
  console.error('Failed to parse GOOGLE_SA_KEY_JSON:', err);
}

// =========================
// Google Auth
// =========================
const auth = new google.auth.GoogleAuth({
  credentials: GOOGLE_SA_KEY_JSON,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

async function getSheetsClient() {
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

// =========================
// دالة لتحويل رقم عمود إلى حرف
// =========================
function colToLetter(col) {
  let s = '';
  while (col >= 0) {
    s = String.fromCharCode((col % 26) + 65) + s;
    col = Math.floor(col / 26) - 1;
  }
  return s;
}

// =========================
// API: جلب كل المستخدمين
// =========================
app.get('/users', async (req, res) => {
  try {
    if (!SHEET_ID) return res.status(500).json({ error: 'SHEET_ID not configured' });

    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!A:Z',
    });

    const rows = r.data.values || [];
    if (rows.length === 0) return res.json({ users: [], headers: [], counts: {} });

    const headers = rows[0].map(h => (h || '').toString().trim());
    const users = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const obj = { __row: i + 1 };
      headers.forEach((h, idx) => (obj[h] = row[idx] || ''));
      users.push(obj);
    }

    const counts = {};
    for (const u of users) {
      const key = (u.LoginNumber || '0').toString();
      counts[key] = (counts[key] || 0) + 1;
    }

    res.json({ users, headers, counts });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to read sheet', details: err.message });
  }
});

// =========================
// API: البحث عن مستخدم
// =========================
app.get('/search', async (req, res) => {
  const personal = (req.query.personalNumber || '').trim();
  if (!personal) return res.status(400).json({ error: 'personalNumber required' });

  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!A:Z',
    });

    const rows = r.data.values || [];
    if (rows.length < 1) return res.json({ found: false });

    const headers = rows[0];

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').toString() === personal) {
        const rowNumber = i + 1;
        const obj = { __row: rowNumber };
        headers.forEach((h, idx) => (obj[h] = rows[i][idx] || ''));
        return res.json({ found: true, user: obj });
      }
    }

    res.json({ found: false });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Search failed', details: err.message });
  }
});

// =========================
// تحديث خلية بناءً على اسم عمود
// =========================
async function updateCellByHeader(headerName, rowNumber, newValue) {
  const sheets = await getSheetsClient();
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'Sheet1!1:1',
  });

  const headers = r.data.values[0] || [];
  const colIndex = headers.findIndex(h => (h || '').trim() === headerName);
  if (colIndex === -1) throw new Error('Header not found: ' + headerName);

  const colLetter = colToLetter(colIndex);
  const range = `Sheet1!${colLetter}${rowNumber}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'RAW',
    requestBody: { values: [[newValue]] },
  });
}

// =========================
// set-login
// =========================
app.post('/action/set-login', async (req, res) => {
  const { personalNumber, loginValue } = req.body;
  if (!personalNumber) return res.status(400).json({ error: 'personalNumber required' });

  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!A:Z',
    });

    const rows = r.data.values || [];

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '') == personalNumber) {
        const rowNumber = i + 1;
        await updateCellByHeader('LoginNumber', rowNumber, loginValue);
        return res.json({ ok: true, row: rowNumber });
      }
    }

    res.status(404).json({ error: 'personalNumber not found' });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'update failed', details: err.message });
  }
});

// =========================
// set-vip
// =========================
app.post('/action/set-vip', async (req, res) => {
  const { personalNumber, vipValue } = req.body;
  if (!personalNumber) return res.status(400).json({ error: 'personalNumber required' });

  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!A:Z',
    });

    const rows = r.data.values || [];

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '') == personalNumber) {
        const rowNumber = i + 1;
        await updateCellByHeader('VIP', rowNumber, vipValue || 'vip');
        return res.json({ ok: true, row: rowNumber });
      }
    }

    res.status(404).json({ error: 'personalNumber not found' });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'update failed', details: err.message });
  }
});

// =========================
// تشغيل السيرفر
// =========================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
