// server.js
require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(cors()); // غيّر origin إذا بدك تقيّد
app.use(bodyParser.json());

// ------------ config from env ----------
const SHEET_ID = process.env.SHEET_ID || '';
const GOOGLE_CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL || '';
let GOOGLE_PRIVATE_KEY = process.env.GOOGLE_PRIVATE_KEY || '';

// normalize private key: if it's escaped like "\\n" -> convert to real newlines
GOOGLE_PRIVATE_KEY = GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n').trim();

if (!SHEET_ID) console.warn('⚠ SHEET_ID not provided in env');
if (!GOOGLE_CLIENT_EMAIL) console.warn('⚠ GOOGLE_CLIENT_EMAIL not provided in env');
if (!GOOGLE_PRIVATE_KEY) console.warn('⚠ GOOGLE_PRIVATE_KEY not provided in env');

// helper: get authenticated sheets client (uses JWT)
async function getSheetsClient() {
  if (!GOOGLE_CLIENT_EMAIL || !GOOGLE_PRIVATE_KEY) {
    throw new Error('Google service account credentials are not configured (GOOGLE_CLIENT_EMAIL / GOOGLE_PRIVATE_KEY).');
  }

  const jwtClient = new google.auth.JWT(
    GOOGLE_CLIENT_EMAIL,
    null,
    GOOGLE_PRIVATE_KEY,
    ['https://www.googleapis.com/auth/spreadsheets'],
    null
  );

  // authorize (will throw if credentials invalid)
  await jwtClient.authorize();
  return google.sheets({ version: 'v4', auth: jwtClient });
}

// helper: 0-based col index -> letter
function colToLetter(col) {
  let s = '';
  while (col >= 0) {
    s = String.fromCharCode((col % 26) + 65) + s;
    col = Math.floor(col / 26) - 1;
  }
  return s;
}

// ---------- endpoints ----------

// GET /users -> returns all rows as objects (headers from row 1)
app.get('/users', async (req, res) => {
  try {
    if (!SHEET_ID) return res.status(500).json({ error: 'SHEET_ID not configured' });

    const sheets = await getSheetsClient();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'profiles!A:Z',
    });

    const rows = result.data.values || [];
    if (!rows.length) return res.json({ users: [], headers: [], counts: {} });

    const headers = rows[0].map(h => (h || '').toString().trim());
    const users = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const obj = { __row: i + 1 };
      headers.forEach((h, idx) => {
        obj[h] = row[idx] !== undefined ? row[idx] : '';
      });
      users.push(obj);
    }

    const counts = {};
    users.forEach(u => {
      const key = (u.LoginNumber || '0').toString();
      counts[key] = (counts[key] || 0) + 1;
    });

    res.json({ users, headers, counts });
  } catch (err) {
    console.error('GET /users error:', err);
    res.status(500).json({ error: 'Failed to read sheet', details: err.message });
  }
});

// GET /search?personalNumber=xxx
app.get('/search', async (req, res) => {
  const personal = (req.query.personalNumber || '').toString().trim();
  if (!personal) return res.status(400).json({ error: 'personalNumber required' });

  try {
    if (!SHEET_ID) return res.status(500).json({ error: 'SHEET_ID not configured' });

    const sheets = await getSheetsClient();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'profiles!A:Z',
    });

    const rows = result.data.values || [];
    if (rows.length < 1) return res.json({ found: false });

    const headers = rows[0];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if ((row[0] || '').toString() === personal) {
        const obj = { __row: i + 1 };
        headers.forEach((h, idx) => (obj[h] = row[idx] || ''));
        return res.json({ found: true, user: obj });
      }
    }
    res.json({ found: false });
  } catch (err) {
    console.error('GET /search error:', err);
    res.status(500).json({ error: 'Search failed', details: err.message });
  }
});

// helper update single cell by header name and row number
async function updateCellByHeader(headerName, rowNumber, newValue) {
  if (!SHEET_ID) throw new Error('SHEET_ID not configured');
  const sheets = await getSheetsClient();
  const hdrsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'profiles!1:1',
  });
  const hdrs = (hdrsRes.data.values && hdrsRes.data.values[0]) || [];
  const colIndex = hdrs.findIndex(h => (h || '').toString().trim() === headerName);
  if (colIndex === -1) throw new Error('Header not found: ' + headerName);

  const colLetter = colToLetter(colIndex);
  const range = `profiles!${colLetter}${rowNumber}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'RAW',
    requestBody: { values: [[newValue]] },
  });
}

// POST /action/set-login { personalNumber, loginValue }
app.post('/action/set-login', async (req, res) => {
  const { personalNumber, loginValue } = req.body || {};
  if (!personalNumber) return res.status(400).json({ error: 'personalNumber required' });
  if (loginValue === undefined) return res.status(400).json({ error: 'loginValue required' });

  try {
    if (!SHEET_ID) return res.status(500).json({ error: 'SHEET_ID not configured' });

    const sheets = await getSheetsClient();
    const result = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'profiles!A:Z' });
    const rows = result.data.values || [];
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '') == personalNumber) {
        const rowNumber = i + 1;
        await updateCellByHeader('LoginNumber', rowNumber, loginValue);
        return res.json({ ok: true, row: rowNumber });
      }
    }
    res.status(404).json({ error: 'personalNumber not found' });
  } catch (err) {
    console.error('POST /action/set-login error:', err);
    res.status(500).json({ error: 'update failed', details: err.message });
  }
});

// POST /action/set-vip { personalNumber, vipValue }
app.post('/action/set-vip', async (req, res) => {
  const { personalNumber, vipValue } = req.body || {};
  if (!personalNumber) return res.status(400).json({ error: 'personalNumber required' });

  try {
    if (!SHEET_ID) return res.status(500).json({ error: 'SHEET_ID not configured' });

    const sheets = await getSheetsClient();
    const result = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'profiles!A:Z' });
    const rows = result.data.values || [];
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '') == personalNumber) {
        const rowNumber = i + 1;
        await updateCellByHeader('VIP', rowNumber, vipValue || 'vip');
        return res.json({ ok: true, row: rowNumber });
      }
    }
    res.status(404).json({ error: 'personalNumber not found' });
  } catch (err) {
    console.error('POST /action/set-vip error:', err);
    res.status(500).json({ error: 'update failed', details: err.message });
  }
});

// start
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
