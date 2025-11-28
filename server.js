// server.js
import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import { google } from "googleapis";

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json());

// تهيئة Google Auth
function getSheets() {
  const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n");
  const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const spreadsheetId = process.env.GOOGLE_SHEET_ID;

  if (!privateKey || !clientEmail || !spreadsheetId) {
    throw new Error("Missing Google Sheets environment variables.");
  }

  const auth = new google.auth.JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  const sheets = google.sheets({ version: "v4", auth });
  return { sheets, spreadsheetId };
}

// دالة تجيب اسم أول ورقة تلقائيًا
async function getFirstSheetName(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  return meta.data.sheets[0].properties.title;
}

// جلب كل الصفوف
async function getAllRows() {
  const { sheets, spreadsheetId } = getSheets();
  const sheetName = await getFirstSheetName(sheets, spreadsheetId);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A:H`,
  });
  return { rows: res.data.values || [], sheetName };
}

// تحويل صف إلى كائن مستخدم
function rowToUser(row) {
  return {
    personalNumber: row[0] || "",
    name: row[1] || "",
    email: row[2] || "",
    password: row[3] || "",
    phone: row[4] || "",
    balance: row[5] || "",
    LoginNumber: row[6] || "",
    VIP: row[7] || "",
  };
}

// جلب كل المستخدمين
app.get("/api/users", async (req, res) => {
  try {
    const { rows } = await getAllRows();
    const header = rows[0] || [];
    const dataRows = rows.slice(1);
    const users = dataRows.map(rowToUser);
    res.json({ header, count: users.length, users });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch users." });
  }
});

// البحث عن مستخدم بالرقم الشخصي
app.get("/api/users/:personalNumber", async (req, res) => {
  try {
    const target = String(req.params.personalNumber).trim();
    const { rows } = await getAllRows();
    const dataRows = rows.slice(1);
    const idx = dataRows.findIndex(r => String(r[0]).trim() === target);

    if (idx === -1) {
      return res.status(404).json({ error: "User not found." });
    }

    const user = rowToUser(dataRows[idx]);
    const sheetRowNumber = idx + 2; // الصف الفعلي في الشيت
    res.json({ user, sheetRowNumber });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Search failed." });
  }
});

// تحديث LoginNumber أو VIP
app.patch("/api/users/:personalNumber", async (req, res) => {
  try {
    const target = String(req.params.personalNumber).trim();
    const { action } = req.body; // "ban_permanent" | "ban_temporary" | "upgrade_vip"

    if (!["ban_permanent", "ban_temporary", "upgrade_vip"].includes(action)) {
      return res.status(400).json({ error: "Invalid action." });
    }

    const { sheets, spreadsheetId } = getSheets();
    const { rows, sheetName } = await getAllRows();
    const dataRows = rows.slice(1);
    const idx = dataRows.findIndex(r => String(r[0]).trim() === target);

    if (idx === -1) {
      return res.status(404).json({ error: "User not found." });
    }

    const sheetRowNumber = idx + 2;

    if (action === "ban_permanent") {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!G${sheetRowNumber}`,
        valueInputOption: "RAW",
        requestBody: { values: [["1"]] },
      });
    } else if (action === "ban_temporary") {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!G${sheetRowNumber}`,
        valueInputOption: "RAW",
        requestBody: { values: [["2"]] },
      });
    } else if (action === "upgrade_vip") {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!H${sheetRowNumber}`,
        valueInputOption: "RAW",
        requestBody: { values: [["vip"]] },
      });
    }

    res.json({ ok: true, sheetRowNumber, action });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Update failed." });
  }
});

// شحن الرصيد
app.post("/api/charge", async (req, res) => {
  try {
    const { personalNumber, amount } = req.body;
    const { sheets, spreadsheetId } = getSheets();
    const { rows, sheetName } = await getAllRows();
    const dataRows = rows.slice(1);
    const idx = dataRows.findIndex(r => String(r[0]).trim() === String(personalNumber).trim());

    if (idx === -1) {
      return res.status(404).json({ error: "User not found." });
    }

    const sheetRowNumber = idx + 2;
    const currentBalance = parseFloat(dataRows[idx][5] || "0");
    const newBalance = currentBalance + parseFloat(amount);

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!F${sheetRowNumber}`, // العمود F هو الرصيد
      valueInputOption: "RAW",
      requestBody: { values: [[String(newBalance)]] },
    });

    res.json({ ok: true, personalNumber, newBalance });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Charge failed." });
  }
});
// تعديل بيانات المستخدم (أي خانة)
app.put("/api/users/:personalNumber", async (req, res) => {
  try {
    const target = String(req.params.personalNumber).trim();
    const updates = req.body; // { name, email, phone, password, balance, ... }

    const { sheets, spreadsheetId } = getSheets();
    const { rows, sheetName } = await getAllRows();
    const dataRows = rows.slice(1);
    const idx = dataRows.findIndex(r => String(r[0]).trim() === target);

    if (idx === -1) {
      return res.status(404).json({ error: "User not found." });
    }

    const sheetRowNumber = idx + 2;

    // بناء القيم الجديدة حسب الأعمدة
    const newRow = [
      updates.personalNumber || dataRows[idx][0],
      updates.name || dataRows[idx][1],
      updates.email || dataRows[idx][2],
      updates.password || dataRows[idx][3],
      updates.phone || dataRows[idx][4],
      updates.balance || dataRows[idx][5],
      updates.LoginNumber || dataRows[idx][6],
      updates.VIP || dataRows[idx][7],
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A${sheetRowNumber}:H${sheetRowNumber}`,
      valueInputOption: "RAW",
      requestBody: { values: [newRow] },
    });

    res.json({ ok: true, updated: newRow });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Update failed." });
  }
});

// تحديث حالة المستخدم حسب الإجراء
app.patch("/api/users/:personalNumber", async (req, res) => {
  try {
    const target = String(req.params.personalNumber).trim();
    const action = req.body.action;

    const { sheets, spreadsheetId } = getSheets();
    const { rows, sheetName } = await getAllRows();
    const dataRows = rows.slice(1);
    const idx = dataRows.findIndex(r => String(r[0]).trim() === target);

    if (idx === -1) {
      return res.status(404).json({ error: "User not found." });
    }

    const sheetRowNumber = idx + 2;
    const row = dataRows[idx];

    if (action === "ban_permanent") {
      row[6] = "1"; // LoginNumber = 1
    } else if (action === "ban_temporary") {
      row[6] = "2"; // LoginNumber = 2
    } else if (action === "upgrade_vip") {
      row[7] = "vip"; // VIP
    } else if (action === "unban") {
      row[6] = "3"; // LoginNumber = 3
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A${sheetRowNumber}:H${sheetRowNumber}`,
      valueInputOption: "RAW",
      requestBody: { values: [row] },
    });

    res.json({ ok: true, updated: row });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Update failed." });
  }
});

// جلب الطلبات من العمود I
app.get('/api/orders-sheet', async (req,res)=>{
  try {
    if (!sheetsClient || !SPREADSHEET_ID) return res.json({ orders: [] });
    const resp = await sheetsClient.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Profiles!I2:I10000'
    });
    const rows = (resp.data.values || []).map(r => r[0]).filter(v => v && v.trim() !== '');
    res.json({ orders: rows });
  } catch(e){
    console.error('orders-sheet error', e);
    res.status(500).json({ orders: [] });
  }
});

// تحديث الطلب (إضافة الرد والحالة)
app.post('/api/orders-update', async (req,res)=>{
  try {
    const { orderText, reply, status } = req.body;
    if (!orderText) return res.status(400).json({ ok:false, error:'missing orderText' });

    // ابحث عن الصف الذي يحتوي الطلب
    const resp = await sheetsClient.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Profiles!I2:I10000'
    });
    const rows = resp.data.values || [];
    let foundRow = null;
    for (let i=0;i<rows.length;i++){
      if (rows[i][0] && rows[i][0].includes(orderText)) {
        foundRow = i+2; // row index
        break;
      }
    }
    if (!foundRow) return res.status(404).json({ ok:false, error:'order_not_found' });

    const newVal = orderText + `\nالرد: ${reply||''}\nالحالة: ${status}`;
    await sheetsClient.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Profiles!I${foundRow}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[ newVal ]] }
    });
    res.json({ ok:true });
  } catch(e){
    console.error('orders-update error', e);
    res.status(500).json({ ok:false, error:e.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
