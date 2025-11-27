// server.js
headers.forEach((h, idx) => {
obj[h] = row[idx] || '';
});
return res.json({found: true, user: obj});
}
return res.json({found: false});
} catch (err) {
console.error(err);
res.status(500).json({error: 'Search failed', details: err.message});
}
});


// helper: update a single cell by column name and row number
async function updateCellByHeader(headerName, rowNumber, newValue) {
const sheets = await getSheetsClient();
// get headers
const r = await sheets.spreadsheets.values.get({
spreadsheetId: SHEET_ID,
range: 'Sheet1!1:1'
});
const hdrs = r.data.values[0];
const colIndex = hdrs.findIndex(h => (h || '').toString().trim() === headerName);
if (colIndex === -1) throw new Error('Header not found: ' + headerName);
const colLetter = colToLetter(colIndex);
const range = `Sheet1!${colLetter}${rowNumber}`;
await sheets.spreadsheets.values.update({
spreadsheetId: SHEET_ID,
range,
valueInputOption: 'RAW',
requestBody: {values: [[newValue]]}
});
}


// set LoginNumber for a personalNumber row — helper endpoint
app.post('/action/set-login', async (req, res) => {
const {personalNumber, loginValue} = req.body;
if (!personalNumber) return res.status(400).json({error: 'personalNumber required'});
if (loginValue === undefined) return res.status(400).json({error: 'loginValue required'});
try {
// find row
const s = await getSheetsClient();
const r = await s.spreadsheets.values.get({spreadsheetId: SHEET_ID, range: 'Sheet1!A:Z'});
const rows = r.data.values || [];
if (rows.length < 1) return res.status(500).json({error: 'sheet empty'});
for (let i = 1; i < rows.length; i++) {
if ((rows[i][0] || '') == personalNumber) {
const rowNumber = i + 1;
await updateCellByHeader('LoginNumber', rowNumber, loginValue);
return res.json({ok: true, row: rowNumber});
}
}
return res.status(404).json({error: 'personalNumber not found'});
} catch (err) {
console.error(err);
res.status(500).json({error: 'update failed', details: err.message});
}
});


// set VIP column value to 'vip' (or any value)
app.post('/action/set-vip', async (req, res) => {
const {personalNumber, vipValue} = req.body;
if (!personalNumber) return res.status(400).json({error: 'personalNumber required'});
try {
const s = await getSheetsClient();
const r = await s.spreadsheets.values.get({spreadsheetId: SHEET_ID, range: 'Sheet1!A:Z'});
const rows = r.data.values || [];
if (rows.length < 1) return res.status(500).json({error: 'sheet empty'});
for (let i = 1; i < rows.length; i++) {
if ((rows[i][0] || '') == personalNumber) {
const rowNumber = i + 1;
await updateCellByHeader('VIP', rowNumber, vipValue || 'vip');
return res.json({ok: true, row: rowNumber});
}
}
return res.status(404).json({error: 'personalNumber not found'});
} catch (err) {
console.error(err);
res.status(500).json({error: 'update failed', details: err.message});
}
});


const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server listening on', PORT));
