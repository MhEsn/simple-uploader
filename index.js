const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.get('/', (req, res) => {
  res.send(`
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="xlsx" accept=".xlsx" required />
      <button type="submit">Upload</button>
    </form>
  `);
});

app.post('/upload', upload.single('xlsx'), (req, res) => {
  try {
    const filePath = path.join(__dirname, req.file.path);
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Create a basic HTML table
    let html = '<table border="1" cellpadding="5">';
    data.forEach(row => {
      html += '<tr>';
      row.forEach(cell => {
        html += `<td>${cell !== undefined ? cell : ''}</td>`;
      });
      html += '</tr>';
    });
    html += '</table>';

    res.send(html);
  } catch (err) {
    res.status(500).send('Error reading file: ' + err.message);
  }
});

app.listen(3000, () => {
  console.log('Server running on http://localhost:3000');
});