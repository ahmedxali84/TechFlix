const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Middleware to parse JSON
app.use(express.json());
app.use(express.static('public'));

// Endpoint to handle order submission
app.post('/submit-order', (req, res) => {
  const { name, order, details, promo } = req.body;

  // Load or create the Excel file
  const filePath = path.join(__dirname, 'orders.xlsx');
  let workbook;
  let sheet;

  if (fs.existsSync(filePath)) {
    workbook = xlsx.readFile(filePath);
    sheet = workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = xlsx.utils.book_new();
    sheet = xlsx.utils.json_to_sheet([], {
      header: ['Name', 'Order', 'Details', 'Promo Code']
    });
    xlsx.utils.book_append_sheet(workbook, sheet, 'Orders');
  }

  // Append new order
  const newRow = [name, order, details, promo];
  xlsx.utils.sheet_add_aoa(sheet, [newRow], { origin: -1 });

  // Save the file
  xlsx.writeFile(workbook, filePath);

  res.status(200).send('Order saved successfully!');
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});