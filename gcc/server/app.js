const express = require('express');
const app = express();
const ExcelJS = require('exceljs');

// Middleware to parse JSON data
app.use(express.json());

// File path for the Excel sheet
const EXCEL_FILE = 'data.xlsx';

// Function to initialize the Excel sheet with headers if it doesn't exist
function initializeExcelFile() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Books');
  sheet.columns = [
    { header: 'Title', key: 'title', width: 30 },
    { header: 'Author', key: 'author', width: 30 },
    { header: 'PublicationYear', key: 'publicationYear', width: 15 },
  ];
  return workbook.xlsx.writeFile(EXCEL_FILE);
}

// Initialize the Excel file if it doesn't exist
initializeExcelFile().catch((error) => {
  console.error('Error initializing Excel file:', error.message);
});

// CRUD endpoints

// Endpoint to add a new book
app.post('/api/books', (req, res) => {
  const { title, author, publicationYear } = req.body;
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx
    .readFile(EXCEL_FILE)
    .then(() => {
      const sheet = workbook.getWorksheet('Books');
      sheet.addRow({ title, author, publicationYear });
      return workbook.xlsx.writeFile(EXCEL_FILE);
    })
    .then(() => {
      res.sendStatus(201);
    })
    .catch((error) => {
      console.error('Error adding book:', error.message);
      res.sendStatus(500);
    });
});

// Endpoint to retrieve all books
app.get('/api/books', (req, res) => {
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx
    .readFile(EXCEL_FILE)
    .then(() => {
      const sheet = workbook.getWorksheet('Books');
      const data = sheet.getSheetValues();
      res.json(data.slice(1)); // Skip the header row
    })
    .catch((error) => {
      console.error('Error retrieving books:', error.message);
      res.sendStatus(500);
    });
});

// Endpoint to update a book by ID
app.put('/api/books/:id', (req, res) => {
  const bookId = parseInt(req.params.id);
  const { title, author, publicationYear } = req.body;
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx
    .readFile(EXCEL_FILE)
    .then(() => {
      const sheet = workbook.getWorksheet('Books');
      const bookRow = sheet.getRow(bookId + 1); // Row number starts from 1
      bookRow.getCell('title').value = title;
      bookRow.getCell('author').value = author;
      bookRow.getCell('publicationYear').value = publicationYear;
      return workbook.xlsx.writeFile(EXCEL_FILE);
    })
    .then(() => {
      res.sendStatus(200);
    })
    .catch((error) => {
      console.error('Error updating book:', error.message);
      res.sendStatus(500);
    });
});

// Endpoint to delete a book by ID
app.delete('/api/books/:id', (req, res) => {
  const bookId = parseInt(req.params.id);
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx
    .readFile(EXCEL_FILE)
    .then(() => {
      const sheet = workbook.getWorksheet('Books');
      sheet.spliceRows(bookId + 1, 1); // Row number starts from 1
      return workbook.xlsx.writeFile(EXCEL_FILE);
    })
    .then(() => {
      res.sendStatus(200);
    })
    .catch((error) => {
      console.error('Error deleting book:', error.message);
      res.sendStatus(500);
    });
});

// Start the server
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
