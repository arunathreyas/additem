const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 10000;

app.use(cors());
app.use(express.json());

const FILE_PATH = "data.xlsx";

// Function to read data from Excel
const readExcel = () => {
    if (!fs.existsSync(FILE_PATH)) return []; // Return empty array if file doesn't exist
    const workbook = XLSX.readFile(FILE_PATH);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet);
};

// Function to write data to Excel
const writeExcel = (data) => {
    console.log("Writing to Excel...");
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
    XLSX.writeFile(workbook, FILE_PATH);
};

// API: Get all items from Excel
app.get("/items", (req, res) => {
    res.json(readExcel());
});

// API: Add new item to Excel
app.post("/items", (req, res) => {
    let items = readExcel();
    const newItem = { id: items.length + 1, name: req.body.name };
    items.push(newItem);
    writeExcel(items);
    res.json(newItem);
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
