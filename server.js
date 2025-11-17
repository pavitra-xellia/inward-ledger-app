// ------------------------------------------------
// Import required modules (ESM syntax)
// ------------------------------------------------
import express from "express";
import bodyParser from "body-parser";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import cors from "cors";

// ------------------------------------------------
// App setup
// ------------------------------------------------
const app = express();
const PORT = process.env.PORT || 5001;
const __dirname = path.resolve();
const EXCEL_FILE = path.join(__dirname, "Inward_Ledger.xlsx");

// ------------------------------------------------
// Middleware
// ------------------------------------------------
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, "public")));

// ------------------------------------------------
// Ensure Excel file exists
// ------------------------------------------------
function ensureExcelFile() {
  if (!fs.existsSync(EXCEL_FILE)) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      [
        "Sl No",
        "Inward Date",
        "Invoice / Delivery Date",
        "To Whom",
        "Department",
        "Item Description",
        "Make",
        "Other Description",
        "IMEI",
        "Comments"
      ]
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Inward Ledger");
    XLSX.writeFile(workbook, EXCEL_FILE);
  }
}

// ------------------------------------------------
// Save form data
// ------------------------------------------------
app.post("/submit", (req, res) => {
  try {
    ensureExcelFile();

    const data = req.body;
    const workbook = XLSX.readFile(EXCEL_FILE);
    const sheetName = "Inward Ledger";
    const sheet = workbook.Sheets[sheetName];

    const existingData = XLSX.utils.sheet_to_json(sheet);

    const newEntry = {
      "Sl No": data.slNo || existingData.length + 1,
      "Inward Date": data.inwardDate || "",
      "Invoice / Delivery Date": data.invoiceDate || "",
      "To Whom": data.personName || "",
      "Department": data.department || "",
      "Item Description": Array.isArray(data.item) ? data.item.join(", ") : "",
      "Make": Array.isArray(data.make) ? data.make.join(", ") : "",
      "Other Description": data.otherDesc || "",
      "IMEI": data.imei || "",
      "Comments": data.comments || ""
    };

    existingData.push(newEntry);

    const newSheet = XLSX.utils.json_to_sheet(existingData);
    workbook.Sheets[sheetName] = newSheet;

    XLSX.writeFile(workbook, EXCEL_FILE);

    res.json({ message: "Data saved successfully!" });
  } catch (error) {
    res.status(500).json({ message: "Server Error: " + error.message });
  }
});

// ------------------------------------------------
// Download Excel
// ------------------------------------------------
app.get("/download", (req, res) => {
  ensureExcelFile();
  res.download(EXCEL_FILE, "Inward_Ledger.xlsx");
});

// ------------------------------------------------
// Start server
// ------------------------------------------------
app.listen(PORT, "0.0.0.0", () => {
  console.log(`Server running on port ${PORT}`);
});