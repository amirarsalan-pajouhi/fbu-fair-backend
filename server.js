const express = require("express");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const app = express();
const port = 3000;
const excelFilePath = path.join(__dirname, "registrations.xlsx");

// Middleware
app.use(cors());
app.use(express.json());

// Initialize Excel file if it doesn't exist
async function initializeExcelFile() {
  if (!fs.existsSync(excelFilePath)) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Registrations");
    worksheet.columns = [
      { header: "Name", key: "name", width: 20 },
      { header: "Surname", key: "surname", width: 20 },
      { header: "Phone Number", key: "phone", width: 15 },
      { header: "Email", key: "email", width: 30 },
      { header: "Program of Interest", key: "program", width: 25 },
      { header: "Submission Date", key: "submissionDate", width: 20 },
    ];
    await workbook.xlsx.writeFile(excelFilePath);
  }
}

// Append data to Excel file
async function appendToExcel(data) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelFilePath);
  const worksheet = workbook.getWorksheet("Registrations");
  worksheet.addRow({
    name: data.name,
    surname: data.surname,
    phone: data.phone,
    email: data.email,
    program: data.program,
    submissionDate: new Date().toISOString(),
  });
  await workbook.xlsx.writeFile(excelFilePath);
}

// Handle form submission
app.post("/submit", async (req, res) => {
  const { name, surname, phone, email, program } = req.body;

  // Basic validation
  if (!name || !surname || !phone || !email || !program) {
    return res.status(400).json({ error: "All fields are required" });
  }

  try {
    await initializeExcelFile();
    await appendToExcel(req.body);
    res.status(200).json({ message: "Data saved successfully" });
  } catch (error) {
    console.error("Error saving data:", error.message);
    res.status(500).json({ error: "Failed to save data" });
  }
});

// Start server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
