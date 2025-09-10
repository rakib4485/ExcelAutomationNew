const xlsx = require("xlsx");
const dotenv = require("dotenv");
const { exec } = require("child_process");
const path = require("path");

dotenv.config();

// Read Excel file (xlsx)
function readExcelFile() {
  const filePath = path.resolve(__dirname, process.env.XLSX_PATH);
  const workbook = xlsx.readFile(filePath);

  const sheetName = "Working sheet";
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    console.error(`❌ Sheet "${sheetName}" not found in ${process.env.XLSX_PATH}`);
    return;
  }

  const data = xlsx.utils.sheet_to_json(sheet);
  console.log(`✅ Data from "${sheetName}":`);
  console.log(data.slice(0, 5));
  res.send(data.slice(0, 5)); // Show first 5 rows
}

// Trigger Macro
function runMacro() {
  exec("powershell.exe -ExecutionPolicy Bypass -File run-macro.ps1", (error, stdout, stderr) => {
    if (error) {
      console.error(`❌ Macro execution failed:\n${stderr}`);
    } else {
      console.log("✅ Macro executed successfully.");
    }
  });
}

// Main
function main() {
  readExcelFile();
  runMacro();
}

main();
