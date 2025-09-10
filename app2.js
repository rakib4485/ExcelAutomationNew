// const express = require("express");
// const xlsx = require("xlsx");
// const dotenv = require("dotenv");
// const { exec } = require("child_process");
// const path = require("path");

// dotenv.config();

// const app = express();
// const PORT = process.env.PORT || 3000;

// // ✅ Excel reading function
// function readExcelFile() {
//   const filePath = path.resolve(__dirname, process.env.XLSX_PATH);
//   const workbook = xlsx.readFile(filePath);

//   const sheetName = "Working sheet";
//   const sheet = workbook.Sheets[sheetName];

//   if (!sheet) {
//     throw new Error(`❌ Sheet "${sheetName}" not found in ${process.env.XLSX_PATH}`);
//   }

//   const data = xlsx.utils.sheet_to_json(sheet);
//   return data;
// }

// // ✅ Macro runner function (optional)
// function runMacro() {
//   const psScriptPath = path.resolve(__dirname, "run-macro.ps1");
//   exec(`powershell.exe -ExecutionPolicy Bypass -File "${psScriptPath}"`, (error, stdout, stderr) => {
//     if (error) {
//       console.error(`❌ Macro execution failed:\n${stderr}`);
//     } else {
//       console.log("✅ Macro executed successfully.");
//       console.log(stdout);
//     }
//   });
// }

// // ✅ API Endpoint: GET /api/data
// app.get("/api/data", (req, res) => {
//   try {
//     const data = readExcelFile();
//     res.json({
//       message: "✅ Data fetched successfully",
//       rows: data, // Limit response (optional)
//     });
//   } catch (err) {
//     res.status(500).json({ error: err.message });
//   }
// });

// // ✅ Optional: Trigger macro via endpoint
// app.get("/api/run-macro", (req, res) => {
//   runMacro();
//   res.send("🌀 Macro triggered.");
// });

// // ✅ Start server
// app.listen(PORT, () => {
//   console.log(`🚀 Server running at http://localhost:${PORT}`);
// });
const express = require("express");
const { exec } = require("child_process");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 5000;

// Path to your macro-enabled Excel file
const xlsmFilePath = path.resolve(__dirname, "250616 Daily Volume Report - Copy.xlsm");

// Path to your PowerShell script
const psScriptPath = path.resolve(__dirname, "run-macro.ps1");

app.get("/api/refresh", (req, res) => {
  console.log("Refrsh api called, please wait");
  const command = `powershell.exe -ExecutionPolicy Bypass -File "${psScriptPath}" -xlsmPath "${xlsmFilePath}"`;

  exec(command, (error, stdout, stderr) => {
    if (error) {
      console.error("Macro execution failed:", stderr);
      return res.status(500).json({
        success: false,
        message: "Failed to run macro",
        error: stderr.trim(),
      });
    }

    console.log("Macro output:", stdout);
    res.json({
      success: true,
      message: "Macro ran successfully",
      details: stdout.trim(),
    });
  });
});

// ✅ Start Api Page
app.get("/", (req, res) => {
  res.send("API Server is Running..........");
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
