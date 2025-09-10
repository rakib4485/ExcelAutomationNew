const express = require("express");
const dotenv = require("dotenv");
const path = require("path");
const { exec } = require("child_process");

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

const XLSM_FILE_PATH = path.resolve(__dirname, process.env.XLSM_PATH);

// API to run macro and download updated Excel
app.get("/api/download-updated-excel", (req, res) => {
  const psScriptPath = path.resolve(__dirname, "run-macro.ps1");

  // Run the PowerShell macro
  exec(`powershell.exe -ExecutionPolicy Bypass -File "${psScriptPath}"`, (error, stdout, stderr) => {
    if (error) {
      console.error(`❌ Macro execution failed:\n${stderr}`);
      return res.status(500).send("Macro failed.");
    } else {
      console.log("✅ Macro executed successfully.");
      console.log(stdout);

      // Wait 2 seconds to ensure file save
      setTimeout(() => {
        res.download(XLSM_FILE_PATH, (err) => {
          if (err) {
            console.error("❌ Error sending file:", err.message);
            res.status(500).send("Failed to send updated file.");
          }
        });
      }, 2000);
    }
  });
});

app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
