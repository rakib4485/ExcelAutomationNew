const express = require("express");
const { exec } = require("child_process");
const path = require("path");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

// 📁 Paths
const xlsmFilePath = path.resolve(__dirname, "250616 Daily Volume Report - Copy2.xlsm");
const macroName = "refresh_All";
const psScriptPath = path.resolve(__dirname, "run-macro.ps1");

app.get("/api/refresh", (req, res) => {
  console.log("📢 /api/refresh called. Running macro...");

  const command = `powershell.exe -ExecutionPolicy Bypass -File "${psScriptPath}" -xlsmPath "${xlsmFilePath}" -macroName "${macroName}"`;

  const child = exec(command);

  child.stdout.on("data", (data) => {
    console.log("📄 PowerShell STDOUT:", data);
  });

  child.stderr.on("data", (data) => {
    console.error("❗ PowerShell STDERR:", data);
  });

  child.on("close", (code) => {
    if (code === 0) {
        console.log("✅ Macro finished.");
        res.json({
            status: "Finished",
            message: `Macro '${macroName}' started. This may take a few hours.`,
        });
    } else {
      console.error(`❌ Macro failed with code ${code}`);
    }
  });

  res.json({
    status: "running",
    message: `Macro '${macroName}' started. This may take a few hours.`,
  });
});

app.get("/", (req, res) => {
  res.send("🚀 Macro API Server is up and running.");
});

app.listen(PORT, () => {
  console.log(`✅ Server is listening at http://localhost:${PORT}`);
});

