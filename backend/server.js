const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const cors = require("cors");
const fs = require("fs");

const app = express();
app.use(cors());

const upload = multer({ dest: "uploads/" });

/* -------------------------------------------------
   Helpers
------------------------------------------------- */

function normalizeValue(value) {
  if (value === null || value === undefined || value === "") return "-";

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed !== "" && !isNaN(trimmed)) return Number(trimmed);
    return trimmed;
  }

  return value;
}

// 🔥 Excel date serial → MMM-YY (IMPORTANT FIX)
function excelSerialToMonthYear(val) {
  if (typeof val === "number") {
    const date = new Date((val - 25569) * 86400 * 1000);
    return date.toLocaleString("en-US", {
      month: "short",
      year: "2-digit"
    });
  }
  return val;
}

function cleanRow(row) {
  const clean = {};
  Object.entries(row).forEach(([key, value]) => {
    if (!key.startsWith("__EMPTY")) {
      clean[key.trim()] = normalizeValue(value);
    }
  });
  return clean;
}

/* -------------------------------------------------
   Upload API
------------------------------------------------- */

app.post("/upload-excel", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetsData = {};

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const lower = sheetName.toLowerCase();

      /* =============================================
         🔥 SUMMARY SHEET (PIVOT – MONTH FIXED)
      ============================================= */
      if (lower === "summary") {
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "-"
        });

        // rows[0] → "Sum of Hours"
        // rows[1] → Column headers
        // rows[2+] → Data
        const headers = rows[1]; // keep Grand Total also

        const tableData = [];

        for (let i = 2; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;

          const month = excelSerialToMonthYear(row[0]);

          const obj = { Month: month };

          for (let c = 1; c < headers.length; c++) {
            obj[headers[c]] = normalizeValue(row[c]);
          }

          tableData.push(obj);
        }

        sheetsData[sheetName] = {
          type: "pivot",
          data: tableData
        };

        return;
      }

      /* =============================================
         MASTER + OTHER SHEETS
      ============================================= */
      let range = 1;
      if (lower === "master") range = 0;

      const rawData = XLSX.utils.sheet_to_json(sheet, {
        range,
        defval: "-",
        raw: false
      });

      if (!rawData.length) {
        sheetsData[sheetName] = { data: [], kpis: {} };
        return;
      }

      const cleanedData = rawData.map(cleanRow);

      const numericColumns = Object.keys(cleanedData[0]).filter(col =>
        cleanedData.some(row => typeof row[col] === "number")
      );

      sheetsData[sheetName] = {
        type: "table",
        data: cleanedData,
        kpis: {
          totalRecords: cleanedData.length,
          numericColumns
        }
      };
    });

    fs.unlinkSync(req.file.path); // cleanup temp file

    res.json({ sheets: sheetsData });

  } catch (err) {
    console.error("❌ Excel processing failed:", err);
    res.status(500).json({ error: "Excel processing failed" });
  }
});

/* ------------------------------------------------- */

app.listen(4000, () => {
  console.log("🚀 Backend running at http://localhost:4000");
});
