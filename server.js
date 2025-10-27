require('dotenv').config();
const express = require("express");
const ExcelJS = require("exceljs");
const { createClient } = require("@supabase/supabase-js");
const path = require("path");
const app = express();
app.use(express.json());
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);
const BUCKET = "visit-logs";

// --- Utility Functions ---
async function uploadFile(fileName, buffer) {
  const { error } = await supabase.storage.from(BUCKET).upload(fileName, buffer, { upsert: true });
  if (error) throw error;
}

async function downloadFile(fileName) {
  const { data, error } = await supabase.storage.from(BUCKET).download(fileName);
  if (error) throw error;
  return Buffer.from(await data.arrayBuffer());
}

// Clone base file (VisitLog.xlsx) preserving Sheet2 data & formatting including styles
async function cloneBaseFile() {
  try {
    const baseFile = "VisitLog.xlsx";
    const { data, error } = await supabase.storage.from(BUCKET).download(baseFile);
    if (error) throw error;
    const buffer = Buffer.from(await data.arrayBuffer());
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // Ensure Sheet2 exists
    const baseSheet = workbook.getWorksheet("Sheet2");
    if (!baseSheet) throw new Error("Sheet2 not found in base file.");

    // Create new workbook with same structure
    const newWorkbook = new ExcelJS.Workbook();
    const newSheet = newWorkbook.addWorksheet("Sheet2");

    baseSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const newRow = newSheet.getRow(rowNumber);
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        newCell.value = cell.value;

        // Deep copy styles (font, fill, border, alignment, etc.)
        if (cell.style) {
          newCell.style = JSON.parse(JSON.stringify(cell.style));
        }
      });
      newRow.commit();
    });

    return newWorkbook;
  } catch (err) {
    console.warn("Base VisitLog.xlsx not found — using default structure.");
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet2");
    sheet.addRow(["SNO", "EXECUTIVE", "VISIT TOOL UTILIZATION", "TOTAL", "TIME", "CD3", "CD5", "CD7", "YB", "MIS", "AFTERNOON"]);
    sheet.addRow(["TOTAL", "", "", 0, "", 0, 0, 0, 0, 0, 0]);
    return workbook;
  }
}

function getTodayFileNames() {
  const today = new Date().toISOString().slice(0, 10);
  return {
    fileName: `VisitLog_${today}.xlsx`,
    viewName: `VisitLog_ViewOnly_${today}.xlsx`
  };
}

// --- Initialize Today's File If Needed ---
async function initTodayFile() {
  const { fileName, viewName } = getTodayFileNames();
  try {
    await downloadFile(fileName); // Exists
  } catch {
    const workbook = await cloneBaseFile(); // Clone from VisitLog.xlsx
    const buffer = await workbook.xlsx.writeBuffer();
    await uploadFile(fileName, buffer);
    await uploadFile(viewName, buffer);
  }
}
initTodayFile();

// --- API Endpoints ---

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "VisitLogger.html"));
});

// Executive Names
app.get("/executives", async (req, res) => {
  try {
    const { fileName: todayFile } = getTodayFileNames();
    const buffer = await downloadFile(todayFile);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    const names = [];
    const lastRow = sheet.lastRow.number;
    for (let i = 2; i < lastRow; i++) {
      const row = sheet.getRow(i);
      const name = row.getCell(2).value?.toString().trim();
      if (name && name.toUpperCase() !== "TOTAL") names.push(name);
    }
    res.json(names);
  } catch (err) {
    res.json([]);
  }
});

// Log a visit
app.post("/log", async (req, res) => {
  const { name, visitType, visitTime } = req.body;
  try {
    const { fileName: todayFile, viewName: viewFile } = getTodayFileNames();
    const buffer = await downloadFile(todayFile);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    if (!sheet) return res.status(404).send("Sheet2 not found");
    let executiveRow;
    const lastRow = sheet.rowCount;
    for (let i = 2; i < lastRow; i++) {
      const row = sheet.getRow(i);
      if (row.getCell(2).value?.toString().trim().toUpperCase() === name.trim().toUpperCase()) {
        executiveRow = row;
        break;
      }
    }
    if (!executiveRow) return res.status(404).send({ error: "Executive not found" });
    const timeCell = executiveRow.getCell(5);
    const newEntry = `${visitType.toUpperCase()}-${visitTime}`;
    const existingTime = timeCell.value ? timeCell.value.toString() : "";
    if (existingTime.includes(newEntry)) return res.status(400).send({ error: "Duplicate entry detected." });
    const updatedTime = existingTime ? `${existingTime}/${newEntry}` : newEntry;
    timeCell.value = updatedTime;
    const visits = updatedTime.split("/").map(v => {
      const [type, time] = v.split("-");
      return { type: type.toUpperCase(), time: parseFloat(time) };
    });
    const totalVisit = visits.length;
    const visitTillAfternoon = visits.filter(v => v.time < 12).length;
    const visitAfterAfternoon = visits.filter(v => v.time >= 12).length;
    const typeCounts = { "CD3": 0, "CD5": 0, "CD7": 0, "YB": 0, "MIS": 0 };
    visits.forEach(v => { if (typeCounts[v.type] !== undefined) typeCounts[v.type]++; });
    executiveRow.getCell(3).value = visitTillAfternoon;
    executiveRow.getCell(4).value = totalVisit;
    executiveRow.getCell(6).value = typeCounts["CD3"];
    executiveRow.getCell(7).value = typeCounts["CD5"];
    executiveRow.getCell(8).value = typeCounts["CD7"];
    executiveRow.getCell(9).value = typeCounts["YB"];
    executiveRow.getCell(10).value = typeCounts["MIS"];
    executiveRow.getCell(11).value = visitAfterAfternoon;
    executiveRow.commit();
    const totalRow = sheet.getRow(lastRow);
    const sum = (col) => {
      let total = 0;
      for (let i = 2; i < lastRow; i++) {
        const val = sheet.getRow(i).getCell(col).value;
        total += typeof val === "number" ? val : 0;
      }
      return total;
    };
    totalRow.getCell(3).value = sum(3);
    totalRow.getCell(4).value = sum(4);
    totalRow.getCell(6).value = sum(6);
    totalRow.getCell(7).value = sum(7);
    totalRow.getCell(8).value = sum(8);
    totalRow.getCell(9).value = sum(9);
    totalRow.getCell(10).value = sum(10);
    totalRow.getCell(11).value = sum(11);
    totalRow.commit();
    const bufferUpdated = await workbook.xlsx.writeBuffer();
    await uploadFile(todayFile, bufferUpdated);
    await uploadFile(viewFile, bufferUpdated);
    res.send({ success: true });
  } catch (err) {
    res.status(500).send({ error: err.message });
  }
});

// Add, Remove, Reset, and New-file endpoints unchanged (use your current versions)

// New-file endpoint (clone base file)
app.post("/new-file", async (req, res) => {
  const { fileName, viewName } = getTodayFileNames();
  try {
    try {
      await downloadFile(fileName);
      return res.json({ error: "File for today already exists." });
    } catch {}
    const workbook = await cloneBaseFile(); // clone from base
    const buffer = await workbook.xlsx.writeBuffer();
    await uploadFile(fileName, buffer);
    await uploadFile(viewName, buffer);
    res.json({ success: true, file: fileName });
  } catch (err) {
    res.json({ success: false, error: err.message });
  }
});

// List last 7 files for history dropdown
app.get("/history", async (req, res) => {
  try {
    const { data: files, error } = await supabase.storage.from(BUCKET).list("", { limit: 30 });
    if (error || !files) return res.json([]);
    const logs = files
      .filter(f => /^VisitLog_\d{4}-\d{2}-\d{2}\.xlsx$/.test(f.name))
      .sort((a, b) => b.name.localeCompare(a.name))
      .map(f => f.name)
      .slice(0, 7);
    res.json(logs);
  } catch {
    res.json([]);
  }
});

// Serve file content as JSON for past file report for use in UI table rendering
app.get("/report/:filename", async (req, res) => {
  try {
    const filename = req.params.filename;
    const buffer = await downloadFile(filename);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");

    if (!sheet) return res.status(404).send("Sheet2 not found");

    let html = `
      <html>
      <head>
        <title>Visit Report - ${filename}</title>
        <style>
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f6f8; padding: 20px; color: #333; }
          table { width: 100%; border-collapse: collapse; font-size: 13px; table-layout: auto; }
          th, td {
            border: 1px solid #ccc;
            padding: 4px 8px;
            text-align: center;
            max-width: 120px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
          }
          th {
            background-color: #0078D4;
            color: white;
            font-weight: 600;
            font-size: 13px;
            padding: 4px 8px;
          }
          h1 { color: #0078D4; }
        </style>
      </head>
      <body>
        <h1>Visit Report: ${filename}</h1>
        <table><thead><tr>`;

    // Add headers
    const headerRow = sheet.getRow(1);
    for (let i = 1; i <= headerRow.cellCount; i++) {
      html += `<th>${headerRow.getCell(i).value || ""}</th>`;
    }
    html += `</tr></thead><tbody>`;

    // Add rows except TOTAL row
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header
      const nameCell = row.getCell(2)?.value;
      if (nameCell && nameCell.toString().toUpperCase().includes("TOTAL")) return; // Skip total

      html += "<tr>";
      for (let i = 1; i <= headerRow.cellCount; i++) {
        html += `<td>${row.getCell(i).value || ""}</td>`;
      }
      html += "</tr>";
    });

    html += "</tbody></table></body></html>";
    res.send(html);

  } catch (err) {
    console.error(err);
    res.status(500).send("Error generating report.");
  }
});



// Recalculate total for filtered report
app.get("/report", async (req, res) => {
  try {
    const { fileName } = getTodayFileNames();
    const buffer = await downloadFile(fileName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    let data = [];
    sheet.eachRow((row) => data.push(row.values.slice(1)));

    const { executive, type, time } = req.query;
    if (data.length < 2) return res.json(data);
    let rows = data.slice(1, data.length - 1);

    if (executive)
      rows = rows.filter(r => r[1]?.toString().trim().toUpperCase() === executive.trim().toUpperCase());
    if (type)
      rows = rows.filter(r => r.includes(type));
    if (time === "morning")
      rows = rows.filter(r => typeof r[4] === "string" && r[4].split("/").some(e => parseFloat(e.split("-")[1]) < 12));
    else if (time === "afternoon")
      rows = rows.filter(r => typeof r[4] === "string" && r[4].split("/").some(e => parseFloat(e.split("-")[1]) >= 12));

    // Recalculate totals for filtered rows
    const totals = new Array(data[0].length).fill("");
    totals[0] = "TOTAL";
    for (let i = 0; i < rows.length; i++) {
      for (let c = 3; c <= 11; c++) {
        const val = rows[i][c - 1];
        if (typeof val === "number") totals[c - 1] = (totals[c - 1] || 0) + val;
      }
    }

    const out = [data[0], ...rows, totals];
    res.json(out);
  } catch {
    res.json([]);
  }
});

// Reset today's sheet clearing data but keeping headers and executive names
app.post("/reset", async (req, res) => {
  try {
    const { fileName, viewName } = getTodayFileNames();
    const buffer = await downloadFile(fileName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    if (!sheet) return res.status(404).send("Sheet2 not found");

    // Clear columns C (3) to K (11) except headers and "TOTAL" row
    sheet.eachRow((row, rowNum) => {
      if (rowNum > 1 && row.getCell(2).value?.toString().toUpperCase() !== "TOTAL") {
        for (let col = 3; col <= 11; col++) {
          row.getCell(col).value = col === 2 ? row.getCell(col).value : null;
        }
      }
    });

    const bufferUpdated = await workbook.xlsx.writeBuffer();
    await uploadFile(fileName, bufferUpdated);
    await uploadFile(viewName, bufferUpdated);

    res.json({ success: true });
  } catch (err) {
    res.status(500).send({ error: err.message });
  }
});

app.listen(3000, () => console.log("✅ Server running at http://localhost:3000"));
