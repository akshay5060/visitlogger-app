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
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet2");
    // Example: starter executives, customize for your team
    sheet.addRow(["SNO", "EXECUTIVE", "VISIT TOOL UTILIZATION", "TOTAL", "TIME", "CD3", "CD5", "CD7", "YB", "MIS", "AFTERNOON"]);
    sheet.addRow(["TOTAL", "", "", 0, "", 0, 0, 0, 0, 0, 0]);
    const buffer = await workbook.xlsx.writeBuffer();
    await uploadFile(fileName, buffer);
    await uploadFile(viewName, buffer);
  }
}
initTodayFile();

// --- API Endpoints ---

// Serve HTML
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
      if (name) names.push(name);
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

// Add Executive
app.post("/add-executive", async (req, res) => {
  const { name } = req.body;
  try {
    const { fileName: todayFile, viewName: viewFile } = getTodayFileNames();
    const buffer = await downloadFile(todayFile);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    const lastRow = sheet.rowCount;
    // Prevent duplicate executive
    for (let i = 2; i < lastRow; i++) {
      const row = sheet.getRow(i);
      if (row.getCell(2).value?.toString().trim().toUpperCase() === name.trim().toUpperCase()) {
        return res.json({ success: false, error: "Executive already exists" });
      }
    }
    sheet.insertRow(lastRow, [lastRow, name, 0, 0, "", 0, 0, 0, 0, 0, 0]);
    const bufferUpdated = await workbook.xlsx.writeBuffer();
    await uploadFile(todayFile, bufferUpdated);
    await uploadFile(viewFile, bufferUpdated);
    res.json({ success: true });
  } catch (err) {
    res.json({ success: false, error: err.message });
  }
});

// Remove Executive
app.post("/remove-executive", async (req, res) => {
  const { name } = req.body;
  try {
    const { fileName: todayFile, viewName: viewFile } = getTodayFileNames();
    const buffer = await downloadFile(todayFile);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    let found = false;
    const lastRow = sheet.rowCount;
    for (let i = 2; i < lastRow; i++) {
      const row = sheet.getRow(i);
      if (row.getCell(2).value?.toString().trim().toUpperCase() === name.trim().toUpperCase()) {
        sheet.spliceRows(i, 1);
        found = true;
        break;
      }
    }
    if (!found) return res.json({ success: false, error: "Executive not found" });
    const bufferUpdated = await workbook.xlsx.writeBuffer();
    await uploadFile(todayFile, bufferUpdated);
    await uploadFile(viewFile, bufferUpdated);
    res.json({ success: true });
  } catch (err) {
    res.json({ success: false, error: err.message });
  }
});

// Reset today's logs (clear columns except SNO/EXECUTIVE)
app.post("/reset", async (req, res) => {
  try {
    const { fileName: todayFile, viewName: viewFile } = getTodayFileNames();
    const buffer = await downloadFile(todayFile);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    const lastRow = sheet.rowCount;
    for (let i = 2; i < lastRow; i++) {
      const row = sheet.getRow(i);
      row.getCell(3).value = 0;
      row.getCell(4).value = 0;
      row.getCell(5).value = "";
      row.getCell(6).value = 0;
      row.getCell(7).value = 0;
      row.getCell(8).value = 0;
      row.getCell(9).value = 0;
      row.getCell(10).value = 0;
      row.getCell(11).value = 0;
      row.commit();
    }
    // Clear totals as well
    const totalRow = sheet.getRow(lastRow);
    for (let c = 3; c <= 11; c++) {
      totalRow.getCell(c).value = 0;
    }
    totalRow.commit();
    const bufferUpdated = await workbook.xlsx.writeBuffer();
    await uploadFile(todayFile, bufferUpdated);
    await uploadFile(viewFile, bufferUpdated);
    res.json({ success: true });
  } catch (err) {
    res.json({ success: false, error: err.message });
  }
});

// Create new file for today
app.post("/new-file", async (req, res) => {
  const { fileName, viewName } = getTodayFileNames();
  try {
    // If file exists, error
    try {
      await downloadFile(fileName);
      return res.json({ error: "File for today already exists." });
    } catch {}
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet2");
    sheet.addRow(["SNO", "EXECUTIVE", "VISIT TOOL UTILIZATION", "TOTAL", "TIME", "CD3", "CD5", "CD7", "YB", "MIS", "AFTERNOON"]);
    sheet.addRow(["TOTAL", "", "", 0, "", 0, 0, 0, 0, 0, 0]);
    const buffer = await workbook.xlsx.writeBuffer();
    await uploadFile(fileName, buffer);
    await uploadFile(viewName, buffer);
    res.json({ success: true, file: fileName });
  } catch (err) {
    res.json({ success: false, error: err.message });
  }
});

// Get history (list filenames from bucket, latest 7)
app.get("/history", async (req, res) => {
  try {
    const {  files, error } = await supabase.storage.from(BUCKET).list("", { limit: 30 });
    if (error || !files) return res.json([]);
    // Only log files for last 7 days, sorted reverse
    const logs = files
      .filter(f => /^VisitLog_\d{4}-\d{2}-\d{2}\.xlsx$/.test(f.name))
      .sort((a, b) => b.name.localeCompare(a.name)) // latest first
      .map(f => f.name)
      .slice(0, 7);
    res.json(logs);
  } catch (err) {
    res.json([]);
  }
});

// Download EXCEL (today's file)
app.get("/download-excel", async (req, res) => {
  try {
    const { fileName } = getTodayFileNames();
    const buffer = await downloadFile(fileName);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=${fileName}`);
    res.send(buffer);
  } catch (err) {
    res.status(500).send("Failed to download.");
  }
});

// View any file (history/report download)
app.get("/history/:filename", async (req, res) => {
  try {
    const buffer = await downloadFile(req.params.filename);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `inline; filename=${req.params.filename}`);
    res.send(buffer);
  } catch (err) {
    res.status(404).send("File not found");
  }
});

// Report endpoint - returns as JSON for the table
app.get("/report", async (req, res) => {
  try {
    const { fileName } = getTodayFileNames();
    const buffer = await downloadFile(fileName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet("Sheet2");
    let data = [];
    sheet.eachRow((row, rowNum) => {
      data.push(row.values.slice(1)); // remove empty first element
    });
    // Filtering support (query params)
    const { executive, type, time } = req.query;
    if (data.length < 2) return res.json(data);
    let rows = data.slice(1, data.length - 1); // skip header + TOTAL row
    if (executive) {
      rows = rows.filter(r => r[1]?.toString().trim().toUpperCase() === executive.trim().toUpperCase());
    }
    if (type) {
      rows = rows.filter(r => r.includes(type));
    }
    if (time === "morning") {
      rows = rows.filter(r => typeof r[4] === "string" && r[4].split("/").some(e => {
        const t = parseFloat(e.split("-")[1]);
        return t && t < 12;
      }));
    } else if (time === "afternoon") {
      rows = rows.filter(r => typeof r[4] === "string" && r[4].split("/").some(e => {
        const t = parseFloat(e.split("-")[1]);
        return t && t >= 12;
      }));
    }
    // Format back for table: header + filtered rows + total row
    const out = [data[0], ...rows, data[data.length - 1]];
    res.json(out);
  } catch (err) {
    res.json([]);
  }
});

app.listen(3000, () => console.log("✅ Server running at http://localhost:3000"));
