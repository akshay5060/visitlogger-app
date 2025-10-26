require('dotenv').config();
const express = require("express");
const ExcelJS = require("exceljs");
const { createClient } = require("@supabase/supabase-js");

const app = express();
app.use(express.json());

// ---------------------
// Supabase setup
// ---------------------
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);
const BUCKET = "visit-logs";

// ---------------------
// Helper functions
// ---------------------
async function uploadFile(fileName, buffer) {
  const { error } = await supabase.storage.from(BUCKET).upload(fileName, buffer, { upsert: true });
  if (error) throw error;
}

async function downloadFile(fileName) {
  const { data, error } = await supabase.storage.from(BUCKET).download(fileName);
  if (error) throw error;
  return Buffer.from(await data.arrayBuffer());
}

async function deleteOldViewOnlyLogs(todayFilename) {
  const { data: files, error } = await supabase.storage.from(BUCKET).list("", { limit: 100 });
  if (error) console.error("Error listing files:", error);
  if (!files) return;

  for (const file of files) {
    const isViewOnly = /^VisitLog_ViewOnly_\d{4}-\d{2}-\d{2}\.xlsx$/.test(file.name);
    if (isViewOnly && file.name !== todayFilename) {
      await supabase.storage.from(BUCKET).remove([file.name]);
      console.log(`Deleted old view-only file: ${file.name}`);
    }
  }
}

// ---------------------
// Get today filenames
// ---------------------
function getTodayFileNames() {
  const today = new Date().toISOString().slice(0, 10);
  return {
    fileName: `VisitLog_${today}.xlsx`,
    viewName: `VisitLog_ViewOnly_${today}.xlsx`
  };
}

// ---------------------
// Initialize today's file if not exists
// ---------------------
async function initTodayFile() {
  const { fileName, viewName } = getTodayFileNames();
  try {
    await downloadFile(fileName); // Already exists
  } catch {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet2");
    sheet.addRow(["SNO", "EXECUTIVE", "VISIT TOOL UTILIZATION", "TOTAL", "TIME", "CD3", "CD5", "CD7", "YB", "MIS", "AFTERNOON"]);
    sheet.addRow(["TOTAL", "", "", 0, "", 0, 0, 0, 0, 0, 0]);
    const buffer = await workbook.xlsx.writeBuffer();
    await uploadFile(fileName, buffer);
    await uploadFile(viewName, buffer);
  }
}

initTodayFile();

// ---------------------
// Routes
// ---------------------
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "VisitLogger.html"));
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
    console.error(err);
    res.status(500).send({ error: err.message });
  }
});

// Get executives
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
    console.error(err);
    res.status(500).send({ error: err.message });
  }
});

// History and report routes
app.get("/history/:filename", async (req, res) => {
  try {
    const buffer = await downloadFile(req.params.filename);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `inline; filename=${req.params.filename}`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(404).send("File not found");
  }
});

// Add alias route for /report/:filename
app.get("/report/:filename", async (req, res) => {
  try {
    const buffer = await downloadFile(req.params.filename);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `inline; filename=${req.params.filename}`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(404).send("File not found");
  }
});

// ---------------------
// Start server
// ---------------------
app.listen(3000, () => console.log("✅ Server running at http://localhost:3000"));
