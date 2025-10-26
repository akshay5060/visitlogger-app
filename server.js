const express = require("express");
const ExcelJS = require("exceljs");
const path = require("path");
const { createClient } = require("@supabase/supabase-js");

const app = express();
app.use(express.json());

const supabase = createClient(
  "https://jjsotbdvooeksoceulbz.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Impqc290YmR2b29la3NvY2V1bGJ6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjE0MDg2ODcsImV4cCI6MjA3Njk4NDY4N30.eZ9bFwCOfpYHdSD_ko-jaR0H28T6u-CnDbJ5BKTCRuk"
);
const BUCKET = "visit-logs"; // Change to your Supabase bucket name

// Helper functions
async function downloadFile(filename) {
  const { data, error } = await supabase.storage.from(BUCKET).download(filename);
  if (error) return null;
  return Buffer.from(await data.arrayBuffer());
}

async function uploadFile(filename, buffer) {
  const { error } = await supabase.storage.from(BUCKET).upload(filename, buffer, { upsert: true });
  if (error) throw error;
}

async function loadWorkbook(filename) {
  const buffer = await downloadFile(filename);
  const workbook = new ExcelJS.Workbook();
  if (buffer) {
    await workbook.xlsx.load(buffer);
  } else {
    const sheet = workbook.addWorksheet("Sheet2");
    sheet.addRow(["SNO", "EXECUTIVE", "VISIT TOOL UTILIZATION", "TOTAL", "TIME", "CD3", "CD5", "CD7", "YB", "MIS", "AFTERNOON"]);
    sheet.addRow(["TOTAL", "", "", 0, "", 0, 0, 0, 0, 0, 0]);
  }
  return workbook;
}

async function saveWorkbook(workbook, filename) {
  const buffer = await workbook.xlsx.writeBuffer();
  await uploadFile(filename, buffer);
}

// Delete old view-only files
async function deleteOldViewOnlyLogs(todayFilename) {
  const { data: files } = await supabase.storage.from(BUCKET).list();
  for (const file of files) {
    const isViewOnly = /^VisitLog_ViewOnly_\d{4}-\d{2}-\d{2}\.xlsx$/.test(file.name);
    if (isViewOnly && file.name !== todayFilename) {
      await supabase.storage.from(BUCKET).remove([file.name]);
      console.log(`Deleted old view-only file: ${file.name}`);
    }
  }
}

// Daily filenames
const today = new Date();
const dateStr = today.toISOString().slice(0, 10);
const fileName = `VisitLog_${dateStr}.xlsx`;
const viewFileName = `VisitLog_ViewOnly_${dateStr}.xlsx`;

// Ensure today's file exists
(async () => {
  const workbook = await loadWorkbook(fileName);
  await saveWorkbook(workbook, fileName);
  await saveWorkbook(workbook, viewFileName);
})();

// --------------------- ROUTES ---------------------

// Log a visit
app.post("/log", async (req, res) => {
  try {
    const { name, visitType, visitTime } = req.body;
    const workbook = await loadWorkbook(fileName);
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

    timeCell.value = existingTime ? `${existingTime}/${newEntry}` : newEntry;

    const visits = timeCell.value.split("/").map(v => {
      const [type, time] = v.split("-");
      return { type: type.toUpperCase(), time: parseFloat(time) };
    });

    const totalVisit = visits.length;
    const visitTillAfternoon = visits.filter(v => v.time < 12).length;
    const visitAfterAfternoon = visits.filter(v => v.time >= 12).length;
    const typeCounts = { CD3: 0, CD5: 0, CD7: 0, YB: 0, MIS: 0 };
    visits.forEach(v => { if (typeCounts[v.type] !== undefined) typeCounts[v.type]++; });

    executiveRow.getCell(3).value = visitTillAfternoon;
    executiveRow.getCell(4).value = totalVisit;
    executiveRow.getCell(6).value = typeCounts.CD3;
    executiveRow.getCell(7).value = typeCounts.CD5;
    executiveRow.getCell(8).value = typeCounts.CD7;
    executiveRow.getCell(9).value = typeCounts.YB;
    executiveRow.getCell(10).value = typeCounts.MIS;
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
    [3,4,6,7,8,9,10,11].forEach(col => totalRow.getCell(col).value = sum(col));
    totalRow.commit();

    await saveWorkbook(workbook, fileName);
    await saveWorkbook(workbook, viewFileName);

    res.send({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).send({ error: "Internal server error" });
  }
});

// View report (with filters)
app.get("/report", async (req, res) => {
  try {
    const { executive, type, time } = req.query;
    const workbook = await loadWorkbook(fileName);
    const sheet = workbook.getWorksheet("Sheet2");

    const data = [];
    const headers = sheet.getRow(1).values.slice(1);
    data.push(headers);

    const lastRow = sheet.lastRow.number;
    for (let i = 2; i <= lastRow; i++) {
      const row = sheet.getRow(i);
      const name = row.getCell(2).value?.toString().trim();
      const timeString = row.getCell(5).value?.toString().trim() || "";

      if (name === "" && row.getCell(1).value?.toString().trim().toUpperCase() === "TOTAL") continue;
      if (executive && name?.toUpperCase() !== executive.toUpperCase()) continue;
      if (type && !timeString.includes(type.toUpperCase())) continue;

      if (time === "morning" && !timeString.split("/").some(v => parseFloat(v.split("-")[1]) < 12)) continue;
      if (time === "afternoon" && !timeString.split("/").some(v => parseFloat(v.split("-")[1]) >= 12)) continue;

      data.push(row.values.slice(1));
    }

    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).send({ error: "Internal server error" });
  }
});

// Reset logs
app.post("/reset", async (req, res) => {
  try {
    const workbook = await loadWorkbook(fileName);
    const sheet = workbook.getWorksheet("Sheet2");

    const lastRow = sheet.lastRow.number;
    for (let i = 2; i < lastRow; i++) {
      const row = sheet.getRow(i);
      for (let col = 3; col <= 11; col++) row.getCell(col).value = "";
      row.getCell(5).value = "";
      row.commit();
    }

    const totalRow = sheet.getRow(lastRow);
    for (let col = 3; col <= 11; col++) totalRow.getCell(col).value = 0;
    totalRow.commit();

    await saveWorkbook(workbook, fileName);
    await saveWorkbook(workbook, viewFileName);
    res.send({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).send({ error: "Internal server error" });
  }
});

// Get executive list
app.get("/executives", async (req, res) => {
  try {
    const workbook = await loadWorkbook(fileName);
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
    res.status(500).send({ error: "Internal server error" });
  }
});

// -----------------------
// The rest of your routes
// `/new-file`, `/add-executive`, `/remove-executive`, `/history`, `/view-excel`, `/download-excel`, `/report/:filename`
// Can all be adapted in the same way using loadWorkbook() and saveWorkbook() instead of fs.
// -----------------------

// Start server
app.listen(3000, () => console.log("✅ Server running at http://localhost:3000"));
