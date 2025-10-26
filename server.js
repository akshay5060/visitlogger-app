const express = require("express");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");
const { createClient } = require('@supabase/supabase-js');
const supabase = createClient('https://jjsotbdvooeksoceulbz.supabase.co', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Impqc290YmR2b29la3NvY2V1bGJ6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjE0MDg2ODcsImV4cCI6MjA3Njk4NDY4N30.eZ9bFwCOfpYHdSD_ko-jaR0H28T6u-CnDbJ5BKTCRuk');
const app = express();
app.use(express.json());
app.use(express.static(__dirname));

function deleteOldViewOnlyLogs(directory, todayFilename) {
  const files = fs.readdirSync(directory);
  files.forEach(file => {
    const isViewOnly = /^VisitLog_ViewOnly_\d{4}-\d{2}-\d{2}\.xlsx$/.test(file);
    if (isViewOnly && file !== todayFilename) {
      fs.unlinkSync(path.join(directory, file));
      console.log(`Deleted old view-only file: ${file}`);
    }
  });
}
const today = new Date();
const dateStr = today.toISOString().slice(0, 10); // e.g., "2025-10-25"
const fileName = `VisitLog_${dateStr}.xlsx`;

const filePath = path.join(__dirname, fileName);
const viewPath = path.join(__dirname, `VisitLog_ViewOnly_${dateStr}.xlsx`);

if (!fs.existsSync(filePath)) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sheet2");

  // Add headers
  sheet.addRow(["SNO", "EXECUTIVE", "VISIT TOOL UTILIZATION", "TOTAL", "TIME", "CD3", "CD5", "CD7", "YB", "MIS", "AFTERNOON"]);
  sheet.addRow(["TOTAL", "", "", 0, "", 0, 0, 0, 0, 0, 0]);

  workbook.xlsx.writeFile(filePath);
}
// Log a visit
app.post("/log", async (req, res) => {
  const { name, visitType, visitTime } = req.body;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
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
  if (existingTime.includes(newEntry)) {
  return res.status(400).send({ error: "Duplicate entry detected." });
}
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
  visits.forEach(v => {
    if (typeCounts[v.type] !== undefined) typeCounts[v.type]++;
  });

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

  await workbook.xlsx.writeFile(filePath);
  fs.copyFile(filePath, viewPath, () => {});
  res.send({ success: true });
});

// View report (with filters)
app.get("/report", async (req, res) => {
  const { executive, type, time } = req.query;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
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

    if (time === "morning" && !timeString.split("/").some(v => {
      const parts = v.split("-");
      return parseFloat(parts[1]) < 12;
    })) continue;

    if (time === "afternoon" && !timeString.split("/").some(v => {
      const parts = v.split("-");
		  return parseFloat(parts[1]) >= 12;
		})) continue;

		const values = row.values.slice(1);
		data.push(values);
	  }

	  res.json(data);
});

// Reset logs
app.post("/reset", async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("Sheet2");

  const lastRow = sheet.lastRow.number;

  for (let i = 2; i < lastRow; i++) {
    const row = sheet.getRow(i);
    for (let col = 3; col <= 11; col++) {
      row.getCell(col).value = "";
    }
    row.getCell(5).value = "";
    row.commit();
  }

  const totalRow = sheet.getRow(lastRow);
  for (let col = 3; col <= 11; col++) {
    totalRow.getCell(col).value = 0;
  }
  totalRow.commit();

  await workbook.xlsx.writeFile(filePath);
  fs.copyFile(filePath, viewPath, () => {});
  res.send({ success: true });
});

// Get executive list
app.get("/executives", async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("Sheet2");

  const names = [];
  const lastRow = sheet.lastRow.number;

  for (let i = 2; i < lastRow; i++) {
    const row = sheet.getRow(i);
    const name = row.getCell(2).value?.toString().trim();
    if (name) names.push(name);
  }

  res.json(names);
});

app.post("/new-file", async (req, res) => {
  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10);
  const newFileName = `VisitLog_${dateStr}.xlsx`;
  const newPath = path.join(__dirname, newFileName);
  

  if (fs.existsSync(newPath)) {
    return res.status(400).send({ error: "File for today already exists." });
  }

  // Find the most recent VisitLog file
  const files = fs.readdirSync(__dirname)
    .filter(f => f.startsWith("VisitLog_") && f.endsWith(".xlsx"))
    .sort()
    .reverse();

  if (files.length === 0) {
    return res.status(500).send({ error: "No previous VisitLog file found to clone." });
  }

  const latestFile = path.join(__dirname, files[0]);

  // Load latest workbook
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(latestFile);
  const sheet = workbook.getWorksheet("Sheet2");

  const lastRow = sheet.lastRow.number;

  // Clear visit data for each executive row
  for (let i = 2; i < lastRow; i++) {
    const row = sheet.getRow(i);
    for (let col = 3; col <= 11; col++) {
      row.getCell(col).value = "";
    }
    row.getCell(5).value = ""; // Clear TIME
    row.commit();
  }

  // Reset TOTAL row
  const totalRow = sheet.getRow(lastRow);
  for (let col = 3; col <= 11; col++) {
    totalRow.getCell(col).value = 0;
  }
  totalRow.commit();

  await workbook.xlsx.writeFile(newPath);
  res.send({ success: true, file: newFileName });
  const todayViewOnly = `VisitLog_ViewOnly_${dateStr}.xlsx`;
deleteOldViewOnlyLogs(__dirname, todayViewOnly);

});

// Add a new executive
app.post("/add-executive", async (req, res) => {
  const { name } = req.body;
  if (!name) return res.status(400).send({ error: "Name required" });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("Sheet2");

  const lastRow = sheet.lastRow.number;
  const totalRow = sheet.getRow(lastRow);

  // Insert above TOTAL
  sheet.insertRow(lastRow, [
    lastRow - 1,
    name.trim().toUpperCase(),
    "", "", "", "", "", "", "", "", "", ""
  ]);

  // Re-number rows
  for (let i = 2; i < lastRow; i++) {
    sheet.getRow(i).getCell(1).value = i - 1;
  }

  await workbook.xlsx.writeFile(filePath);
  fs.copyFile(filePath, viewPath, () => {});
  res.send({ success: true });
});

app.get("/history", (req, res) => {
  const files = fs.readdirSync(__dirname)
    .filter(f => f.startsWith("VisitLog_") && f.endsWith(".xlsx") && !f.includes("ViewOnly"))
    .sort()
    .reverse()
    .slice(0, 7);

  res.json(	files);
});


app.get("/history/:filename", (req, res) => {
  const file = path.join(__dirname, req.params.filename);
  if (fs.existsSync(file)) {
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "inline; filename=" + req.params.filename);
    res.sendFile(file);
  } else {
    res.status(404).send("File not found");
  }
});

	
// Remove an executive
app.post("/remove-executive", async (req, res) => {
  const { name } = req.body;
  if (!name) return res.status(400).send({ error: "Name required" });

  const target = name.trim().toUpperCase();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("Sheet2");

  const lastRow = sheet.rowCount;
  let found = false;

  for (let i = 2; i <= lastRow; i++) {
    const row = sheet.getRow(i);
    const cell = row.getCell(2);
    const value = typeof cell.value === "string"
      ? cell.value.trim().toUpperCase()
      : cell.value?.toString().trim().toUpperCase();

    if (value === target) {
      // Subtract visit counts from TOTAL row
      const totalRow = sheet.getRow(sheet.rowCount);
      for (let col = 3; col <= 11; col++) {
        const execVal = parseInt(row.getCell(col).value || 0);
        const totalVal = parseInt(totalRow.getCell(col).value || 0);
        totalRow.getCell(col).value = totalVal - execVal;
      }

      sheet.spliceRows(i, 1);
      found = true;
      break;
    }
  }

  if (!found) {
    console.log(`❌ Executive "${target}" not found.`);
    return res.status(404).send({ error: "Executive not found" });
  }

  // Re-number rows
  for (let i = 2; i < sheet.rowCount; i++) {
    sheet.getRow(i).getCell(1).value = i - 1;
  }

  await workbook.xlsx.writeFile(filePath);
  fs.copyFile(filePath, viewPath, () => {});
  res.send({ success: true });
});

async function loadExecutives() {
  const res = await fetch("/executives");
  const names = await res.json();
  names.sort((a, b) => a.localeCompare(b)); // Sort alphabetically

  const dropdowns = [document.getElementById("employee"), document.getElementById("filterExecutive")];
  dropdowns.forEach(drop => {
    drop.innerHTML = drop.id === "employee"
      ? `<option value="" disabled selected>-- Select Executive --</option>`
      : `<option value="">-- Filter by Executive --</option>`;
    names.forEach(name => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      drop.appendChild(option.cloneNode(true));
    });
  });
}
// Serve view-only Excel inline
app.get("/view-excel", (req, res) => {
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "inline; filename=VisitLog_ViewOnly.xlsx");
  res.sendFile(viewPath);
});
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "VisitLogger.html"));
});

app.get("/download-excel", async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("Sheet2");

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=Visit_Report.xlsx");

  await workbook.xlsx.write(res);
  res.end();
});

app.get("/report/:filename", async (req, res) => {
  const filePath = path.join(__dirname, req.params.filename);
  if (!fs.existsSync(filePath)) return res.status(404).send("File not found");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("Sheet2");
  if (!sheet) return res.status(404).send("Sheet2 not found");

  let html = `
  <html>
  <head>
    <title>VisitLogger Report</title>
    <style>
      body {
        font-family: 'Segoe UI', sans-serif;
        background: #f4f6f8;
        padding: 20px;
        color: #333;
      }
      h1 {
        color: #0078D4;
        margin-bottom: 10px;
        font-size: 20px;
      }
      h3 {
        margin-bottom: 10px;
        font-size: 16px;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        background: #fff;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        font-size: 13px;
      }
      th, td {
        border: 1px solid #ccc;
        padding: 4px 8px;
        text-align: left;
        vertical-align: middle;
      }
      th {
        background-color: #e6f0ff;
        font-weight: 600;
      }
      .top-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
      }
      .btn {
        background: #0078D4;
        color: white;
        padding: 6px 10px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 12px;
        margin-right: 8px;
      }
      #rowCount {
        margin-bottom: 10px;
        font-size: 13px;
        color: #555;
      }
    </style>
  </head>
  <body>
    <div class="top-bar">
      <h1>VisitLogger Report</h1>
      <button class="btn" onclick="window.print()">🖨️ Print</button>
    </div>
    <h3>${req.params.filename}</h3>

    <div id="rowCount">
      Showing <span id="visibleCount">0</span> rows
    </div>

    <table>
  `;

  const headerRow = sheet.getRow(1);
  const headers = [];

  html += "<tr>";
  for (let i = 1; i <= headerRow.cellCount; i++) {
    const cell = headerRow.getCell(i);
    const headerText = cell.value || "";
    headers.push(headerText);
    html += `<th>${headerText}</th>`;
  }
  html += "</tr>";

  let visibleRowCount = 0;

  sheet.eachRow((row, rowIndex) => {
    if (rowIndex === 1) return;
    const nameCell = row.getCell(2)?.value?.toString().trim().toLowerCase();
    const isTotalRow = nameCell === "total" || nameCell?.includes("total");
    if (isTotalRow) return;

    html += "<tr>";
    for (let i = 1; i <= headers.length; i++) {
      const cell = row.getCell(i);
      html += `<td>${cell.value || ""}</td>`;
    }
    html += "</tr>";
    visibleRowCount++;
  });

  html += `
    </table>

    <script>
      document.getElementById("visibleCount").textContent = ${visibleRowCount};
    </script>
  </body>
  </html>
  `;

  res.send(html);
});
// Upload file
async function uploadFile(fileName, buffer) {
  const { data, error } = await supabase
    .storage
    .from('visitlogs')
    .upload(fileName, buffer, { upsert: true });
  if (error) console.error(error);
}

// Download file
async function downloadFile(fileName) {
  const { data, error } = await supabase
    .storage
    .from('visitlogs')
    .download(fileName);
  if (error) throw error;
  return data.arrayBuffer(); // convert to buffer
}
// Start server
app.listen(3000, () => console.log("✅ Server running at http://localhost:3000"));
