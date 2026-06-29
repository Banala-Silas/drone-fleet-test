// ─────────────────────────────────────────────
// Drone Fleet Manager — Google Apps Script v2
// Supports: Operator App + Technician App
// Adds: DroneHistory audit log
// ─────────────────────────────────────────────

const SHEET_NAME = "Drones";
const HISTORY_SHEET_NAME = "DroneHistory";

// "ajay7836899826@gmail.com,Krtripathi.ashish@gmail.com,francisjaladi13@gmail.com";

const EMAIL_TO  = "ajay7836899826@gmail.com,Krtripathi.ashish@gmail.com,sfthings0@gmail.com";
const EMAIL_CC  = "silasbanala@gmail.com";

// ─────────────────────────────────────────────
// WEB APP ENTRY POINTS
// ─────────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action;

  try {
    let result;

    if (action === "getDrones") {
      result = getDrones();

    } else if (action === "updateDrone") {
      result = updateDrone(
        params.droneId,
        params.status,
        params.reason || "",
        params.notes || ""
      );

    } else if (action === "getFailsByDate") {
      result = getFailsByDate();

    } else if (action === "batchUpdate") {
      result = batchUpdate(
        params.droneIds,
        params.status
      );

    } else if (action === "autoTransition") {
      result = autoTransition();

    } else if (action === "sendDailyReport") {
      result = sendDailyReport();

    }

    // ==========================================
    // CONFIGURATION MODULE
    // ==========================================

    else if (action === "getConfigDrones") {
      result = getConfigDrones();

    } else if (action === "createNextBatch") {
      result = createNextBatch(params.portId);

    } else if (action === "batchConfigUpdate") {
      result = batchConfigUpdate(
        params.droneIds,
        params.status
      );

    }

    // ==========================================

    else {
      result = {
        error: "Unknown action: " + action
      };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────
// SHARED HELPERS
// ─────────────────────────────────────────────
function findCol(headers, candidates) {
  for (const c of candidates) {
    const cc = String(c).toLowerCase().trim();
    const idx = headers.findIndex(h => h === cc || h.includes(cc));
    if (idx >= 0) return idx;
  }
  return -1;
}

function getSpreadsheetTimeZone() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
}

function getSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) throw new Error(`Sheet "${SHEET_NAME}" is empty`);

  const headers = data[0].map(h => String(h).toLowerCase().trim());

  const col = {
    id:       findCol(headers, ["droneid", "drone id", "id"]),
    name:     findCol(headers, ["drone name", "dronename", "name"]),
    status:   findCol(headers, ["status"]),
    reason:   findCol(headers, ["reason", "fail reason"]),
    notes:    findCol(headers, ["notes", "note", "remarks"]),
    updated:  findCol(headers, ["updated", "last updated", "timestamp"]),
    fixed_at: findCol(headers, ["fixed_at", "fixed at", "fixedat"]),
  };

  return { sheet, data, headers, col };
}

function getHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
  }

  const headers = [
    "event_id",
    "drone_id",
    "drone_name",
    "from_status",
    "to_status",
    "reason",
    "changed_by",
    "changed_at"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 8)).getValues()[0]
      .map(h => String(h).toLowerCase().trim());
    const missing = headers.some((h, i) => existingHeaders[i] !== h);
    if (missing) {
      sheet.clearContents();
      sheet.appendRow(headers);
    }
  }

  return sheet;
}

function getNextEventId(historySheet) {
  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) return "EVT-00001";

  const lastId = String(historySheet.getRange(lastRow, 1).getValue()).trim();
  const match = lastId.match(/EVT-(\d+)/i);

  if (match) {
    const nextNum = parseInt(match[1], 10) + 1;
    return "EVT-" + String(nextNum).padStart(5, "0");
  }

  return "EVT-" + String(lastRow).padStart(5, "0");
}

function appendDroneHistory({ droneId, droneName, fromStatus, toStatus, reason, changedBy }) {
  const historySheet = getHistorySheet();
  const eventId = getNextEventId(historySheet);

  historySheet.appendRow([
    eventId,
    droneId || "",
    droneName || "",
    fromStatus || "",
    toStatus || "",
    reason || "",
    changedBy || "",
    new Date()
  ]);
}

function findDroneRow(data, col, droneId) {
  for (let i = 1; i < data.length; i++) {
    const rowId = col.id >= 0 ? String(data[i][col.id]).trim() : "";
    const rowName = col.name >= 0 ? String(data[i][col.name]).trim() : String(data[i][0]).trim();
    if (rowId === droneId || rowName === droneId) return i + 1;
  }
  return -1;
}

function toDateKey(val) {
  if (!val) return null;
  try {
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d)) return null;
    return Utilities.formatDate(d, getSpreadsheetTimeZone(), "yyyy-MM-dd");
  } catch (e) {
    return null;
  }
}

function formatDate(val) {
  if (!val) return "";
  try {
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d)) return String(val);
    return Utilities.formatDate(d, getSpreadsheetTimeZone(), "dd MMM yyyy, HH:mm");
  } catch (e) {
    return String(val);
  }
}

function formatDisplayDate(dateKey) {
  try {
    const [y, m, d] = dateKey.split("-");
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return `${d} ${months[parseInt(m, 10) - 1]} ${y}`;
  } catch (e) {
    return dateKey;
  }
}

function getDroneMeta(sheet, data, col, rowIndex) {
  const row = data[rowIndex - 1];
  return {
    droneId: col.id >= 0 ? String(row[col.id]).trim() : `D${String(rowIndex - 1).padStart(2, "0")}`,
    droneName: col.name >= 0 ? String(row[col.name]).trim() : String(row[0]).trim(),
    status: col.status >= 0 ? String(row[col.status]).trim() : "",
    reason: col.reason >= 0 ? String(row[col.reason]).trim() : "",
    notes: col.notes >= 0 ? String(row[col.notes]).trim() : "",
    updated: col.updated >= 0 ? row[col.updated] : "",
    fixed_at: col.fixed_at >= 0 ? row[col.fixed_at] : "",
  };
}

// ─────────────────────────────────────────────
// OPERATOR: GET ALL DRONES
// ─────────────────────────────────────────────
function getDrones() {
  const { data, col } = getSheetData();
  if (data.length < 2) return { drones: [] };

  const drones = data.slice(1)
    .filter(row => row[col.name >= 0 ? col.name : 0])
    .map((row, i) => ({
      rowIndex: i + 2,
      id:       col.id       >= 0 ? String(row[col.id]).trim()     : `D${String(i+1).padStart(2,"0")}`,
      name:     col.name     >= 0 ? String(row[col.name]).trim()   : String(row[0]),
      status:   col.status   >= 0 ? String(row[col.status]).trim() : "Unknown",
      reason:   col.reason   >= 0 ? String(row[col.reason]).trim() : "",
      notes:    col.notes    >= 0 ? String(row[col.notes]).trim()  : "",
      updated:  col.updated   >= 0 ? formatDate(row[col.updated])   : "",
      fixed_at: col.fixed_at  >= 0 ? formatDate(row[col.fixed_at])  : "",
    }));

  return { drones };
}

// ─────────────────────────────────────────────
// OPERATOR: UPDATE SINGLE DRONE
// ─────────────────────────────────────────────
function updateDrone(droneId, status, reason, notes) {
  const { sheet, data, col } = getSheetData();
  const rowIndex = findDroneRow(data, col, droneId);

  if (rowIndex === -1) {
    return { error: `Drone "${droneId}" not found` };
  }

  const now = new Date();

  const fromStatus = col.status >= 0 ? String(sheet.getRange(rowIndex, col.status + 1).getValue()).trim() : "";
  const droneName  = col.name   >= 0 ? String(sheet.getRange(rowIndex, col.name + 1).getValue()).trim() : String(sheet.getRange(rowIndex, 1).getValue()).trim();

  if (col.status >= 0) sheet.getRange(rowIndex, col.status + 1).setValue(status);
  if (col.reason >= 0) sheet.getRange(rowIndex, col.reason + 1).setValue(reason);
  if (col.notes  >= 0) sheet.getRange(rowIndex, col.notes + 1).setValue(notes);
  if (col.updated >= 0) sheet.getRange(rowIndex, col.updated + 1).setValue(now);

  appendDroneHistory({
    droneId: droneId,
    droneName: droneName,
    fromStatus: fromStatus,
    toStatus: status,
    reason: reason || notes || "",
    changedBy: "Operator"
  });

  return { success: true };
}

// ─────────────────────────────────────────────
// TECHNICIAN: GET FAILS GROUPED BY DATE
// ─────────────────────────────────────────────
function getFailsByDate() {
  const { data, col } = getSheetData();
  if (data.length < 2) return { dates: [] };

  const dateMap = {};

  data.slice(1).forEach((row, i) => {
    const status = col.status >= 0 ? String(row[col.status]).trim() : "";
    const updatedRaw = col.updated >= 0 ? row[col.updated] : null;
    if (!updatedRaw || !status) return;

    if (status === "Fail" || status === "Work In Progress") {
      // include
    } else if (status === "Good" && col.fixed_at >= 0 && row[col.fixed_at]) {
      // include good only when fixed
    } else {
      return;
    }

    const dateKey = toDateKey(updatedRaw);
    if (!dateKey) return;

    if (!dateMap[dateKey]) {
      dateMap[dateKey] = { fail: 0, wip: 0, good: 0, drones: [] };
    }

    if (status === "Fail") dateMap[dateKey].fail++;
    else if (status === "Work In Progress") dateMap[dateKey].wip++;
    else if (status === "Good") dateMap[dateKey].good++;

    dateMap[dateKey].drones.push({
      rowIndex: i + 2,
      id:       col.id >= 0 ? String(row[col.id]).trim() : `D${String(i+1).padStart(2,"0")}`,
      name:     col.name >= 0 ? String(row[col.name]).trim() : String(row[0]),
      status,
      reason:   col.reason >= 0 ? String(row[col.reason]).trim() : "",
      updated:  formatDate(updatedRaw),
      fixed_at: col.fixed_at >= 0 ? formatDate(row[col.fixed_at]) : "",
    });
  });

  const dates = Object.entries(dateMap)
    .sort((a, b) => b[0].localeCompare(a[0]))
    .map(([date, val]) => ({
      date,
      displayDate: formatDisplayDate(date),
      fail: val.fail,
      wip: val.wip,
      good: val.good,
      drones: val.drones.sort((a, b) => a.id.localeCompare(b.id))
    }));

  const totalFail = dates.reduce((s, d) => s + d.fail, 0);
  const totalWip  = dates.reduce((s, d) => s + d.wip, 0);

  return { dates, totalFail, totalWip };
}

// ─────────────────────────────────────────────
// TECHNICIAN: BATCH UPDATE MULTIPLE DRONES
// ─────────────────────────────────────────────
function batchUpdate(droneIdsJson, status) {
  const droneIds = JSON.parse(droneIdsJson || "[]");
  if (!droneIds.length) return { error: "No drone IDs provided" };

  const { sheet, data, col } = getSheetData();
  let updated = 0;
  const errors = [];

  droneIds.forEach(droneId => {
    const rowIndex = findDroneRow(data, col, droneId);
    if (rowIndex === -1) {
      errors.push(`Not found: ${droneId}`);
      return;
    }

    const fromStatus = col.status >= 0 ? String(sheet.getRange(rowIndex, col.status + 1).getValue()).trim() : "";
    const droneName  = col.name   >= 0 ? String(sheet.getRange(rowIndex, col.name + 1).getValue()).trim() : String(sheet.getRange(rowIndex, 1).getValue()).trim();

    if (col.status >= 0) sheet.getRange(rowIndex, col.status + 1).setValue(status);
    if (col.fixed_at >= 0) sheet.getRange(rowIndex, col.fixed_at + 1).setValue(new Date());

    appendDroneHistory({
      droneId: droneId,
      droneName: droneName,
      fromStatus: fromStatus,
      toStatus: status,
      reason: "",
      changedBy: "Technician"
    });

    updated++;
  });

  return { success: true, updated, errors };
}

// ─────────────────────────────────────────────
// AUTO-TRANSITION: YESTERDAY'S FAIL -> WIP
// ─────────────────────────────────────────────
function autoTransition() {
  const { sheet, data, col } = getSheetData();
  const now = new Date();
  const tz = getSpreadsheetTimeZone();

  const y = new Date(now);
  y.setDate(y.getDate() - 1);
  const yesterday = Utilities.formatDate(y, tz, "yyyy-MM-dd");

  let transitioned = 0;
  const wip = "Work In Progress";

  data.slice(1).forEach((row, i) => {
    const status = col.status >= 0 ? String(row[col.status]).trim() : "";
    const updatedRaw = col.updated >= 0 ? row[col.updated] : null;
    if (status !== "Fail" || !updatedRaw) return;

    const rowDate = Utilities.formatDate(new Date(updatedRaw), tz, "yyyy-MM-dd");

    if (rowDate === yesterday) {
      const rowIndex = i + 2;
      const droneId   = col.id   >= 0 ? String(row[col.id]).trim()   : `D${String(i+1).padStart(2,"0")}`;
      const droneName = col.name >= 0 ? String(row[col.name]).trim() : String(row[0]).trim();
      const fromStatus = status;

      if (col.status >= 0) sheet.getRange(rowIndex, col.status + 1).setValue(wip);
      if (col.fixed_at >= 0) sheet.getRange(rowIndex, col.fixed_at + 1).setValue(now);

      appendDroneHistory({
        droneId: droneId,
        droneName: droneName,
        fromStatus: fromStatus,
        toStatus: wip,
        reason: "Auto transition at end of day",
        changedBy: "auto-transition"
      });

      transitioned++;
    }
  });

  return { success: true, transitioned, yesterday };
}
// email trigger

// ═══════════════════════════════════════════════════════════════
// DAILY EMAIL REPORT — 9:00 PM trigger
// ═══════════════════════════════════════════════════════════════

// const EMAIL_TO  = "ajay7836899826@gmail.com,Krtripathi.ashish@gmail.com,francisjaladi13@gmail.com";
// const EMAIL_CC  = "silasbanala@gmail.com";

// ── MAIN TRIGGER FUNCTION — set this as 9 PM daily trigger ──
function sendDailyReport() {
  const today      = toDateKey(new Date());
  const displayDay = formatDisplayDate(today);
  const failedToday = getTodayFailed(today);

 if (failedToday.length === 0) {
  Logger.log(`No failed drones found for ${displayDay}. No email will be sent.`);
  return;
}

  // Build PDF and send report
  const pdf = buildReportPdf(failedToday, displayDay);
  sendReportEmail(failedToday, displayDay, pdf);
}

// ── GET TODAY'S FAILED DRONES ──
function getTodayFailed(todayKey) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const rawHeaders = data[0];
  Logger.log("Sheet headers: " + JSON.stringify(rawHeaders));
  const headers = rawHeaders.map(h => String(h).toLowerCase().trim());

  // Match exact header names from your sheet
  // Falls back to findCol logic if exact match not found
  function getCol(exact, fallbacks) {
    let idx = headers.indexOf(exact);
    if (idx >= 0) return idx;
    for (const f of fallbacks) {
      idx = headers.findIndex(h => h.includes(f));
      if (idx >= 0) return idx;
    }
    return -1;
  }

  const idCol      = getCol("droneid",    ["drone id","id"]);
  const nameCol    = getCol("drone name", ["dronename","name"]);
  const statusCol  = getCol("status",     ["status"]);
  const reasonCol  = getCol("reason",     ["fail reason","reason"]);
  const updatedCol = getCol("updated",    ["last updated","timestamp"]);

  Logger.log("Cols → id:" + idCol + " name:" + nameCol + " status:" + statusCol + " reason:" + reasonCol + " updated:" + updatedCol);

  const failed = [];
  data.slice(1).forEach((row, i) => {
    const status     = statusCol  >= 0 ? String(row[statusCol]).trim()  : "";
    const updatedRaw = updatedCol >= 0 ? row[updatedCol]                : null;
    if (status !== "Fail" || !updatedRaw) return;
    if (toDateKey(updatedRaw) !== todayKey) return;
    const id     = idCol     >= 0 ? String(row[idCol]).trim()     : ("D" + String(i+1).padStart(2,"0"));
    const name   = nameCol   >= 0 ? String(row[nameCol]).trim()   : String(row[0]);
    const reason = reasonCol >= 0 ? String(row[reasonCol]).trim() : "—";
    const time   = formatDate(updatedRaw);
    Logger.log("Drone → id:" + id + " name:" + name + " reason:" + reason);
    failed.push({ id, name, reason, time, status });
  });

  Logger.log("Total failed today: " + failed.length);
  return failed;
}

// ── BUILD PDF AS GOOGLE DOC → EXPORT ──
function buildReportPdf(drones, displayDay) {
  // Build HTML table that will become the PDF
  const rows = drones.map((d, i) => `
    <tr style="background:${i%2===0?'#ffffff':'#fafafa'}">
      <td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:600;color:#1a1a1a">${d.id}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#333">${d.name}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee">
        <span style="background:#fdf0f0;color:#e24b4a;padding:2px 8px;border-radius:12px;font-size:12px;font-weight:600">Fail</span>
      </td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#555">${d.reason}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#888;font-size:13px">${d.time}</td>
    </tr>`).join('');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8"/>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; color: #1a1a1a; }
    .header { border-bottom: 3px solid #e24b4a; padding-bottom: 16px; margin-bottom: 24px; }
    .logo-row { display: flex; align-items: center; gap: 12px; margin-bottom: 8px; }
    .logo-box { width: 40px; height: 40px; background: #1a1a2e; border-radius: 10px;
                display: flex; align-items: center; justify-content: center;
                font-size: 20px; }
    h1 { font-size: 22px; color: #1a1a2e; margin: 0; }
    .subtitle { font-size: 14px; color: #666; margin-top: 4px; }
    .summary-box { background: #fff8f8; border: 1px solid #f5c6c6; border-radius: 10px;
                   padding: 16px 20px; margin-bottom: 24px; display: flex; gap: 40px; }
    .summary-item .num { font-size: 28px; font-weight: 800; color: #e24b4a; line-height: 1; }
    .summary-item .lbl { font-size: 12px; color: #666; text-transform: uppercase;
                          letter-spacing: .05em; margin-top: 4px; }
    table { width: 100%; border-collapse: collapse; }
    thead tr { background: #1a1a2e; }
    thead th { padding: 10px 12px; text-align: left; color: white; font-size: 13px;
               font-weight: 600; letter-spacing: .04em; }
    .footer { margin-top: 32px; padding-top: 16px; border-top: 1px solid #eee;
              font-size: 12px; color: #999; }
    .generated { font-size: 11px; color: #bbb; margin-top: 8px; }
  </style>
</head>
<body>
  <div class="header">
    <div class="logo-row">
      <div class="logo-box">🚁</div>
      <div>
        <h1>Daily Drone Failure Report</h1>
        <div class="subtitle">Report Date: ${displayDay}</div>
      </div>
    </div>
  </div>

  <div class="summary-box">
    <div class="summary-item">
      <div class="num">${drones.length}</div>
      <div class="lbl">Total Failed Drones</div>
    </div>
    <div class="summary-item">
      <div class="num">${displayDay}</div>
      <div class="lbl">Report Date</div>
    </div>
  </div>

  <table>
    <thead>
      <tr>
        <th>Drone ID</th>
        <th>Drone Name</th>
        <th>Status</th>
        <th>Failure Reason</th>
        <th>Inspection Time</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>

  <div class="footer">
    <strong>Drone Fleet Management System</strong><br/>
    This is an automated report generated at 9:00 PM.
    <div class="generated">Generated: ${new Date().toLocaleString("en-IN")}</div>
  </div>
</body>
</html>`;

  // Create a temporary Google Doc, export as PDF, then delete it
  const docName  = `DroneReport_${toDateKey(new Date())}`;
  const doc      = DocumentApp.create(docName);
  const body     = doc.getBody();

  // Write HTML-like content into the doc (Apps Script can't render HTML directly in Doc)
  // Instead we build the doc programmatically for clean PDF output
  body.clear();

  // Title
  const title = body.appendParagraph("🚨 Daily Drone Failure Report");
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.editAsText().setForegroundColor("#1a1a2e");

  body.appendParagraph(`Report Date: ${displayDay}`)
    .editAsText().setForegroundColor("#666666");

  body.appendParagraph(""); // spacer

  // Summary
  const summaryPara = body.appendParagraph(`Total Failed Drones: ${drones.length}   |   Date: ${displayDay}`);
  summaryPara.editAsText().setBold(true).setForegroundColor("#e24b4a");
  body.appendParagraph("");

  // Table header
  const table = body.appendTable();
  const headerRow = table.appendTableRow();
  ["Drone ID","Drone Name","Status","Failure Reason","Inspection Time"].forEach(h => {
    const cell = headerRow.appendTableCell(h);
    cell.setBackgroundColor("#1a1a2e");
    cell.editAsText().setBold(true).setForegroundColor("#ffffff").setFontSize(11);
  });

  // Table rows — explicitly set black text to avoid white-on-white bug
  drones.forEach((d, i) => {
    const row = table.appendTableRow();
    [d.id || "—", d.name || "—", "Fail", d.reason || "No reason given", d.time || "—"].forEach(val => {
      const cell = row.appendTableCell(val);
      cell.editAsText().setForegroundColor("#000000").setBold(false);
      if (i % 2 !== 0) cell.setBackgroundColor("#f5f5f5");
      else cell.setBackgroundColor("#ffffff");
    });
  });

  // Footer
  body.appendParagraph("");
  body.appendParagraph("Drone Fleet Management System")
    .editAsText().setItalic(true).setForegroundColor("#999999");
  body.appendParagraph("This is an automated report generated at 9:00 PM.")
    .editAsText().setForegroundColor("#bbbbbb");

  doc.saveAndClose();

  // Export as PDF blob
  const docId  = doc.getId();
  const pdfBlob = DriveApp.getFileById(docId)
    .getAs("application/pdf")
    .setName(`Drone_Failure_Report_${toDateKey(new Date())}.pdf`);

  // Delete the temp doc
  DriveApp.getFileById(docId).setTrashed(true);

  return pdfBlob;
}

// ── SEND REPORT EMAIL (with failures) ──
function sendReportEmail(drones, displayDay, pdfBlob) {
  const subject = `🚨 Daily Drone Failure Report – ${displayDay}`;
  const body = `Hi Team,

Please find attached today's drone failure report.

Summary:
• Total Failed Drones: ${drones.length}
• Report Date: ${displayDay}

Regards,
Drone Fleet Management System`;

  MailApp.sendEmail({
    to:          EMAIL_TO,
    cc:          EMAIL_CC,
    subject:     subject,
    body:        body,
    attachments: [pdfBlob],
  });

  Logger.log(`✅ Report sent for ${displayDay} — ${drones.length} failed drones`);
}

// ── SEND NO-FAILURE EMAIL (clean day) ──
function sendNoFailureEmail(displayDay) {
  const subject = `✅ Daily Drone Report – ${displayDay} — No Failures`;
  const body = `Hi Team,

No drones were failed in today's inspections.

• Total Failed Drones: 0
• Report Date: ${displayDay}

All systems are operational. ✅

Regards,
Drone Fleet Management System`;

  MailApp.sendEmail({
    to:      EMAIL_TO,
    cc:      EMAIL_CC,
    subject: subject,
    body:    body,
  });

  Logger.log(`✅ No-failure email sent for ${displayDay}`);
}

// ── HOW TO SET 9 PM TRIGGER ──
// 1. Apps Script editor → Triggers (clock icon)
// 2. + Add Trigger
// 3. Function: sendDailyReport
// 4. Time-based → Day timer → 9 PM to 10 PM
// 5. Save

// ── DEBUG FUNCTION — run this manually to see exact sheet data ──
function debugSheetData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log("Sheet not found: " + SHEET_NAME); return; }

  const data = sheet.getDataRange().getValues();
  Logger.log("Total rows: " + data.length);
  Logger.log("Headers (row 1): " + JSON.stringify(data[0]));

  // Show first 5 data rows
  for (let i = 1; i <= Math.min(5, data.length - 1); i++) {
    Logger.log("Row " + (i+1) + ": " + JSON.stringify(data[i]));
  }

  // Show today's date key
  const today = toDateKey(new Date());
  Logger.log("Today key: " + today);

  // Check each row's updated value and date key
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const statusCol  = headers.indexOf("status");
  const updatedCol = headers.indexOf("updated");
  Logger.log("statusCol index: " + statusCol + "  updatedCol index: " + updatedCol);

  data.slice(1).forEach((row, i) => {
    const status     = statusCol  >= 0 ? String(row[statusCol])  : "N/A";
    const updatedRaw = updatedCol >= 0 ? row[updatedCol]         : "N/A";
    const dateKey    = toDateKey(updatedRaw);
    Logger.log("Row " + (i+2) + " → status: [" + status + "] updated: [" + updatedRaw + "] dateKey: [" + dateKey + "] matchestoday: " + (dateKey === today));
  });
}

// ── TEST FUNCTION — run this to test with yesterday's data ──
// Use this when you want to test the email without waiting for today's failures
function testSendReportYesterday() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const dateKey    = toDateKey(yesterday);
  const displayDay = formatDisplayDate(dateKey);

  Logger.log("Testing with date: " + dateKey);

  // Temporarily override getTodayFailed to use yesterday
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());

  const idCol      = headers.indexOf("droneid");
  const nameCol    = headers.indexOf("drone name");
  const statusCol  = headers.indexOf("status");
  const reasonCol  = headers.indexOf("reason");
  const updatedCol = headers.indexOf("updated");

  const failed = [];
  data.slice(1).forEach((row, i) => {
    const status     = statusCol  >= 0 ? String(row[statusCol]).trim()  : "";
    const updatedRaw = updatedCol >= 0 ? row[updatedCol]                : null;
    if (status !== "Fail" || !updatedRaw) return;
    if (toDateKey(updatedRaw) !== dateKey) return;
    failed.push({
      id:     idCol     >= 0 ? String(row[idCol]).trim()     : "—",
      name:   nameCol   >= 0 ? String(row[nameCol]).trim()   : "—",
      reason: reasonCol >= 0 ? String(row[reasonCol]).trim() || "No reason given" : "No reason given",
      time:   formatDate(updatedRaw),
      status,
    });
  });

  Logger.log("Found " + failed.length + " failed drones for " + displayDay);
  failed.forEach(d => Logger.log("  → " + d.id + " | " + d.name + " | " + d.reason));

  if (failed.length === 0) {
  Logger.log("No failed drones found for yesterday. No email will be sent.");
  return;
}

  const pdf = buildReportPdf(failed, displayDay);
  sendReportEmail(failed, displayDay, pdf);
  Logger.log("✅ Test email sent for " + displayDay);
}


// ------------------------------------------------------------------------------------------------------

// Configuration Screen



// ─────────────────────────────────────────────
// CONFIG SCREEN — NEW FUNCTIONS
// ─────────────────────────────────────────────

// Add these to handleRequest() in your existing Code.gs:
// else if (action === "getConfigDrones")  result = getConfigDrones();
// else if (action === "createNextBatch")  result = createNextBatch(params.portId);
// else if (action === "batchConfigUpdate") result = batchConfigUpdate(params.droneIds, params.status);

// ── GET CONFIG DRONES (Unconfigured + Configured only) ──
function getConfigDrones() {
  const { data, col } = getSheetData();
  if (data.length < 2) return { drones: [], summary: { unconfigured: 0, configured: 0, total: 0 } };

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const headers = data[0].map(h => String(h).toLowerCase().trim());

  // Find port_id and configured_at columns
  const portCol    = headers.indexOf("port_id");
  const confAtCol  = headers.indexOf("configured_at");

  const drones = [];
  let unconfigured = 0, configured = 0;

  data.slice(1).forEach((row, i) => {
    const status = col.status >= 0 ? String(row[col.status]).trim() : "";
    if (status !== "Unconfigured" && status !== "Configured") return;

    if (status === "Unconfigured") unconfigured++;
    if (status === "Configured")   configured++;

    drones.push({
      rowIndex:      i + 2,
      id:            col.id     >= 0 ? String(row[col.id]).trim()     : `D${String(i+1).padStart(2,"0")}`,
      name:          col.name   >= 0 ? String(row[col.name]).trim()   : String(row[0]),
      status,
      port_id:       portCol   >= 0 ? String(row[portCol]).trim()    : "",
      configured_at: confAtCol >= 0 ? formatDate(row[confAtCol])     : "",
    });
  });

  // Sort: Unconfigured first, then Configured; within each group sort by id
  drones.sort((a, b) => {
    if (a.status !== b.status) return a.status === "Unconfigured" ? -1 : 1;
    return a.id.localeCompare(b.id);
  });

  // Find next batch range
  const allIds = data.slice(1)
    .map(row => col.id >= 0 ? String(row[col.id]).trim() : "")
    .filter(id => /^DA\d+$/i.test(id))
    .map(id => parseInt(id.replace(/^DA/i, ""), 10))
    .filter(n => !isNaN(n));

  const maxNum      = allIds.length ? Math.max(...allIds) : 0;
  const nextStart   = maxNum + 1;
  const nextEnd     = maxNum + 200;
  const canCreate   = unconfigured === 0 && (configured + unconfigured) > 0 || (configured + unconfigured) === 0;

  return {
    drones,
    summary: { unconfigured, configured, total: unconfigured + configured },
    nextBatch: { start: nextStart, end: nextEnd, startId: `DA${String(nextStart).padStart(3,"0")}`, endId: `DA${String(nextEnd).padStart(3,"0")}` },
    canCreate
  };
}

// ── CREATE NEXT BATCH OF 200 DRONES ──
function createNextBatch(portId) {
  if (!portId || !portId.trim()) return { error: "Port ID is required" };

  const { sheet, data, col } = getSheetData();
  const headers = data[0].map(h => String(h).toLowerCase().trim());

  const portCol    = headers.indexOf("port_id");
  const confAtCol  = headers.indexOf("configured_at");

  // Check if any Unconfigured drones exist — block if so
  const hasUnconfigured = data.slice(1).some(row => {
    const status = col.status >= 0 ? String(row[col.status]).trim() : "";
    return status === "Unconfigured";
  });

  if (hasUnconfigured) {
    return { error: "Current batch has unconfigured drones. Please configure all drones before creating a new batch." };
  }

  // Find highest DA number
  const allNums = data.slice(1)
    .map(row => col.id >= 0 ? String(row[col.id]).trim() : "")
    .filter(id => /^DA\d+$/i.test(id))
    .map(id => parseInt(id.replace(/^DA/i, ""), 10))
    .filter(n => !isNaN(n));

  const maxNum    = allNums.length ? Math.max(...allNums) : 0;
  const startNum  = maxNum + 1;
  const endNum    = maxNum + 200;

  // Build 200 new rows
  const newRows = [];
  for (let n = startNum; n <= endNum; n++) {
    const droneId   = `DA${String(n).padStart(3, "0")}`;
    const droneName = `D-${String(n).padStart(4, "0")}`;
    const row = new Array(data[0].length).fill("");

    if (col.id     >= 0) row[col.id]     = droneId;
    if (col.name   >= 0) row[col.name]   = droneName;
    if (col.status >= 0) row[col.status] = "Unconfigured";
    if (portCol    >= 0) row[portCol]    = portId.trim();
    if (confAtCol  >= 0) row[confAtCol]  = "";

    newRows.push(row);
  }

  // Batch append all 200 rows at once (fast)
  if (newRows.length > 0) {
    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      newRows.length,
      data[0].length
    ).setValues(newRows);
  }

  return {
    success: true,
    created: newRows.length,
    startId: `DA${String(startNum).padStart(3,"0")}`,
    endId:   `DA${String(endNum).padStart(3,"0")}`,
    portId:  portId.trim()
  };
}

// ── BATCH CONFIG UPDATE (Unconfigured ↔ Configured) ──
function batchConfigUpdate(droneIdsJson, status) {
  if (status !== "Configured" && status !== "Unconfigured") {
    return { error: "Invalid status. Must be Configured or Unconfigured." };
  }

  const droneIds = JSON.parse(droneIdsJson || "[]");
  if (!droneIds.length) return { error: "No drone IDs provided" };

  const { sheet, data, col } = getSheetData();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const confAtCol = headers.indexOf("configured_at");

  const now = new Date();
  let updated = 0;
  const errors = [];

  droneIds.forEach(droneId => {
    const rowIndex = findDroneRow(data, col, droneId);
    if (rowIndex === -1) { errors.push(`Not found: ${droneId}`); return; }

    const fromStatus = col.status >= 0 ? String(sheet.getRange(rowIndex, col.status + 1).getValue()).trim() : "";
    const droneName  = col.name   >= 0 ? String(sheet.getRange(rowIndex, col.name + 1).getValue()).trim()   : "";

    if (col.status >= 0) sheet.getRange(rowIndex, col.status + 1).setValue(status);
    if (confAtCol  >= 0) sheet.getRange(rowIndex, confAtCol + 1).setValue(status === "Configured" ? now : "");

    // Log to DroneHistory
    appendDroneHistory({
      droneId,
      droneName,
      fromStatus,
      toStatus: status,
      reason: "",
      changedBy: "Operator"
    });

    updated++;
  });

  return { success: true, updated, errors };
}
