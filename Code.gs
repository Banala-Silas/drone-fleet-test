// ─────────────────────────────────────────────
// Drone Fleet Manager — Google Apps Script
// Paste this in Extensions > Apps Script
// Deploy as Web App (Anyone can access)
// ─────────────────────────────────────────────

const SHEET_NAME = "Drones"; // ← Change to your tab name

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = e.parameter;
  const action = params.action;
  try {
    let result;
    if (action === "getDrones") result = getDrones();
    else if (action === "updateDrone") result = updateDrone(params.droneId, params.status, params.reason || "", params.notes || "");
    else result = { error: "Unknown action" };
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getDrones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { error: `Sheet "${SHEET_NAME}" not found` };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { drones: [] };
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col = {
    id:     findCol(headers, ["id","drone id","droneid"]),
    name:   findCol(headers, ["name","drone name","dronename"]),
    group:  findCol(headers, ["group","fleet","team"]),
    status: findCol(headers, ["status"]),
    reason: findCol(headers, ["reason","fail reason","reason for fail"]),
    notes:  findCol(headers, ["notes","note","remarks"]),
  };
  const drones = data.slice(1)
    .filter(row => row[col.name >= 0 ? col.name : 0])
    .map((row, i) => ({
      rowIndex: i + 2,
      id:     col.id     >= 0 ? String(row[col.id])     : `D${String(i+1).padStart(2,"0")}`,
      name:   col.name   >= 0 ? String(row[col.name])   : String(row[0]),
      group:  col.group  >= 0 ? String(row[col.group])  : "",
      status: col.status >= 0 ? String(row[col.status]) : "Unknown",
      reason: col.reason >= 0 ? String(row[col.reason]) : "",
      notes:  col.notes  >= 0 ? String(row[col.notes])  : "",
    }));
  return { drones };
}

function updateDrone(droneId, status, reason, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { error: `Sheet "${SHEET_NAME}" not found` };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const idCol     = findCol(headers, ["id","drone id","droneid"]);
  const nameCol   = findCol(headers, ["name","drone name"]);
  const statCol   = findCol(headers, ["status"]);
  const reasonCol = findCol(headers, ["reason","fail reason","reason for fail"]);
  const notesCol  = findCol(headers, ["notes","note","remarks"]);
  const updatedCol= findCol(headers, ["updated","last updated","timestamp"]);
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    const rowId   = idCol   >= 0 ? String(data[i][idCol]).trim()   : "";
    const rowName = nameCol >= 0 ? String(data[i][nameCol]).trim() : String(data[i][0]).trim();
    if (rowId === droneId || rowName === droneId) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return { error: `Drone "${droneId}" not found` };
  if (statCol   >= 0) sheet.getRange(rowIndex, statCol + 1).setValue(status);
  if (reasonCol >= 0) sheet.getRange(rowIndex, reasonCol + 1).setValue(reason);
  if (notesCol  >= 0) sheet.getRange(rowIndex, notesCol + 1).setValue(notes);
  if (updatedCol>= 0) sheet.getRange(rowIndex, updatedCol + 1).setValue(new Date().toLocaleString());
  return { success: true };
}

function findCol(headers, candidates) {
  for (const c of candidates) {
    const idx = headers.findIndex(h => h === c || h.includes(c));
    if (idx >= 0) return idx;
  }
  return -1;
}
