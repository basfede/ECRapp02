// ════════════════════════════════════════════════════════════
//  ECRApp — Google Apps Script Backend
// ════════════════════════════════════════════════════════════

const SS_ID = "1Qqp_udakFBVPu_yV-6dj9rPPQb6dhFi9fTCeeCFZpts";

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = {
      usuarios:      ["id","name","colorBg","colorFg"],
      contactos:     ["id","name","address","phone","notes","userId","createdAt"],
      interacciones: ["id","contactId","userId","type","result","reason","notes","date"],
      recordatorios: ["id","userId","title","notes","date","done","createdAt"],
    };
    if (headers[name]) sheet.appendRow(headers[name]);
  }
  return sheet;
}

function uid() { return Utilities.getUuid(); }

function sheetToArray(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ""; });
    return obj;
  });
}

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch(action) {
      case "getUsers":           result = getUsers(); break;
      case "addUser":            result = addUser(e.parameter); break;
      case "deleteUser":         result = deleteRow("usuarios", e.parameter.id); break;
      case "getContacts":        result = getContacts(); break;
      case "addContact":         result = addContact(e.parameter); break;
      case "updateContact":      result = updateContact(e.parameter); break;
      case "deleteContact":      result = deleteRow("contactos", e.parameter.id); break;
      case "getInteractions":    result = getInteractions(); break;
      case "addInteraction":     result = addInteraction(e.parameter); break;
      case "updateInteraction":  result = updateInteraction(e.parameter); break;
      case "deleteInteraction":  result = deleteRow("interacciones", e.parameter.id); break;
      case "getReminders":       result = getReminders(); break;
      case "addReminder":        result = addReminder(e.parameter); break;
      case "updateReminder":     result = updateReminder(e.parameter); break;
      case "deleteReminder":     result = deleteRow("recordatorios", e.parameter.id); break;
      default: result = { ok:false, error:"Acción desconocida: "+action };
    }
  } catch(err) { result = { ok:false, error:err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ── UTIL ─────────────────────────────────────────────────────
function deleteRow(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) { sheet.deleteRow(i+1); break; }
  }
  return { ok:true };
}

// ── USUARIOS ─────────────────────────────────────────────────
function getUsers() { return { ok:true, data:sheetToArray(getSheet("usuarios")) }; }
function addUser(p) {
  const id = uid();
  getSheet("usuarios").appendRow([id, p.name, p.colorBg||"#c8f04a", p.colorFg||"#0e0f13"]);
  return { ok:true, id };
}

// ── CONTACTOS ─────────────────────────────────────────────────
function getContacts() { return { ok:true, data:sheetToArray(getSheet("contactos")) }; }
function addContact(p) {
  const id = uid(), now = new Date().toISOString().split("T")[0];
  getSheet("contactos").appendRow([id, p.name, p.address||"", p.phone||"", p.notes||"", p.userId||"", now]);
  return { ok:true, id };
}
function updateContact(p) {
  const sheet = getSheet("contactos");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === p.id) {
      if(p.name    !== undefined) sheet.getRange(i+1,2).setValue(p.name);
      if(p.address !== undefined) sheet.getRange(i+1,3).setValue(p.address);
      if(p.phone   !== undefined) sheet.getRange(i+1,4).setValue(p.phone);
      if(p.notes   !== undefined) sheet.getRange(i+1,5).setValue(p.notes);
      if(p.userId  !== undefined) sheet.getRange(i+1,6).setValue(p.userId);
      break;
    }
  }
  return { ok:true };
}

// ── INTERACCIONES ─────────────────────────────────────────────
function getInteractions() { return { ok:true, data:sheetToArray(getSheet("interacciones")) }; }
function addInteraction(p) {
  const id = uid();
  getSheet("interacciones").appendRow([id, p.contactId, p.userId, p.type, p.result, p.reason||"", p.notes||"", p.date]);
  return { ok:true, id };
}
function updateInteraction(p) {
  const sheet = getSheet("interacciones");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === p.id) {
      if(p.type   !== undefined) sheet.getRange(i+1,4).setValue(p.type);
      if(p.result !== undefined) sheet.getRange(i+1,5).setValue(p.result);
      if(p.reason !== undefined) sheet.getRange(i+1,6).setValue(p.reason);
      if(p.notes  !== undefined) sheet.getRange(i+1,7).setValue(p.notes);
      break;
    }
  }
  return { ok:true };
}

// ── RECORDATORIOS ─────────────────────────────────────────────
function getReminders() { return { ok:true, data:sheetToArray(getSheet("recordatorios")) }; }
function addReminder(p) {
  const id = uid(), now = new Date().toISOString().split("T")[0];
  getSheet("recordatorios").appendRow([id, p.userId||"", p.title, p.notes||"", p.date||"", p.done||"false", now]);
  return { ok:true, id };
}
function updateReminder(p) {
  const sheet = getSheet("recordatorios");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === p.id) {
      if(p.title !== undefined) sheet.getRange(i+1,3).setValue(p.title);
      if(p.notes !== undefined) sheet.getRange(i+1,4).setValue(p.notes);
      if(p.date  !== undefined) sheet.getRange(i+1,5).setValue(p.date);
      if(p.done  !== undefined) sheet.getRange(i+1,6).setValue(p.done);
      break;
    }
  }
  return { ok:true };
}
