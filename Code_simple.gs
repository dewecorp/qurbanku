/**
 * QurbanKu - Backend Google Apps Script sederhana dan stabil.
 * Tempel seluruh isi file ini ke Code.gs Apps Script, menggantikan kode lama.
 */

var SPREADSHEET_ID = "";

var HEADERS = {
  Tahap: ["ID Tahap", "Nama Tahap", "Status", "UpdatedAt"],
  Kelompok: ["ID Kelompok", "Nama Kelompok", "Tipe Hewan", "Target Total", "ID Tahap", "UpdatedAt"],
  Anggota: ["ID Anggota", "Nama Lengkap", "WhatsApp", "ID Kelompok", "Target Tabungan", "ID Tahap", "UpdatedAt", "PIN"],
  Setoran: ["ID Transaksi", "Tanggal", "ID Anggota", "Nominal Setor", "Pencatat", "Catatan", "ID Tahap", "Bulan", "UpdatedAt"],
  Deleted: ["ID Terhapus", "Tipe", "UpdatedAt"],
  Pengaturan: ["Kunci Pengaturan", "Nilai"]
};

function safeString(value) {
  return value === null || value === undefined ? "" : value.toString().trim();
}

function safeNumber(value, fallback) {
  var number = Number(value);
  return isNaN(number) ? (fallback || 0) : number;
}

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Spreadsheet tidak ditemukan. Jalankan script dari Extensions > Apps Script di Google Sheet, atau isi SPREADSHEET_ID.");
  }
  return ss;
}

function jsonOutput(payload, callback) {
  var text = JSON.stringify(payload);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + text + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

function setupDatabase() {
  return "Jangan jalankan setupDatabase. Buat tab dan header Google Sheet secara manual, lalu deploy Web App.";
}

function testSpreadsheetConnection() {
  try {
    var ss = getSpreadsheet();
    return "OK: Terhubung ke spreadsheet " + ss.getName() + " / " + ss.getId();
  } catch (error) {
    return "GAGAL koneksi spreadsheet: " + error.message;
  }
}

function ensureSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  var headers = HEADERS[name];
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
    return safeString(header);
  });
  for (var i = 0; i < headers.length; i++) {
    if (existing.indexOf(headers[i]) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(headers[i]);
    }
  }
  return sheet;
}

function ensureSetting(ss, key, value) {
  var sheet = ensureSheet(ss, "Pengaturan");
  var row = findRowById(sheet, key, 1, 2);
  if (row === -1) {
    sheet.appendRow([key, value]);
  }
}

function doGet(e) {
  var callback = e && e.parameter ? e.parameter.callback : "";
  try {
    var ss = getSpreadsheet();

    if (e.parameter && (e.parameter.action === "sync" || e.parameter.action === "write") && e.parameter.data) {
      var raw = e.parameter.data;
      if (raw.indexOf("%7B") === 0 || raw.indexOf("%5B") === 0) {
        raw = decodeURIComponent(raw);
      }
      saveRequest(ss, JSON.parse(raw));
    }

    if (e.parameter && e.parameter.test === "ping") {
      return jsonOutput({ status: "success", message: "Koneksi Google Sheet OK" }, callback);
    }

    return jsonOutput(buildResponse(ss), callback);
  } catch (error) {
    return jsonOutput({ status: "error", message: error.message }, callback);
  }
}

function doPost(e) {
  try {
    var ss = getSpreadsheet();
    var request = JSON.parse(e.postData.contents || "{}");
    if (request.testPing) {
      return jsonOutput({ status: "success", message: "Tes tulis OK" });
    }
    saveRequest(ss, request);
    return jsonOutput(buildResponse(ss));
  } catch (error) {
    return jsonOutput({ status: "error", message: error.message });
  }
}

function saveRequest(ss, request) {
  request = request || {};
  var action = safeString(request.action);
  var data = request.data || {};
  var now = Date.now().toString();
  var activeTahapId = getActiveTahapId(ss);

  if (action === "saveTahap") {
    upsertRow(ss.getSheetByName("Tahap"), safeString(data.id) || generateId("t"), [
      safeString(data.id) || generateId("t"),
      safeString(data.nama || data.name),
      safeString(data.status || "Aktif"),
      now
    ], 1);
  }

  if (action === "saveKelompok") {
    var groupId = safeString(data.id) || generateId("g");
    upsertRow(ss.getSheetByName("Kelompok"), groupId, [
      groupId,
      safeString(data.nama || data.name),
      safeString(data.tipe || data.type || "Sapi"),
      safeNumber(data.target || data.targetTotal, 0),
      safeString(data.tahapId) || activeTahapId,
      now
    ], 1);
  }

  if (action === "saveAnggota") {
    var participantId = safeString(data.id) || generateId("p");
    upsertRow(ss.getSheetByName("Anggota"), participantId, [
      participantId,
      safeString(data.nama || data.name),
      safeString(data.whatsapp),
      safeString(data.kelompok_id || data.groupId),
      safeNumber(data.target_saving || data.targetAmount, 3000000),
      safeString(data.tahapId) || activeTahapId,
      now,
      safeString(data.pin || "1234")
    ], 1);
  }

  if (action === "saveSetoran") {
    var depositId = safeString(data.id) || generateId("d");
    upsertRow(ss.getSheetByName("Setoran"), depositId, [
      depositId,
      safeString(data.tanggal || data.date),
      safeString(data.anggota_id || data.participantId),
      safeNumber(data.jumlah || data.amount, 0),
      safeString(data.recorded_by || data.recordedBy),
      safeString(data.catatan || data.note),
      safeString(data.tahapId) || activeTahapId,
      safeNumber(data.bulan, 1),
      now
    ], 1);
  }

  if (action === "saveCloudUrl") {
    upsertSetting(ss, "cloudUrl", safeString(data.cloudUrl || request.settings && request.settings.cloudUrl));
  }

  if (action === "saveSettings") {
    upsertSetting(ss, "ketuaName", safeString(data.ketuaName || request.settings && request.settings.ketuaName));
    upsertSetting(ss, "bendaharaName", safeString(data.bendaharaName || request.settings && request.settings.bendaharaName));
    upsertSetting(ss, "monthlyInstallment", safeNumber(data.monthlyInstallment || request.settings && request.settings.monthlyInstallment, 150000).toString());
  }

  if (action === "deleteTahap") deleteRow(ss.getSheetByName("Tahap"), data.id);
  if (action === "deleteKelompok") deleteRow(ss.getSheetByName("Kelompok"), data.id);
  if (action === "deleteAnggota") deleteRow(ss.getSheetByName("Anggota"), data.id);
  if (action === "deleteSetoran") deleteRow(ss.getSheetByName("Setoran"), data.id);
}

function upsertRow(sheet, id, values, idColumn) {
  var row = findRowById(sheet, id, idColumn || 1, 2);
  if (row === -1) {
    sheet.appendRow(values);
  } else {
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
  }
}

function deleteRow(sheet, id) {
  var row = findRowById(sheet, safeString(id), 1, 2);
  if (row !== -1) {
    sheet.deleteRow(row);
  }
}

function findRowById(sheet, id, column, startRow) {
  id = safeString(id);
  if (!id) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return -1;
  var values = sheet.getRange(startRow, column, lastRow - startRow + 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (safeString(values[i][0]) === id) return startRow + i;
  }
  return -1;
}

function generateId(prefix) {
  return prefix + "-" + Date.now() + "-" + Math.floor(Math.random() * 100000);
}

function getActiveTahapId(ss) {
  var tahaps = parseTahap(ss.getSheetByName("Tahap"));
  for (var i = 0; i < tahaps.length; i++) {
    if (tahaps[i].status === "Aktif") return tahaps[i].id;
  }
  return tahaps.length ? tahaps[0].id : "t-1";
}

function buildResponse(ss) {
  var settings = parsePengaturan(ss.getSheetByName("Pengaturan"));
  var raw = {
    status: "success",
    tahaps: parseTahap(ss.getSheetByName("Tahap")),
    groups: parseKelompok(ss.getSheetByName("Kelompok")),
    participants: parseAnggota(ss.getSheetByName("Anggota")),
    deposits: parseSetoran(ss.getSheetByName("Setoran")),
    deletedRecords: [],
    settings: settings,
    lastUpdated: settings.lastUpdated || Date.now()
  };
  raw.data = toClientData(raw);
  return raw;
}

function toClientData(raw) {
  return {
    tahap: raw.tahaps.map(function(t) { return { id: t.id, nama: t.name, status: t.status, updatedAt: t.updatedAt }; }),
    kelompok: raw.groups.map(function(g) { return { id: g.id, nama: g.name, tipe: g.type, target: g.targetTotal, tahapId: g.tahapId, updatedAt: g.updatedAt }; }),
    anggota: raw.participants.map(function(p) { return { id: p.id, nama: p.name, whatsapp: p.whatsapp, pin: p.pin, kelompok_id: p.groupId, target_saving: p.targetAmount, tahapId: p.tahapId, updatedAt: p.updatedAt }; }),
    setoran: raw.deposits.map(function(d) { return { id: d.id, tanggal: d.date, anggota_id: d.participantId, jumlah: d.amount, recorded_by: d.recordedBy, catatan: d.note, tahapId: d.tahapId, bulan: d.bulan, updatedAt: d.updatedAt }; })
  };
}

function parseTahap(sheet) {
  var rows = getRows(sheet);
  return rows.map(function(r) { return { id: safeString(r[0]), name: safeString(r[1]), status: safeString(r[2]), updatedAt: safeNumber(r[3], 0) }; });
}

function parseKelompok(sheet) {
  var rows = getRows(sheet);
  return rows.map(function(r) { return { id: safeString(r[0]), name: safeString(r[1]), type: safeString(r[2]), targetTotal: safeNumber(r[3], 0), tahapId: safeString(r[4]), updatedAt: safeNumber(r[5], 0) }; });
}

function parseAnggota(sheet) {
  var rows = getRows(sheet);
  return rows.map(function(r) { return { id: safeString(r[0]), name: safeString(r[1]), whatsapp: safeString(r[2]), groupId: safeString(r[3]), targetAmount: safeNumber(r[4], 3000000), tahapId: safeString(r[5]), updatedAt: safeNumber(r[6], 0), pin: safeString(r[7] || "1234") }; });
}

function parseSetoran(sheet) {
  var rows = getRows(sheet);
  return rows.map(function(r) {
    var dateValue = r[1];
    if (dateValue instanceof Date) {
      dateValue = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    return { id: safeString(r[0]), date: safeString(dateValue), participantId: safeString(r[2]), amount: safeNumber(r[3], 0), recordedBy: safeString(r[4]), note: safeString(r[5]), tahapId: safeString(r[6]), bulan: safeNumber(r[7], 1), updatedAt: safeNumber(r[8], 0) };
  });
}

function parsePengaturan(sheet) {
  var settings = { ketuaName: "", bendaharaName: "", monthlyInstallment: 150000, cloudUrl: "", gatheringDates: [], lastUpdated: 0 };
  var rows = getRows(sheet);
  for (var i = 0; i < rows.length; i++) {
    var key = safeString(rows[i][0]);
    var value = rows[i][1];
    if (key === "ketuaName") settings.ketuaName = safeString(value);
    if (key === "bendaharaName") settings.bendaharaName = safeString(value);
    if (key === "monthlyInstallment") settings.monthlyInstallment = safeNumber(value, 150000);
    if (key === "cloudUrl") settings.cloudUrl = safeString(value);
    if (key === "lastUpdated") settings.lastUpdated = safeNumber(value, 0);
    if (key === "gatheringDates") {
      try { settings.gatheringDates = JSON.parse(safeString(value)); } catch (error) { settings.gatheringDates = []; }
    }
  }
  return settings;
}

function getRows(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function upsertSetting(ss, key, value) {
  var sheet = ss.getSheetByName("Pengaturan");
  var row = findRowById(sheet, key, 1, 2);
  if (row === -1) {
    sheet.appendRow([key, value]);
  } else {
    sheet.getRange(row, 2).setValue(value);
  }
  upsertSettingTimestamp(sheet);
}

function upsertSettingTimestamp(sheet) {
  var row = findRowById(sheet, "lastUpdated", 1, 2);
  if (row === -1) {
    sheet.appendRow(["lastUpdated", Date.now().toString()]);
  } else {
    sheet.getRange(row, 2).setValue(Date.now().toString());
  }
}
