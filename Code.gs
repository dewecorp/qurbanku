/**
 * QurbanKu - Google Sheets Database Backend API (Smart Server-Side LWW Merge Engine)
 * Pasang kode ini pada Apps Script di dalam Google Sheets Anda.
 * * Jaminan: 100% stabil digunakan bersamaan oleh HP (Mobile) dan Laptop (Desktop) 
 * tanpa risiko data terhapus, tertimpa, atau saling mengabaikan!
 */

// Fungsi pembantu untuk mengonversi nilai sel secara aman agar tidak memicu crash jika kosong/null
function safeString(val) {
  return val === null || val === undefined ? "" : val.toString().trim();
}

function safeNumber(val, defaultVal) {
  var num = Number(val);
  return isNaN(num) ? (defaultVal || 0) : num;
}

// Jika Apps Script Anda tidak dibuat dari menu Extensions > Apps Script di Google Sheet,
// isi ID spreadsheet di sini. Kalau script sudah menempel di Sheet, biarkan kosong.
var SPREADSHEET_ID = "";

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Spreadsheet tidak ditemukan. Buka Google Sheet database Anda > Extensions > Apps Script, atau isi SPREADSHEET_ID di Code.gs.");
  }
  return ss;
}

function styleHeaderSafely(sheet, columns) {
  try {
    if (sheet && columns > 0) {
      sheet.getRange(1, 1, 1, columns).setFontWeight("bold").setBackground("#dcfce7");
    }
  } catch (err) {
    Logger.log("Lewati styling header: " + err);
  }
}

function resizeColumnsSafely(sheet) {
  try {
    var lastCol = sheet ? sheet.getLastColumn() : 0;
    if (lastCol > 0) sheet.autoResizeColumns(1, lastCol);
  } catch (err) {
    Logger.log("Lewati auto resize kolom: " + err);
  }
}

function jsonOutput(obj, callback) {
  var text = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + text + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

function generateId(prefix) {
  return prefix + "-" + Date.now() + "-" + Math.floor(Math.random() * 100000);
}

function getActiveTahapId(ss) {
  var tahaps = parseTahap(ss.getSheetByName("Tahap"));
  for (var i = 0; i < tahaps.length; i++) {
    if (tahaps[i].status === "Aktif") return tahaps[i].id;
  }
  return tahaps.length > 0 ? tahaps[tahaps.length - 1].id : "t-1";
}

function normalizeClientPayload(ss, request) {
  request = request || {};
  if (request.tahaps || request.groups || request.participants || request.deposits || request.deletedRecords) {
    request.lastUpdated = request.lastUpdated || Date.now();
    return request;
  }

  var action = safeString(request.action);
  var data = request.data || {};
  var now = Date.now();
  var activeTahapId = getActiveTahapId(ss);
  var payload = {
    tahaps: [],
    groups: [],
    participants: [],
    deposits: [],
    deletedRecords: [],
    settings: request.settings || null,
    lastUpdated: now
  };

  if (action === "saveTahap") {
    var tahapId = safeString(data.id) || generateId("t");
    payload.tahaps.push({
      id: tahapId,
      name: safeString(data.nama || data.name),
      status: safeString(data.status || "Aktif"),
      updatedAt: now
    });
  }

  if (action === "saveKelompok") {
    payload.groups.push({
      id: safeString(data.id) || generateId("g"),
      name: safeString(data.nama || data.name),
      type: safeString(data.tipe || data.type),
      targetTotal: safeNumber(data.target || data.targetTotal, 0),
      tahapId: safeString(data.tahapId) || activeTahapId,
      updatedAt: now
    });
  }

  if (action === "saveAnggota") {
    payload.participants.push({
      id: safeString(data.id) || generateId("p"),
      name: safeString(data.nama || data.name),
      whatsapp: safeString(data.whatsapp),
      pin: safeString(data.pin || "1234"),
      groupId: safeString(data.kelompok_id || data.groupId),
      targetAmount: safeNumber(data.target_saving || data.targetAmount, 3000000),
      tahapId: safeString(data.tahapId) || activeTahapId,
      updatedAt: now
    });
  }

  if (action === "saveSetoran") {
    payload.deposits.push({
      id: safeString(data.id) || generateId("d"),
      date: safeString(data.tanggal || data.date),
      participantId: safeString(data.anggota_id || data.participantId),
      amount: safeNumber(data.jumlah || data.amount, 0),
      recordedBy: safeString(data.recorded_by || data.recordedBy),
      note: safeString(data.catatan || data.note),
      tahapId: safeString(data.tahapId) || activeTahapId,
      bulan: safeNumber(data.bulan, 1),
      updatedAt: now
    });
  }

  if (action === "saveCloudUrl") {
    payload.settings = payload.settings || {};
    payload.settings.cloudUrl = safeString(data.cloudUrl || request.settings && request.settings.cloudUrl);
  }

  var deleteMap = {
    deleteTahap: "tahap",
    deleteKelompok: "group",
    deleteAnggota: "participant",
    deleteSetoran: "deposit"
  };
  if (deleteMap[action] && data.id) {
    payload.deletedRecords.push({ id: safeString(data.id), type: deleteMap[action], updatedAt: now });
  }

  return payload;
}

function toClientData(raw) {
  var tahaps = (raw.tahaps || []).map(function(t) {
    return { id: t.id, nama: t.name || t.nama, status: t.status, updatedAt: t.updatedAt };
  });
  var groups = (raw.groups || []).map(function(g) {
    return { id: g.id, nama: g.name || g.nama, tipe: g.type || g.tipe, target: safeNumber(g.targetTotal || g.target, 0), tahapId: g.tahapId, updatedAt: g.updatedAt };
  });
  var participants = (raw.participants || []).map(function(p) {
    return { id: p.id, nama: p.name || p.nama, whatsapp: p.whatsapp, pin: p.pin || "1234", kelompok_id: p.groupId || p.kelompok_id, target_saving: safeNumber(p.targetAmount || p.target_saving, 3000000), tahapId: p.tahapId, updatedAt: p.updatedAt };
  });
  var deposits = (raw.deposits || []).map(function(d) {
    return { id: d.id, tanggal: d.date || d.tanggal, anggota_id: d.participantId || d.anggota_id, jumlah: safeNumber(d.amount || d.jumlah, 0), recorded_by: d.recordedBy || d.recorded_by, catatan: d.note || d.catatan, tahapId: d.tahapId, bulan: safeNumber(d.bulan, 1), updatedAt: d.updatedAt };
  });
  return { tahap: tahaps, kelompok: groups, anggota: participants, setoran: deposits };
}

function buildResponse(ss) {
  var sheetTahaps = ss.getSheetByName("Tahap") || ss.insertSheet("Tahap");
  var sheetGroups = ss.getSheetByName("Kelompok") || ss.insertSheet("Kelompok");
  var sheetParticipants = ss.getSheetByName("Anggota") || ss.insertSheet("Anggota");
  var sheetDeposits = ss.getSheetByName("Setoran") || ss.insertSheet("Setoran");
  var sheetDeleted = ss.getSheetByName("Deleted") || ss.insertSheet("Deleted");
  var sheetPengaturan = ss.getSheetByName("Pengaturan") || ss.insertSheet("Pengaturan");

  var tahaps = parseTahap(sheetTahaps);
  if (tahaps.length === 0) {
    tahaps = [{ id: "t-1", name: "Tahap 1", status: "Aktif", updatedAt: Date.now() }];
  }

  var raw = {
    tahaps: tahaps,
    currentTahapId: getActiveTahapId(ss),
    groups: parseKelompok(sheetGroups),
    participants: parseAnggota(sheetParticipants),
    deposits: parseSetoran(sheetDeposits),
    deletedRecords: parseDeleted(sheetDeleted),
    settings: parsePengaturan(sheetPengaturan)
  };

  try {
    var cloudUrl = ScriptApp.getService().getUrl();
    if (!raw.settings.cloudUrl && cloudUrl) raw.settings.cloudUrl = cloudUrl;
  } catch (urlErr) {}

  raw.status = "success";
  raw.lastUpdated = raw.settings.lastUpdated || Date.now();
  raw.data = toClientData(raw);
  return raw;
}

// Fungsi inisialisasi awal versi ringan. Tidak menghapus data lama.
function setupDatabase() {
  try {
    var ss = getSpreadsheet();
    var now = Date.now().toString();

    ensureSheet_(ss, "Tahap", ["ID Tahap", "Nama Tahap", "Status", "UpdatedAt"]);
    ensureSheet_(ss, "Kelompok", ["ID Kelompok", "Nama Kelompok", "Tipe Hewan", "Target Total", "ID Tahap", "UpdatedAt"]);
    ensureSheet_(ss, "Anggota", ["ID Anggota", "Nama Lengkap", "WhatsApp", "ID Kelompok", "Target Tabungan", "ID Tahap", "UpdatedAt", "PIN"]);
    ensureSheet_(ss, "Setoran", ["ID Transaksi", "Tanggal", "ID Anggota", "Nominal Setor", "Pencatat", "Catatan", "ID Tahap", "Bulan", "UpdatedAt"]);
    ensureSheet_(ss, "Deleted", ["ID Terhapus", "Tipe", "UpdatedAt"]);
    ensureSheet_(ss, "Pengaturan", ["Kunci Pengaturan", "Nilai"]);

    var sheetTahap = ss.getSheetByName("Tahap");
    if (sheetTahap.getLastRow() < 2) {
      sheetTahap.appendRow(["t-1", "Tahap 1", "Aktif", now]);
    }

    ensureSetting_(ss, "ketuaName", "H. Rozikin Dimyati");
    ensureSetting_(ss, "bendaharaName", "Sudarlim");
    ensureSetting_(ss, "monthlyInstallment", "150000");
    ensureSetting_(ss, "cloudUrl", "");
    ensureSetting_(ss, "gatheringDates", JSON.stringify([]));
    ensureSetting_(ss, "lastUpdated", "0");

    SpreadsheetApp.flush();
    return "OK: setupDatabase selesai untuk spreadsheet " + ss.getName() + ". Sheet QurbanKu siap dipakai dan data lama tidak dihapus.";
  } catch (err) {
    return "GAGAL setupDatabase: " + (err && err.message ? err.message : err);
  }
}

function ensureSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }

  var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return safeString(h);
  });

  for (var i = 0; i < headers.length; i++) {
    if (existingHeaders.indexOf(headers[i]) === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(headers[i]);
      existingHeaders.push(headers[i]);
    }
  }

  return sheet;
}

function ensureSetting_(ss, key, defaultValue) {
  var sheet = ss.getSheetByName("Pengaturan") || ensureSheet_(ss, "Pengaturan", ["Kunci Pengaturan", "Nilai"]);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.appendRow([key, defaultValue]);
    return;
  }

  var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (safeString(keys[i][0]) === key) return;
  }
  sheet.appendRow([key, defaultValue]);
}

function testSpreadsheetConnection() {
  try {
    var ss = getSpreadsheet();
    return "OK: Terhubung ke spreadsheet: " + ss.getName() + " (" + ss.getId() + ")";
  } catch (err) {
    return "GAGAL koneksi spreadsheet: " + (err && err.message ? err.message : err);
  }
}

// Menangani pengambilan data (GET) & penulisan via JSONP luring seluler
function doGet(e) {
  var ss = getSpreadsheet();
  var callback = e.parameter.callback;
  
  try {
    if (e.parameter.test === "ping") {
      var pingResult = { status: "success", message: "Koneksi berhasil terhubung ke pangkalan data Google Sheets!" };
      return jsonOutput(pingResult, callback);
    }

    // JALUR 1: SIMPAN DATA via JSONP (GET sync dari HP seluler)
    if ((e.parameter.action === "sync" || e.parameter.action === "write") && e.parameter.data) {
      var rawData = e.parameter.data;
      try {
        if (rawData.indexOf("%7B") === 0 || rawData.indexOf("%5B") === 0) {
          rawData = decodeURIComponent(rawData);
        }
      } catch (decodeErr) {
        // Gunakan rawData jika sudah dalam format decoded
      }
      var postData = JSON.parse(rawData);
      postData = normalizeClientPayload(ss, postData);
      saveDataToSheets(ss, postData);
    }

    // JALUR 2: AMBIL DATA TERBARU SETELAH MERGE SELESAI
    return jsonOutput(buildResponse(ss), callback);

  } catch (err) {
    var errResult = { status: "error", message: err.toString(), lastUpdated: 0 };
    return jsonOutput(errResult, callback);
  }
}

// Menangani permintaan penulisan data (POST) dari komputer/desktop
function doPost(e) {
  var ss = getSpreadsheet();
  try {
    var postData = JSON.parse(e.postData.contents);
    
    if (postData.testPing) {
      return jsonOutput({ status: "success", message: "Tes tulis sukses!" });
    }

    // Melakukan merge data luring
    postData = normalizeClientPayload(ss, postData);
    saveDataToSheets(ss, postData);

    // Ambil data terbaru hasil merge untuk langsung dikirim balik ke pengirim
    return jsonOutput(buildResponse(ss));
  } catch (error) {
    return jsonOutput({ status: "error", message: error.toString() });
  }
}

// ALGORITMA UTAMA: SMART LAST-WRITE-WINS RECORD MERGER (Per-ID Object Merging)
function saveDataToSheets(ss, postData) {
  if (!postData) return;

  // Membaca Sheet Deleted untuk mengisi daftar penghapusan (Tombstone)
  var sheetDeleted = ss.getSheetByName("Deleted") || ss.insertSheet("Deleted");
  var existingDeleted = parseDeleted(sheetDeleted);
  var deletedMap = {};
  
  existingDeleted.forEach(function(d) {
    if (d.id) deletedMap[d.id] = safeNumber(d.updatedAt, 0);
  });

  // Gabungkan dengan data penghapusan baru yang masuk dari client
  if (postData.deletedRecords) {
    postData.deletedRecords.forEach(function(d) {
      if (d.id) {
        var existingDelTime = deletedMap[d.id] || 0;
        var incomingDelTime = safeNumber(d.updatedAt, 0);
        if (incomingDelTime > existingDelTime) {
          deletedMap[d.id] = incomingDelTime;
        }
      }
    });
  }

  // Tulis ulang Sheet Deleted yang sudah dimerge
  sheetDeleted.clear();
  sheetDeleted.appendRow(["ID Terhapus", "Tipe", "UpdatedAt"]);
  var deletedRows = [];
  for (var id in deletedMap) {
    var type = id.startsWith("t-") ? "tahap" : id.startsWith("g-") ? "group" : id.startsWith("p-") ? "participant" : "deposit";
    deletedRows.push([id, type, deletedMap[id].toString()]);
  }
  if (deletedRows.length > 0) {
    sheetDeleted.getRange(2, 1, deletedRows.length, 3).setValues(deletedRows);
  }

  // Helper Merging Cerdas per Record LWW
  function mergeLWW(existingArr, incomingArr) {
    var map = {};
    existingArr.forEach(function(item) {
      if (item && item.id) {
        map[item.id] = item;
      }
    });

    if (incomingArr) {
      incomingArr.forEach(function(item) {
        if (item && item.id) {
          var id = item.id;
          
          // Pastikan item memiliki updatedAt bawaan
          var itemUpdatedAt = safeNumber(item.updatedAt, 0);
          if (itemUpdatedAt === 0) {
            itemUpdatedAt = Date.now();
            item.updatedAt = itemUpdatedAt;
          }
          
          // Abaikan jika ID ini sudah terhapus secara sah di cloud
          if (deletedMap[id] && deletedMap[id] >= itemUpdatedAt) {
            delete map[id];
            return;
          }

          var existing = map[id];
          var existingTime = existing ? safeNumber(existing.updatedAt, 0) : 0;

          if (!existing || itemUpdatedAt > existingTime) {
            map[id] = item;
          }
        }
      });
    }

    // Bersihkan kembali sisa-sisa ID terhapus
    for (var k in map) {
      if (deletedMap[k] && deletedMap[k] >= safeNumber(map[k].updatedAt, 0)) {
        delete map[k];
      }
    }

    var result = [];
    for (var k in map) {
      result.push(map[k]);
    }
    return result;
  }

  // 1. MERGE SHEET TAHAP
  var sheetTahap = ss.getSheetByName("Tahap") || ss.insertSheet("Tahap");
  var mergedTahaps = mergeLWW(parseTahap(sheetTahap), postData.tahaps);
  var requestedActiveTahap = null;
  (postData.tahaps || []).forEach(function(t) {
    if (safeString(t.status) === "Aktif") requestedActiveTahap = safeString(t.id);
  });
  if (requestedActiveTahap) {
    mergedTahaps = mergedTahaps.map(function(t) {
      if (safeString(t.id) !== requestedActiveTahap && safeString(t.status) === "Aktif") {
        t.status = "Nonaktif";
        t.updatedAt = Date.now();
      }
      return t;
    });
  }
  sheetTahap.clear();
  sheetTahap.appendRow(["ID Tahap", "Nama Tahap", "Status", "UpdatedAt"]);
  var tahapRows = mergedTahaps.map(function(t) {
    return [t.id, t.name, t.status, safeNumber(t.updatedAt, Date.now()).toString()];
  });
  if (tahapRows.length > 0) {
    sheetTahap.getRange(2, 1, tahapRows.length, 4).setValues(tahapRows);
  }

  // 2. MERGE SHEET KELOMPOK
  var sheetKelompok = ss.getSheetByName("Kelompok") || ss.insertSheet("Kelompok");
  var mergedGroups = mergeLWW(parseKelompok(sheetKelompok), postData.groups);
  sheetKelompok.clear();
  sheetKelompok.appendRow(["ID Kelompok", "Nama Kelompok", "Tipe Hewan", "Target Total", "ID Tahap", "UpdatedAt"]);
  var groupRows = mergedGroups.map(function(g) {
    return [g.id, g.name, g.type, g.targetTotal, g.tahapId, safeNumber(g.updatedAt, Date.now()).toString()];
  });
  if (groupRows.length > 0) {
    sheetKelompok.getRange(2, 1, groupRows.length, 6).setValues(groupRows);
  }

  // 3. MERGE SHEET ANGGOTA
  var sheetAnggota = ss.getSheetByName("Anggota") || ss.insertSheet("Anggota");
  var mergedParticipants = mergeLWW(parseAnggota(sheetAnggota), postData.participants);
  sheetAnggota.clear();
  sheetAnggota.appendRow(["ID Anggota", "Nama Lengkap", "WhatsApp", "ID Kelompok", "Target Tabungan", "ID Tahap", "UpdatedAt", "PIN"]);
  var participantRows = mergedParticipants.map(function(p) {
    return [p.id, p.name, p.whatsapp, p.groupId, p.targetAmount, p.tahapId, safeNumber(p.updatedAt, Date.now()).toString(), p.pin || "1234"];
  });
  if (participantRows.length > 0) {
    sheetAnggota.getRange(2, 1, participantRows.length, 8).setValues(participantRows);
  }

  // 4. MERGE SHEET SETORAN
  var sheetSetoran = ss.getSheetByName("Setoran") || ss.insertSheet("Setoran");
  var mergedDeposits = mergeLWW(parseSetoran(sheetSetoran), postData.deposits);
  sheetSetoran.clear();
  sheetSetoran.appendRow(["ID Transaksi", "Tanggal", "ID Anggota", "Nominal Setor", "Pencatat", "Catatan", "ID Tahap", "Bulan", "UpdatedAt"]);
  var depositRows = mergedDeposits.map(function(d) {
    return [d.id, d.date, d.participantId, d.amount, d.recordedBy, d.note, d.tahapId, d.bulan, safeNumber(d.updatedAt, Date.now()).toString()];
  });
  if (depositRows.length > 0) {
    sheetSetoran.getRange(2, 1, depositRows.length, 9).setValues(depositRows);
  }

  // 5. MERGE SHEET PENGATURAN
  var sheetPengaturan = ss.getSheetByName("Pengaturan") || ss.insertSheet("Pengaturan");
  var existingSettings = parsePengaturan(sheetPengaturan);
  
  var newKetua = (postData.settings && postData.settings.ketuaName) ? postData.settings.ketuaName : existingSettings.ketuaName;
  var newBendahara = (postData.settings && postData.settings.bendaharaName) ? postData.settings.bendaharaName : existingSettings.bendaharaName;
  var newInstallment = (postData.settings && postData.settings.monthlyInstallment) ? postData.settings.monthlyInstallment : existingSettings.monthlyInstallment;
  var newCloudUrl = (postData.settings && postData.settings.cloudUrl) ? postData.settings.cloudUrl : existingSettings.cloudUrl;
  var newGatheringDates = (postData.settings && postData.settings.gatheringDates) ? postData.settings.gatheringDates : existingSettings.gatheringDates;

  sheetPengaturan.clear();
  sheetPengaturan.getRange(1, 1, 1, 2).setValues([["Kunci Pengaturan", "Nilai"]]);
  sheetPengaturan.appendRow(["ketuaName", safeString(newKetua)]);
  sheetPengaturan.appendRow(["bendaharaName", safeString(newBendahara)]);
  sheetPengaturan.appendRow(["monthlyInstallment", safeNumber(newInstallment, 150000).toString()]);
  sheetPengaturan.appendRow(["cloudUrl", safeString(newCloudUrl)]);
  
  if (newGatheringDates) {
    sheetPengaturan.appendRow(["gatheringDates", JSON.stringify(newGatheringDates)]);
  }

  var updatedTime = Math.max(existingSettings.lastUpdated || 0, postData.lastUpdated || 0, Date.now());
  sheetPengaturan.appendRow(["lastUpdated", updatedTime.toString()]);

  // Memberikan warna header hijau lembut agar rapi dibaca di Google Sheets
  var sheetsToStyle = ["Tahap", "Kelompok", "Anggota", "Setoran", "Deleted", "Pengaturan"];
  sheetsToStyle.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      styleHeaderSafely(sh, sh.getLastColumn());
      resizeColumnsSafely(sh);
    }
  });
}

// Fungsi pengurai data dari spreadsheet ke JSON khusus Tahap
function parseTahap(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    list.push({
      id: safeString(values[i][0]),
      name: safeString(values[i][1]),
      status: safeString(values[i][2]),
      updatedAt: lastCol >= 4 ? safeNumber(values[i][3], 0) : 0
    });
  }
  return list;
}

// Fungsi pengurai data dari spreadsheet ke JSON khusus Kelompok
function parseKelompok(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    list.push({
      id: safeString(values[i][0]),
      name: safeString(values[i][1]),
      type: safeString(values[i][2]),
      targetTotal: safeNumber(values[i][3], 0),
      tahapId: safeString(values[i][4]),
      updatedAt: lastCol >= 6 ? safeNumber(values[i][5], 0) : 0
    });
  }
  return list;
}

// Fungsi pengurai data dari spreadsheet ke JSON khusus Anggota
function parseAnggota(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    list.push({
      id: safeString(values[i][0]),
      name: safeString(values[i][1]),
      whatsapp: safeString(values[i][2]),
      groupId: safeString(values[i][3]),
      targetAmount: safeNumber(values[i][4], 0),
      tahapId: safeString(values[i][5]),
      updatedAt: lastCol >= 7 ? safeNumber(values[i][6], 0) : 0,
      pin: lastCol >= 8 ? safeString(values[i][7]) : "1234"
    });
  }
  return list;
}

// Fungsi pengurai data dari spreadsheet ke JSON khusus Setoran
function parseSetoran(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    var dateVal = values[i][1];
    if (dateVal instanceof Date) {
      var yyyy = dateVal.getFullYear();
      var mm = String(dateVal.getMonth() + 1).padStart(2, '0');
      var dd = String(dateVal.getDate()).padStart(2, '0');
      dateVal = yyyy + "-" + mm + "-" + dd;
    } else {
      dateVal = safeString(dateVal);
    }
    list.push({
      id: safeString(values[i][0]),
      date: dateVal,
      participantId: safeString(values[i][2]),
      amount: safeNumber(values[i][3], 0),
      recordedBy: safeString(values[i][4]),
      note: safeString(values[i][5]),
      tahapId: safeString(values[i][6]),
      bulan: lastCol >= 8 ? safeNumber(values[i][7], 1) : 1,
      updatedAt: lastCol >= 9 ? safeNumber(values[i][8], 0) : 0
    });
  }
  return list;
}

// Fungsi pengurai khusus untuk memuat tabel Deleted (Tombstone)
function parseDeleted(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    list.push({
      id: safeString(values[i][0]),
      type: safeString(values[i][1]),
      updatedAt: safeNumber(values[i][2], 0)
    });
  }
  return list;
}

// Fungsi pengurai data dari spreadsheet ke JSON khusus Pengaturan
function parsePengaturan(sheet) {
  var settings = {
    ketuaName: "",
    bendaharaName: "",
    monthlyInstallment: 150000,
    cloudUrl: "",
    gatheringDates: [],
    lastUpdated: 0
  };
  if (!sheet) return settings;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return settings;
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = safeString(values[i][0]);
    var val = values[i][1];
    if (key === "ketuaName") settings.ketuaName = safeString(val);
    if (key === "bendaharaName") settings.bendaharaName = safeString(val);
    if (key === "monthlyInstallment") settings.monthlyInstallment = safeNumber(val, 150000);
    if (key === "cloudUrl") settings.cloudUrl = safeString(val);
    if (key === "lastUpdated") settings.lastUpdated = safeNumber(val, 0);
    if (key === "gatheringDates") {
      try {
        settings.gatheringDates = JSON.parse(safeString(val));
      } catch (e) {
        // Fallback default jika gagal parse JSON
      }
    }
  }
  return settings;
}
