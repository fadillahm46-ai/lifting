// ==========================================
// BACKEND API CRANE TRUCK MONITORING (VERCEL)
// ==========================================

function doPost(e) {
  try {
    ensureSheetsExist();

    // Parse request dari frontend yang sekarang 100% menggunakan POST
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const payload = request.payload || {};

    let result;
    if (action === 'getAppData') result = getAppData();
    else if (action === 'saveOrder') result = saveOrder(payload);
    else if (action === 'updateJobRecord') result = updateJobRecord(payload);
    else if (action === 'updatePasswords') result = updatePasswords(payload.adminPass, payload.ctPass);
    else if (action === 'addMasterItem') result = addMasterItem(payload.category, payload.value);
    else if (action === 'deleteMasterItem') result = deleteMasterItem(payload.category, payload.value);
    else if (action === 'deleteJobRecord') result = deleteJobRecord(payload.id);
    else throw new Error("Aksi tidak valid atau tidak ditemukan");

    // Return respons JSON ke Frontend
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Menangani klik URL Script secara langsung dari browser agar tidak muncul tulisan Error
function doGet(e) {
  ensureSheetsExist();
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Backend API Crane Truck Aktif dan Berjalan Normal. Database siap digunakan!"
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// FUNGSI AUTO-SETUP DATABASE
// ==========================================
function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Orders") || !ss.getSheetByName("Master")) {
    setupDatabase();
  }
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) {
    orderSheet = ss.insertSheet("Orders");
    const headers = ["ID_Order", "Timestamp", "Nama_Pemohon", "No_WA", "Perusahaan", "Departemen", "Section", "Tgl_Pelaksanaan", "Shift", "Waktu_Request", "Durasi_Request", "Lokasi", "Tujuan", "Deskripsi", "Foto_Base64", "Status", "Unit_CT", "GL", "Operator", "Rigger", "Waktu_Start_Aktual", "Waktu_End_Aktual", "Durasi_Aktual", "Alasan_Delay"];
    orderSheet.appendRow(headers);
    orderSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d2e3fc");
    orderSheet.setFrozenRows(1);
  }

  let masterSheet = ss.getSheetByName("Master");
  if (!masterSheet) {
    masterSheet = ss.insertSheet("Master");
    const masterHeaders = ["Kategori", "Nilai"];
    masterSheet.appendRow(masterHeaders);
    masterSheet.getRange(1, 1, 1, masterHeaders.length).setFontWeight("bold").setBackground("#d2e3fc");
    masterSheet.setFrozenRows(1);

    const defaultData = [
      ["Perusahaan", "PT. PPA"], ["Perusahaan", "PT. BIB"], ["Departemen", "PLANT"], ["Departemen", "PROD"],
      ["Section", "Mechanic"], ["Section", "Warehouse"], ["Unit CT", "CT2514"], ["Unit CT", "CT2515"],
      ["GL", "ACHMAD HARIYANTO"], ["Operator", "ISKANDAR"], ["Rigger", "KHOLIK SYAIFUDIN"],
      ["Password", "Admin|101010"], ["Password", "CT|191919"]
    ];
    masterSheet.getRange(2, 1, defaultData.length, 2).setValues(defaultData);
  }

  let sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1) ss.deleteSheet(sheet1);
}

// ==========================================
// FUNGSI PENGELOLAAN DATA (CRUD)
// ==========================================

function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const masterSheet = ss.getSheetByName("Master");
  const masterData = masterSheet ? masterSheet.getDataRange().getValues().slice(1) : [];

  const orderSheet = ss.getSheetByName("Orders");
  let orders = [];
  if (orderSheet) {
    const data = orderSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      for (let i = 1; i < data.length; i++) {
        let row = data[i];
        let obj = {};
        for (let j = 0; j < headers.length; j++) {
          obj[headers[j]] = row[j];
        }

        if (obj.Timestamp && obj.Timestamp instanceof Date) {
          obj.Timestamp = Utilities.formatDate(obj.Timestamp, "GMT+8", "dd/MM/yyyy HH:mm");
        }
        if (obj.Tgl_Pelaksanaan && obj.Tgl_Pelaksanaan instanceof Date) {
          obj.Tgl_Pelaksanaan = Utilities.formatDate(obj.Tgl_Pelaksanaan, "GMT+8", "yyyy-MM-dd");
        }
        if (obj.Waktu_Request && obj.Waktu_Request instanceof Date) {
          obj.Waktu_Request = Utilities.formatDate(obj.Waktu_Request, "GMT+8", "HH:mm");
        }
        orders.push(obj);
      }
    }
  }

  return { master: masterData, orders: orders.reverse() };
}

function saveOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Orders");

  const now = new Date();
  const dateStr = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");

  const data = sheet.getDataRange().getValues();
  let countToday = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).includes(`LIFT-${dateStr}`)) {
      countToday++;
    }
  }
  const idOrder = `LIFT-${dateStr}-${countToday + 1}_${now.getTime()}`;

  const newRow = [
    idOrder, now, payload.nama, payload.wa, payload.perusahaan, payload.departemen, payload.section,
    payload.tanggal, payload.shift, payload.waktu, payload.durasi, payload.lokasi, payload.tujuan,
    payload.deskripsi, payload.foto, "Menunggu Validasi", "", "", "", "", "", "", "", ""
  ];

  sheet.appendRow(newRow);
  return { success: true, id: idOrder };
}

function updateJobRecord(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === payload.id) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) return { success: false, message: "ID Job tidak ditemukan di database" };

  if (payload.isResuming) {
    let oldData = data[targetRow - 1];
    sheet.getRange(targetRow, 16).setValue('Completed');

    let baseId = payload.id.split('_')[0];
    let maxSplit = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).startsWith(baseId + '_')) {
        let c = parseInt(String(data[i][0]).split('_')[1]);
        if (!isNaN(c) && c > maxSplit) maxSplit = c;
      }
    }
    const newId = `${baseId}_${maxSplit + 1}`;

    const newRow = [
      newId, oldData[1], oldData[2], oldData[3], oldData[4], oldData[5], oldData[6],
      payload.tglReq || oldData[7], payload.shift || oldData[8], payload.waktuReq || oldData[9],
      payload.durasiReq || oldData[10], oldData[11], oldData[12], payload.deskripsiReq || oldData[13],
      oldData[14], payload.status, payload.unit || "", payload.gl || "", payload.operator || "",
      payload.rigger || "", "", "", "", ""
    ];
    sheet.appendRow(newRow);
    return { success: true, message: "Job berhasil divalidasi dan di-split." };
  }

  if (payload.status !== undefined) sheet.getRange(targetRow, 16).setValue(payload.status);
  if (payload.unit !== undefined) sheet.getRange(targetRow, 17).setValue(payload.unit);
  if (payload.gl !== undefined) sheet.getRange(targetRow, 18).setValue(payload.gl);
  if (payload.operator !== undefined) sheet.getRange(targetRow, 19).setValue(payload.operator);
  if (payload.rigger !== undefined) sheet.getRange(targetRow, 20).setValue(payload.rigger);
  if (payload.startAktual !== undefined) sheet.getRange(targetRow, 21).setValue(payload.startAktual);
  if (payload.endAktual !== undefined) sheet.getRange(targetRow, 22).setValue(payload.endAktual);
  if (payload.durasiAktual !== undefined) sheet.getRange(targetRow, 23).setValue(payload.durasiAktual);
  if (payload.alasanDelay !== undefined) sheet.getRange(targetRow, 24).setValue(payload.alasanDelay);

  if (payload.tglReq !== undefined) sheet.getRange(targetRow, 8).setValue(payload.tglReq);
  if (payload.shift !== undefined) sheet.getRange(targetRow, 9).setValue(payload.shift);
  if (payload.waktuReq !== undefined) sheet.getRange(targetRow, 10).setValue(payload.waktuReq);
  if (payload.durasiReq !== undefined) sheet.getRange(targetRow, 11).setValue(payload.durasiReq);
  if (payload.deskripsiReq !== undefined) sheet.getRange(targetRow, 14).setValue(payload.deskripsiReq);

  return { success: true };
}

function deleteJobRecord(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1); return { success: true };
    }
  }
  return { success: false, message: "Data tidak ditemukan" };
}

function addMasterItem(category, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master");
  sheet.appendRow([category, value]);
  return { success: true };
}

function deleteMasterItem(category, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === category && data[i][1] === value) {
      sheet.deleteRow(i + 1); return { success: true };
    }
  }
  return { success: false, message: "Data master tidak ditemukan" };
}

function updatePasswords(adminPass, ctPass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master");
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === 'Password') sheet.deleteRow(i + 1);
  }
  sheet.appendRow(['Password', 'Admin|' + adminPass]);
  sheet.appendRow(['Password', 'CT|' + ctPass]);
  return { success: true };
}