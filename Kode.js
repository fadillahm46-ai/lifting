// ==========================================
// BACKEND API CRANE TRUCK MONITORING (FULL VERSION)
// ==========================================

// KUNCI UTAMA: ID SPREADSHEET ANDA DITANAM DI SINI AGAR ANTI-GAGAL
const SHEET_ID = "1C1Ec3hVGJxGxDpZfmUGXm8lWVvBD5-WxRAKFINakSvo";

function doPost(e) {
  try {
    // Pastikan struktur tabel aman sebelum melakukan apapun
    ensureSheetsExist();

    // Parsing data yang dikirim dari web/APK (Vercel)
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const payload = request.payload || {};

    let result;
    // Rute aksi (Routing)
    if (action === 'saveOrder') {
      result = saveOrder(payload);
    } else if (action === 'updateJobRecord') {
      result = updateJobRecord(payload);
    } else if (action === 'updatePasswords') {
      result = updatePasswords(payload.adminPass, payload.ctPass);
    } else if (action === 'addMasterItem') {
      result = addMasterItem(payload.category, payload.value);
    } else if (action === 'deleteMasterItem') {
      result = deleteMasterItem(payload.category, payload.value);
    } else if (action === 'deleteJobRecord') {
      result = deleteJobRecord(payload.id);
    } else {
      throw new Error("Aksi POST tidak valid atau tidak ditemukan");
    }

    // Kembalikan hasil eksekusi ke aplikasi
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Tangkap dan kirim pesan error jika gagal
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    // Pastikan struktur tabel aman
    ensureSheetsExist();

    // Ambil parameter URL (contoh: ?action=getAppData)
    const action = e.parameter.action;

    // Jika aplikasi web meminta data untuk ditampilkan (Load)
    if (action === 'getAppData') {
      const result = getAppData();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Jika web dibuka langsung secara manual di browser (tanpa parameter action)
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "API Aktif. Sistem Monitoring Berjalan. Gunakan aplikasi Web untuk melihat data."
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// FUNGSI AUTO-SETUP & CEK KOLOM (PENCEGAH ERROR UI)
// ==========================================
function ensureSheetsExist() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Header baku (Wajib 24 Kolom) agar UI web tidak error
  const headers = [
    "ID_Order", "Timestamp", "Nama_Pemohon", "No_WA", "Perusahaan",
    "Departemen", "Section", "Tgl_Pelaksanaan", "Shift", "Waktu_Request",
    "Durasi_Request", "Lokasi", "Tujuan", "Deskripsi", "Foto_Base64",
    "Status", "Unit_CT", "GL", "Operator", "Rigger",
    "Waktu_Start_Aktual", "Waktu_End_Aktual", "Durasi_Aktual", "Alasan_Delay"
  ];

  let orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) {
    setupDatabase(ss, headers);
  } else {
    // Cek jika ada kolom header yang terhapus secara tidak sengaja oleh user di Spreadsheet
    const currentCols = orderSheet.getLastColumn();
    if (currentCols < headers.length) {
      orderSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d2e3fc");
    }
  }

  if (!ss.getSheetByName("Master")) {
    setupDatabase(ss, headers);
  }
}

function setupDatabase(ss, headers) {
  let orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) {
    orderSheet = ss.insertSheet("Orders");
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

    // Inject Data Master Default
    const defaultData = [
      ["Perusahaan", "PT. PPA"], ["Departemen", "PLANT"], ["Section", "Mechanic"],
      ["Unit CT", "CT2514"], ["GL", "ACHMAD HARIYANTO"], ["Operator", "ISKANDAR"], ["Rigger", "KHOLIK SYAIFUDIN"],
      ["Password", "Admin|101010"], ["Password", "CT|191919"]
    ];
    masterSheet.getRange(2, 1, defaultData.length, 2).setValues(defaultData);
  }

  // Hapus Sheet1 bawaan Google Sheets jika ada
  let sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1) ss.deleteSheet(sheet1);
}

// ==========================================
// FUNGSI PENGELOLAAN DATA UTAMA (CRUD)
// ==========================================

function getAppData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Ambil Data Master
  const masterSheet = ss.getSheetByName("Master");
  // Abaikan baris kosong di Data Master
  const masterData = masterSheet ? masterSheet.getDataRange().getValues().slice(1).filter(row => row[0] && row[1]) : [];

  // 2. Ambil Data Orders
  const orderSheet = ss.getSheetByName("Orders");
  let orders = [];

  if (orderSheet) {
    const data = orderSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      for (let i = 1; i < data.length; i++) {
        let row = data[i];

        // PENCEGAH ERROR: Lewati (skip) baris kosong yang ada di tengah-tengah spreadsheet
        if (!row[0]) continue;

        let obj = {};
        for (let j = 0; j < headers.length; j++) {
          obj[headers[j]] = row[j];
        }

        // FORMATTING TANGGAL & WAKTU AGAR TIDAK ERROR ZONA WAKTU (GMT)
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

  // Kirim data orders secara terbalik (descending/terbaru di atas)
  return {
    status: "success",
    master: masterData,
    orders: orders.reverse()
  };
}

function saveOrder(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Orders");
  const now = new Date();
  const dateStr = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");

  // Cek nomor urut hari ini
  const data = sheet.getDataRange().getValues();
  let countToday = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).includes(`LIFT-${dateStr}`)) countToday++;
  }

  // Generate ID Unik (contoh: LIFT-20260821-1_1700000000)
  const idOrder = `LIFT-${dateStr}-${countToday + 1}_${now.getTime()}`;

  const newRow = [
    idOrder,
    now,
    payload.nama,
    payload.wa,
    payload.perusahaan,
    payload.departemen,
    payload.section,
    payload.tanggal,
    payload.shift,
    payload.waktu,
    payload.durasi,
    payload.lokasi,
    payload.tujuan,
    payload.deskripsi,
    payload.foto,
    "Menunggu Validasi", // Status Awal
    "", "", "", "", "", "", "", "" // Kolom CT dan Evaluasi dikosongkan
  ];

  sheet.appendRow(newRow);
  return { success: true, id: idOrder };
}

function updateJobRecord(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();

  // Cari baris berdasarkan ID
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === payload.id) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return { success: false, message: "ID Job tidak ditemukan di database" };
  }

  // LOGIKA SPLIT JOB (Jika Melanjutkan Job Pending)
  if (payload.isResuming) {
    let oldData = data[targetRow - 1];

    // Tutup job lama menjadi Completed
    sheet.getRange(targetRow, 16).setValue('Completed');

    // Hitung prefix ID untuk split job (ex: _1, _2)
    let baseId = payload.id.split('_')[0];
    let maxSplit = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).startsWith(baseId + '_')) {
        let c = parseInt(String(data[i][0]).split('_')[1]);
        if (!isNaN(c) && c > maxSplit) maxSplit = c;
      }
    }

    const newId = `${baseId}_${maxSplit + 1}`;

    // Buat baris baru untuk job lanjutan
    const newRow = [
      newId,
      oldData[1], oldData[2], oldData[3], oldData[4], oldData[5], oldData[6],
      payload.tglReq || oldData[7],
      payload.shift || oldData[8],
      payload.waktuReq || oldData[9],
      payload.durasiReq || oldData[10],
      oldData[11], oldData[12],
      payload.deskripsiReq || oldData[13],
      oldData[14],
      payload.status,
      payload.unit || "",
      payload.gl || "",
      payload.operator || "",
      payload.rigger || "",
      "", "", "", "" // Kosongkan waktu aktual
    ];
    sheet.appendRow(newRow);
    return { success: true };
  }

  // UPDATE JOB REGULER (Validasi, Selesai, Pending, Cancel)
  if (payload.status !== undefined) sheet.getRange(targetRow, 16).setValue(payload.status);
  if (payload.unit !== undefined) sheet.getRange(targetRow, 17).setValue(payload.unit);
  if (payload.gl !== undefined) sheet.getRange(targetRow, 18).setValue(payload.gl);
  if (payload.operator !== undefined) sheet.getRange(targetRow, 19).setValue(payload.operator);
  if (payload.rigger !== undefined) sheet.getRange(targetRow, 20).setValue(payload.rigger);
  if (payload.startAktual !== undefined) sheet.getRange(targetRow, 21).setValue(payload.startAktual);
  if (payload.endAktual !== undefined) sheet.getRange(targetRow, 22).setValue(payload.endAktual);
  if (payload.durasiAktual !== undefined) sheet.getRange(targetRow, 23).setValue(payload.durasiAktual);
  if (payload.alasanDelay !== undefined) sheet.getRange(targetRow, 24).setValue(payload.alasanDelay);

  // Update Req (Jika dirubah oleh Admin saat validasi)
  if (payload.tglReq !== undefined) sheet.getRange(targetRow, 8).setValue(payload.tglReq);
  if (payload.shift !== undefined) sheet.getRange(targetRow, 9).setValue(payload.shift);
  if (payload.waktuReq !== undefined) sheet.getRange(targetRow, 10).setValue(payload.waktuReq);
  if (payload.durasiReq !== undefined) sheet.getRange(targetRow, 11).setValue(payload.durasiReq);
  if (payload.deskripsiReq !== undefined) sheet.getRange(targetRow, 14).setValue(payload.deskripsiReq);

  return { success: true };
}

function deleteJobRecord(id) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "Data tidak ditemukan untuk dihapus" };
}

// ==========================================
// FUNGSI DATA MASTER
// ==========================================

function addMasterItem(category, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheetByName("Master").appendRow([category, value]);
  return { success: true };
}

function deleteMasterItem(category, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Master");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === category && data[i][1] === value) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "Data master tidak ditemukan" };
}

function updatePasswords(adminPass, ctPass) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Master");
  const data = sheet.getDataRange().getValues();

  // Bersihkan password lama (dibalik dari bawah agar urutan tidak kacau saat dihapus)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === 'Password') {
      sheet.deleteRow(i + 1);
    }
  }

  // Masukkan password baru
  sheet.appendRow(['Password', 'Admin|' + adminPass]);
  sheet.appendRow(['Password', 'CT|' + ctPass]);

  return { success: true };
}