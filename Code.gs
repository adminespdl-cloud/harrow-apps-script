// ==============================================
// KONFIGURASI AWAL (WAJIB DISESUAIKAN)
// ==============================================
const FOLDER_ID = "MASUKKAN_ID_FOLDER_DRIVE_DI_SINI"; 
const SHEET_NAME = "Data_HARROW";
const USERS_SHEET_NAME = "Users";

// ==============================================
// FUNGSI UTAMA
// ==============================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Aplikasi HARROW')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function processForm(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMain = getOrCreateMainSheet(ss);
    let fileUrl = "";

    if (payload.fotoKolase && payload.fotoKolase.length > 0) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const fileName = `HARROW_${payload.penyulang}_${payload.tanggal}_${new Date().getTime()}.jpg`;
      const contentType = 'image/jpeg';
      const b64Data = payload.fotoKolase.split(',')[1]; 
      const blob = Utilities.newBlob(Utilities.base64Decode(b64Data), contentType, fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
    }

    const lastRow = sheetMain.getLastRow();
    const no = lastRow;
    const imgFormula = fileUrl ? `=IMAGE("${fileUrl}")` : "";

    const rowMain = [
      no,
      new Date(),
      payload.tanggal,
      payload.namaPetugas,
      payload.penyulang,
      payload.noWo,
      payload.koordinat,
      payload.ruasSegmen,
      payload.volume,
      imgFormula,
      payload.ruasFull || payload.ruasSegmen
    ];
    sheetMain.appendRow(rowMain);
    sheetMain.hideColumns(11);

    return "Sukses";
  } catch (error) {
    throw new Error("Gagal menyimpan data: " + error.message);
  }
}

function getRiwayatData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMain = getOrCreateMainSheet(ss);
  const lastRow = sheetMain.getLastRow();
  
  if (lastRow <= 1) return [];

  const data = sheetMain.getRange(2, 1, lastRow - 1, 11).getValues();
  const formulas = sheetMain.getRange(2, 1, lastRow - 1, 11).getFormulas();
  const result = [];
  
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const frm = formulas[i];
    
    let tgl = row[2];
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    let picUrl = "";
    if (frm[9] && frm[9].toUpperCase().includes("IMAGE")) {
      const match = frm[9].match(/"([^"]+)"/);
      if (match) picUrl = match[1];
    }

    result.push({
      noUrut: row[0] || (i + 1),
      tanggal: tgl,
      namaPetugas: row[3] || "-",
      penyulang: row[4] || "-",
      noWo: row[5] || "-",
      koordinat: row[6] || "-",
      ruasFull: row[10] || row[7] || "",
      volume: row[8] || 0,
      fotoUrl: picUrl
    });
  }
  
  return result;
}

// ==============================================
// USER MANAGEMENT
// ==============================================

function authenticateUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateUsersSheet(ss);
    let lastRow = sheet.getLastRow();
    
    // Failsafe: Jika sheet kosong, tambahkan admin default
    if (lastRow <= 1) {
      const id = new Date().getTime().toString();
      sheet.appendRow([id, 'admin', 'harrow2024', 'Administrator', 'Admin', 'AKTIF', new Date()]);
      lastRow = 2;
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[1].toString().toLowerCase() === username.toLowerCase() && 
          row[2].toString() === password && 
          row[5].toString().toUpperCase() === 'AKTIF') {
        return { id: row[0].toString(), nama: row[3], role: row[4], username: row[1] };
      }
    }
    return null;
  } catch(e) {
    throw new Error('Login gagal: ' + e.message);
  }
}

function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateUsersSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  return data.map((row) => ({
    id: row[0].toString(),
    username: row[1],
    password: row[2],
    nama: row[3],
    role: row[4],
    status: row[5],
    dibuat: row[6] instanceof Date ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[6].toString()
  }));
}

function saveUser(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateUsersSheet(ss);
  const lastRow = sheet.getLastRow();
  
  if (payload.id) {
    if (lastRow <= 1) throw new Error('User tidak ditemukan');
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString() === payload.id.toString()) {
        const rowNum = i + 2;
        sheet.getRange(rowNum, 2, 1, 5).setValues([[
          payload.username, payload.password, payload.nama, payload.role, payload.status
        ]]);
        return 'Sukses';
      }
    }
    throw new Error('User tidak ditemukan');
  } else {
    // Cek duplikat username
    if (lastRow > 1) {
      const usernameData = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < usernameData.length; i++) {
        if (usernameData[i][0].toString().toLowerCase() === payload.username.toLowerCase()) {
          throw new Error('Username sudah digunakan');
        }
      }
    }
    const id = new Date().getTime().toString();
    sheet.appendRow([id, payload.username, payload.password, payload.nama, payload.role, 'AKTIF', new Date()]);
    return 'Sukses';
  }
}

function deleteUser(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateUsersSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) throw new Error('User tidak ditemukan');
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() === userId.toString()) {
      sheet.deleteRow(i + 2);
      return 'Sukses';
    }
  }
  throw new Error('User tidak ditemukan');
}

// ==============================================
// FUNGSI BANTUAN
// ==============================================

function getOrCreateMainSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "NO", "Waktu Submit", "Tanggal Eksekusi", "Nama Petugas", "Penyulang", 
      "No. WO", "Titik Koordinat", "Detail Ruas (Filter)", "Volume (KMS)", "Foto Dokumentasi", "Detail Ruas (Full)"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0f2fe");
  } else {
    const firstHeader = sheet.getRange(1, 1).getValue();
    if (firstHeader !== "NO") {
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue("NO").setFontWeight("bold").setBackground("#e0f2fe");
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const numValues = [];
        for (let i = 2; i <= lastRow; i++) numValues.push([i - 1]);
        sheet.getRange(2, 1, lastRow - 1, 1).setValues(numValues);
      }
    }
    const lc = sheet.getLastColumn();
    if (lc < 11) {
      sheet.getRange(1, 11).setValue("Detail Ruas (Full)").setFontWeight("bold").setBackground("#e0f2fe");
    }
    sheet.hideColumns(11);
  }
  return sheet;
}

function getOrCreateUsersSheet(ss) {
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    const headers = ['ID', 'Username', 'Password', 'Nama Lengkap', 'Role', 'Status', 'Dibuat'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0d3b6e').setFontColor('#ffffff');
    // Tambah akun admin default
    const id = new Date().getTime().toString();
    sheet.appendRow([id, 'admin', 'harrow2024', 'Administrator', 'Admin', 'AKTIF', new Date()]);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 80);
    sheet.setColumnWidth(7, 130);
  }
  return sheet;
}
