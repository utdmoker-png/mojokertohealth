const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_DATA = "Mojokerto Health V2";
const SHEET_USERS = "Users";

// Inisialisasi Sheet & Database Automatis
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Sheet Data Utama
  let sheetData = ss.getSheetByName(SHEET_DATA);
  if (!sheetData) {
    sheetData = ss.insertSheet(SHEET_DATA);
    sheetData.appendRow([
      "ID", "Username", "Nama Rumah Sakit", "Nama Petugas", "Nama Dokter yang Meminta", 
      "No.Register/RM", "Nama Pasien", "Alamat", "Tanggal Lahir", "Jenis Kelamin", 
      "Ruangan/Bagian", "Pelayanan", "Nomor SEP", "Golongan Darah", "Rhesus", 
      "Diagnosa", "Indikasi Transfusi", "Jumlah Trombosit", "Hb Terakhir", 
      "Jenis Darah yang Diminta", "Jumlah Kantong", "Sifat Permintaan", 
      "Tanggal Dibutuhkan", "Transfusi Sebelumnya", "Reaksi Transfusi", "Timestamp"
    ]);
  }
  
  // Sheet Pengguna
  let sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(SHEET_USERS);
    sheetUsers.appendRow(["Username", "Password", "Nama Rumah Sakit"]);
    // Akun default (Username: admin, Pass: admin123)
    sheetUsers.appendRow(["admin", "admin123", "RSUD Dr. Wahidin Sudirohusodo"]);
  }
}

function doGet() {
  setupDatabase();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Formulir Permintaan Darah PMI Kota Mojokerto')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi Autentikasi User
function loginUser(username, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      return { success: true, username: data[i][0], rs: data[i][2] };
    }
  }
  return { success: false, message: "Username atau Password salah!" };
}

function editUser(oldUsername, newPassword, newRS) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === oldUsername) {
      if (newPassword) sheet.getRange(i + 1, 2).setValue(newPassword);
      if (newRS) sheet.getRange(i + 1, 3).setValue(newRS);
      return { success: true, message: "Profil berhasil diperbarui!" };
    }
  }
  return { success: false, message: "User tidak ditemukan!" };
}

// CRUD Data Permintaan Darah
function saveData(formData, currentUser) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATA);
  
  if (formData.id) {
    // Edit Data
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === formData.id && data[i][1] === currentUser) {
        const row = i + 1;
        const updateRow = [
          formData.id, currentUser, formData.rs, formData.petugas, formData.dokter,
          formData.rm, formData.pasien, formData.alamat, formData.tgllahir, formData.jk,
          formData.ruangan, formData.pelayanan, formData.sep, formData.goldar, formData.rhesus,
          formData.diagnosa, formData.indikasi, formData.trombosit, formData.hb,
          formData.jenisdarah, formData.kantong, formData.sifat, formData.tgldibutuhkan,
          formData.transfusi_sebelumnya, formData.reaksi, new Date()
        ];
        sheet.getRange(row, 1, 1, updateRow.length).setValues([updateRow]);
        return { success: true, message: "Data berhasil diperbarui!" };
      }
    }
  } else {
    // Tambah Data Baru
    const newId = "REQ-" + new Date().getTime();
    const newRow = [
      newId, currentUser, formData.rs, formData.petugas, formData.dokter,
      formData.rm, formData.pasien, formData.alamat, formData.tgllahir, formData.jk,
      formData.ruangan, formData.pelayanan, formData.sep, formData.goldar, formData.rhesus,
      formData.diagnosa, formData.indikasi, formData.trombosit, formData.hb,
      formData.jenisdarah, formData.kantong, formData.sifat, formData.tgldibutuhkan,
      formData.transfusi_sebelumnya, formData.reaksi, new Date()
    ];
    sheet.appendRow(newRow);
    return { success: true, message: "Data berhasil disimpan!" };
  }
}

function getUserData(username) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATA);
  const data = sheet.getDataRange().getValues();
  const userRecords = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      userRecords.push({
        id: data[i][0], username: data[i][1], rs: data[i][2], petugas: data[i][3], dokter: data[i][4],
        rm: data[i][5], pasien: data[i][6], alamat: data[i][7], tgllahir: formatDate(data[i][8]), jk: data[i][9],
        ruangan: data[i][10], pelayanan: data[i][11], sep: data[i][12], goldar: data[i][13], rhesus: data[i][14],
        diagnosa: data[i][15], indikasi: data[i][16], trombosit: data[i][17], hb: data[i][18],
        jenisdarah: data[i][19], kantong: data[i][20], sifat: data[i][21], tgldibutuhkan: formatDate(data[i][22]),
        transfusi_sebelumnya: data[i][23], reaksi: data[i][24]
      });
    }
  }
  return userRecords;
}

function deleteData(id, username) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATA);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id && data[i][1] === username) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Data berhasil dihapus!" };
    }
  }
  return { success: false, message: "Gagal menghapus data." };
}

function formatDate(dateVal) {
  if (!dateVal) return "";
  try {
    const d = new Date(dateVal);
    return d.toISOString().split('T')[0];
  } catch(e) {
    return dateVal;
  }
}