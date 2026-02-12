const SPREADSHEET_ID = '1vXjhXxfgecTxwiFwymK5kTFLTZx4a7L_ANerjCxWnh8';

// Mendapatkan spreadsheet
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Fungsi untuk inisialisasi sheet dan data sample
function initializeSheets() {
  try {
    const ss = getSpreadsheet();
    
    // 1. Inisialisasi Sheet Users
    let usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) {
      usersSheet = ss.insertSheet('Users');
    } else {
      usersSheet.clear();
    }
    
    // Header Users
    usersSheet.getRange(1, 1, 1, 5).setValues([
      ['id', 'nisn', 'password', 'role', 'nama_lengkap','kelas','status_nilai']
    ]);
    usersSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
    
    // Sample data Users
    usersSheet.getRange(2, 1, 5, 5).setValues([
      [1, '1234567890', 'siswa123', 'siswa', 'Ahmad Fauzi'],
      [2, '1234567891', 'siswa123', 'siswa', 'Siti Nurhaliza'],
      [3, '1234567892', 'siswa123', 'siswa', 'Budi Santoso'],
      [4, '1234567893', 'siswa123', 'siswa', 'Dewi Lestari'],
      [5, '0987654321', 'admin123', 'admin', 'Admin Sekolah']
    ]);
    
    // 2. Inisialisasi Sheet Kelulusan
    let kelulusanSheet = ss.getSheetByName('Kelulusan');
    if (!kelulusanSheet) {
      kelulusanSheet = ss.insertSheet('Kelulusan');
    } else {
      kelulusanSheet.clear();
    }
    
    // Header Kelulusan
    kelulusanSheet.getRange(1, 1, 1, 6).setValues([
      ['id', 'nama_lengkap', 'nisn', 'kelas', 'jurusan', 'status_kelulusan']
    ]);
    kelulusanSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
    
    // Sample data Kelulusan
    kelulusanSheet.getRange(2, 1, 4, 6).setValues([
      [1, 'Ahmad Fauzi', '1234567890', 'XII IPA 1', 'IPA', 'LULUS'],
      [2, 'Siti Nurhaliza', '1234567891', 'XII IPA 2', 'IPA', 'LULUS'],
      [3, 'Budi Santoso', '1234567892', 'XII IPS 1', 'IPS', 'TIDAK LULUS'],
      [4, 'Dewi Lestari', '1234567893', 'XII IPA 1', 'IPA', 'LULUS']
    ]);
    
    // 3. Inisialisasi Sheet Announcements
    let announcementsSheet = ss.getSheetByName('Announcements');
    if (!announcementsSheet) {
      announcementsSheet = ss.insertSheet('Announcements');
    } else {
      announcementsSheet.clear();
    }
    
    // Header Announcements
    announcementsSheet.getRange(1, 1, 1, 5).setValues([
      ['id', 'title', 'message', 'countdown_iso', 'published']
    ]);
    announcementsSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
    
    // Sample data Announcements
    // Buat countdown untuk 1 jam dari sekarang untuk testing
    const now = new Date();
    const countdownDate = new Date(now.getTime() + (60 * 60 * 1000)); // 1 jam dari sekarang
    const countdownISO = Utilities.formatDate(countdownDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
    const formattedCountdown = Utilities.formatDate(countdownDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    
    announcementsSheet.getRange(2, 1, 1, 5).setValues([
      [1, 'Pengumuman Kelulusan Tahun Ajaran 2025/2026', 'Selamat kepada siswa-siswi yang telah dinyatakan LULUS. Bagi yang belum berhasil, jangan berkecil hati dan tetap semangat untuk mencoba lagi!', countdownISO, true]
    ]);
    
    // Auto-resize columns untuk semua sheet
    usersSheet.autoResizeColumns(1, 5);
    kelulusanSheet.autoResizeColumns(1, 6);
    announcementsSheet.autoResizeColumns(1, 5);
    
    return {
      success: true,
      message: 'Semua sheet berhasil diinisialisasi dengan data sample!\n\nSheet yang dibuat:\n- Users (5 user: 4 siswa, 1 admin)\n- Kelulusan (4 data siswa)\n- Announcements (1 pengumuman dengan countdown pada: ' + formattedCountdown + ')\n\nLogin Demo:\nSiswa: NISN 1234567890 | Password siswa123\nAdmin: NISN 0987654321 | Password admin123'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Mendapatkan sheet berdasarkan nama
function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

// Fungsi untuk menampilkan halaman HTML
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Aplikasi NILAI MAPEL PRODUKTIF TKJT')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setFaviconUrl('https://w7.pngwing.com/pngs/184/833/png-transparent-exam-test-checklist-online-learning-education-online-document-online-learning-icon.png');
}

// Fungsi login
function login(nisn, password) {
  try {
    const sheet = getSpreadsheet().getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    // Cari user berdasarkan NISN dan password
    for (let i = 1; i < data.length; i++) {
      // Kolom 1 = NISN, Kolom 2 = Password
      // Pastikan konversi ke string agar aman
      if (String(data[i][1]) === String(nisn) && String(data[i][2]) === String(password)) {
        
        // Kita HAPUS PropertiesService dari sini.
        // Langsung kembalikan data user ke browser.
        return {
          success: true,
          user: {
            id: data[i][0],
            nisn: data[i][1],
            role: data[i][3],
            nama_lengkap: data[i][4]
          }
        };
      }
    }
    
    return { success: false, message: 'NISN atau password salah' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Fungsi logout
function logout() {
  PropertiesService.getUserProperties().deleteProperty('logged_in_nisn');
  PropertiesService.getUserProperties().deleteProperty('logged_in_role');
  PropertiesService.getUserProperties().deleteProperty('logged_in_nama');
  return { success: true };
}

// Mendapatkan user yang sedang login
function getCurrentUser() {
  const nisn = PropertiesService.getUserProperties().getProperty('logged_in_nisn');
  const role = PropertiesService.getUserProperties().getProperty('logged_in_role');
  const nama_lengkap = PropertiesService.getUserProperties().getProperty('logged_in_nama');
  
  if (nisn) {
    return {
      success: true,
      user: {
        nisn: nisn,
        role: role,
        nama_lengkap: nama_lengkap
      }
    };
  }
  
  return { success: false };
}

// Mendapatkan data siswa berdasarkan NISN
function getStudentDataByNISN(nisn) {
  try {
    const sheet = getSheet('Users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === nisn) {
        return {
          success: true,
          data: {
            id: data[i][0],
            nisn: data[i][1],
            role: data[i][3],
            nama_lengkap: data[i][4],
            status_nilai: data[i][5]
          }
        };
      }
    }
    
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Mendapatkan data kelulusan siswa berdasarkan NISN
function getStudentGraduationData(nisn) {
  try {
    const sheet = getSheet('Kelulusan');
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Get header row
    const result = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === nisn) { // Check NISN in the 3rd column (index 2)
        // Populate result object dynamically based on headers
        headers.forEach((header, index) => {
          result[header] = data[i][index];
        });

        return {
          success: true,
          data: result
        };
      }
    }
    
    return { success: false, message: 'Data kelulusan tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Mendapatkan semua data kelulusan (untuk admin)
function getAllGraduationData() {
  try {
    const sheet = getSheet('Kelulusan');
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0],
        nama_lengkap: data[i][1],
        nisn: data[i][2],
        kelas: data[i][3],
        jurusan: data[i][4],
        status_kelulusan: data[i][5]
      });
    }
    
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Membuat data siswa baru
function createStudent(data) {
  try {
    const sheet = getSheet('Kelulusan');
    const lastRow = sheet.getLastRow();
    const newId = lastRow > 0 ? lastRow : 1;
    
    sheet.appendRow([
      newId,
      data.nama_lengkap,
      data.nisn,
      data.kelas,
      data.jurusan,
      data.status_kelulusan
    ]);
    
    return { success: true, message: 'Data berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Update data siswa
function updateStudent(id, data) {
  try {
    const sheet = getSheet('Kelulusan');
    const dataRange = sheet.getDataRange().getValues();
    
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == id) {
        sheet.getRange(i + 1, 2).setValue(data.nama_lengkap);
        sheet.getRange(i + 1, 3).setValue(data.nisn);
        sheet.getRange(i + 1, 4).setValue(data.kelas);
        sheet.getRange(i + 1, 5).setValue(data.jurusan);
        sheet.getRange(i + 1, 6).setValue(data.status_kelulusan);
        return { success: true, message: 'Data berhasil diupdate' };
      }
    }
    
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Hapus data siswa
function deleteStudent(id) {
  try {
    const sheet = getSheet('Kelulusan');
    const dataRange = sheet.getDataRange().getValues();
    
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data berhasil dihapus' };
      }
    }
    
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Mendapatkan semua pengumuman
function getAnnouncements() {
  try {
    const sheet = getSheet('Announcements');
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0],
        title: data[i][1],
        message: data[i][2],
        countdown_iso: data[i][3],
        published: data[i][4]
      });
    }
    
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Membuat pengumuman baru
function createAnnouncement(data) {
  try {
    const sheet = getSheet('Announcements');
    const lastRow = sheet.getLastRow();
    const newId = lastRow > 0 ? lastRow : 1;
    
    sheet.appendRow([
      newId,
      data.title,
      data.message,
      data.countdown_iso,
      data.published
    ]);
    
    return { success: true, message: 'Pengumuman berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Update pengumuman
function updateAnnouncement(id, data) {
  try {
    const sheet = getSheet('Announcements');
    const dataRange = sheet.getDataRange().getValues();
    
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == id) {
        sheet.getRange(i + 1, 2).setValue(data.title);
        sheet.getRange(i + 1, 3).setValue(data.message);
        sheet.getRange(i + 1, 4).setValue(data.countdown_iso);
        sheet.getRange(i + 1, 5).setValue(data.published);
        return { success: true, message: 'Pengumuman berhasil diupdate' };
      }
    }
    
    return { success: false, message: 'Pengumuman tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Hapus pengumuman
function deleteAnnouncement(id) {
  try {
    const sheet = getSheet('Announcements');
    const dataRange = sheet.getDataRange().getValues();
    
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Pengumuman berhasil dihapus' };
      }
    }
    
    return { success: false, message: 'Pengumuman tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
// Mendapatkan statistik siswa
function getStudentStats() {
  try {
    const sheet = getSheet('Kelulusan');
    const data = sheet.getDataRange().getValues();
    
    const totalStudents = data.length - 1; // Subtract header row
    let passedStudents = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === 'Aman') { // status_kelulusan is in the 6th column (index 5)
        passedStudents++;
      }
    }
    
    const failedStudents = totalStudents - passedStudents;
    
    return { 
      success: true, 
      data: {
        total: totalStudents,
        passed: passedStudents,
        failed: failedStudents
      } 
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
