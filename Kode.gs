/**
 * LIBRARY SPENTIGS - DIGITAL LIBRARY SYSTEM (SMPN 39 SURABAYA)
 * Backend Google Apps Script (Code.gs) - Fitur Lengkap
 */

function getDB() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Library Spentigs - Perpustakaan Digital SMPN 39 Surabaya')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper: Membersihkan input teks agar aman dari formula injection di Google Sheets
 * @param {any} val
 * @return {string}
 */
function sanitizeSheetText(val) {
  const str = val === null || val === undefined ? '' : String(val).trim();
  if (/^[=+\-@]/.test(str)) {
    return "'" + str;
  }
  return str;
}

/**
 * Helper: Memformat ID / NIS / NIP / Tanggal agar tersimpan sebagai teks murni di Sheets
 * @param {any} val
 * @return {string}
 */
function formatAsTextCell(val) {
  const clean = sanitizeSheetText(val);
  if (!clean) return '';
  return clean.startsWith("'") ? clean : "'" + clean;
}

/**
 * Helper: Memverifikasi apakah username & password memiliki role ADMIN
 * @param {string} adminUsername
 * @param {string} adminPassword
 * @return {boolean}
 */
function verifyAdminAuth(adminUsername, adminPassword) {
  try {
    const ss = getDB();
    const sheetUsers = ss.getSheetByName('Users');
    if (!sheetUsers) return false;

    const data = sheetUsers.getDataRange().getDisplayValues();
    const inputUser = String(adminUsername || '').trim().toLowerCase();
    const inputPass = String(adminPassword || '').trim();
    if (!inputUser || !inputPass) return false;

    for (let i = 1; i < data.length; i++) {
      const dbUser = String(data[i][0] || '').trim().toLowerCase();
      const dbPass = String(data[i][1] || '').trim();
      const dbRole = String(data[i][3] || '').trim().toUpperCase();
      if (dbUser === inputUser && dbPass === inputPass && dbRole === 'ADMIN') {
        return true;
      }
    }
    return false;
  } catch (err) {
    Logger.log('verifyAdminAuth error: ' + err.toString());
    return false;
  }
}

/**
 * Inisialisasi Struktur Database Lengkap di Google Sheets
 * Jalankan fungsi ini sekali secara manual saat setup atau pembaruan struktur!
 */
function setupDatabase() {
  const ss = getDB();
  
  // 1. Setup Sheet Buku
  let sheetBuku = ss.getSheetByName('Buku');
  if (!sheetBuku) {
    sheetBuku = ss.insertSheet('Buku');
    sheetBuku.appendRow(['ID_Buku', 'Judul', 'Penulis', 'Kategori', 'Tahun', 'Stok', 'Cover_URL', 'Deskripsi']);
    sheetBuku.getRange('A1:H1').setFontWeight('bold').setBackground('#e2e8f0');
    
    sheetBuku.appendRow([
      'BK-001', 'Laskar Pelangi', 'Andrea Hirata', 'Fiksi', '2005', 5, 
      'https://images.unsplash.com/photo-1544947950-fa07a98d237f?auto=format&fit=crop&q=80&w=400',
      'Kisah inspiratif 10 anak dari keluarga miskin yang bersekolah di sebuah sekolah Muhammadiyah di Belitung.'
    ]);
    sheetBuku.appendRow([
      'BK-002', 'Filosofi Teras', 'Henry Manampiring', 'Pengembangan Diri', '2018', 3, 
      'https://images.unsplash.com/photo-1512820790803-83ca734da794?auto=format&fit=crop&q=80&w=400',
      'Pengantar filsafat Stoa untuk mengatasi emosi negatif dan menghasilkan mental yang tangguh.'
    ]);
    sheetBuku.appendRow([
      'BK-003', 'Bumi Manusia', 'Pramoedya Ananta Toer', 'Sejarah & Sastra', '1980', 4, 
      'https://images.unsplash.com/photo-1589829085413-56de8ae18c73?auto=format&fit=crop&q=80&w=400',
      'Roman sejarah yang mengambil latar belakang kebangkitan nasional Indonesia di awal abad ke-20.'
    ]);
    sheetBuku.appendRow([
      'BK-004', 'IPA Terpadu SMP Kelas VIII', 'Tim Kemdikbud', 'Pelajaran', '2022', 12, 
      'https://images.unsplash.com/photo-1532012197267-da84d127e765?auto=format&fit=crop&q=80&w=400',
      'Buku paket Ilmu Pengetahuan Alam Kurikulum Merdeka untuk siswa kelas 8 SMPN 39 Surabaya.'
    ]);
  }

  // 2. Setup Sheet Peminjaman
  let sheetPinjam = ss.getSheetByName('Peminjaman');
  if (!sheetPinjam) {
    sheetPinjam = ss.insertSheet('Peminjaman');
    sheetPinjam.appendRow(['ID_Pinjam', 'NIS_NIP', 'Nama_Peminjam', 'Kelas_Jabatan', 'ID_Buku', 'Judul_Buku', 'Tgl_Pinjam', 'Tgl_Kembali', 'Status']);
    sheetPinjam.getRange('A1:I1').setFontWeight('bold').setBackground('#e2e8f0');
  }

  // 3. Setup Sheet Users (Admin & Siswa/Pengunjung/Guru)
  let sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['Username_NIS', 'Password', 'Nama_Lengkap', 'Role', 'Kelas_Jabatan']);
    sheetUsers.getRange('A1:E1').setFontWeight('bold').setBackground('#e2e8f0');
    // Akun Default Admin, Guru & Contoh Siswa
    sheetUsers.appendRow(["'admin", "'admin123", 'Pustakawan Spentigs', 'ADMIN', 'Staf Perpustakaan']);
    sheetUsers.appendRow(["'2026001", "'siswa123", 'Ahmad Rizky Pratama', 'SISWA', 'VIII-A']);
    sheetUsers.appendRow(["'2026002", "'siswa123", 'Siti Nurhaliza', 'SISWA', 'IX-C']);
    sheetUsers.appendRow(["'198504122010011005", "'guru123", 'Bapak Hendra Wijaya, S.Pd.', 'GURU', 'Guru Matematika']);
  }

  // 4. Setup Sheet Request Pengadaan Buku
  let sheetReq = ss.getSheetByName('Requests');
  if (!sheetReq) {
    sheetReq = ss.insertSheet('Requests');
    sheetReq.appendRow(['ID_Request', 'Nama_Pengusul', 'NIS_Kelas', 'Judul_Buku', 'Penulis', 'Kategori', 'Alasan', 'Tanggal', 'Status']);
    sheetReq.getRange('A1:I1').setFontWeight('bold').setBackground('#e2e8f0');
    sheetReq.appendRow([
      'REQ-001', 'Ahmad Rizky Pratama', '2026001 (VIII-A)', 'Negeri 5 Menara', 'Ahmad Fuadi', 'Fiksi', 
      'Sangat bagus untuk memotivasi teman-teman kelas 8 dan 9 dalam belajar.', '2026-11-20', 'Menunggu'
    ]);
  }

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
  
  SpreadsheetApp.flush();
  return "Database Library Spentigs (SMPN 39 Surabaya) berhasil diinisialisasi!";
}

/**
 * Mengambil seluruh data untuk inisialisasi aplikasi
 */
function getInitialData() {
  try {
    const ss = getDB();
    let sheetBuku = ss.getSheetByName('Buku');
    let sheetPinjam = ss.getSheetByName('Peminjaman');
    let sheetUsers = ss.getSheetByName('Users');
    let sheetReq = ss.getSheetByName('Requests');
    
    if (!sheetBuku || !sheetPinjam || !sheetUsers || !sheetReq) {
      setupDatabase();
      sheetBuku = ss.getSheetByName('Buku');
      sheetPinjam = ss.getSheetByName('Peminjaman');
      sheetUsers = ss.getSheetByName('Users');
      sheetReq = ss.getSheetByName('Requests');
    }

    const booksData = sheetBuku.getDataRange().getDisplayValues();
    const loansData = sheetPinjam.getDataRange().getDisplayValues();
    const usersData = sheetUsers.getDataRange().getDisplayValues();
    const reqData = sheetReq.getDataRange().getDisplayValues();

    const books = [];
    for (let i = 1; i < booksData.length; i++) {
      if (booksData[i][0] && String(booksData[i][0]).trim() !== '') {
        const parsedStok = parseInt(booksData[i][5], 10);
        books.push({
          id: String(booksData[i][0]).trim(),
          judul: String(booksData[i][1]),
          penulis: String(booksData[i][2]),
          kategori: String(booksData[i][3]),
          tahun: String(booksData[i][4]),
          stok: isNaN(parsedStok) ? 0 : parsedStok,
          cover: String(booksData[i][6]).trim() || 'https://images.unsplash.com/photo-1543002588-bfa74002ed7e?auto=format&fit=crop&q=80&w=400',
          deskripsi: String(booksData[i][7]) || 'Tidak ada deskripsi.'
        });
      }
    }

    const loans = [];
    for (let i = 1; i < loansData.length; i++) {
      if (loansData[i][0] && String(loansData[i][0]).trim() !== '') {
        loans.push({
          id: String(loansData[i][0]).trim(),
          nis: String(loansData[i][1]).trim(),
          nama: String(loansData[i][2]),
          kelas: String(loansData[i][3]),
          bookId: String(loansData[i][4]).trim(),
          judulBuku: String(loansData[i][5]),
          tglPinjam: String(loansData[i][6]),
          tglKembali: String(loansData[i][7]),
          status: String(loansData[i][8]).trim()
        });
      }
    }

    // Daftar user tanpa menyertakan password untuk keamanan
    const users = [];
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][0] && String(usersData[i][0]).trim() !== '') {
        users.push({
          username: String(usersData[i][0]).trim(),
          nama: String(usersData[i][2]),
          role: String(usersData[i][3]).trim(),
          kelas: String(usersData[i][4])
        });
      }
    }

    const requests = [];
    for (let i = 1; i < reqData.length; i++) {
      if (reqData[i][0] && String(reqData[i][0]).trim() !== '') {
        requests.push({
          id: String(reqData[i][0]).trim(),
          pengusul: String(reqData[i][1]),
          nisKelas: String(reqData[i][2]),
          judul: String(reqData[i][3]),
          penulis: String(reqData[i][4]),
          kategori: String(reqData[i][5]),
          alasan: String(reqData[i][6]),
          tanggal: String(reqData[i][7]),
          status: String(reqData[i][8]).trim()
        });
      }
    }

    return {
      status: 'success',
      books: books,
      loans: loans.reverse(),
      users: users,
      requests: requests.reverse()
    };
  } catch (error) {
    Logger.log('getInitialData error: ' + error.toString());
    return { status: 'error', message: error.toString() };
  }
}

/**
 * Login Terpadu (Admin, Guru & Siswa/Pengunjung)
 */
function loginUser(username, password) {
  try {
    if (!username || !password) {
      return { status: 'error', message: 'Username/NIS dan Password wajib diisi.' };
    }

    const ss = getDB();
    let sheetUsers = ss.getSheetByName('Users');
    if (!sheetUsers) {
      setupDatabase();
      sheetUsers = ss.getSheetByName('Users');
    }

    const data = sheetUsers.getDataRange().getDisplayValues();
    const inputUser = String(username).trim().toLowerCase();
    const inputPass = String(password).trim();

    for (let i = 1; i < data.length; i++) {
      const dbUser = String(data[i][0]).trim().toLowerCase();
      const dbPass = String(data[i][1]).trim();
      if (dbUser === inputUser && dbPass === inputPass) {
        return {
          status: 'success',
          user: {
            username: String(data[i][0]).trim(),
            nama: String(data[i][2]),
            role: String(data[i][3]).trim().toUpperCase(),
            kelas: String(data[i][4])
          }
        };
      }
    }
    return { status: 'error', message: 'Username/NIS atau Password salah!' };
  } catch (e) {
    Logger.log('loginUser error: ' + e.toString());
    return { status: 'error', message: 'Gagal memproses login: ' + e.toString() };
  }
}

/**
 * Tambah Akun Siswa / Guru / Pengguna Baru (Admin)
 */
function addUserAccount(userObj, auth) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (auth && !verifyAdminAuth(auth.username, auth.password)) {
      return { status: 'error', message: 'Akses ditolak: kredensial Admin tidak valid.' };
    }

    const newNis = sanitizeSheetText(userObj && userObj.username);
    const newNama = sanitizeSheetText(userObj && userObj.nama);
    const newKelas = sanitizeSheetText(userObj && userObj.kelas);
    const newPass = sanitizeSheetText((userObj && userObj.password) || 'spentigs123');
    const newRole = sanitizeSheetText((userObj && userObj.role) || 'SISWA').toUpperCase();

    if (!newNis || !newNama || !newPass) {
      return { status: 'error', message: 'NIS/Username, Nama Lengkap, dan Password wajib diisi.' };
    }

    const ss = getDB();
    let sheetUsers = ss.getSheetByName('Users');
    if (!sheetUsers) {
      setupDatabase();
      sheetUsers = ss.getSheetByName('Users');
    }

    const data = sheetUsers.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === newNis.toLowerCase()) {
        return { status: 'error', message: 'NIS/Username "' + newNis + '" sudah terdaftar!' };
      }
    }

    sheetUsers.appendRow([
      formatAsTextCell(newNis),
      formatAsTextCell(newPass),
      newNama,
      ['SISWA', 'GURU', 'ADMIN'].includes(newRole) ? newRole : 'SISWA',
      newKelas || 'Umum'
    ]);

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Akun berhasil ditambahkan!' };
  } catch (e) {
    Logger.log('addUserAccount error: ' + e.toString());
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Import Massal Akun Siswa / Guru dari Excel/CSV (Admin)
 */
function importUsersBulk(usersArray, auth) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (auth && !verifyAdminAuth(auth.username, auth.password)) {
      return { status: 'error', message: 'Akses ditolak: kredensial Admin tidak valid.' };
    }

    if (!Array.isArray(usersArray) || usersArray.length === 0) {
      return { status: 'error', message: 'Tidak ada data akun yang dikirim untuk diimpor.' };
    }

    const ss = getDB();
    let sheetUsers = ss.getSheetByName('Users');
    if (!sheetUsers) {
      setupDatabase();
      sheetUsers = ss.getSheetByName('Users');
    }

    const existingData = sheetUsers.getDataRange().getDisplayValues();
    const existingSet = new Set();
    for (let i = 1; i < existingData.length; i++) {
      const uname = String(existingData[i][0]).trim().toLowerCase();
      if (uname) existingSet.add(uname);
    }

    let insertedCount = 0;
    let skippedCount = 0;

    usersArray.forEach(item => {
      const cleanUser = sanitizeSheetText(item && item.username);
      const cleanName = sanitizeSheetText(item && item.nama);
      if (!cleanUser || !cleanName) return;

      const lowerUser = cleanUser.toLowerCase();
      if (existingSet.has(lowerUser)) {
        skippedCount++;
        return;
      }

      let cleanRole = sanitizeSheetText((item && item.role) || 'SISWA').toUpperCase();
      if (!['GURU', 'SISWA', 'ADMIN'].includes(cleanRole)) {
        cleanRole = 'SISWA';
      }
      const cleanKelas = sanitizeSheetText(item && item.kelas) || (cleanRole === 'GURU' ? 'Guru' : 'Siswa');
      const cleanPass = sanitizeSheetText((item && item.password) || 'spentigs123') || 'spentigs123';

      sheetUsers.appendRow([
        formatAsTextCell(cleanUser),
        formatAsTextCell(cleanPass),
        cleanName,
        cleanRole,
        cleanKelas
      ]);
      existingSet.add(lowerUser);
      insertedCount++;
    });

    SpreadsheetApp.flush();
    return {
      status: 'success',
      message: `Import selesai: ${insertedCount} akun baru ditambahkan` + (skippedCount > 0 ? `, ${skippedCount} akun dilewati karena sudah terdaftar.` : '.')
    };
  } catch (e) {
    Logger.log('importUsersBulk error: ' + e.toString());
    return { status: 'error', message: 'Gagal mengimpor akun: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hapus Akun Siswa / Guru (Admin)
 */
function deleteUserAccount(username, auth) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (auth && !verifyAdminAuth(auth.username, auth.password)) {
      return { status: 'error', message: 'Akses ditolak: kredensial Admin tidak valid.' };
    }

    const targetUser = String(username || '').trim();
    if (!targetUser) {
      return { status: 'error', message: 'Username/NIS tidak valid.' };
    }
    if (targetUser.toLowerCase() === 'admin') {
      return { status: 'error', message: 'Akun utama admin tidak dapat dihapus.' };
    }

    const ss = getDB();
    const sheetUsers = ss.getSheetByName('Users');
    if (!sheetUsers) {
      return { status: 'error', message: 'Sheet Users tidak ditemukan.' };
    }

    const data = sheetUsers.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetUser.toLowerCase()) {
        sheetUsers.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { status: 'success', message: 'Akun pengguna berhasil dihapus.' };
      }
    }
    return { status: 'error', message: 'Akun tidak ditemukan.' };
  } catch (e) {
    Logger.log('deleteUserAccount error: ' + e.toString());
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menambah Buku Baru (Admin)
 */
function addBook(bookObj, auth) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (auth && !verifyAdminAuth(auth.username, auth.password)) {
      return { status: 'error', message: 'Akses ditolak: kredensial Admin tidak valid.' };
    }

    const judul = sanitizeSheetText(bookObj && bookObj.judul);
    const penulis = sanitizeSheetText(bookObj && bookObj.penulis);
    const kategori = sanitizeSheetText(bookObj && bookObj.kategori);
    if (!judul || !penulis || !kategori) {
      return { status: 'error', message: 'Judul, Penulis, dan Kategori buku wajib diisi.' };
    }

    const ss = getDB();
    let sheet = ss.getSheetByName('Buku');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Buku');
    }

    const customId = bookObj && bookObj.id && String(bookObj.id).trim() !== ''
      ? sanitizeSheetText(bookObj.id)
      : 'BK-' + Math.floor(100 + Math.random() * 900) + '-' + new Date().getSeconds();

    const stokVal = Math.max(0, parseInt(bookObj && bookObj.stok, 10) || 1);
    const coverVal = bookObj && bookObj.cover && String(bookObj.cover).trim() !== ''
      ? String(bookObj.cover).trim()
      : 'https://images.unsplash.com/photo-1543002588-bfa74002ed7e?auto=format&fit=crop&q=80&w=400';

    sheet.appendRow([
      customId,
      judul,
      penulis,
      kategori,
      sanitizeSheetText(bookObj && bookObj.tahun) || String(new Date().getFullYear()),
      stokVal,
      coverVal,
      sanitizeSheetText(bookObj && bookObj.deskripsi) || 'Tidak ada deskripsi.'
    ]);

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Buku berhasil ditambahkan dengan ID: ' + customId };
  } catch (error) {
    Logger.log('addBook error: ' + error.toString());
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus Buku (Admin)
 */
function deleteBook(bookId, auth) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (auth && !verifyAdminAuth(auth.username, auth.password)) {
      return { status: 'error', message: 'Akses ditolak: kredensial Admin tidak valid.' };
    }

    const ss = getDB();
    const sheet = ss.getSheetByName('Buku');
    if (!sheet) {
      return { status: 'error', message: 'Sheet Buku tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getDisplayValues();
    const targetId = String(bookId || '').trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { status: 'success', message: 'Buku berhasil dihapus.' };
      }
    }
    return { status: 'error', message: 'Buku tidak ditemukan.' };
  } catch (error) {
    Logger.log('deleteBook error: ' + error.toString());
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Proses Peminjaman Buku (Bisa dilakukan oleh Siswa/Guru yang login maupun Admin)
 */
function borrowBook(loanData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const ss = getDB();
    const sheetBuku = ss.getSheetByName('Buku');
    const sheetPinjam = ss.getSheetByName('Peminjaman');
    
    const booksData = sheetBuku.getDataRange().getDisplayValues();
    const targetBookId = String(loanData.bookId).trim();
    
    let bookRowIndex = -1;
    let currentStock = 0;
    let bookTitle = '';

    for (let i = 1; i < booksData.length; i++) {
      if (String(booksData[i][0]).trim().toLowerCase() === targetBookId.toLowerCase()) {
        bookRowIndex = i + 1;
        bookTitle = String(booksData[i][1]);
        currentStock = parseInt(booksData[i][5], 10);
        if (isNaN(currentStock)) currentStock = 0;
        break;
      }
    }

    if (bookRowIndex === -1) {
      return { status: 'error', message: 'Kode/ID Buku "' + targetBookId + '" tidak ditemukan.' };
    }
    if (currentStock <= 0) {
      return { status: 'error', message: 'Maaf, stok buku "' + bookTitle + '" sedang habis.' };
    }

    const newStock = currentStock - 1;
    sheetBuku.getRange(bookRowIndex, 6).setValue(newStock);

    const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
    const today = new Date();
    const dueDate = new Date();
    dueDate.setDate(today.getDate() + 7);
    
    const tglPinjamStr = Utilities.formatDate(today, tz, "yyyy-MM-dd");
    const tglKembaliStr = Utilities.formatDate(dueDate, tz, "yyyy-MM-dd");
    const loanId = 'TRX-' + Utilities.formatDate(today, tz, "yyMMdd") + '-' + Math.floor(100 + Math.random() * 900);

    sheetPinjam.appendRow([
      loanId,
      formatAsTextCell(loanData.nis),
      sanitizeSheetText(loanData.nama),
      sanitizeSheetText(loanData.kelas),
      booksData[bookRowIndex - 1][0],
      bookTitle,
      formatAsTextCell(tglPinjamStr),
      formatAsTextCell(tglKembaliStr),
      'Dipinjam'
    ]);

    SpreadsheetApp.flush();

    return { 
      status: 'success', 
      message: 'Peminjaman "' + bookTitle + '" berhasil! Batas kembali: ' + tglKembaliStr,
      updatedBookId: booksData[bookRowIndex - 1][0],
      newStock: newStock
    };
  } catch (error) {
    Logger.log('borrowBook error: ' + error.toString());
    return { status: 'error', message: 'Gagal memproses peminjaman: ' + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Proses Pengembalian Buku (Admin)
 */
function returnBook(loanId, auth) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (auth && !verifyAdminAuth(auth.username, auth.password)) {
      return { status: 'error', message: 'Akses ditolak: kredensial Admin tidak valid.' };
    }

    const ss = getDB();
    const sheetBuku = ss.getSheetByName('Buku');
    const sheetPinjam = ss.getSheetByName('Peminjaman');
    
    const loansData = sheetPinjam.getDataRange().getDisplayValues();
    const targetLoanId = String(loanId).trim();
    
    let loanRowIndex = -1;
    let bookId = '';
    let currentStatus = '';

    for (let i = 1; i < loansData.length; i++) {
      if (String(loansData[i][0]).trim() === targetLoanId) {
        loanRowIndex = i + 1;
        bookId = String(loansData[i][4]).trim();
        currentStatus = String(loansData[i][8]).trim();
        break;
      }
    }

    if (loanRowIndex === -1) {
      return { status: 'error', message: 'Data transaksi tidak ditemukan.' };
    }
    if (currentStatus === 'Dikembalikan') {
      return { status: 'error', message: 'Buku ini sudah dikembalikan sebelumnya.' };
    }

    sheetPinjam.getRange(loanRowIndex, 9).setValue('Dikembalikan');

    const booksData = sheetBuku.getDataRange().getDisplayValues();
    for (let i = 1; i < booksData.length; i++) {
      if (String(booksData[i][0]).trim() === bookId) {
        let currentStock = parseInt(booksData[i][5], 10);
        if (isNaN(currentStock)) currentStock = 0;
        sheetBuku.getRange(i + 1, 6).setValue(currentStock + 1);
        break;
      }
    }

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Buku berhasil dikembalikan dan stok bertambah!' };
  } catch (error) {
    Logger.log('returnBook error: ' + error.toString());
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mengajukan Request Pengadaan Buku Baru (Pengunjung / Siswa / Guru)
 */
function submitBookRequest(reqObj) {
  try {
    const judul = sanitizeSheetText(reqObj && reqObj.judul);
    const penulis = sanitizeSheetText(reqObj && reqObj.penulis);
    const pengusul = sanitizeSheetText(reqObj && reqObj.pengusul);
    if (!judul || !penulis || !pengusul) {
      return { status: 'error', message: 'Nama pengusul, judul buku, dan penulis wajib diisi.' };
    }

    const ss = getDB();
    let sheetReq = ss.getSheetByName('Requests');
    if (!sheetReq) {
      setupDatabase();
      sheetReq = ss.getSheetByName('Requests');
    }

    const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const reqId = 'REQ-' + Math.floor(100 + Math.random() * 900);

    sheetReq.appendRow([
      reqId,
      pengusul,
      sanitizeSheetText(reqObj && reqObj.nisKelas) || '-',
      judul,
      penulis,
      sanitizeSheetText(reqObj && reqObj.kategori) || 'Umum',
      sanitizeSheetText(reqObj && reqObj.alasan) || '-',
      formatAsTextCell(todayStr),
      'Menunggu'
    ]);

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Usulan pengadaan buku berhasil dikirim ke Pustakawan!' };
  } catch (e) {
    Logger.log('submitBookRequest error: ' + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

/**
 * Memperbarui Status Request Pengadaan Buku (Admin)
 */
function updateRequestStatus(reqId, newStatus, auth) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (auth && !verifyAdminAuth(auth.username, auth.password)) {
      return { status: 'error', message: 'Akses ditolak: kredensial Admin tidak valid.' };
    }

    const allowedStatus = ['Menunggu', 'Disetujui', 'Ditolak'];
    const cleanStatus = String(newStatus || '').trim();
    if (!allowedStatus.includes(cleanStatus)) {
      return { status: 'error', message: 'Status usulan tidak valid.' };
    }

    const ss = getDB();
    const sheetReq = ss.getSheetByName('Requests');
    if (!sheetReq) {
      return { status: 'error', message: 'Sheet Requests tidak ditemukan.' };
    }

    const data = sheetReq.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(reqId || '').trim()) {
        sheetReq.getRange(i + 1, 9).setValue(cleanStatus);
        SpreadsheetApp.flush();
        return { status: 'success', message: 'Status usulan buku diperbarui menjadi: ' + cleanStatus };
      }
    }
    return { status: 'error', message: 'Data usulan tidak ditemukan.' };
  } catch (e) {
    Logger.log('updateRequestStatus error: ' + e.toString());
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
