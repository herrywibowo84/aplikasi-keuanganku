/**
 * APLIKASI KEUANGANKU - Backend (Code.gs)
 * Google Apps Script + Google Sheets Database
 */

const APP_NAME    = 'DB_KeuanganKu';
const APP_VERSION = '3.49'; // AUTO-UPDATED by deploy.sh — jangan edit manual

// ── LISENSI ──────────────────────────────────────────────────
// Ganti dengan email pemilik aplikasi — selalu Pro secara otomatis
const OWNER_EMAIL = 'herry.wibowo84@gmail.com';

// Ambil/buat spreadsheet admin lisensi (1 sheet untuk semua user)
function getAdminDB_() {
  const props = PropertiesService.getScriptProperties(); // admin = script-level, bukan per-user
  let id = props.getProperty('ADMIN_DB_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  // Buat baru jika belum ada
  const ss = SpreadsheetApp.create('KeuanganKu_AdminLisensi');
  id = ss.getId();
  props.setProperty('ADMIN_DB_ID', id);
  const sheet = ss.getSheets()[0];
  sheet.setName('Lisensi');
  const headers = ['Email','Tier','TanggalMulai','TanggalExpired','Catatan'];
  sheet.getRange(1,1,1,headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#1e40af').setFontColor('#fff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  // Tambah owner sebagai lifetime otomatis
  sheet.appendRow([OWNER_EMAIL,'lifetime',new Date().toISOString().split('T')[0],'9999-12-31','Owner otomatis']);
  return ss;
}

function getLicenseInfo() {
  const email = getUserEmail_();
  if (!email || email === 'anonymous') return { tier:'free', expired:false, daysLeft:null, email:'' };
  if (email === OWNER_EMAIL) return { tier:'lifetime', expired:false, daysLeft:null, email };

  try {
    const ss = getAdminDB_();
    const sheet = ss.getSheetByName('Lisensi');
    if (!sheet || sheet.getLastRow() <= 1) return { tier:'free', expired:false, daysLeft:null, email };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxEmail   = headers.indexOf('Email');
    const idxTier    = headers.indexOf('Tier');
    const idxExpired = headers.indexOf('TanggalExpired');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxEmail]).toLowerCase() === email.toLowerCase()) {
        const tier = data[i][idxTier] || 'free';
        const expStr = data[i][idxExpired];
        if (tier === 'lifetime') return { tier, expired:false, daysLeft:null, email };
        const expDate = new Date(expStr);
        const now = new Date();
        const expired = expDate < now;
        const daysLeft = Math.ceil((expDate - now) / 86400000);
        return { tier: expired ? 'free' : tier, expired, daysLeft: expired ? 0 : daysLeft, email };
      }
    }
  } catch(e) {}
  return { tier:'free', expired:false, daysLeft:null, email };
}

// Dipanggil admin dari GAS editor: activateLicense('user@gmail.com','pro',30)
function activateLicense(email, tier, durationDays) {
  const ss = getAdminDB_();
  const sheet = ss.getSheetByName('Lisensi');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxEmail = headers.indexOf('Email');
  const start = new Date();
  const end   = new Date(start); end.setDate(end.getDate() + (durationDays || 30));
  const endStr = end.toISOString().split('T')[0];
  const startStr = start.toISOString().split('T')[0];
  // Update jika sudah ada
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxEmail]).toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i+1, 1, 1, 5).setValues([[email, tier, startStr, endStr, 'Updated '+startStr]]);
      return 'Updated: '+email+' → '+tier+' s/d '+endStr;
    }
  }
  // Tambah baru
  sheet.appendRow([email, tier, startStr, endStr, 'Activated '+startStr]);
  return 'Activated: '+email+' → '+tier+' s/d '+endStr;
}

function getLicenseList() {
  try {
    const ss = getAdminDB_();
    const sheet = ss.getSheetByName('Lisensi');
    const data = sheet.getDataRange().getValues();
    return data;
  } catch(e) { return []; }
}

// Schema spreadsheet
const SCHEMAS = {
  'Dompet':        ['ID', 'Nama', 'Kategori', 'SaldoAwal', 'SaldoSaatIni', 'Limit', 'IsCC'],
  'Kategori':      ['ID', 'Jenis', 'Nama', 'ParentID'],
  'Anggaran':      ['ID', 'BulanTahun', 'Kategori', 'Nominal'],
  'Transaksi':     ['ID', 'Tanggal', 'Waktu', 'Jenis', 'Kategori', 'Nominal', 'Biaya', 'KategoriBiaya', 'Keterangan', 'DompetAsal'],
  'Transfer':      ['ID', 'Tanggal', 'DariDompet', 'KeDompet', 'Jumlah', 'Biaya', 'Catatan'],
  'HutangPiutang': ['ID', 'Jenis', 'Nama', 'Nominal', 'Tanggal', 'JatuhTempo', 'Status', 'Keterangan'],
  'Investasi':     ['ID', 'NamaInvestasi', 'JenisInvestasi', 'BeratGram', 'HargaBeliGram', 'TotalModal', 'NilaiSaatIni', 'ReturnRate', 'Tanggal'],
  'Recurring':     ['ID', 'Nama', 'Jenis', 'Kategori', 'Nominal', 'DompetAsal', 'Frekuensi', 'TanggalMulai', 'Aktif', 'TerakhirDijalankan']
};

// ==================== ENTRY POINT ====================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aplikasi Keuanganku')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== USER ====================

function getUserProfile() {
  try {
    return Session.getActiveUser().getEmail() || 'Pengguna';
  } catch(e) {
    return 'Pengguna';
  }
}

// ==================== DB SETUP ====================

// Ambil email user aktif (sebagai identifier unik per user)
function getUserEmail_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || 'anonymous';
  } catch(e) { return 'anonymous'; }
}

function getDB() {
  // getUserProperties() → terisolasi per-user, tidak bocor ke user lain
  const dbId = PropertiesService.getUserProperties().getProperty('DB_ID');
  if (!dbId) return null;
  try { return SpreadsheetApp.openById(dbId); }
  catch(e) { return null; }
}

function setupDB() {
  const props = PropertiesService.getUserProperties(); // per-user isolation
  let dbId = props.getProperty('DB_ID');
  let ss;

  if (dbId) {
    try { ss = SpreadsheetApp.openById(dbId); }
    catch(e) { ss = null; }
  }

  if (!ss) {
    // Beri nama spreadsheet dengan email user agar mudah dikelola di Drive owner
    const userEmail = getUserEmail_();
    const sheetName = APP_NAME + (userEmail !== 'anonymous' ? '_' + userEmail.split('@')[0] : '');
    ss = SpreadsheetApp.create(sheetName);
    dbId = ss.getId();
    props.setProperty('DB_ID', dbId);
  }

  // Create sheets that don't exist yet
  for (const sheetName in SCHEMAS) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = SCHEMAS[sheetName];
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }

  // Seed default categories if Kategori sheet is empty
  const katSheet = ss.getSheetByName('Kategori');
  if (katSheet && katSheet.getLastRow() <= 1) {
    const defaults = [
      [Utilities.getUuid(), 'Pemasukan', 'Gaji / Upah', ''],
      [Utilities.getUuid(), 'Pemasukan', 'Bonus / THR', ''],
      [Utilities.getUuid(), 'Pemasukan', 'Usaha / Bisnis', ''],
      [Utilities.getUuid(), 'Pemasukan', 'Investasi', ''],
      [Utilities.getUuid(), 'Pemasukan', 'Lainnya', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Makanan & Minuman', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Belanja Dapur', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Transportasi', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Tagihan & Utilitas', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Kesehatan', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Pendidikan', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Hiburan', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Pakaian & Fashion', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Penyesuaian Saldo', ''],
      [Utilities.getUuid(), 'Pengeluaran', 'Lainnya', '']
    ];
    katSheet.getRange(2, 1, defaults.length, 4).setValues(defaults);
    // Seed Biaya categories
    seedBiayaKategori_(ss);
  }

  // Remove default "Sheet1" if it still exists
  try {
    const defSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1');
    if (defSheet && ss.getSheets().length > 1) ss.deleteSheet(defSheet);
  } catch(e) {}

  return { success: true, dbId };
}

function resetAndSetupDB() {
  const props = PropertiesService.getUserProperties(); // per-user isolation
  const dbId = props.getProperty('DB_ID');

  if (dbId) {
    try {
      const ss = SpreadsheetApp.openById(dbId);
      const sheets = ss.getSheets();

      // Keep one sheet (can't delete all), rename it
      for (let i = 1; i < sheets.length; i++) ss.deleteSheet(sheets[i]);
      sheets[0].setName('_temp');
      sheets[0].clearContents();
    } catch(e) {
      props.deleteProperty('DB_ID');
    }
  }

  return setupDB();
}

// ==================== DATA READ ====================

function getAppData() {
  const ss = getDB();
  const empty = { Dompet: [], Kategori: [], Anggaran: [], Transaksi: [], Transfer: [], HutangPiutang: [], Investasi: [], Recurring: [] };
  if (!ss) return empty;

  const result = {};
  for (const name in SCHEMAS) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) {
      result[name] = [];
      continue;
    }
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    result[name] = data.slice(1)
      .filter(row => row.some(cell => cell !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
  }
  result._version = APP_VERSION;
  result._license = getLicenseInfo(); // tier, expired, daysLeft
  return result;
}

function getDatabaseUrl() {
  const ss = getDB();
  return ss ? ss.getUrl() : null;
}

// ==================== DATA WRITE ====================

function saveRecord(tableName, recordObj) {
  const ss = getDB();
  if (!ss) return { success: false, message: 'Database belum disetup. Silakan refresh halaman.' };

  const sheet = ss.getSheetByName(tableName);
  if (!sheet) return { success: false, message: 'Sheet tidak ditemukan: ' + tableName };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('ID');

  // UPDATE: if ID exists, find and update the row
  if (recordObj['ID']) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(recordObj['ID'])) {

        // Revert old Transaksi balance effect before update
        if (tableName === 'Transaksi') {
          const oldJenis = data[i][headers.indexOf('Jenis')];
          const oldNominal = parseFloat(data[i][headers.indexOf('Nominal')]) || 0;
          const oldDompet = data[i][headers.indexOf('DompetAsal')];
          updateWalletBalance_(ss, oldDompet, oldNominal, oldJenis === 'Pengeluaran' ? 'Pemasukan' : 'Pengeluaran');
        }

        const rowData = headers.map(h => (recordObj[h] !== undefined ? recordObj[h] : data[i][headers.indexOf(h)]));
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);

        // Apply new Transaksi balance effect
        if (tableName === 'Transaksi') {
          updateWalletBalance_(ss, recordObj['DompetAsal'], parseFloat(recordObj['Nominal']) || 0, recordObj['Jenis']);
        }

        return { success: true };
      }
    }
  }

  // INSERT: new record
  const newId = Utilities.getUuid();
  recordObj['ID'] = newId;

  // Special handling: Transfer auto-updates wallet balances
  if (tableName === 'Transfer') {
    updateWalletBalance_(ss, recordObj['DariDompet'], parseFloat(recordObj['Jumlah']) || 0, 'Pengeluaran');
    updateWalletBalance_(ss, recordObj['KeDompet'], parseFloat(recordObj['Jumlah']) || 0, 'Pemasukan');
  }

  const rowData = headers.map(h => (recordObj[h] !== undefined ? recordObj[h] : ''));
  sheet.appendRow(rowData);

  // Transaksi: update wallet balance on new transaction
  if (tableName === 'Transaksi') {
    updateWalletBalance_(ss, recordObj['DompetAsal'], parseFloat(recordObj['Nominal']) || 0, recordObj['Jenis']);
  }

  // New Dompet: set SaldoSaatIni = SaldoAwal
  if (tableName === 'Dompet') {
    const allData = sheet.getDataRange().getValues();
    const lastRow = allData.length;
    const saldoIdx = headers.indexOf('SaldoSaatIni');
    const saldoAwal = parseFloat(recordObj['SaldoAwal']) || 0;
    sheet.getRange(lastRow, saldoIdx + 1).setValue(saldoAwal);
  }

  return { success: true };
}

// ==================== DATA DELETE ====================

function deleteRecord(tableName, id) {
  const ss = getDB();
  if (!ss) return { success: false, message: 'DB tidak ditemukan' };

  const sheet = ss.getSheetByName(tableName);
  if (!sheet) return { success: false, message: 'Sheet tidak ada' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('ID');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id)) {

      // Revert Transaksi balance
      if (tableName === 'Transaksi') {
        const oldJenis = data[i][headers.indexOf('Jenis')];
        const oldNominal = parseFloat(data[i][headers.indexOf('Nominal')]) || 0;
        const oldDompet = data[i][headers.indexOf('DompetAsal')];
        updateWalletBalance_(ss, oldDompet, oldNominal, oldJenis === 'Pengeluaran' ? 'Pemasukan' : 'Pengeluaran');
      }

      // Revert Transfer balance
      if (tableName === 'Transfer') {
        const dari = data[i][headers.indexOf('DariDompet')];
        const ke = data[i][headers.indexOf('KeDompet')];
        const jumlah = parseFloat(data[i][headers.indexOf('Jumlah')]) || 0;
        updateWalletBalance_(ss, dari, jumlah, 'Pemasukan');
        updateWalletBalance_(ss, ke, jumlah, 'Pengeluaran');
      }

      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, message: 'Data tidak ditemukan' };
}

// ==================== INTERNAL HELPERS ====================

function updateWalletBalance_(ss, walletName, nominal, jenis) {
  if (!walletName || !nominal) return;
  const sheet = ss.getSheetByName('Dompet');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = headers.indexOf('Nama');
  const saldoIdx = headers.indexOf('SaldoSaatIni');

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameIdx] === walletName) {
      let saldo = parseFloat(data[i][saldoIdx]) || 0;
      saldo = jenis === 'Pemasukan' ? saldo + nominal : saldo - nominal;
      sheet.getRange(i + 1, saldoIdx + 1).setValue(saldo);
      break;
    }
  }
}

// ==================== HELPERS ====================

function seedBiayaKategori_(ss) {
  const sheet = ss.getSheetByName('Kategori');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const namaIdx = headers.indexOf('Nama');
  const names = data.slice(1).map(r => String(r[namaIdx]).trim());
  // Check if 'Jenis Biaya' parent already exists
  if (names.includes('Jenis Biaya')) return;
  const parentId = Utilities.getUuid();
  const biayaRows = [
    [parentId, 'Biaya', 'Jenis Biaya', ''],
    [Utilities.getUuid(), 'Biaya', 'Biaya Admin', parentId],
    [Utilities.getUuid(), 'Biaya', 'Biaya Transfer', parentId],
    [Utilities.getUuid(), 'Biaya', 'Kode Unik', parentId],
    [Utilities.getUuid(), 'Biaya', 'Lainnya', parentId]
  ];
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, biayaRows.length, 4).setValues(biayaRows);
}

// ==================== RECURRING TRANSACTIONS ====================

/**
 * Run all active recurring transactions that are due today.
 * Called by a daily time-driven trigger (set up via setupDailyTrigger).
 */
function runRecurringTransactions() {
  const ss = getDB();
  if (!ss) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const recSheet = ss.getSheetByName('Recurring');
  if (!recSheet || recSheet.getLastRow() <= 1) return;

  const recData = recSheet.getDataRange().getValues();
  const recHeaders = recData[0];
  const idIdx        = recHeaders.indexOf('ID');
  const aktifIdx     = recHeaders.indexOf('Aktif');
  const frekIdx      = recHeaders.indexOf('Frekuensi');
  const mulaiIdx     = recHeaders.indexOf('TanggalMulai');
  const terakhirIdx  = recHeaders.indexOf('TerakhirDijalankan');
  const jenisIdx     = recHeaders.indexOf('Jenis');
  const katIdx       = recHeaders.indexOf('Kategori');
  const nomIdx       = recHeaders.indexOf('Nominal');
  const dompetIdx    = recHeaders.indexOf('DompetAsal');
  const namaIdx      = recHeaders.indexOf('Nama');

  for (let i = 1; i < recData.length; i++) {
    const row = recData[i];
    if (!row[aktifIdx] || row[aktifIdx] === 'false' || row[aktifIdx] === false) continue;

    const frekuensi   = String(row[frekIdx]).toLowerCase();
    const tanggalMulai = row[mulaiIdx] ? new Date(row[mulaiIdx]) : null;
    const terakhir    = row[terakhirIdx] ? new Date(row[terakhirIdx]) : null;

    if (!tanggalMulai) continue;

    let isDue = false;
    if (frekuensi === 'harian') {
      isDue = !terakhir || daysBetween(terakhir, today) >= 1;
    } else if (frekuensi === 'mingguan') {
      isDue = !terakhir || daysBetween(terakhir, today) >= 7;
    } else if (frekuensi === 'bulanan') {
      isDue = !terakhir || (today.getDate() === tanggalMulai.getDate() &&
              (today.getFullYear() > terakhir.getFullYear() || today.getMonth() > terakhir.getMonth()));
    }

    if (!isDue) continue;

    // Create transaction
    const txObj = {
      ID: '', Tanggal: Utilities.formatDate(today, 'Asia/Jakarta', 'yyyy-MM-dd'),
      Waktu: '', Jenis: row[jenisIdx], Kategori: row[katIdx],
      Nominal: parseFloat(row[nomIdx]) || 0, Biaya: 0, KodeUnik: '',
      Keterangan: '[Auto] ' + String(row[namaIdx]),
      DompetAsal: row[dompetIdx]
    };
    saveRecord('Transaksi', txObj);

    // Update TerakhirDijalankan
    recSheet.getRange(i + 1, terakhirIdx + 1).setValue(
      Utilities.formatDate(today, 'Asia/Jakarta', 'yyyy-MM-dd')
    );
    console.log('Recurring run: ' + row[namaIdx]);
  }
}

function daysBetween(d1, d2) {
  const ms = Math.abs(d2.getTime() - d1.getTime());
  return Math.floor(ms / (1000 * 60 * 60 * 24));
}

/**
 * Create a daily trigger for runRecurringTransactions.
 * Run once manually from GAS editor (select setupDailyTrigger → Run).
 */
function setupDailyTrigger() {
  // Remove existing triggers for this function first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runRecurringTransactions') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create new daily trigger at 7am Jakarta time
  ScriptApp.newTrigger('runRecurringTransactions')
    .timeBased().everyDays(1).atHour(7).inTimezone('Asia/Jakarta').create();
  console.log('Daily trigger set for runRecurringTransactions at 07:00 WIB');
}

// ==================== MIGRATION ====================

/**
 * migrateToV2: Run once to:
 * 1. Rename KodeUnik → KategoriBiaya in Transaksi sheet
 * 2. Add ParentID column to Kategori sheet
 * 3. Seed Biaya sub-categories
 */
function migrateToV2() {
  const ss = getDB();
  if (!ss) { console.log('DB not found'); return; }

  // --- Migrate Transaksi: KodeUnik → KategoriBiaya ---
  const trxSheet = ss.getSheetByName('Transaksi');
  if (trxSheet) {
    const trxData = trxSheet.getDataRange().getValues();
    const trxHeaders = trxData[0].map(h => String(h).trim());
    const kuIdx = trxHeaders.indexOf('KodeUnik');
    if (kuIdx >= 0) {
      trxSheet.getRange(1, kuIdx + 1).setValue('KategoriBiaya');
      console.log('Transaksi: KodeUnik → KategoriBiaya pada kolom ' + (kuIdx + 1));
    } else if (trxHeaders.indexOf('KategoriBiaya') >= 0) {
      console.log('Transaksi: KategoriBiaya sudah ada, skip.');
    } else {
      console.log('Transaksi: kolom KodeUnik tidak ditemukan.');
    }
  }

  // --- Migrate Kategori: tambah kolom ParentID ---
  const katSheet = ss.getSheetByName('Kategori');
  if (katSheet) {
    const katData = katSheet.getDataRange().getValues();
    const katHeaders = katData[0].map(h => String(h).trim());
    if (!katHeaders.includes('ParentID')) {
      const newCol = katHeaders.length + 1;
      katSheet.getRange(1, newCol).setValue('ParentID');
      // Fill empty string for all existing rows
      if (katSheet.getLastRow() > 1) {
        katSheet.getRange(2, newCol, katSheet.getLastRow() - 1, 1).setValue('');
      }
      console.log('Kategori: kolom ParentID ditambahkan di kolom ' + newCol);
    } else {
      console.log('Kategori: ParentID sudah ada, skip.');
    }
    // Seed Biaya categories
    seedBiayaKategori_(ss);
    console.log('Kategori Biaya: seed selesai.');
  }

  console.log('migrateToV2 selesai.');
}

/**
 * Migrate Transaksi sheet from old 8-col schema to new 10-col schema.
 * Old: [ID, Tanggal, Jenis, KategoriNama, Nominal, Biaya, Keterangan, DompetAsal]
 * New: [ID, Tanggal, Waktu, Jenis, Kategori, Nominal, Biaya, KodeUnik, Keterangan, DompetAsal]
 * Run once via GAS editor (select migrateTransaksiSchema → Run).
 */
function migrateTransaksiSchema() {
  const ss = getDB();
  if (!ss) { console.log('DB not found'); return; }

  const sheet = ss.getSheetByName('Transaksi');
  if (!sheet) { console.log('Sheet Transaksi not found'); return; }

  const data = sheet.getDataRange().getValues();
  const oldHeaders = data[0].map(h => String(h).trim());
  const newHeaders = SCHEMAS['Transaksi'];

  console.log('Old headers (' + oldHeaders.length + '):', JSON.stringify(oldHeaders));
  console.log('New headers (' + newHeaders.length + '):', JSON.stringify(newHeaders));

  if (JSON.stringify(oldHeaders) === JSON.stringify(newHeaders)) {
    console.log('Headers already match — no migration needed.');
    return;
  }

  // Field rename map: old name → new name
  const fieldMap = { 'KategoriNama': 'Kategori', 'KategoriID': 'Kategori' };

  // Read & migrate all data rows
  const rows = data.slice(1).filter(r => r.some(c => c !== ''));
  const migratedRows = rows.map(row => {
    const obj = {};
    oldHeaders.forEach((h, i) => {
      const key = fieldMap[h] || h;
      obj[key] = row[i];
    });
    return obj;
  });

  // Clear sheet & rewrite with new headers
  sheet.clearContents();
  const headerRange = sheet.getRange(1, 1, 1, newHeaders.length);
  headerRange.setValues([newHeaders]);
  headerRange.setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (migratedRows.length > 0) {
    const rowData = migratedRows.map(obj =>
      newHeaders.map(h => (obj[h] !== undefined && obj[h] !== null ? obj[h] : ''))
    );
    sheet.getRange(2, 1, rowData.length, newHeaders.length).setValues(rowData);
  }

  console.log('Migration done. ' + migratedRows.length + ' rows migrated.');
  if (migratedRows.length > 0) {
    console.log('Sample row 1:', JSON.stringify(migratedRows[0]));
  }
}
