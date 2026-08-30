/**
 * APLIKASI KEUANGANKU - Backend (Code.gs)
 * Google Apps Script + Google Sheets Database
 */

const APP_NAME    = 'DB_KeuanganKu';
const APP_VERSION = '3.80'; // AUTO-UPDATED by deploy.sh — jangan edit manual

// ── LISENSI ──────────────────────────────────────────────────
// Email owner diambil dari ScriptProperties agar tidak hardcoded di source code
// Set di GAS editor: Project Settings → Script Properties → tambah OWNER_EMAIL = emailmu
const OWNER_EMAIL = PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') || '';

// Lisensi disimpan di ScriptProperties (bukan Google Sheet)
// agar bisa dibaca semua user tanpa perlu akses spreadsheet pemilik.
// Key: 'LIC_email@domain.com'  Value: JSON {tier, expiresAt, activatedAt}

function getLicenseInfo() {
  const email = getUserEmail_();
  if (!email || email === 'anonymous') return { tier:'free', expired:false, daysLeft:null, email:'' };
  if (email.toLowerCase() === OWNER_EMAIL.toLowerCase())
    return { tier:'lifetime', expired:false, daysLeft:null, email };

  try {
    const raw = PropertiesService.getScriptProperties().getProperty('LIC_' + email.toLowerCase());
    if (!raw) return { tier:'free', expired:false, daysLeft:null, email };
    const lic = JSON.parse(raw);
    if (lic.tier === 'lifetime') return { tier:'lifetime', expired:false, daysLeft:null, email };
    const expDate = new Date(lic.expiresAt);
    const now = new Date();
    const expired = expDate < now;
    const daysLeft = Math.ceil((expDate - now) / 86400000);
    return { tier: expired ? 'free' : lic.tier, expired, daysLeft: expired ? 0 : daysLeft, email };
  } catch(e) {}
  return { tier:'free', expired:false, daysLeft:null, email };
}

// Dipanggil admin dari GAS editor: activateLicense('user@gmail.com','pro',30)
// Untuk lifetime: activateLicense('user@gmail.com','lifetime',0)
function activateLicense(email, tier, durationDays) {
  const sp = PropertiesService.getScriptProperties();
  const today = new Date().toISOString().split('T')[0];
  const expiresAt = tier === 'lifetime' ? '9999-12-31'
    : new Date(Date.now() + (durationDays || 30) * 86400000).toISOString().split('T')[0];
  sp.setProperty('LIC_' + email.toLowerCase(), JSON.stringify({ tier, expiresAt, activatedAt: today }));
  Logger.log('✅ Lisensi ' + tier + ' untuk ' + email + ' aktif sampai ' + expiresAt);
  return 'OK: ' + email + ' → ' + tier + ' s/d ' + expiresAt;
}

// Dipanggil admin untuk melihat daftar lisensi
function getLicenseList() {
  const sp = PropertiesService.getScriptProperties();
  const all = sp.getProperties();
  const list = [['Email','Tier','Aktif Sampai','Tanggal Aktifasi']];
  Object.entries(all)
    .filter(([k]) => k.startsWith('LIC_'))
    .forEach(([k, v]) => {
      try {
        const lic = JSON.parse(v);
        list.push([k.replace('LIC_',''), lic.tier, lic.expiresAt, lic.activatedAt||'-']);
      } catch(e) {}
    });
  Logger.log(JSON.stringify(list));
  return list;
}

// Hapus lisensi: revokeLicense('user@gmail.com')
function revokeLicense(email) {
  PropertiesService.getScriptProperties().deleteProperty('LIC_' + email.toLowerCase());
  Logger.log('❌ Lisensi ' + email + ' dicabut');
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

// ==================== BATCH SAVE (untuk import CSV) ====================

// Simpan banyak transaksi sekaligus dalam 1 GAS call — jauh lebih cepat dari satu-per-satu
function saveTransaksiBatch(records) {
  const ss = getDB();
  if (!ss) {
    setupDB();
    return saveTransaksiBatch(records); // retry setelah setup
  }
  const sheet = ss.getSheetByName('Transaksi');
  if (!sheet) return { success: false, saved: 0 };

  const schema = SCHEMAS['Transaksi'];
  const rows = records.map(rec => {
    if (!rec.ID) rec.ID = Utilities.getUuid();
    return schema.map(col => rec[col] !== undefined ? rec[col] : '');
  });

  if (rows.length === 0) return { success: true, saved: 0 };

  // Tulis semua baris sekaligus
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, schema.length).setValues(rows);

  // Update saldo dompet: hitung net per dompet lalu apply sekali
  const netPerDompet = {};
  records.forEach(rec => {
    const dompet = rec.DompetAsal;
    if (!dompet) return;
    const nom = parseFloat(rec.Nominal) || 0;
    const sign = rec.Jenis === 'Pemasukan' ? 1 : -1;
    netPerDompet[dompet] = (netPerDompet[dompet] || 0) + sign * nom;
  });

  const dompetSheet = ss.getSheetByName('Dompet');
  if (dompetSheet && dompetSheet.getLastRow() > 1) {
    const dData = dompetSheet.getDataRange().getValues();
    const dHeaders = dData[0];
    const iNama = dHeaders.indexOf('Nama');
    const iSaldo = dHeaders.indexOf('SaldoSaatIni');
    for (let i = 1; i < dData.length; i++) {
      const nama = dData[i][iNama];
      if (netPerDompet[nama] !== undefined) {
        const saldoBaru = (parseFloat(dData[i][iSaldo]) || 0) + netPerDompet[nama];
        dompetSheet.getRange(i + 1, iSaldo + 1).setValue(saldoBaru);
      }
    }
  }

  return { success: true, saved: rows.length };
}

// ==================== BULK FIELD UPDATE ====================

// Bulk-update satu field di banyak baris Transaksi sekaligus.
// Contoh: updateTransaksiField(['id1','id2'], 'DompetAsal', 'Jago Syariah')
function updateTransaksiField(ids, field, value) {
  if (!ids || !ids.length || !field) return { success: false, error: 'Parameter tidak valid' };
  const ss = getDB();
  if (!ss) return { success: false, error: 'DB tidak ditemukan' };
  const sheet = ss.getSheetByName('Transaksi');
  if (!sheet || sheet.getLastRow() < 2) return { success: true, updated: 0 };

  const range = sheet.getDataRange();
  const data  = range.getValues();
  const headers = data[0];
  const iID    = headers.indexOf('ID');
  const iField = headers.indexOf(field);
  if (iID < 0 || iField < 0) return { success: false, error: 'Kolom tidak ditemukan: ' + field };

  const idSet = new Set(ids);
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    if (idSet.has(data[i][iID])) {
      data[i][iField] = value;
      updated++;
    }
  }
  if (updated > 0) range.setValues(data);
  Logger.log('updateTransaksiField: ' + field + '=' + value + ' → ' + updated + ' baris diperbarui');
  return { success: true, updated };
}

// Bulk-update beberapa field sekaligus untuk banyak baris.
// changes = { DompetAsal: 'Jago', Jenis: 'Pemasukan', Kategori: 'Gaji / Upah' }
// Hanya field yang ada di object changes yang diupdate.
function updateTransaksiFields(ids, changes) {
  if (!ids || !ids.length || !changes || !Object.keys(changes).length)
    return { success: false, error: 'Parameter tidak valid' };
  const ss = getDB();
  if (!ss) return { success: false, error: 'DB tidak ditemukan' };
  const sheet = ss.getSheetByName('Transaksi');
  if (!sheet || sheet.getLastRow() < 2) return { success: true, updated: 0 };

  const range   = sheet.getDataRange();
  const data    = range.getValues();
  const headers = data[0];
  const iID     = headers.indexOf('ID');
  if (iID < 0) return { success: false, error: 'Kolom ID tidak ditemukan' };

  // Petakan field name → index kolom
  const fieldIdx = {};
  for (const field of Object.keys(changes)) {
    const i = headers.indexOf(field);
    if (i >= 0) fieldIdx[field] = i;
  }

  const idSet = new Set(ids);
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    if (!idSet.has(data[i][iID])) continue;
    for (const [field, val] of Object.entries(changes)) {
      if (fieldIdx[field] !== undefined) data[i][fieldIdx[field]] = val;
    }
    updated++;
  }
  if (updated > 0) range.setValues(data);
  Logger.log('updateTransaksiFields: ' + JSON.stringify(changes) + ' → ' + updated + ' baris');
  return { success: true, updated };
}

// ==================== BATCH SAVE KATEGORI ====================

// Simpan banyak kategori sekaligus dalam 1 GAS call.
// items: [{ID:'', Jenis, Nama, ParentID}, ...]
function saveKategoriBatch(items) {
  if (!items || !items.length) return { success: false, error: 'Tidak ada data' };
  const ss = getDB();
  if (!ss) return { success: false, error: 'DB tidak ditemukan' };
  const sheet = ss.getSheetByName('Kategori');
  if (!sheet) return { success: false, error: 'Sheet Kategori tidak ditemukan' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = items.map(item => {
    const id = Utilities.getUuid();
    return headers.map(h => {
      if (h === 'ID') return id;
      return item[h] !== undefined ? item[h] : '';
    });
  });

  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
  }

  return { success: true, saved: rows.length };
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

// ==================== ADMIN PANEL (OWNER ONLY) ====================
// Semua fungsi di bawah ini hanya bisa dijalankan oleh OWNER_EMAIL.
// Dipanggil dari frontend via google.script.run.

function _isOwner() {
  if (!OWNER_EMAIL) return false;
  return getUserEmail_().toLowerCase() === OWNER_EMAIL.toLowerCase();
}

/** Ambil semua data lisensi untuk tabel admin */
function adminGetLicenseList() {
  if (!_isOwner()) return { error: 'Unauthorized' };
  const sp  = PropertiesService.getScriptProperties();
  const all = sp.getProperties();
  const now = new Date();
  const list = [];
  Object.entries(all)
    .filter(([k]) => k.startsWith('LIC_'))
    .forEach(([k, v]) => {
      try {
        const lic     = JSON.parse(v);
        const email   = k.replace('LIC_', '');
        const expDate = new Date(lic.expiresAt);
        const expired = expDate < now;
        const daysLeft = expired ? 0 : Math.ceil((expDate - now) / 86400000);
        list.push({ email, tier: lic.tier, expiresAt: lic.expiresAt,
                    activatedAt: lic.activatedAt || '-', expired, daysLeft });
      } catch(e) {}
    });
  // Urutkan: yang akan expired dulu, lalu alfabet
  list.sort((a, b) => new Date(a.expiresAt) - new Date(b.expiresAt));
  return { success: true, list };
}

/** Aktifkan atau perpanjang lisensi Pro */
function adminActivateLicense(email, days) {
  if (!_isOwner()) return { error: 'Unauthorized' };
  if (!email || !days) return { error: 'Email dan durasi wajib diisi' };
  const result = activateLicense(email.trim().toLowerCase(), 'pro', parseInt(days));
  return { success: true, message: result };
}

/** Cabut lisensi */
function adminRevokeLicense(email) {
  if (!_isOwner()) return { error: 'Unauthorized' };
  revokeLicense(email.trim().toLowerCase());
  return { success: true, message: 'Lisensi ' + email + ' berhasil dicabut' };
}
