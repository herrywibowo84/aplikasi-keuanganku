/**
 * APLIKASI KEUANGANKU - Backend (Code.gs)
 * Google Apps Script + Google Sheets Database
 */

const APP_NAME = 'DB_KeuanganKu';

// Schema spreadsheet
const SCHEMAS = {
  'Dompet':        ['ID', 'Nama', 'Kategori', 'SaldoAwal', 'SaldoSaatIni', 'Limit', 'IsCC'],
  'Kategori':      ['ID', 'Jenis', 'Nama'],
  'Anggaran':      ['ID', 'BulanTahun', 'Kategori', 'Nominal'],
  'Transaksi':     ['ID', 'Tanggal', 'Jenis', 'Kategori', 'Nominal', 'Biaya', 'Keterangan', 'DompetAsal'],
  'Transfer':      ['ID', 'Tanggal', 'DariDompet', 'KeDompet', 'Jumlah', 'Biaya', 'Catatan'],
  'HutangPiutang': ['ID', 'Jenis', 'Nama', 'Nominal', 'Tanggal', 'JatuhTempo', 'Status', 'Keterangan'],
  'Investasi':     ['ID', 'NamaInvestasi', 'JenisInvestasi', 'BeratGram', 'HargaBeliGram', 'TotalModal', 'NilaiSaatIni', 'ReturnRate', 'Tanggal']
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

function getDB() {
  const dbId = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!dbId) return null;
  try { return SpreadsheetApp.openById(dbId); }
  catch(e) { return null; }
}

function setupDB() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty('DB_ID');
  let ss;

  if (dbId) {
    try { ss = SpreadsheetApp.openById(dbId); }
    catch(e) { ss = null; }
  }

  if (!ss) {
    ss = SpreadsheetApp.create(APP_NAME);
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
      [Utilities.getUuid(), 'Pemasukan', 'Gaji / Upah'],
      [Utilities.getUuid(), 'Pemasukan', 'Bonus / THR'],
      [Utilities.getUuid(), 'Pemasukan', 'Usaha / Bisnis'],
      [Utilities.getUuid(), 'Pemasukan', 'Investasi'],
      [Utilities.getUuid(), 'Pemasukan', 'Lainnya'],
      [Utilities.getUuid(), 'Pengeluaran', 'Makanan & Minuman'],
      [Utilities.getUuid(), 'Pengeluaran', 'Belanja Dapur'],
      [Utilities.getUuid(), 'Pengeluaran', 'Transportasi'],
      [Utilities.getUuid(), 'Pengeluaran', 'Tagihan & Utilitas'],
      [Utilities.getUuid(), 'Pengeluaran', 'Kesehatan'],
      [Utilities.getUuid(), 'Pengeluaran', 'Pendidikan'],
      [Utilities.getUuid(), 'Pengeluaran', 'Hiburan'],
      [Utilities.getUuid(), 'Pengeluaran', 'Pakaian & Fashion'],
      [Utilities.getUuid(), 'Pengeluaran', 'Penyesuaian Saldo'],
      [Utilities.getUuid(), 'Pengeluaran', 'Lainnya']
    ];
    katSheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }

  // Remove default "Sheet1" if it still exists
  try {
    const defSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1');
    if (defSheet && ss.getSheets().length > 1) ss.deleteSheet(defSheet);
  } catch(e) {}

  return { success: true, dbId };
}

function resetAndSetupDB() {
  const props = PropertiesService.getScriptProperties();
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
  const empty = { Dompet: [], Kategori: [], Anggaran: [], Transaksi: [], Transfer: [], HutangPiutang: [], Investasi: [] };
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
  return result;
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
