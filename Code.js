/**
 * NAMA FILE: Code.gs
 */

const APP_NAME = 'DB_Catatan_Keuangan_App';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aplikasi Catatan Keuangan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Tambahan fungsi untuk mengambil email user (Mencegah error saat loading)
function getUserProfile() {
  try {
    return Session.getActiveUser().getEmail() || "Pengguna";
  } catch(e) {
    return "Pengguna";
  }
}

function setupDB() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty('DB_ID');
  
  let ss;
  if (dbId) {
    try { ss = SpreadsheetApp.openById(dbId); } 
    catch(e) { ss = SpreadsheetApp.create(APP_NAME); dbId = ss.getId(); props.setProperty('DB_ID', dbId); }
  } else {
    ss = SpreadsheetApp.create(APP_NAME);
    dbId = ss.getId();
    props.setProperty('DB_ID', dbId);
  }
  
  const sheets = {
    'Dompet': ['ID', 'Nama', 'Kategori', 'SaldoAwal', 'SaldoSaatIni'],
    'Kategori': ['ID', 'Jenis', 'Nama', 'IndukID'],
    'Anggaran': ['ID', 'BulanTahun', 'Kategori', 'Nominal', 'Rollover'],
    'Transaksi': ['ID', 'Tanggal', 'Jenis', 'KategoriNama', 'Nominal', 'Biaya', 'Keterangan', 'DompetAsal'],
    'Transfer': ['ID', 'Tanggal', 'Jumlah', 'DariDompet', 'KeDompet'],
    'HutangPiutang': ['ID', 'Jenis', 'Nama', 'Nominal', 'Tanggal', 'Status'],
    'Investasi': ['ID', 'Nama', 'Nominal', 'Tanggal', 'ReturnRate']
  };
  
  for (let sheetName in sheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, sheets[sheetName].length).setValues([sheets[sheetName]]);
      sheet.getRange(1, 1, 1, sheets[sheetName].length).setFontWeight("bold").setBackground("#f3f4f6");
    }
  }
  return { success: true };
}

function getDB() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID'));
}

function getAppData() {
  const ss = getDB();
  const result = {};
  const sheetNames = ['Dompet', 'Kategori', 'Anggaran', 'Transaksi', 'Transfer', 'HutangPiutang', 'Investasi'];
  
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length > 0) {
        const headers = data.shift();
        result[name] = data.map(row => {
          let obj = {};
          headers.forEach((h, i) => obj[h] = row[i]);
          return obj;
        });
      } else {
        result[name] = [];
      }
    } else {
      result[name] = [];
    }
  });
  return result;
}

function saveRecord(tableName, recordObj) {
  const ss = getDB();
  const sheet = ss.getSheetByName(tableName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('ID');
  
  if (tableName === 'Transfer' && !recordObj['ID']) {
    updateWalletBalance(recordObj['DariDompet'], recordObj['Jumlah'], 'Pengeluaran');
    updateWalletBalance(recordObj['KeDompet'], recordObj['Jumlah'], 'Pemasukan');
  }

  // LOGIKA UPDATE JIKA ADA ID
  if (recordObj['ID']) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === recordObj['ID']) {
        if (tableName === 'Transaksi') {
           const oldJenis = data[i][headers.indexOf('Jenis')];
           const oldNominal = data[i][headers.indexOf('Nominal')];
           const oldDompet = data[i][headers.indexOf('DompetAsal')];
           const revertJenis = oldJenis === 'Pengeluaran' ? 'Pemasukan' : 'Pengeluaran';
           updateWalletBalance(oldDompet, oldNominal, revertJenis);
        }

        const rowData = headers.map(h => recordObj[h] !== undefined ? recordObj[h] : data[i][headers.indexOf(h)]);
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);

        if (tableName === 'Transaksi') {
           updateWalletBalance(recordObj['DompetAsal'], recordObj['Nominal'], recordObj['Jenis']);
        }
        return { success: true };
      }
    }
  }

  // LOGIKA INSERT JIKA ID TIDAK ADA
  recordObj['ID'] = Utilities.getUuid();
  sheet.appendRow(headers.map(h => recordObj[h] || ""));
  
  if (tableName === 'Transaksi') {
    updateWalletBalance(recordObj['DompetAsal'], recordObj['Nominal'], recordObj['Jenis']);
  }
  return { success: true };
}

function deleteRecord(tableName, id) {
  const ss = getDB();
  const sheet = ss.getSheetByName(tableName);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf('ID');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      if (tableName === 'Transaksi') {
        const oldJenis = data[i][data[0].indexOf('Jenis')];
        const oldNominal = data[i][data[0].indexOf('Nominal')];
        const oldDompet = data[i][data[0].indexOf('DompetAsal')];
        const revertJenis = oldJenis === 'Pengeluaran' ? 'Pemasukan' : 'Pengeluaran';
        updateWalletBalance(oldDompet, oldNominal, revertJenis);
      }
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Data tidak ditemukan' };
}

function updateWalletBalance(walletName, nominal, jenis) {
  const sheet = getDB().getSheetByName('Dompet');
  const data = sheet.getDataRange().getValues();
  const nameIdx = data[0].indexOf('Nama');
  const saldoIdx = data[0].indexOf('SaldoSaatIni');
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][nameIdx] === walletName) {
      let currentSaldo = parseFloat(data[i][saldoIdx]) || 0;
      if(jenis === 'Pemasukan') currentSaldo += parseFloat(nominal);
      else currentSaldo -= parseFloat(nominal);
      sheet.getRange(i + 1, saldoIdx + 1).setValue(currentSaldo);
      break;
    }
  }
}