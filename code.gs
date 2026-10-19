const TX_SHEET = 'Transaksi';
const ACCOUNT_SHEET = 'Rekening';
const ADJ_SHEET = 'PenyesuaianSaldo';
const BUDGET_SHEET = 'Budget';
const CATEGORY_SHEET = 'Kategori';
const DEBT_SHEET = 'HutangPiutang';

const TX_HEADERS = [
  'Tanggal', 'Waktu Input', 'Waktu Transaksi', 'Rekening', 'Kategori',
  'Keterangan', 'Pemasukan', 'Pengeluaran', 'Saldo Rekening', 'Split ID',
  'Ref Hutang-Piutang'
];
const ACCOUNT_HEADERS = ['Nama Rekening', 'Jenis', 'Saldo Awal', 'Tanggal Saldo Awal', 'Aktif'];
const ADJ_HEADERS = ['Tanggal', 'Rekening', 'Saldo Sebenarnya', 'Saldo Sistem', 'Selisih', 'Keterangan'];
const BUDGET_HEADERS = [
  'Bulan', 'Total Uang Bulanan', 'Makan 45%', 'Operasional 30%',
  'Nabung 15%', 'Sedekah / Investasi 10%'
];
const CATEGORY_HEADERS = ['Kategori', 'Kelompok'];
const DEBT_HEADERS = [
  'ID', 'Tanggal', 'Jenis', 'Nama Orang', 'Rekening', 'Kategori', 'Nominal',
  'Keterangan', 'Jatuh Tempo', 'Status', 'Tanggal Lunas'
];

const GROUPS = [
  { key: 'makan', label: 'Uang Makan', percent: 0.45 },
  { key: 'operasional', label: 'Operasional', percent: 0.30 },
  { key: 'nabung', label: 'Nabung', percent: 0.15 },
  { key: 'sedekahInvestasi', label: 'Sedekah / Investasi', percent: 0.10 }
];
const INCOME = 'pemasukan';

const DEFAULT_CATEGORIES = [
  ['Uang dari Orang Tua', INCOME],
  ['Beasiswa', INCOME],
  ['Penghasilan', INCOME],
  ['Transfer Masuk', INCOME],
  ['Transfer Keluar', INCOME],
  ['Makanan', 'makan'],
  ['Minuman', 'makan'],
  ['Jajan', 'makan'],
  ['Transportasi', 'operasional'],
  ['Bensin', 'operasional'],
  ['Listrik', 'operasional'],
  ['Pulsa', 'operasional'],
  ['Kuota / Internet', 'operasional'],
  ['Servis', 'operasional'],
  ['Kos / Kontrakan', 'operasional'],
  ['Kuliah', 'operasional'],
  ['Nabung', 'nabung'],
  ['Tabungan', 'nabung'],
  ['Investasi', 'nabung'],
  ['Sedekah', 'sedekahInvestasi'],
  ['Donasi', 'sedekahInvestasi'],
  ['Meminjamkan Uang', 'sedekahInvestasi']
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Keuangan Mahasiswa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  sheet_(ss, TX_SHEET, TX_HEADERS);
  sheet_(ss, ACCOUNT_SHEET, ACCOUNT_HEADERS);
  sheet_(ss, ADJ_SHEET, ADJ_HEADERS);
  sheet_(ss, BUDGET_SHEET, BUDGET_HEADERS);
  sheet_(ss, DEBT_SHEET, DEBT_HEADERS);

  const cat = sheet_(ss, CATEGORY_SHEET, CATEGORY_HEADERS);

  if (cat.getLastRow() < 2) {
    cat.getRange(2, 1, DEFAULT_CATEGORIES.length, 2).setValues(DEFAULT_CATEGORIES);
  }
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (headers.some((h, i) => current[i] !== h)) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }

  sh.setFrozenRows(1);
  return sh;
}

function txSheet_() { return sheet_(SpreadsheetApp.getActiveSpreadsheet(), TX_SHEET, TX_HEADERS); }
function accSheet_() { return sheet_(SpreadsheetApp.getActiveSpreadsheet(), ACCOUNT_SHEET, ACCOUNT_HEADERS); }
function adjSheet_() { return sheet_(SpreadsheetApp.getActiveSpreadsheet(), ADJ_SHEET, ADJ_HEADERS); }
function budgetSheet_() { return sheet_(SpreadsheetApp.getActiveSpreadsheet(), BUDGET_SHEET, BUDGET_HEADERS); }
function catSheet_() { return sheet_(SpreadsheetApp.getActiveSpreadsheet(), CATEGORY_SHEET, CATEGORY_HEADERS); }
function debtSheet_() { return sheet_(SpreadsheetApp.getActiveSpreadsheet(), DEBT_SHEET, DEBT_HEADERS); }

function tz_() { return Session.getScriptTimeZone(); }
function fmtDate_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function fmtMonth_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM');
  return String(v || '').trim().substring(0, 7);
}
function fmtShort_(d) { return Utilities.formatDate(d, tz_(), 'dd MMM yyyy'); }
function fmtTime_(d) { return Utilities.formatDate(d, tz_(), 'HH:mm'); }
function monthNow_() { return fmtDate_(new Date()).substring(0, 7); }
function newId_(p) { return p + '-' + new Date().getTime().toString(36).toUpperCase(); }
function rp_(v) { return 'Rp' + Math.round(Number(v || 0)).toLocaleString('id-ID'); }
function dayStart_(s) { return new Date(s + 'T00:00:00'); }

function groupLabel_(key) {
  if (key === INCOME) return 'Pemasukan';
  const g = GROUPS.find(x => x.key === key);
  return g ? g.label : key;
}

function isValidGroup_(key) {
  return key === INCOME || GROUPS.some(g => g.key === key);
}

function isSplitDetail_(row) {
  return String(row[9] || '') !== '' && String(row[4]) !== 'Split Transaksi';
}

/* ================= KATEGORI ================= */

function getCategories() {
  const sh = catSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 2).getValues()
    .map((r, i) => ({
      rowNumber: i + 2,
      nama: String(r[0]).trim(),
      kelompok: String(r[1]).trim()
    }))
    .filter(c => c.nama && c.kelompok);
}

function findCategory_(nama) {
  return getCategories().find(c => c.nama === nama);
}

function addCategory(data) {
  try {
    const nama = String(data.nama || '').trim();
    const kelompok = String(data.kelompok || '').trim();

    if (!nama) throw new Error('Nama kategori wajib diisi.');
    if (!isValidGroup_(kelompok)) throw new Error('Kelompok wajib dipilih.');
    if (getCategories().some(c => c.nama.toLowerCase() === nama.toLowerCase())) {
      throw new Error('Kategori dengan nama tersebut sudah ada.');
    }

    catSheet_().appendRow([nama, kelompok]);

    return {
      success: true,
      message: `Kategori "${nama}" ditambahkan ke ${groupLabel_(kelompok)}.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateCategoryGroup(rowNumber, kelompok) {
  try {
    const sh = catSheet_();
    if (!rowNumber || rowNumber < 2 || rowNumber > sh.getLastRow()) {
      throw new Error('Kategori tidak valid.');
    }
    if (!isValidGroup_(kelompok)) throw new Error('Kelompok tidak valid.');

    const nama = sh.getRange(rowNumber, 1).getValue();
    sh.getRange(rowNumber, 2).setValue(kelompok);

    return {
      success: true,
      message: `Kategori "${nama}" dipindah ke ${groupLabel_(kelompok)}.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteCategory(rowNumber) {
  try {
    const sh = catSheet_();
    if (!rowNumber || rowNumber < 2 || rowNumber > sh.getLastRow()) {
      throw new Error('Kategori tidak valid.');
    }

    sh.deleteRow(rowNumber);

    return {
      success: true,
      message: 'Kategori dihapus. Transaksi lama tetap tersimpan.'
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/* ================= REKENING ================= */

function getAccounts_() {
  const sh = accSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => String(r[0]).trim() && String(r[4]).toUpperCase() === 'YA')
    .map(r => ({
      nama: String(r[0]).trim(),
      jenis: String(r[1]).trim() || 'Lainnya',
      saldoAwal: Number(r[2]) || 0,
      tanggalSaldoAwal: r[3] instanceof Date ? fmtDate_(r[3]) : String(r[3] || '')
    }));
}

function getAccountList() {
  return getAccounts_();
}

function addAccount(data) {
  try {
    const nama = String(data.nama || '').trim();
    const jenis = String(data.jenis || 'Lainnya').trim();
    const saldoAwal = Number(data.saldoAwal || 0);

    if (!nama) throw new Error('Nama rekening wajib diisi.');
    if (saldoAwal < 0) throw new Error('Saldo awal tidak boleh negatif.');
    if (!data.tanggal) throw new Error('Tanggal saldo awal wajib diisi.');
    if (getAccounts_().some(a => a.nama.toLowerCase() === nama.toLowerCase())) {
      throw new Error('Rekening dengan nama tersebut sudah ada.');
    }

    accSheet_().appendRow([nama, jenis, saldoAwal, dayStart_(data.tanggal), 'YA']);
    recalc_();

    return {
      success: true,
      message: `Rekening ${nama} ditambahkan dengan saldo ${rp_(saldoAwal)}.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function computeLedger_() {
  const accounts = getAccounts_();
  const balances = {};
  const starts = {};

  accounts.forEach(a => {
    balances[a.nama] = a.saldoAwal;
    starts[a.nama] = a.tanggalSaldoAwal ? dayStart_(a.tanggalSaldoAwal).getTime() : 0;
  });

  const tx = txSheet_();
  const txLast = tx.getLastRow();
  const entries = [];

  if (txLast >= 2) {
    tx.getRange(2, 1, txLast - 1, 11).getValues().forEach((r, i) => {
      if (isSplitDetail_(r)) return;

      const t = r[2] instanceof Date
        ? r[2]
        : (r[0] instanceof Date ? r[0] : new Date(r[0]));

      entries.push({
        type: 'tx',
        row: i + 2,
        time: t.getTime(),
        rekening: String(r[3]).trim(),
        inc: Number(r[6]) || 0,
        out: Number(r[7]) || 0
      });
    });
  }

  const adj = adjSheet_();
  const adjLast = adj.getLastRow();

  if (adjLast >= 2) {
    adj.getRange(2, 1, adjLast - 1, 6).getValues().forEach((r, i) => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);

      entries.push({
        type: 'adj',
        row: i + 2,
        time: d.getTime() + 86399000,
        rekening: String(r[1]).trim(),
        actual: Number(r[2]) || 0
      });
    });
  }

  entries.sort((a, b) =>
    a.time - b.time ||
    (a.type === b.type ? a.row - b.row : (a.type === 'tx' ? -1 : 1))
  );

  const saldoByRow = {};
  const adjBefore = {};

  entries.forEach(e => {
    if (!balances.hasOwnProperty(e.rekening)) return;

    if (e.time < starts[e.rekening]) {
      if (e.type === 'tx') saldoByRow[e.row] = '';
      return;
    }

    if (e.type === 'tx') {
      balances[e.rekening] += e.inc - e.out;
      saldoByRow[e.row] = balances[e.rekening];
    } else {
      adjBefore[e.row] = balances[e.rekening];
      balances[e.rekening] = e.actual;
    }
  });

  return { balances, saldoByRow, adjBefore, txLast, adjLast };
}

function recalc_() {
  const ledger = computeLedger_();

  if (ledger.txLast >= 2) {
    const out = [];
    for (let r = 2; r <= ledger.txLast; r++) {
      out.push([ledger.saldoByRow.hasOwnProperty(r) ? ledger.saldoByRow[r] : '']);
    }
    txSheet_().getRange(2, 9, out.length, 1).setValues(out);
  }

  if (ledger.adjLast >= 2) {
    const adj = adjSheet_();
    const vals = adj.getRange(2, 1, ledger.adjLast - 1, 6).getValues();

    const out = vals.map((r, i) => {
      const before = ledger.adjBefore.hasOwnProperty(i + 2) ? ledger.adjBefore[i + 2] : '';
      return [before, before === '' ? '' : (Number(r[2]) || 0) - before];
    });

    adj.getRange(2, 4, out.length, 2).setValues(out);
  }

  return ledger;
}

function saveBalanceAdjustment(data) {
  try {
    const rekening = String(data.rekening || '').trim();
    const saldo = Number(data.saldoSebenarnya);

    if (!data.tanggal) throw new Error('Tanggal penyesuaian wajib diisi.');
    if (isNaN(saldo) || saldo < 0) throw new Error('Saldo sebenarnya tidak valid.');
    if (!getAccounts_().some(a => a.nama === rekening)) throw new Error('Rekening tidak valid.');

    adjSheet_().appendRow([
      dayStart_(data.tanggal), rekening, saldo, '', '',
      String(data.keterangan || '').trim() || 'Penyesuaian saldo'
    ]);

    const ledger = recalc_();

    return {
      success: true,
      message: `Saldo ${rekening} disesuaikan menjadi ${rp_(saldo)}.`,
      saldoRekening: rp_(ledger.balances[rekening])
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/* ================= TRANSAKSI ================= */

function appendTx_(d) {
  const rekening = String(d.rekening || '').trim();
  const kategori = String(d.kategori || '').trim();
  const pemasukan = Number(d.pemasukan || 0);
  const pengeluaran = Number(d.pengeluaran || 0);

  if (!d.tanggal) throw new Error('Tanggal wajib diisi.');
  if (!getAccounts_().some(a => a.nama === rekening)) {
    throw new Error('Rekening tidak valid atau tidak aktif.');
  }

  const cat = findCategory_(kategori);
  if (!cat) throw new Error('Kategori tidak ditemukan. Tambahkan di halaman Budget.');

  if (pemasukan < 0 || pengeluaran < 0) throw new Error('Nominal tidak boleh negatif.');
  if (pemasukan > 0 && pengeluaran > 0) throw new Error('Isi salah satu: pemasukan atau pengeluaran.');
  if (pemasukan === 0 && pengeluaran === 0) throw new Error('Nominal harus lebih dari 0.');

  const now = new Date();
  const when = d.waktuTransaksi ? new Date(`${d.tanggal}T${d.waktuTransaksi}`) : now;

  if (isNaN(when.getTime())) throw new Error('Waktu transaksi tidak valid.');

  txSheet_().appendRow([
    dayStart_(d.tanggal),
    now,
    when,
    rekening,
    kategori,
    String(d.keterangan || '').trim(),
    pemasukan,
    pengeluaran,
    '',
    d.splitId || '',
    d.refId || ''
  ]);
}

function addTransaction(data) {
  try {
    appendTx_({
      tanggal: data.tanggal,
      waktuTransaksi: data.waktuTransaksi,
      rekening: data.rekening,
      kategori: data.kategori,
      keterangan: data.keterangan,
      pemasukan: data.pemasukan,
      pengeluaran: data.pengeluaran
    });

    const ledger = recalc_();
    const rekening = String(data.rekening).trim();
    const total = Object.values(ledger.balances).reduce((a, b) => a + b, 0);

    return {
      success: true,
      message: `Transaksi disimpan. Saldo ${rekening}: ${rp_(ledger.balances[rekening])} | Total: ${rp_(total)}`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getRecentTransactions(limit) {
  const sh = txSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 11).getValues()
    .map((r, i) => ({
      rowNumber: i + 2,
      tanggal: r[0] instanceof Date ? fmtDate_(r[0]) : String(r[0]),
      tanggalTampil: r[0] instanceof Date ? fmtShort_(r[0]) : String(r[0]),
      waktuTransaksi: r[2] instanceof Date ? fmtTime_(r[2]) : '',
      rekening: r[3],
      kategori: r[4],
      keterangan: r[5],
      pemasukan: Number(r[6]) || 0,
      pengeluaran: Number(r[7]) || 0,
      saldo: r[8] === '' ? '' : Number(r[8]),
      splitId: String(r[9] || ''),
      refId: String(r[10] || ''),
      isDetail: isSplitDetail_(r)
    }))
    .reverse()
    .slice(0, limit || 30);
}

function deleteTransaction(rowNumber) {
  try {
    const sh = txSheet_();
    const last = sh.getLastRow();

    if (!rowNumber || rowNumber < 2 || rowNumber > last) {
      throw new Error('Transaksi tidak valid.');
    }

    const row = sh.getRange(rowNumber, 1, 1, 11).getValues()[0];

    if (String(row[10] || '')) {
      throw new Error('Transaksi terkait hutang-piutang tidak bisa dihapus dari sini.');
    }

    const splitId = String(row[9] || '');

    if (splitId) {
      const all = sh.getRange(2, 1, last - 1, 11).getValues();
      const rows = [];

      all.forEach((r, i) => {
        if (String(r[9]) === splitId) rows.push(i + 2);
      });

      rows.reverse().forEach(r => sh.deleteRow(r));
    } else {
      sh.deleteRow(rowNumber);
    }

    recalc_();
    return { success: true, message: 'Transaksi dihapus.' };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/* ================= TRANSFER ================= */

function transferBetweenAccounts(data) {
  try {
    const tanggal = data.tanggal;
    const dari = String(data.dari || '').trim();
    const ke = String(data.ke || '').trim();
    const nominal = Number(data.nominal || 0);
    const keterangan = String(data.keterangan || '').trim();

    if (!tanggal) throw new Error('Tanggal wajib diisi.');
    if (dari === ke) throw new Error('Rekening asal dan tujuan tidak boleh sama.');
    if (nominal <= 0) throw new Error('Nominal harus lebih dari 0.');

    const accounts = getAccounts_();
    if (!accounts.some(a => a.nama === dari)) throw new Error('Rekening asal tidak valid.');
    if (!accounts.some(a => a.nama === ke)) throw new Error('Rekening tujuan tidak valid.');

    appendTx_({
      tanggal,
      waktuTransaksi: data.waktuTransaksi || '',
      rekening: dari,
      kategori: 'Transfer Keluar',
      keterangan: `Transfer ke ${ke}${keterangan ? ' - ' + keterangan : ''}`,
      pemasukan: 0,
      pengeluaran: nominal
    });

    appendTx_({
      tanggal,
      waktuTransaksi: data.waktuTransaksi || '',
      rekening: ke,
      kategori: 'Transfer Masuk',
      keterangan: `Transfer dari ${dari}${keterangan ? ' - ' + keterangan : ''}`,
      pemasukan: nominal,
      pengeluaran: 0
    });

    const ledger = recalc_();

    return {
      success: true,
      message: `Transfer ${rp_(nominal)} dari ${dari} ke ${ke} berhasil.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/* ================= SPLIT ================= */

function getSplitCandidates() {
  return getRecentTransactions(60)
    .filter(t =>
      t.pengeluaran > 0 &&
      !t.splitId &&
      !t.refId &&
      t.kategori !== 'Split Transaksi' &&
      t.kategori !== 'Transfer Keluar'
    )
    .slice(0, 30);
}

function getTransactionById(rowNumber) {
  const sh = txSheet_();

  if (!rowNumber || rowNumber < 2 || rowNumber > sh.getLastRow()) {
    throw new Error('Transaksi tidak ditemukan.');
  }

  const r = sh.getRange(rowNumber, 1, 1, 11).getValues()[0];

  return {
    rowNumber,
    tanggal: r[0] instanceof Date ? fmtDate_(r[0]) : String(r[0]),
    waktuTransaksi: r[2] instanceof Date ? fmtTime_(r[2]) : '',
    rekening: String(r[3]),
    kategori: String(r[4]),
    keterangan: String(r[5]),
    pengeluaran: Number(r[7]) || 0,
    splitId: String(r[9] || ''),
    refId: String(r[10] || '')
  };
}

function addSplitTransaction(data) {
  try {
    const items = data.items || [];
    const tx = getTransactionById(Number(data.transactionRowNumber || 0));

    if (tx.pengeluaran <= 0) throw new Error('Hanya transaksi pengeluaran yang bisa di-split.');
    if (tx.splitId) throw new Error('Transaksi ini sudah pernah di-split.');
    if (tx.refId) throw new Error('Transaksi hutang-piutang tidak bisa di-split.');
    if (!items.length) throw new Error('Tambahkan minimal satu bagian split.');

    let total = 0;

    items.forEach((it, i) => {
      const nominal = Number(it.nominal || 0);
      const cat = findCategory_(it.kategori);

      if (!cat) throw new Error(`Kategori bagian ${i + 1} tidak valid.`);
      if (cat.kelompok === INCOME) throw new Error(`Bagian ${i + 1}: kategori pemasukan tidak bisa dipakai di split.`);
      if (nominal <= 0) throw new Error(`Nominal bagian ${i + 1} harus lebih dari 0.`);

      total += nominal;
    });

    if (total !== tx.pengeluaran) {
      throw new Error(
        `Total split (${rp_(total)}) harus sama dengan transaksi awal (${rp_(tx.pengeluaran)}).`
      );
    }

    const sh = txSheet_();
    const splitId = newId_('SPL');
    const time = data.waktuTransaksi || tx.waktuTransaksi;
    const when = time ? new Date(`${tx.tanggal}T${time}`) : new Date();

    sh.getRange(tx.rowNumber, 5).setValue('Split Transaksi');
    sh.getRange(tx.rowNumber, 6).setValue(tx.keterangan || `Split ${items.length} bagian`);
    sh.getRange(tx.rowNumber, 10).setValue(splitId);

    if (data.waktuTransaksi) {
      sh.getRange(tx.rowNumber, 3).setValue(when);
    }

    items.forEach(it => {
      sh.appendRow([
        dayStart_(tx.tanggal),
        new Date(),
        when,
        tx.rekening,
        it.kategori,
        String(it.keterangan || '').trim(),
        0,
        Number(it.nominal),
        '',
        splitId,
        ''
      ]);
    });

    recalc_();

    return {
      success: true,
      message: `Transaksi ${rp_(total)} di-split menjadi ${items.length} bagian.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/* ================= BUDGET ================= */

function saveBudget(data) {
  try {
    const bulan = String(data.bulan || '').trim();
    const total = Number(data.total || 0);

    if (!/^\d{4}-\d{2}$/.test(bulan)) throw new Error('Format bulan tidak valid.');
    if (total <= 0) throw new Error('Total uang bulanan harus lebih dari 0.');

    const sh = budgetSheet_();
    const row = [
      bulan,
      total,
      total * 0.45,
      total * 0.30,
      total * 0.15,
      total * 0.10
    ];

    const last = sh.getLastRow();
    let idx = -1;

    if (last >= 2) {
      const months = sh.getRange(2, 1, last - 1, 1).getValues();

      for (let i = 0; i < months.length; i++) {
        if (fmtMonth_(months[i][0]) === bulan) {
          idx = i + 2;
          break;
        }
      }
    }

    if (idx === -1) {
      sh.appendRow(row);
    } else {
      sh.getRange(idx, 1, 1, 6).setValues([row]);
    }

    return {
      success: true,
      message: `Budget ${bulan} disimpan (${rp_(total)}).`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getBudget_(month) {
  const sh = budgetSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;

  const vals = sh.getRange(2, 1, last - 1, 6).getValues();

  for (let i = vals.length - 1; i >= 0; i--) {
    if (fmtMonth_(vals[i][0]) === month) {
      return {
        total: Number(vals[i][1]) || 0,
        makan: Number(vals[i][2]) || 0,
        operasional: Number(vals[i][3]) || 0,
        nabung: Number(vals[i][4]) || 0,
        sedekahInvestasi: Number(vals[i][5]) || 0
      };
    }
  }

  return null;
}

function getDashboard(monthKey) {
  const month = monthKey || monthNow_();
  const ledger = recalc_();
  const cats = getCategories();
  const groupOf = {};

  cats.forEach(c => groupOf[c.nama] = c.kelompok);

  let totalIn = 0;
  let totalOut = 0;
  const used = {
    makan: 0,
    operasional: 0,
    nabung: 0,
    sedekahInvestasi: 0
  };

  const sh = txSheet_();
  const last = sh.getLastRow();

  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 11).getValues().forEach(r => {
      const rowMonth = r[0] instanceof Date ? fmtDate_(r[0]).substring(0, 7) : '';
      if (rowMonth !== month) return;

      const inc = Number(r[6]) || 0;
      const out = Number(r[7]) || 0;
      const kategori = String(r[4]);
      const ref = String(r[10] || '');
      const detail = isSplitDetail_(r);

      if (!detail) {
        totalIn += inc;
        totalOut += out;
      }

      const group = groupOf[kategori];

      if (
        kategori === 'Split Transaksi' ||
        kategori === 'Transfer Masuk' ||
        kategori === 'Transfer Keluar' ||
        !used.hasOwnProperty(group)
      ) return;

      if (out > 0) used[group] += out;
      if (inc > 0 && ref.endsWith('|LUNAS-PIUTANG')) used[group] -= inc;
    });
  }

  const budget = getBudget_(month);

  const usage = budget
    ? GROUPS.map(g => ({
        key: g.key,
        label: g.label,
        percent: g.percent,
        budget: budget[g.key],
        used: Math.max(used[g.key], 0)
      }))
    : [];

  return {
    month,
    budget,
    budgetUsage: usage,
    categories: cats,
    accounts: getAccounts_().map(a => ({
      nama: a.nama,
      jenis: a.jenis,
      saldo: ledger.balances[a.nama] || 0
    })),
    totalSaldo: Object.values(ledger.balances).reduce((a, b) => a + b, 0),
    totalPemasukan: totalIn,
    totalPengeluaran: totalOut
  };
}

function getRemainingBudgetForCategory(kategori) {
  const cat = findCategory_(kategori);

  if (!cat) {
    return { status: 'error', message: 'Kategori tidak ditemukan.' };
  }

  if (cat.kelompok === INCOME) {
    return { status: 'income', message: 'Kategori pemasukan/transfer tidak memakai budget.' };
  }

  const dash = getDashboard(monthNow_());

  if (!dash.budget) {
    return {
      status: 'nobudget',
      message: 'Budget bulan ini belum diatur. Atur di halaman Budget.',
      label: groupLabel_(cat.kelompok)
    };
  }

  const u = dash.budgetUsage.find(x => x.key === cat.kelompok);

  return {
    status: 'ok',
    label: u.label,
    budget: u.budget,
    terpakai: u.used,
    sisa: u.budget - u.used
  };
}

/* ================= HUTANG-PIUTANG ================= */

function addDebt(data) {
  try {
    const jenis = String(data.jenis || '').trim();
    const nama = String(data.namaOrang || '').trim();
    const rekening = String(data.rekening || '').trim();
    const kategori = String(data.kategori || '').trim();
    const nominal = Number(data.nominal || 0);
    const ket = String(data.keterangan || '').trim();

    if (!data.tanggal) throw new Error('Tanggal wajib diisi.');
    if (!['Hutang', 'Piutang'].includes(jenis)) throw new Error('Jenis tidak valid.');
    if (!nama) throw new Error('Nama orang wajib diisi.');
    if (nominal <= 0) throw new Error('Nominal harus lebih dari 0.');

    const cat = findCategory_(kategori);
    if (!cat) throw new Error('Pilih kategori yang terhubung ke Budget.');
    if (cat.kelompok === INCOME) throw new Error('Pilih kategori pengeluaran, bukan kategori pemasukan.');

    const id = newId_('DBT');

    appendTx_({
      tanggal: data.tanggal,
      waktuTransaksi: data.waktuTransaksi,
      rekening,
      kategori,
      keterangan: `${jenis} ${nama}${ket ? ' - ' + ket : ''}`,
      pemasukan: jenis === 'Hutang' ? nominal : 0,
      pengeluaran: jenis === 'Piutang' ? nominal : 0,
      refId: `${id}|${jenis.toUpperCase()}`
    });

    debtSheet_().appendRow([
      id,
      dayStart_(data.tanggal),
      jenis,
      nama,
      rekening,
      kategori,
      nominal,
      ket,
      data.jatuhTempo ? dayStart_(data.jatuhTempo) : '',
      'Belum Lunas',
      ''
    ]);

    recalc_();

    return {
      success: true,
      message: `${jenis} ${nama} sebesar ${rp_(nominal)} dicatat.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function settleDebt(rowNumber) {
  try {
    const sh = debtSheet_();

    if (!rowNumber || rowNumber < 2 || rowNumber > sh.getLastRow()) {
      throw new Error('Data tidak valid.');
    }

    const r = sh.getRange(rowNumber, 1, 1, 11).getValues()[0];

    if (String(r[9]) === 'Lunas') throw new Error('Sudah lunas.');

    const id = String(r[0]);
    const jenis = String(r[2]);
    const nama = String(r[3]);
    const rekening = String(r[4]);
    const kategori = String(r[5]);
    const nominal = Number(r[6]);

    appendTx_({
      tanggal: fmtDate_(new Date()),
      rekening,
      kategori,
      keterangan: `Pelunasan ${jenis.toLowerCase()} ${nama}`,
      pemasukan: jenis === 'Piutang' ? nominal : 0,
      pengeluaran: jenis === 'Hutang' ? nominal : 0,
      refId: `${id}|LUNAS-${jenis.toUpperCase()}`
    });

    sh.getRange(rowNumber, 10).setValue('Lunas');
    sh.getRange(rowNumber, 11).setValue(new Date());

    recalc_();

    return {
      success: true,
      message: `${jenis} ${nama} (${rp_(nominal)}) ditandai lunas.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getDebts() {
  const sh = debtSheet_();
  const last = sh.getLastRow();
  const items = [];
  let totalHutang = 0;
  let totalPiutang = 0;

  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 11).getValues().forEach((r, i) => {
      const status = String(r[9]);
      const nominal = Number(r[6]) || 0;

      if (status !== 'Lunas') {
        if (r[2] === 'Hutang') totalHutang += nominal;
        else totalPiutang += nominal;
      }

      items.push({
        rowNumber: i + 2,
        tanggal: r[1] instanceof Date ? fmtShort_(r[1]) : String(r[1]),
        jenis: r[2],
        nama: r[3],
        rekening: r[4],
        kategori: r[5],
        nominal,
        jatuhTempo: r[8] instanceof Date ? fmtShort_(r[8]) : '-',
        status
      });
    });
  }

  return {
    items: items.reverse(),
    totalHutang,
    totalPiutang
  };
}
