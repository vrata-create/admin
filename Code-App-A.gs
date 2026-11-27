/** Code App A: backend (API) penghubung ke Google Sheets. Dipanggil dari aplikasi di GitHub Pages dengan token rahasia. */
const SS_ID = '1t6yg7r19OkML5AeItzKfCeqUcjHsnLYr1UcIGOnssl0';
const SHEET = 'Sheet1';

// ---------- API (dipanggil aplikasi di GitHub Pages) ----------
// Hanya permintaan yang membawa token rahasia (disimpan di Properti Skrip) yang diproses.
function doGet() {
  return json_({ ok: true, app: 'Aplikasi A', info: 'API aktif. Buka aplikasi lewat alamat GitHub Pages.' });
}
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return json_(api_(body));
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function api_(b) {
  const tok = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (!tok || !b || b.token !== tok) return { ok: false, error: 'Token salah' };
  const allowed = { getData: getData, saveTransaction: saveTransaction, updateRow: updateRow, deleteRow: deleteRow, getHistory: getHistory };
  if (!Object.prototype.hasOwnProperty.call(allowed, b.fn)) return { ok: false, error: 'Fungsi tidak dikenal' };
  try { return { ok: true, data: allowed[b.fn].apply(null, Array.isArray(b.args) ? b.args : []) }; }
  catch (err) { return { ok: false, error: String(err && err.message ? err.message : err) }; }
}
// Jalankan SEKALI dari editor, lalu salin token dari Log eksekusi.
// Menjalankan lagi membuat token baru dan mematikan token lama.
function buatToken() {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_TOKEN', t);
  Logger.log('TOKEN APLIKASI A: ' + t);
}

// ---------- Pembaca data ----------
function headerRow_(sh) {
  const v = sh.getRange(1, 1, 10, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim().toLowerCase() === 'tanggal') return i + 1;
  throw new Error('Header "Tanggal" tidak ditemukan di Sheet1');
}
function num_(s) {
  if (typeof s === 'number') return s;
  s = String(s || '').trim();
  if (!s) return 0;
  const n = parseFloat(s.replace(/\./g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function iso_(s) {
  const m = String(s).match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  return m ? m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0') : String(s);
}
function sig_(r) { return [r.tanggal, r.sl, r.produk, r.qtt].join('|'); }

function read_() {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET);
  if (!sh) throw new Error('Sheet "' + SHEET + '" tidak ditemukan');
  const hr = headerRow_(sh);
  const n = sh.getMaxRows() - hr;
  const vals = n > 0 ? sh.getRange(hr + 1, 1, n, 8).getDisplayValues() : [];
  const rows = [];
  vals.forEach(function (v, i) {
    if (!v[0] && !v[2]) return;
    const r = {
      row: hr + 1 + i, tanggal: iso_(v[0]), sl: v[1].trim(), produk: v[2].trim(),
      qtt: num_(v[3]), reb: num_(v[4]), sisa: num_(v[5]),
      qc: v[6].trim().toLowerCase(), status: v[7].trim()
    };
    r.sig = sig_(r);
    rows.push(r);
  });
  return { sh: sh, hr: hr, rows: rows };
}

// ---------- Cache sementara (otomatis dihapus setiap ada perubahan) ----------
const RANGES = [30, 60, 90, 180, 0]; // rentang data dalam hari; 0 = semua
function cget_(k) {
  try {
    const cc = CacheService.getScriptCache(), n = Number(cc.get(k + '_n') || 0);
    if (!n) return null;
    let s = '';
    for (let i = 0; i < n; i++) { const x = cc.get(k + '_' + i); if (x === null) return null; s += x; }
    return JSON.parse(s);
  } catch (err) { return null; }
}
function cput_(k, o) {
  try {
    const s = JSON.stringify(o), n = Math.ceil(s.length / 90000), m = {};
    for (let i = 0; i < n; i++) m[k + '_' + i] = s.substr(i * 90000, 90000);
    m[k + '_n'] = String(n);
    CacheService.getScriptCache().putAll(m, 120);
  } catch (err) {}
}
function cclear_() {
  try {
    const cc = CacheService.getScriptCache(), ks = [];
    RANGES.forEach(function (g) {
      const k = 'data' + g, n = Number(cc.get(k + '_n') || 0);
      ks.push(k + '_n');
      for (let i = 0; i < n; i++) ks.push(k + '_' + i);
    });
    cc.removeAll(ks);
  } catch (err) {}
}

// Memuat data per rentang tanggal produksi (default 30 hari).
// Batch yang belum selesai (on process, atau pass dengan sisa > 0) SELALU ikut dimuat,
// berapa pun umurnya, supaya daftar QC dan Rebagging tetap lengkap.
function getData(fresh, days) {
  days = RANGES.indexOf(Number(days)) >= 0 ? Number(days) : 30;
  const key = 'data' + days;
  if (!fresh) { const c = cget_(key); if (c) return c; }
  const d = read_();
  const cut = days ? Utilities.formatDate(new Date(Date.now() - days * 86400000), 'Asia/Jakarta', 'yyyy-MM-dd') : '';
  const out = {
    rows: d.rows.filter(function (r) {
      return !days || r.tanggal >= cut || r.qc === 'on process' || (r.qc === 'pass' && r.sisa > 0);
    }),
    produk: Array.from(new Set(d.rows.map(function (r) { return r.produk; }))).sort(),
    url: d.sh.getParent().getUrl() + '#gid=' + d.sh.getSheetId(),
    days: days, from: cut
  };
  cput_(key, out);
  return out;
}

// ---------- Aturan data ----------
function clean_(p) {
  const e = [];
  const sl = String(p.sl || '').trim(), produk = String(p.produk || '').trim(), qtt = Number(p.qtt);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.tanggal || '')) e.push('Tanggal wajib diisi');
  if (!/^[12]\/[1-5]$/.test(sl)) e.push('Shift/Line tidak valid');
  if (!produk) e.push('Produk wajib diisi');
  if (!(qtt > 0)) e.push('Qtt wajib diisi dan lebih dari 0');
  if (e.length) throw new Error(e.join(', '));
  return { tanggal: p.tanggal, sl: sl, produk: produk, qtt: qtt };
}
function derive_(r) {
  if (r.qc === 'reject') { r.reb = 0; r.sisa = 0; r.status = 'Tidak'; }
  else if (r.qc === 'pass') { r.sisa = r.qtt - r.reb; r.status = r.sisa === 0 ? 'selesai' : 'belum'; }
  else { r.qc = 'on process'; r.reb = 0; r.sisa = r.qtt; r.status = ''; }
  return r;
}
function same_(r, c) {
  return r.tanggal === c.tanggal && r.sl === c.sl && r.produk.toLowerCase() === c.produk.toLowerCase();
}
function lock_(f) {
  const l = LockService.getScriptLock();
  l.waitLock(15000);
  try { return f(); } finally { cclear_(); l.releaseLock(); }
}

// ---------- Nama produk dirapikan otomatis ----------
function normName_(s) { return String(s).trim().replace(/\s+/g, ' ').replace(/\s*([\/,])\s*/g, '$1'); }
function key_(s) { return normName_(s).toLowerCase().replace(/\s/g, ''); }
function canon_(name, rows) { // pakai ejaan yang sudah ada; produk baru ditulis HURUF BESAR
  const n = normName_(name), k = key_(n);
  for (let i = 0; i < rows.length; i++) if (key_(rows[i].produk) === k) return rows[i].produk;
  return n.toUpperCase();
}

// ---------- Penulisan (tidak pernah menimpa baris lain) ----------
function writeFull_(sh, row, r, isNew) {
  const [y, m, d] = r.tanggal.split('-').map(Number);
  const serial = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  if (isNew) { // format hanya untuk baris baru; dilewati jika kolom bertipe (Table) tidak mengizinkan
    try { sh.getRange(row, 1).setNumberFormat('dd/MM/yyyy'); } catch (e) {}
    try { sh.getRange(row, 2, 1, 2).setNumberFormat('@'); } catch (e) {} // Shift/Line & Produk sebagai teks
  }
  sh.getRange(row, 1, 1, 8).setValues([[serial, r.sl, r.produk, r.qtt, r.reb || '', r.sisa, r.qc, r.status]]);
}
function writeStatus_(sh, r) { // hanya kolom Rebagging, Sisa, QC, Status
  sh.getRange(r.row, 5, 1, 4).setValues([[r.reb || '', r.sisa, r.qc, r.status]]);
}

// ---------- Histori (sheet "Histori", dibuat otomatis) ----------
const LOG = 'Histori';
function fm_(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function desc_(r) { return r.tanggal.split('-').reverse().join('/') + ' · ' + r.sl + ' · ' + r.produk + ' · ' + fm_(r.qtt) + ' Kg'; }
function diff_(a, b) {
  const o = [];
  if (a.tanggal !== b.tanggal) o.push('Tanggal ' + a.tanggal.split('-').reverse().join('/') + ' → ' + b.tanggal.split('-').reverse().join('/'));
  if (a.sl !== b.sl) o.push('Shift/Line ' + a.sl + ' → ' + b.sl);
  if (a.produk !== b.produk) o.push('Produk ' + a.produk + ' → ' + b.produk);
  if (a.qtt !== b.qtt) o.push('Qtt ' + fm_(a.qtt) + ' → ' + fm_(b.qtt));
  if (a.reb !== b.reb) o.push('Rebagging ' + fm_(a.reb) + ' → ' + fm_(b.reb));
  if (a.qc !== b.qc) o.push('QC ' + a.qc + ' → ' + b.qc);
  return o.length ? o.join('; ') : 'Tidak ada perubahan';
}
function logTo_(ss, aksi, data, detail, dev) {
  try {
    let sh = ss.getSheetByName(LOG);
    if (!sh) {
      sh = ss.insertSheet(LOG);
      sh.getRange('A:E').setNumberFormat('@');
      sh.getRange(1, 1, 1, 5).setValues([['Waktu', 'Aksi', 'Data', 'Detail', 'Perangkat']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const t = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    // teks yang diawali = + - @ dianggap rumus oleh Sheets (#ERROR!), jadi diberi awalan dan ditulis sebagai teks
    const safe = function (x) { x = String(x === undefined || x === null ? '' : x); return /^[=+\-@]/.test(x) ? '\u2192 ' + x : x; };
    const rg = sh.getRange(sh.getLastRow() + 1, 1, 1, 5);
    rg.setNumberFormat('@');
    rg.setValues([[t, safe(aksi), safe(data), safe(detail), safe(String(dev || 'Tidak diketahui').slice(0, 80))]]);
  } catch (err) {}
}
function log_(aksi, data, detail, dev) { logTo_(SpreadsheetApp.openById(SS_ID), aksi, data, detail, dev); }

function getHistory() {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(300, sh.getLastRow() - 1), start = sh.getLastRow() - n + 1;
  return sh.getRange(start, 1, n, 5).getDisplayValues().reverse().map(function (v) {
    return { t: v[0], aksi: v[1], data: v[2], detail: v[3], dev: v[4] };
  });
}

// ---------- Aksi dari aplikasi ----------
function saveTransaction(p) {
  return lock_(function () {
    const d = read_();
    if (p.jenis === 'Produksi') {
      const c = clean_(p);
      c.produk = canon_(c.produk, d.rows);
      let row = d.hr;
      d.rows.forEach(function (r) { if (r.row > row) row = r.row; });
      row += 1;
      writeFull_(d.sh, row, derive_({ tanggal: c.tanggal, sl: c.sl, produk: c.produk, qtt: c.qtt, reb: 0, qc: 'on process' }), true);
      log_('Tambah Produksi', desc_(c), '', p.dev);
      return 'Produksi tersimpan';
    }
    if (p.jenis !== 'QC' && p.jenis !== 'Rebagging') throw new Error('Jenis transaksi tidak dikenal');
    // QC & Rebagging: batch dipilih dari daftar (row + sig), jadi tidak ada salah ketik
    const t = d.rows.find(function (r) { return r.row === Number(p.row); });
    if (!t || t.sig !== p.sig) throw new Error('Batch sudah berubah. Muat ulang lalu pilih batch lagi');
    if (p.jenis === 'QC') {
      if (p.qc !== 'pass' && p.qc !== 'reject') throw new Error('Pilih status QC');
      if (t.qc !== 'on process') throw new Error('Data telah dimasukkan. Status QC tercatat: ' + t.qc);
      t.qc = p.qc; derive_(t); writeStatus_(d.sh, t);
      log_('QC ' + p.qc, desc_(t), 'Dari on process', p.dev);
      return 'QC ' + p.qc + ' tersimpan';
    }
    const q = Number(p.qtt);
    if (!(q > 0)) throw new Error('Qtt rebagging wajib diisi dan lebih dari 0');
    if (t.qc !== 'pass' || t.sisa <= 0) throw new Error('Batch ini belum QC pass atau sudah selesai');
    if (q > t.sisa) throw new Error('Qtt rebagging melebihi sisa (' + fm_(t.sisa) + ' Kg)');
    t.reb += q; derive_(t); writeStatus_(d.sh, t);
    log_('Rebagging', desc_(t), 'Tambah ' + fm_(q) + ' Kg, sisa ' + fm_(t.sisa) + (t.status === 'selesai' ? ', selesai' : ''), p.dev);
    return t.status === 'selesai' ? 'Rebagging tersimpan, produk selesai' : 'Rebagging tersimpan';
  });
}

function updateRow(p) {
  return lock_(function () {
    const c = clean_(p), d = read_();
    const t = d.rows.find(function (r) { return r.row === Number(p.row); });
    if (!t || t.sig !== p.sig) throw new Error('Data sudah berubah. Muat ulang lalu coba lagi');
    c.produk = canon_(c.produk, d.rows);
    const reb = Number(p.reb) || 0;
    if (reb < 0) throw new Error('Rebagging tidak boleh negatif');
    if (p.qc === 'reject' && reb > 0) throw new Error('Rebagging sudah terisi ' + reb + ' Kg. Kosongkan Rebagging dulu sebelum mengubah QC menjadi reject');
    if (p.qc === 'pass' && reb > c.qtt) throw new Error('Rebagging melebihi Qtt');
    const nr = derive_({ tanggal: c.tanggal, sl: c.sl, produk: c.produk, qtt: c.qtt, reb: reb, qc: p.qc });
    writeFull_(d.sh, t.row, nr);
    log_('Edit', desc_(nr), diff_(t, nr), p.dev);
    return 'Data diperbarui';
  });
}

function deleteRow(p) {
  return lock_(function () {
    const d = read_();
    const t = d.rows.find(function (r) { return r.row === Number(p.row); });
    if (!t || t.sig !== p.sig) throw new Error('Data sudah berubah. Muat ulang lalu coba lagi');
    log_('Hapus', desc_(t), 'QC ' + t.qc + ', sisa ' + fm_(t.sisa), p.dev);
    d.sh.deleteRow(t.row);
    return 'Data dihapus';
  });
}

// ---------- Otomatis saat kamu mengedit langsung di spreadsheet ----------
// Jika kolom Qtt (D), Rebagging (E), atau QC (G) diubah di sheet,
// kolom Sisa (F) dan Status (H) dihitung ulang dengan aturan yang sama seperti aplikasi.
function onEdit(e) {
  try {
    const rg = e.range, sh = rg.getSheet();
    if (sh.getName() !== SHEET) return;
    cclear_();
    const hr = headerRow_(sh);
    const c1 = rg.getColumn(), c2 = rg.getLastColumn();
    const first = Math.max(rg.getRow(), hr + 1), last = rg.getLastRow();
    if (first > last || c1 > 8) return;
    // catat ke Histori
    const H = ['Tanggal', 'Shift/Line', 'Produk', 'Qtt', 'Rebagging', 'Sisa', 'QC', 'Status'];
    const one = rg.getNumRows() === 1 && rg.getNumColumns() === 1;
    const prod = sh.getRange(first, 3).getDisplayValue();
    const nv = function (x) { return x === undefined || x === '' ? '(kosong)' : x; };
    logTo_(e.source, 'Edit di Spreadsheet', 'Baris ' + first + (last > first ? '-' + last : '') + (prod ? ' · ' + prod : ''),
      one ? H[c1 - 1] + ': ' + nv(e.oldValue) + ' → ' + nv(e.value) : 'Mengubah ' + rg.getNumRows() + ' baris x ' + rg.getNumColumns() + ' kolom', 'Spreadsheet');
    // hanya Qtt (D), Rebagging (E), QC (G) yang memicu hitung ulang; Sisa/Status bisa kamu timpa manual
    if (![4, 5, 7].some(function (x) { return x >= c1 && x <= c2; })) return;
    for (let row = first; row <= last; row++) {
      const v = sh.getRange(row, 1, 1, 8).getDisplayValues()[0];
      if (!v[0] && !v[2]) continue;
      const qc = String(v[6]).trim().toLowerCase();
      if (['pass', 'reject', 'on process'].indexOf(qc) < 0) continue;
      const qtt = num_(v[3]), reb = num_(v[4]);
      let sisa, status;
      if (qc === 'reject') { sisa = 0; status = 'Tidak'; }
      else if (qc === 'pass') {
        if (reb > qtt) { SpreadsheetApp.getActive().toast('Baris ' + row + ': Rebagging melebihi Qtt', 'Periksa data', 6); continue; }
        sisa = qtt - reb; status = sisa === 0 ? 'selesai' : 'belum';
      } else { sisa = qtt; status = ''; }
      sh.getRange(row, 6).setValue(sisa);
      sh.getRange(row, 8).setValue(status);
    }
  } catch (err) {}
}

// ---------- Backup otomatis harian ----------
// Isi dengan ID folder tujuan (lihat panduan). Kosongkan untuk memakai folder "Backup Gudang".
const BACKUP_FOLDER_ID = '';
function backupFolder_() {
  if (BACKUP_FOLDER_ID) return DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const it = DriveApp.getFoldersByName('Backup Gudang');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Backup Gudang');
}
// Jalankan fungsi aktifkanBackupHarian SEKALI dari editor (pilih fungsi, klik Jalankan).
function backupHarian() {
  const folder = backupFolder_();
  const name = 'Backup ' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  DriveApp.getFileById(SS_ID).makeCopy(name, folder);
  const files = [], fi = folder.getFiles();
  while (fi.hasNext()) { const f = fi.next(); if (f.getName().indexOf('Backup ') === 0) files.push(f); }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(14).forEach(function (f) { f.setTrashed(true); }); // simpan 14 salinan terbaru
}
function aktifkanBackupHarian() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupHarian') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupHarian').timeBased().everyDays(1).atHour(2).create();
  backupHarian(); // salinan pertama langsung dibuat
}

// Jalankan SEKALI dari editor untuk memperbaiki catatan Histori lama yang tampil #ERROR!
function perbaikiHistori() {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(LOG);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1, f = sh.getRange(2, 1, n, 5).getFormulas();
  let fixed = 0;
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < 5; j++) {
      const fm = f[i][j];
      if (!fm) continue;
      const t = /^=\+/.test(fm) ? 'Tambah ' + fm.replace(/^=\+/, '') : fm.replace(/^=/, '');
      sh.getRange(i + 2, j + 1).setNumberFormat('@').setValue(t);
      fixed++;
    }
  }
  SpreadsheetApp.getActive().toast(fixed + ' catatan diperbaiki', 'Histori', 5);
}
