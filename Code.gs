// Time zone project wajib diset ke Asia/Jakarta di appsscript.json atau Project Settings.
// Kosongkan SHEET_ID jika script dibuat langsung dari menu Extensions > Apps Script di dalam Google Sheet.
const SHEET_ID = '';
// Nama lengkap sesuai mutasi bank (untuk deteksi transfer antar-rekening milik sendiri saat auto-sync email).
const OWNER_NAME = '';
const QRIS_DEFAULT_CATEGORY = 'Makan & Minum';

// Dicek berurutan dari atas; nama kategori harus ada di daftar kategori Pengaturan.
const MERCHANT_RULES = [
  [/gojek|grab|gocar|goride|maxim|indrive|parkir|parking|\btol\b|spbu|pertamina|\bshell\b|\bbp\b|krl|\bmrt\b|transjakarta|kereta api|\bkai\b|traveloka|tiket/i, 'Transportasi'],
  [/netflix|spotify|youtube|google|apple|disney|vidio|canva|chatgpt|openai|anthropic|icloud/i, 'Langganan Digital'],
  [/kopi|coffee|starbucks|janji jiwa|kenangan|\bfore\b|tomoro|cafe|caffe|excelso|j\.co|kopitiam/i, 'Ngopi / Cafe'],
  [/indomaret|\bidm\b|alfamart|alfamidi|\balfa\b|superindo|lawson|circle k|familymart|hypermart|transmart|giant|tokopedia|shopee/i, 'Belanja']
];

const SHEETS = { out: 'Pengeluaran', in: 'Pemasukan', set: 'Pengaturan', inv: 'Investasi' };
const INVEST_ACCOUNT = 'Reksa Dana';
const HEADERS = ['ID', 'Tanggal', 'Kategori', 'Keterangan', 'Rekening', 'Nominal'];

const DEFAULT_SETTINGS = {
  limit: 35000,
  accs: ['Mandiri', 'Bank MAS', 'BCA', 'GoPay', 'Tunai'],
  catKeluar: ['Makan & Minum', 'Ngopi / Cafe', 'Transportasi', 'Langganan Digital', 'Belanja', 'Rokok', 'Penyesuaian Saldo', 'Lain-lain'],
  catMasuk: ['Gaji', 'Hasil Usaha', 'Bonus', 'Penyesuaian Saldo', 'Lain-lain'],
  investYield: 4.13,
  investName: 'Sucorinvest Sharia Money Market Fund'
};

function getResolvedSheetId() {
  if (SHEET_ID && !SHEET_ID.startsWith('GANTI_')) return SHEET_ID.trim();
  try {
    const p = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (p) return p.trim();
  } catch (e) {}
  return '';
}

function getResolvedOwnerName() {
  if (OWNER_NAME && !OWNER_NAME.startsWith('GANTI_')) return OWNER_NAME.trim().toUpperCase();
  try {
    const p = PropertiesService.getScriptProperties().getProperty('OWNER_NAME');
    if (p) return p.trim().toUpperCase();
  } catch (e) {}
  return '';
}

let _ss = null;
function ss() {
  if (_ss) return _ss;
  const id = getResolvedSheetId();
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_ss) {
    throw new Error('Spreadsheet belum terhubung. Buat script melalui menu Extensions > Apps Script di Google Sheet, atau isi SHEET_ID di Code.gs.');
  }
  return _ss;
}
const getSheet = (name) => ss().getSheetByName(name);
const tz = () => Session.getScriptTimeZone();
const fmt = (d, pattern) => Utilities.formatDate(d, tz(), pattern);
const toIso = (d) => fmt(d, "yyyy-MM-dd'T'HH:mm:ss");
const toDate = (v) => (v instanceof Date ? v : new Date(v));
const isValidDate = (d) => d instanceof Date && !isNaN(d.getTime());
const num = (v) => parseFloat(v) || 0;

function readRowsLimit(name, limit) {
  const sh = getSheet(name);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];
  const total = lastRow - 1;
  const count = limit && total > limit ? limit : total;
  const start = lastRow - count + 1;
  return sh.getRange(start, 1, count, 6).getValues();
}

function buildDate(tgl) {
  if (!tgl) return new Date();
  if (tgl instanceof Date) return tgl;
  const m = String(tgl).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const n = new Date();
    return new Date(+m[1], +m[2] - 1, +m[3], n.getHours(), n.getMinutes(), n.getSeconds());
  }
  const d = new Date(tgl);
  return isValidDate(d) ? d : new Date();
}

function parseRupiah(s) {
  const clean = String(s).replace(/(?:Rp\.?|IDR)/gi, '').trim().replace(/[.,]\d{1,2}$/, '').replace(/\D/g, '');
  return parseInt(clean, 10) || 0;
}

const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, mei: 4, may: 4, jun: 5, jul: 6, agu: 7, agt: 7, aug: 7, sep: 8, okt: 9, oct: 9, nov: 10, des: 11, dec: 11 };
function parseIdDate(dateStr, timeStr) {
  const m = String(dateStr || '').trim().match(/(\d{1,2})[\s-]+([A-Za-z]+)[\s-]+(\d{4})/);
  if (!m) return null;
  const mon = MONTHS[m[2].substring(0, 3).toLowerCase()];
  if (mon === undefined) return null;
  const t = String(timeStr || '').match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return new Date(+m[3], mon, +m[1], t ? +t[1] : 0, t ? +t[2] : 0, t && t[3] ? +t[3] : 0);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { status: 'error', message: 'Server sibuk, coba lagi sebentar.' }; }
  try { return fn(); }
  catch (e) { return { status: 'error', message: String(e && e.message ? e.message : e) }; }
  finally { lock.releaseLock(); }
}

function setupSheets() {
  [SHEETS.out, SHEETS.in].forEach(n => {
    const sh = getSheet(n) || ss().insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  });
  if (!getSheet(SHEETS.set)) ss().insertSheet(SHEETS.set).appendRow(['Key', 'Value']);
  invSheet();
}

function ensureSheets() {
  if (!getSheet(SHEETS.out) || !getSheet(SHEETS.in) || !getSheet(SHEETS.set) || !getSheet(SHEETS.inv)) {
    setupSheets();
  }
}

function getSettingsData() {
  const set = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  set.rekeningSaldo = {};

  const sh = getSheet(SHEETS.set);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([k, v]) => {
      if (k === 'daily_limit') { set.limit = parseFloat(v) || DEFAULT_SETTINGS.limit; return; }
      if (k === 'invest_yield') { const y = parseFloat(v); if (isFinite(y) && y >= 0) set.investYield = y; return; }
      if (k === 'invest_name') { if (v) set.investName = String(v); return; }
      if (k === 'sync_start_date') { if (v) set.syncStartDate = String(v).trim(); return; }
      if (k === 'tot_pemasukan') { set.totPemasukan = num(v); return; }
      if (k === 'tot_pengeluaran') { set.totPengeluaran = num(v); return; }
      if (String(k).startsWith('saldo_')) { set.rekeningSaldo[String(k).replace('saldo_', '')] = num(v); return; }

      if (!v) return;
      try {
        const parsed = JSON.parse(v);
        if (!Array.isArray(parsed) || !parsed.length) return;
        if (k === 'accounts') set.accs = parsed;
        else if (k === 'cat_pengeluaran') set.catKeluar = parsed;
        else if (k === 'cat_pemasukan') set.catMasuk = parsed;
      } catch (e) {}
    });
  }
  return set;
}

function writeSettingsRows(newSet) {
  ensureSheets();
  const prev = getSettingsData();
  const sh = getSheet(SHEETS.set) || ss().insertSheet(SHEETS.set);
  if (sh.getLastRow() === 0) sh.appendRow(['Key', 'Value']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();

  const accs = Array.isArray(newSet.accounts) && newSet.accounts.length ? newSet.accounts : DEFAULT_SETTINGS.accs;
  const catOut = Array.isArray(newSet.catPengeluaran) && newSet.catPengeluaran.length ? newSet.catPengeluaran : DEFAULT_SETTINGS.catKeluar;
  const catIn = Array.isArray(newSet.catPemasukan) && newSet.catPemasukan.length ? newSet.catPemasukan : (Array.isArray(newSet.catMasuk) ? newSet.catMasuk : DEFAULT_SETTINGS.catMasuk);
  const syncStart = newSet.syncStartDate || prev.syncStartDate || dayKey(new Date());

  // Formula SUMIFS dihitung langsung oleh mesin Google Sheets agar pembacaan saldo instan.
  const rows = [
    ['daily_limit', num(newSet.dailyLimit) || DEFAULT_SETTINGS.limit],
    ['accounts', JSON.stringify(accs)],
    ['cat_pengeluaran', JSON.stringify(catOut)],
    ['cat_pemasukan', JSON.stringify(catIn)],
    ['invest_yield', isFinite(parseFloat(newSet.investYield)) ? parseFloat(newSet.investYield) : DEFAULT_SETTINGS.investYield],
    ['invest_name', newSet.investName || DEFAULT_SETTINGS.investName],
    ['sync_start_date', syncStart],
    ['tot_pemasukan', `=SUMIFS('${SHEETS.in}'!F:F, '${SHEETS.in}'!C:C, "<>Transfer Masuk")`],
    ['tot_pengeluaran', `=SUMIFS('${SHEETS.out}'!F:F, '${SHEETS.out}'!C:C, "<>Transfer Keluar")`]
  ];

  accs.forEach(acc => {
    if (acc === 'Tunai') {
      rows.push(['saldo_Tunai', `=SUMIFS('${SHEETS.in}'!F:F, '${SHEETS.in}'!E:E, "Tunai") - SUMIFS('${SHEETS.out}'!F:F, '${SHEETS.out}'!E:E, "Tunai") + SUMIFS('${SHEETS.in}'!F:F, '${SHEETS.in}'!E:E, "") - SUMIFS('${SHEETS.out}'!F:F, '${SHEETS.out}'!E:E, "")`]);
    } else {
      rows.push(['saldo_' + acc, `=SUMIFS('${SHEETS.in}'!F:F, '${SHEETS.in}'!E:E, "${acc}") - SUMIFS('${SHEETS.out}'!F:F, '${SHEETS.out}'!E:E, "${acc}")`]);
    }
  });

  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  SpreadsheetApp.flush();
}

// Menggabungkan pembacaan Pengaturan, Ringkasan Dashboard, dan Riwayat dalam satu eksekusi.
function buildPayload() {
  ensureSheets();
  SpreadsheetApp.flush();
  let set = getSettingsData();
  if (set.totPemasukan === undefined || !set.syncStartDate) {
    writeSettingsRows({
      dailyLimit: set.limit, accounts: set.accs, catPengeluaran: set.catKeluar,
      catPemasukan: set.catMasuk, investYield: set.investYield, investName: set.investName,
      syncStartDate: set.syncStartDate || dayKey(new Date())
    });
    set = getSettingsData();
  }

  const now = new Date();
  const todayStr = fmt(now, 'yyyy-MM-dd');
  const thisM = fmt(now, 'yyyy-MM');
  const lastM = fmt(new Date(now.getFullYear(), now.getMonth() - 1, 1), 'yyyy-MM');
  const thresholdDate = new Date(now.getTime() - 60 * 86400000);
  const thresholdIso = toIso(thresholdDate);

  let outToday = 0, outThisMonth = 0, outLastMonth = 0;
  const catMap = {}, dailyOut = {}, activeDays = new Set();
  const hist = [];

  const outRows = readRowsLimit(SHEETS.out, 600);
  const inRows = readRowsLimit(SHEETS.in, 600);

  outRows.forEach(r => {
    if (!r[0] && !r[5]) return;
    const v = num(r[5]), d = toDate(r[1]), valid = isValidDate(d);
    const tglIso = valid ? toIso(d) : String(r[1]);

    if (r[0] && tglIso >= thresholdIso) {
      hist.push({ id: String(r[0]), tanggal: tglIso, kategori: r[2], keterangan: r[3], rekening: r[4], nominal: v, jenis: 'pengeluaran' });
    }
    if (r[2] === 'Transfer Keluar' || !valid) return;

    const day = tglIso.substring(0, 10), mon = tglIso.substring(0, 7);
    if (mon === thisM) {
      outThisMonth += v;
      catMap[r[2]] = (catMap[r[2]] || 0) + v;
      activeDays.add(day);
      dailyOut[day] = (dailyOut[day] || 0) + v;
      if (day === todayStr) outToday += v;
    } else if (mon === lastM) {
      outLastMonth += v;
    }
  });

  inRows.forEach(r => {
    if (!r[0]) return;
    const d = toDate(r[1]);
    const tglIso = isValidDate(d) ? toIso(d) : String(r[1]);
    if (tglIso >= thresholdIso) {
      hist.push({ id: String(r[0]), tanggal: tglIso, kategori: r[2], keterangan: r[3], rekening: r[4], nominal: num(r[5]), jenis: 'pemasukan' });
    }
  });

  hist.sort((a, b) => (a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : 0));

  let limitHarian = set.limit;
  const kemarin = new Date(now);
  kemarin.setDate(now.getDate() - 1);
  if (fmt(kemarin, 'yyyy-MM') === thisM) {
    limitHarian += set.limit - (dailyOut[fmt(kemarin, 'yyyy-MM-dd')] || 0);
  }
  limitHarian = Math.max(limitHarian, 0);

  const inv = computeInvestasi(set, now);
  let maxCat = '-', maxNom = 0;
  Object.keys(catMap).forEach(c => { if (catMap[c] > maxNom) { maxNom = catMap[c]; maxCat = c; } });
  const diff = outLastMonth > 0 ? Math.abs(Math.round(((outThisMonth - outLastMonth) / outLastMonth) * 100)) : (outThisMonth > 0 ? 100 : 0);

  const totIn = set.totPemasukan || 0;
  const totOut = set.totPengeluaran || 0;
  const _cr = String.fromCharCode(80, 111, 108, 97, 114, 105, 122, 101, 95, 95);

  return {
    status: 'success',
    _cr: _cr,
    settings: {
      dailyLimit: set.limit,
      accounts: set.accs,
      catPengeluaran: set.catKeluar,
      catPemasukan: set.catMasuk,
      investYield: set.investYield,
      investName: set.investName
    },
    dash: {
      status: 'success',
      saldo: totIn - totOut + (inv.hasData ? inv.untung : 0),
      investasi: inv,
      pemasukan: totIn,
      pengeluaran: totOut,
      rekeningSaldo: set.rekeningSaldo || {},
      pengeluaranHariIni: outToday,
      effectiveDailyLimit: limitHarian,
      insight: {
        terbesarKat: maxCat,
        terbesarNom: maxNom,
        rataHarian: Math.round(outThisMonth / Math.max(activeDays.size, 1)),
        diffPercent: diff,
        isIncrease: outThisMonth > outLastMonth
      }
    },
    history: hist
  };
}

function ensureAutoSyncTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const has = triggers.some(t => t.getHandlerFunction() === 'autoSyncJob');
    if (!has) ScriptApp.newTrigger('autoSyncJob').timeBased().everyMinutes(10).create();
  } catch (e) {}
}

function getBootData() {
  try {
    ensureAutoSyncTrigger();
    return buildPayload();
  } catch (e) {
    return { status: 'error', message: String(e) };
  }
}

function getAppSettings() {
  try {
    const s = getSettingsData();
    return { status: 'success', settings: { dailyLimit: s.limit, accounts: s.accs, catPengeluaran: s.catKeluar, catPemasukan: s.catMasuk, investYield: s.investYield, investName: s.investName } };
  } catch (e) { return { status: 'error', message: String(e) }; }
}

function saveAppSettings(newSet) {
  return withLock(() => {
    writeSettingsRows(newSet);
    return buildPayload();
  });
}

function getIosProfileContent(customUrl) {
  try {
    return { status: 'success', xml: getMobileConfigXml(customUrl), filename: 'FinansialKu.mobileconfig' };
  } catch (e) {
    return { status: 'error', message: String(e) };
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ios === '1') {
    return ContentService.createTextOutput(getMobileConfigXml())
      .setMimeType(ContentService.MimeType.XML)
      .downloadAsFile('FinansialKu.mobileconfig');
  }
  const sig = String.fromCharCode(80, 111, 108, 97, 114, 105, 122, 101, 95, 95);
  const rawHtml = HtmlService.createHtmlOutputFromFile('Index').getContent();
  if (rawHtml.indexOf(sig) === -1) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:40px;text-align:center;"><h3>Error Integritas Aplikasi</h3><p>Kredit pembuat asli (<b>' + sig + '</b>) telah dihapus atau diubah. Kembalikan atribusi asli pada Index.html untuk menjalankan FinansialKu.</p></div>'
    ).setTitle('FinansialKu - Integrity Error');
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('FinansialKu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function addTransaction(data) {
  return withLock(() => {
    ensureSheets();
    const nom = num(data.nominal);
    if (nom <= 0) throw new Error('Nominal tidak valid');
    const sheetName = data.jenis === 'pemasukan' ? SHEETS.in : SHEETS.out;
    getSheet(sheetName).appendRow([Utilities.getUuid(), buildDate(data.tanggal), data.kategori, data.keterangan || '', data.rekening, nom]);
    return buildPayload();
  });
}

// ID transfer berpasangan (_o dan _i) agar penghapusan salah satu ikut menghapus pasangannya.
function appendTransferPair(tgl, ket, asal, tujuan, nom) {
  ensureSheets();
  const base = Utilities.getUuid();
  getSheet(SHEETS.out).appendRow([base + '_o', tgl, 'Transfer Keluar', ket, asal, nom]);
  getSheet(SHEETS.in).appendRow([base + '_i', tgl, 'Transfer Masuk', ket, tujuan, nom]);
  return base;
}

function transferFunds(data) {
  return withLock(() => {
    const nom = num(data.nominal);
    if (nom <= 0) throw new Error('Nominal tidak valid');
    if (data.asal === data.tujuan) throw new Error('Rekening asal dan tujuan sama');
    const ket = data.keterangan ? `Transfer: ${data.keterangan}` : 'Transfer antar rekening';
    appendTransferPair(buildDate(data.tanggal), `${ket} (${data.asal} ke ${data.tujuan})`, data.asal, data.tujuan, nom);
    return buildPayload();
  });
}

function adjustBalance(data) {
  return withLock(() => {
    ensureSheets();
    if (!data.rekening) throw new Error('Rekening tidak valid');
    const target = num(data.targetSaldo);
    const current = num(data.currentSaldo);
    const selisih = target - current;
    if (selisih === 0) return Object.assign({ selisih: 0 }, buildPayload());
    const tgl = new Date();
    const ket = data.catatan ? `Penyesuaian: ${data.catatan}` : `Penyesuaian saldo ${data.rekening}`;
    if (selisih > 0) {
      getSheet(SHEETS.in).appendRow([Utilities.getUuid(), tgl, 'Penyesuaian Saldo', ket, data.rekening, selisih]);
    } else {
      getSheet(SHEETS.out).appendRow([Utilities.getUuid(), tgl, 'Penyesuaian Saldo', ket, data.rekening, Math.abs(selisih)]);
    }
    return Object.assign({ selisih: selisih }, buildPayload());
  });
}

function updateTransaction(data) {
  return withLock(() => {
    if (!data.id) throw new Error('ID tidak valid');
    const nom = num(data.nominal);
    if (nom <= 0) throw new Error('Nominal tidak valid');
    const sheetName = data.jenis === 'pemasukan' ? SHEETS.in : SHEETS.out;
    const sh = getSheet(sheetName);
    if (!sh || sh.getLastRow() <= 1) return { status: 'error', message: 'Transaksi tidak ditemukan' };
    const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const idx = col.findIndex(r => String(r[0]) === String(data.id));
    if (idx < 0) return { status: 'error', message: 'Transaksi tidak ditemukan' };
    sh.getRange(idx + 2, 2, 1, 5).setValues([[buildDate(data.tanggal), data.kategori, data.keterangan || '', data.rekening, nom]]);
    return buildPayload();
  });
}

function deleteTransaction(id) {
  return withLock(() => {
    const ids = new Set([String(id)]);
    const m = String(id).match(/^(.*)_(o|i)$/);
    if (m) { ids.add(m[1] + '_o'); ids.add(m[1] + '_i'); }
    let deleted = 0;
    [SHEETS.out, SHEETS.in].forEach(name => {
      const sh = getSheet(name);
      if (!sh || sh.getLastRow() <= 1) return;
      const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (let i = col.length - 1; i >= 0; i--) {
        if (ids.has(String(col[i][0]))) { sh.deleteRow(i + 2); deleted++; }
      }
    });
    return deleted ? buildPayload() : { status: 'error', message: 'Transaksi tidak ditemukan' };
  });
}

function getDashboardSummary() {
  try { return buildPayload().dash; }
  catch (e) { return { status: 'error', message: String(e) }; }
}

function getHistoryData() {
  try { return { status: 'success', data: buildPayload().history }; }
  catch (e) { return { status: 'error', message: String(e) }; }
}

function checkDupe(known, idOrRef, nom, rek, tglDate) {
  const key = isValidDate(tglDate) ? toIso(tglDate).substring(0, 16) : '';
  const idStr = String(idOrRef || '');
  const tokens = idStr ? idStr.split(/[-_]/).filter(s => s.length >= 6) : [];
  return known.some(r => {
    const desc = String(r[3] || '');
    if (idStr && desc.includes(`[${idStr}]`)) return true;
    if (tokens.length && tokens.some(tok => desc.includes(tok))) return true;
    const d = toDate(r[1]);
    return num(r[5]) === nom && r[4] === rek && isValidDate(d) && toIso(d).substring(0, 16) === key;
  });
}

const cleanText = (t) => String(t || '').replace(/\u00a0/g, ' ').replace(/\r/g, '');
const toLines = (t) => cleanText(t).split('\n').map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
const hasOwner = (t) => {
  const owner = getResolvedOwnerName();
  if (!owner) return true;
  return cleanText(t).toUpperCase().replace(/\s+/g, ' ').includes(owner);
};
const isSelfTransfer = (penerima) => {
  const owner = getResolvedOwnerName();
  if (!owner) return false;
  return String(penerima || '').toUpperCase().replace(/\s+/g, ' ').includes(owner);
};

function mailText(m) {
  let t = cleanText(m.getPlainBody());
  if (t.length < 40 || !/Rp\s?\d/i.test(t)) {
    const html = String(m.getBody())
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<(br|\/p|\/div|\/tr|\/td|\/th|\/h[1-6]|\/li)[^>]*>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'")
      .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>');
    t = cleanText(html);
  }
  return t;
}

function fieldValue(lines, label) {
  const re = new RegExp('^' + label + '\\s*:?\\s*(.*)$', 'i');
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(re);
    if (!m) continue;
    if (m[1]) return m[1].trim();
    if (lines[i + 1]) return lines[i + 1].trim();
  }
  return '';
}

function eachMail(query, fn) {
  GmailApp.search(query, 0, 30).forEach(t => t.getMessages().forEach(m => fn(m, mailText(m), m.getId())));
}

function skip(ctx, tag, id, why) { ctx.notes.push(`${tag} ...${String(id).slice(-5)}: ${why}`); }

function syncEmailTransactions() {
  return withLock(() => {
    ensureSheets();
    const set = getSettingsData();
    const ctx = {
      known: readRowsLimit(SHEETS.out, 300).concat(readRowsLimit(SHEETS.in, 300)),
      syncStart: set.syncStartDate || dayKey(new Date()),
      added: 0, checked: 0, dupes: 0, notes: []
    };

    eachMail('(from:bankmas OR "Bebas by Bank MAS") newer_than:2d', (m, text, id) => {
      if (!/informasi transaksi sukses/i.test(text)) return;
      ctx.checked++;
      if (!hasOwner(text)) return skip(ctx, 'Bank MAS', id, 'nama pemilik tidak ditemukan di email');
      if (/Nama Penerima/i.test(text) || /Jenis Transaksi\s*:\s*Transfer/i.test(text)) processMasTF(ctx, text, id);
      else if (/Merchant/i.test(text)) processQRIS(ctx, text, 'Bank MAS', id);
      else skip(ctx, 'Bank MAS', id, 'format email tidak dikenali');
    });

    eachMail('(from:bankmandiri.co.id OR "Pembayaran Berhasil" OR "Transfer Berhasil" OR "Top-up Berhasil") newer_than:2d', (m, text, id) => {
      if (/bankmas/i.test(m.getFrom()) || /Bebas by Bank MAS/i.test(text)) return;
      ctx.checked++;
      if (!hasOwner(text)) return skip(ctx, 'Mandiri', id, 'nama pemilik tidak ditemukan di email');
      processMandiriMail(ctx, text, id);
    });

    return Object.assign(
      { added: ctx.added, checked: ctx.checked, dupes: ctx.dupes, notes: ctx.notes.slice(0, 4) },
      buildPayload()
    );
  });
}

function setupAutoSync() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoSyncJob').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoSyncJob').timeBased().everyMinutes(10).create();
}
function autoSyncJob() { syncEmailTransactions(); }

function pushRow(ctx, sheetName, row) {
  getSheet(sheetName).appendRow(row);
  ctx.known.push(row);
  ctx.added++;
}

function masDate(body) {
  const tm = body.match(/Tanggal\s*-\s*Jam\s*:\s*(\d{1,2}[\s-][A-Za-z]+[\s-]\d{4})\s+(\d{1,2}:\d{2})/i);
  return tm ? parseIdDate(tm[1], tm[2]) : null;
}

function processMasTF(ctx, body, msgId) {
  const nomMatch = body.match(/Jumlah\s*:\s*Rp\.?\s*([\d,.]+)/i);
  if (!nomMatch) return;
  const nom = parseRupiah(nomMatch[1]);
  if (nom <= 0) return;

  const penerima = ((body.match(/Nama Penerima\s*:\s*([^\n\r]+)/i) || [])[1] || 'Orang Lain').trim();
  const bankTj = ((body.match(/Jenis Transaksi\s*:\s*Transfer ke Rekening\s+([^\n\r]+)/i) || [])[1] || 'Bank Lain').trim();
  const tgl = masDate(body) || new Date();
  if (isValidDate(tgl) && dayKey(tgl) < ctx.syncStart) return;

  if (checkDupe(ctx.known, msgId, nom, 'Bank MAS', tgl)) return;

  if (isSelfTransfer(penerima)) {
    const up = bankTj.toUpperCase();
    const tj = up.includes('MANDIRI') ? 'Mandiri' : (up.includes('BCA') ? 'BCA' : 'Tunai');
    const ket = `Transfer ke ${tj} [${msgId}]`;
    const base = Utilities.getUuid();
    pushRow(ctx, SHEETS.out, [base + '_o', tgl, 'Transfer Keluar', ket, 'Bank MAS', nom]);
    pushRow(ctx, SHEETS.in, [base + '_i', tgl, 'Transfer Masuk', ket, tj, nom]);
  } else {
    pushRow(ctx, SHEETS.out, [Utilities.getUuid(), tgl, 'Lain-lain', `TF ke ${penerima} (${bankTj}) [${msgId}]`, 'Bank MAS', nom]);
  }
}

function processMandiriMail(ctx, text, msgId) {
  const lines = toLines(text), flat = lines.join(' ');
  const d = flat.match(/Tanggal\s*:?\s*(\d{1,2}[\s-]+[A-Za-z]+[\s-]+\d{4})/i);
  const j = flat.match(/Jam\s*:?\s*(\d{1,2}:\d{2}(?::\d{2})?)/i);
  const tgl = (d ? parseIdDate(d[1], j ? j[1] : '') : null) || new Date();
  if (isValidDate(tgl) && dayKey(tgl) < ctx.syncStart) {
    return skip(ctx, 'Mandiri', msgId, 'transaksi sebelum tanggal mulai diabaikan');
  }

  const refM = flat.match(/No\.?\s*Referensi\s*:?\s*(\d+)/i);
  const refNo = refM ? refM[1] : '';
  const tag = refNo ? `${refNo}-${msgId}` : msgId;

  if (/Top-up Berhasil/i.test(flat) || /Nominal Top-up/i.test(flat)) {
    const totM = flat.match(/Total\s+Transaksi\s*:?\s*(?:Rp\.?|IDR)\s*([\d.,]+)/i);
    const nomM = flat.match(/Nominal\s+Top-up\s*:?\s*(?:Rp\.?|IDR)\s*([\d.,]+)/i);
    const feeM = flat.match(/Biaya\s+Transaksi\s*:?\s*(?:Rp\.?|IDR)\s*([\d.,]+)/i);
    const total = totM ? parseRupiah(totM[1]) : 0;
    const nominal = nomM ? parseRupiah(nomM[1]) : 0;
    const fee = feeM ? parseRupiah(feeM[1]) : 0;
    const netNom = nominal || total;
    if (netNom <= 0) return skip(ctx, 'Mandiri', msgId, 'nominal top-up tidak terbaca');

    if (checkDupe(ctx.known, tag, total || netNom, 'Mandiri', tgl)) { ctx.dupes++; return; }

    const pjM = flat.match(/Penyedia\s+Jasa\s+(.+?)(?:\s+\*{2,}|\s+Tanggal|\s+Jam|\s+Nominal)/i);
    const provider = pjM ? pjM[1].trim() : 'E-Wallet';
    const upProv = provider.toUpperCase();
    let destAcc = null;
    const set = getSettingsData();
    for (const acc of set.accs) {
      if (upProv.includes(acc.toUpperCase())) { destAcc = acc; break; }
    }

    if (destAcc) {
      const ket = `Top-up ${destAcc} [${tag}]`;
      const base = Utilities.getUuid();
      pushRow(ctx, SHEETS.out, [base + '_o', tgl, 'Transfer Keluar', ket, 'Mandiri', netNom]);
      pushRow(ctx, SHEETS.in, [base + '_i', tgl, 'Transfer Masuk', ket, destAcc, netNom]);
      if (fee > 0) {
        pushRow(ctx, SHEETS.out, [Utilities.getUuid(), tgl, 'Lain-lain', `Biaya Top-up ${destAcc} [${tag}]`, 'Mandiri', fee]);
      }
    } else {
      pushRow(ctx, SHEETS.out, [Utilities.getUuid(), tgl, 'Belanja', `Top-up ${provider} [${tag}]`, 'Mandiri', total || netNom]);
    }
    return;
  }

  if (/Transfer Berhasil/i.test(flat) || /Jumlah Transfer/i.test(flat)) {
    const jm = flat.match(/Jumlah\s+Transfer\s*:?\s*(?:Rp\.?|IDR)\s*([\d.,]+)/i);
    const nom = jm ? parseRupiah(jm[1]) : 0;
    if (nom <= 0) return skip(ctx, 'Mandiri', msgId, 'nominal transfer tidak terbaca');

    if (checkDupe(ctx.known, tag, nom, 'Mandiri', tgl)) { ctx.dupes++; return; }

    const pm = flat.match(/Penerima\s+(.+?)(?:\s+(?:Bank\s+[A-Za-z]+|Tanggal|Jam|Jumlah))/i);
    const penerima = pm ? pm[1].trim() : 'Penerima Transfer';
    const bankM = flat.match(/Penerima\s+.+?\s+((?:Bank\s+[^\n\r-]+|\w+)\s*-\s*\d+)/i);
    const bankTujuan = bankM ? bankM[1].trim() : '';

    const ketM = flat.match(/Keterangan\s+(.+?)(?:\s+Rekening\s+Sumber|\s+Simpan\s+email)/i);
    let userKet = ketM ? ketM[1].trim() : '';
    if (userKet === '-') userKet = '';

    if (isSelfTransfer(penerima)) {
      const upBank = bankTujuan.toUpperCase();
      let destAcc = null;
      const set = getSettingsData();
      for (const acc of set.accs) {
        if (acc !== 'Mandiri' && upBank.includes(acc.toUpperCase())) { destAcc = acc; break; }
      }
      if (destAcc) {
        const ket = `Transfer ke ${destAcc} [${tag}]`;
        const base = Utilities.getUuid();
        pushRow(ctx, SHEETS.out, [base + '_o', tgl, 'Transfer Keluar', ket, 'Mandiri', nom]);
        pushRow(ctx, SHEETS.in, [base + '_i', tgl, 'Transfer Masuk', ket, destAcc, nom]);
        return;
      }
    }

    const detail = userKet ? `${penerima} (${userKet})` : (bankTujuan ? `${penerima} (${bankTujuan})` : penerima);
    pushRow(ctx, SHEETS.out, [Utilities.getUuid(), tgl, 'Lain-lain', `TF ke ${detail} [${tag}]`, 'Mandiri', nom]);
    return;
  }

  if (/Pembayaran Berhasil/i.test(flat) || /detail transaksi/i.test(flat)) {
    const totM = flat.match(/Total\s+Transaksi\s*:?\s*(?:Rp\.?|IDR)\s*([\d.,]+)/i);
    const nm = flat.match(/Nominal(?:\s+Transaksi)?\s*:?\s*(?:Rp\.?|IDR)\s*([\d.,]+)/i);
    const total = totM ? parseRupiah(totM[1]) : (nm ? parseRupiah(nm[1]) : 0);
    if (total <= 0) return skip(ctx, 'Mandiri', msgId, 'nominal pembayaran tidak terbaca');

    if (checkDupe(ctx.known, tag, total, 'Mandiri', tgl)) { ctx.dupes++; return; }

    let merchant = '';
    const pm = flat.match(/Penerima\s+(.+?)(?:\s+(?:BATAM\s*-\s*ID|[A-Z]{3,}\s*-\s*ID|\d{8,}|Tanggal|Jam|Nominal))/i);
    if (pm) merchant = pm[1].trim();
    if (!merchant) {
      const pm2 = flat.match(/Penerima\s+(.+?)(?:\s+Tanggal|\s+Jam|\s+Nominal)/i);
      if (pm2) merchant = pm2[1].trim();
    }
    merchant = (merchant || 'Belanja').replace(/\s+(Tanggal|Jam|Nominal|No\.?\s*Referensi)\b.*$/i, '').trim();

    pushRow(ctx, SHEETS.out, [Utilities.getUuid(), tgl, guessCategory(ctx, merchant), `${merchant} [${tag}]`, 'Mandiri', total]);
    return;
  }

  skip(ctx, 'Mandiri', msgId, 'format email Livin tidak dikenali');
}

function parseQR(text, rek) {
  const lines = toLines(text), flat = lines.join(' ');
  const nm = flat.match(/Nominal(?:\s+Transaksi)?\s*:?\s*Rp\.?\s*([\d.,]+)/i) || flat.match(/Rp\.?\s*([\d.,]+)/i);
  const nom = nm ? parseRupiah(nm[1]) : 0;

  let merchant = '';
  if (rek === 'Mandiri') {
    merchant = fieldValue(lines, 'Penerima');
    if (!merchant) { const pm = flat.match(/Penerima\s+(.+?)\s+(?:Tanggal|Jam|Nominal)\b/i); merchant = pm ? pm[1] : ''; }
  } else {
    const mc = cleanText(text).match(/Merchant\s*:\s*([^\n\r]+)/i);
    merchant = mc ? mc[1].trim() : '';
  }
  merchant = merchant.replace(/\s+(Tanggal|Jam|Nominal|No\.?\s*Referensi)\b.*$/i, '').trim();

  let tgl = null;
  if (rek === 'Mandiri') {
    const d = flat.match(/Tanggal\s*:?\s*(\d{1,2}[\s-]+[A-Za-z]+[\s-]+\d{4})/i);
    const j = flat.match(/Jam\s*:?\s*(\d{1,2}:\d{2}(?::\d{2})?)/i);
    if (d) tgl = parseIdDate(d[1], j ? j[1] : '');
  } else {
    tgl = masDate(flat);
  }
  return { nom: nom, merchant: merchant || 'Belanja QRIS', tgl: tgl || new Date() };
}

function guessCategory(ctx, merchant) {
  if (!ctx.cats) ctx.cats = getSettingsData().catKeluar;
  for (const [re, cat] of MERCHANT_RULES) if (re.test(merchant) && ctx.cats.indexOf(cat) > -1) return cat;
  return QRIS_DEFAULT_CATEGORY;
}

function processQRIS(ctx, body, rek, msgId) {
  const q = parseQR(body, rek);
  if (q.nom <= 0) return skip(ctx, rek, msgId, 'nominal tidak terbaca');
  if (isValidDate(q.tgl) && dayKey(q.tgl) < ctx.syncStart) return;
  if (checkDupe(ctx.known, msgId, q.nom, rek, q.tgl)) { ctx.dupes++; return; }
  pushRow(ctx, SHEETS.out, [Utilities.getUuid(), q.tgl, guessCategory(ctx, q.merchant), `${q.merchant} [${msgId}]`, rek, q.nom]);
}

// NAB reksa dana terbit sekali per hari kerja; jika NAB belum diinput, nilai diestimasi dari yield tahunan.
const dayKey = (d) => fmt(d, 'yyyy-MM-dd');
const dayNum = (key) => Math.round(Date.parse(key + 'T00:00:00Z') / 86400000);
const dayStr = (n) => Utilities.formatDate(new Date(n * 86400000), 'UTC', 'yyyy-MM-dd');

function invSheet() {
  let sh = getSheet(SHEETS.inv);
  if (!sh) {
    sh = ss().insertSheet(SHEETS.inv);
    sh.appendRow(['ID', 'Tanggal', 'Jenis', 'Produk', 'Nominal', 'Unit', 'NAB', 'Ref']);
  }
  return sh;
}

function readInv() {
  const sh = getSheet(SHEETS.inv);
  if (!sh || sh.getLastRow() <= 1) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues()
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), d: toDate(r[1]), jenis: String(r[2]), nominal: num(r[4]), unit: num(r[5]), nab: num(r[6]) }))
    .filter(r => isValidDate(r.d));
}

function computeInvestasi(set, now) {
  const rows = readInv();
  const entri = rows.slice().sort((a, b) => b.d - a.d).slice(0, 40)
    .map(r => ({ id: r.id, tanggal: toIso(r.d), jenis: r.jenis, nominal: r.nominal, unit: r.unit, nab: r.nab }));
  const base = { produk: set.investName, yield: set.investYield, entri: entri };

  const beli = rows.filter(r => r.jenis === 'Beli');
  if (!beli.length) return Object.assign({ hasData: false }, base);

  const daily = (set.investYield || 0) / 100 / 365;
  const today = dayNum(dayKey(now));
  const modal = beli.reduce((a, r) => a + r.nominal, 0);
  const buys = beli.map(r => ({
    day: dayNum(dayKey(r.d)), nominal: r.nominal,
    unit: r.unit > 0 ? r.unit : (r.nab > 0 ? r.nominal / r.nab : 0)
  }));
  const unitKnown = buys.every(b => b.unit > 0);
  const totalUnit = buys.reduce((a, b) => a + b.unit, 0);
  const nabs = rows.filter(r => r.jenis === 'NAB' && r.nab > 0).sort((a, b) => a.d - b.d);

  const valueAt = (day) => buys.reduce((a, b) => day >= b.day ? a + b.nominal * Math.pow(1 + daily, day - b.day) : a, 0);
  const gainAt = (day) => buys.reduce((a, b) => day > b.day ? a + b.nominal * Math.pow(1 + daily, day - 1 - b.day) * daily : a, 0);

  let mode, nilai, perHari, nabTerakhir = null, tglNab = null;
  const riwayat = [];
  const estNilai = valueAt(today);

  if (nabs.length && unitKnown) {
    mode = 'NAB';
    const last = nabs[nabs.length - 1];
    nabTerakhir = last.nab; tglNab = dayKey(last.d);
    nilai = totalUnit * last.nab;
    perHari = estNilai * daily;
    for (let i = Math.max(1, nabs.length - 7); i < nabs.length; i++) {
      const a = nabs[i - 1], b = nabs[i];
      const gap = Math.max(1, dayNum(dayKey(b.d)) - dayNum(dayKey(a.d)));
      riwayat.push({ tgl: dayKey(b.d), tambah: totalUnit * (b.nab - a.nab), hari: gap });
    }
    if (riwayat.length) {
      const l = riwayat[riwayat.length - 1];
      perHari = l.tambah / l.hari;
    }
  } else {
    mode = 'ESTIMASI';
    nilai = estNilai;
    perHari = estNilai * daily;
    const first = Math.min.apply(null, buys.map(b => b.day));
    for (let i = 6; i >= 0; i--) {
      const d = today - i;
      if (d <= first) continue;
      riwayat.push({ tgl: dayStr(d), tambah: gainAt(d), hari: 1 });
    }
  }

  return Object.assign({
    hasData: true, mode: mode, modal: modal, nilai: nilai, untung: nilai - modal,
    persen: modal > 0 ? ((nilai - modal) / modal) * 100 : 0,
    unit: unitKnown ? totalUnit : 0, nabTerakhir: nabTerakhir, tglNab: tglNab,
    perHari: perHari, perBulan: estNilai * daily * 30, riwayat: riwayat
  }, base);
}

function addInvestasi(data) {
  return withLock(() => {
    const nom = num(data.nominal);
    if (nom <= 0) throw new Error('Nominal tidak valid');
    const tgl = buildDate(data.tanggal), set = getSettingsData();
    let ref = '';
    if (data.rekening) ref = appendTransferPair(tgl, `Beli reksa dana: ${set.investName}`, data.rekening, INVEST_ACCOUNT, nom);
    invSheet().appendRow([Utilities.getUuid(), tgl, 'Beli', set.investName, nom, num(data.unit) || '', num(data.nab) || '', ref]);
    return buildPayload();
  });
}

function addNab(data) {
  return withLock(() => {
    const nab = num(data.nab);
    if (nab <= 0) throw new Error('NAB tidak valid');
    const sh = invSheet(), tgl = buildDate(data.tanggal), key = dayKey(tgl);
    const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues() : [];

    let done = false;
    for (let i = 0; i < vals.length; i++) {
      const d = toDate(vals[i][1]);
      if (vals[i][2] === 'NAB' && isValidDate(d) && dayKey(d) === key) { sh.getRange(i + 2, 7).setValue(nab); done = true; break; }
    }
    if (!done) sh.appendRow([Utilities.getUuid(), tgl, 'NAB', getSettingsData().investName, '', '', nab, '']);

    // Lengkapi kolom unit pembelian yang masih kosong secara proporsional terhadap nominal beli.
    const totalUnit = num(data.unit) > 0 ? num(data.unit) : (num(data.nilai) > 0 ? num(data.nilai) / nab : 0);
    if (totalUnit > 0) {
      const lacking = []; let known = 0, nomLacking = 0;
      vals.forEach((r, i) => {
        if (r[2] !== 'Beli') return;
        if (num(r[5]) > 0) known += num(r[5]); else { lacking.push(i); nomLacking += num(r[4]); }
      });
      const remaining = totalUnit - known;
      if (lacking.length && remaining > 0 && nomLacking > 0) {
        lacking.forEach(i => sh.getRange(i + 2, 6).setValue(remaining * num(vals[i][4]) / nomLacking));
      }
    }
    return buildPayload();
  });
}

function deleteInvestasi(id) {
  return withLock(() => {
    const sh = invSheet();
    if (sh.getLastRow() <= 1) return { status: 'error', message: 'Data tidak ditemukan' };
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
    const idx = vals.findIndex(r => String(r[0]) === String(id));
    if (idx < 0) return { status: 'error', message: 'Data tidak ditemukan' };
    const ref = String(vals[idx][7] || '');
    sh.deleteRow(idx + 2);
    if (ref) {
      const ids = new Set([ref + '_o', ref + '_i']);
      [SHEETS.out, SHEETS.in].forEach(name => {
        const t = getSheet(name);
        if (!t || t.getLastRow() <= 1) return;
        const col = t.getRange(2, 1, t.getLastRow() - 1, 1).getValues();
        for (let i = col.length - 1; i >= 0; i--) if (ids.has(String(col[i][0]))) t.deleteRow(i + 2);
      });
    }
    return buildPayload();
  });
}

function getMobileConfigXml(customUrl) {
  let targetUrl = customUrl || "";
  if (!targetUrl) { try { targetUrl = ScriptApp.getService().getUrl() || ""; } catch (e) {} }
  if (!targetUrl) targetUrl = "https://script.google.com/macros/s/GANTI_DENGAN_DEPLOYMENT_ID_KAMU/exec";
  return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n<plist version=\"1.0\">\n<dict>\n  <key>PayloadContent</key>\n  <array>\n    <dict>\n      <key>FullScreen</key>\n      <true/>\n      <key>Icon</key>\n      <data>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</data>\n      <key>IsRemovable</key>\n      <true/>\n      <key>Label</key>\n      <string>FinansialKu</string>\n      <key>PayloadDescription</key>\n      <string>FinansialKu Standalone WebClip</string>\n      <key>PayloadDisplayName</key>\n      <string>FinansialKu</string>\n      <key>PayloadIdentifier</key>\n      <string>com.finansialku.webclip</string>\n      <key>PayloadType</key>\n      <string>com.apple.webClip.managed</string>\n      <key>PayloadUUID</key>\n      <string>8A10C35B-11E2-4E87-9DE1-D6A8D2F8246E</string>\n      <key>PayloadVersion</key>\n      <integer>1</integer>\n      <key>Precomposed</key>\n      <true/>\n      <key>URL</key>\n      <string>" + targetUrl + "</string>\n    </dict>\n  </array>\n  <key>PayloadDescription</key>\n  <string>Installs FinansialKu Standalone WebClip on Home Screen</string>\n  <key>PayloadDisplayName</key>\n  <string>FinansialKu</string>\n  <key>PayloadIdentifier</key>\n  <string>com.finansialku.profile</string>\n  <key>PayloadOrganization</key>\n  <string>FinansialKu</string>\n  <key>PayloadRemovalDisallowed</key>\n  <false/>\n  <key>PayloadType</key>\n  <string>Configuration</string>\n  <key>PayloadUUID</key>\n  <string>F7A25D01-9876-46DF-BF8A-EEB32AC83D21</string>\n  <key>PayloadVersion</key>\n  <integer>1</integer>\n</dict>\n</plist>";
}
