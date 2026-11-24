/**
 * ABSENSI WAJAH KIOSK — Backend (Google Apps Script)
 * Pasang: Spreadsheet > Extensions > Apps Script > tempel file ini
 * Deploy: Deploy > New deployment > Web app > Execute as: Me, Access: Anyone
 * Salin URL /exec ke konstanta GAS_URL di index.html
 */
const SHEET_ID = '';  // ID Spreadsheet database (dari URL .../d/ID/edit). Kosong = pakai spreadsheet tempat script ini dipasang
const CFG = {
  PIN: '1234',               // PIN admin: daftar wajah & ubah jadwal
  TZ: 'Asia/Jakarta',
  DEF: {masuk: '07:00', terlambat: '07:15', pulang: '15:00'}  // jadwal awal (bisa diubah dari menu Jadwal di kiosk / sheet Pengaturan)
};
const HARI = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'];
const H_SISWA = ['NIS', 'Nama', 'Kelas', 'Wajah'];
const H_ABSEN = ['Tanggal', 'NIS', 'Nama', 'Kelas', 'Jam Masuk', 'Status', 'Jam Pulang'];
const H_CFG = ['Hari', 'Masuk', 'Batas Terlambat', 'Pulang'];

function sh_(name, head) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange(1, 1, s.getMaxRows(), head.length).setNumberFormat('@'); // cegah auto-konversi tanggal/angka
  }
  if (s.getLastColumn() < head.length) {   // sheet lama: tambah kolom baru (mis. Jam Pulang)
    for (let c = s.getLastColumn() + 1; c <= head.length; c++) {
      if (s.getMaxColumns() < c) s.insertColumnAfter(s.getMaxColumns());
      s.getRange(1, c).setValue(head[c - 1]).setFontWeight('bold');
      s.getRange(1, c, s.getMaxRows(), 1).setNumberFormat('@');
    }
  }
  return s;
}
const now_ = f => Utilities.formatDate(new Date(), CFG.TZ, f);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() {
  try { return out_(init_()); } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = JSON.parse(e.postData.contents);
    if (p.action === 'absen') return out_(absen_(String(p.nis), p.jenis));
    if (p.action === 'setCfg') return out_(setCfg_(p));
    if (p.action === 'rekap') return out_(rekap_(p));
    if (p.action === 'daftar') return out_(daftar_(p));
    return out_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function cfg_() {
  const s = sh_('Pengaturan', H_CFG);
  if (s.getLastRow() < 2) s.getRange(2, 1, 5, 4).setValues(HARI.map(h => [h, CFG.DEF.masuk, CFG.DEF.terlambat, CFG.DEF.pulang]));
  const tz = s.getParent().getSpreadsheetTimeZone();
  const t = v => v instanceof Date ? Utilities.formatDate(v, tz, 'HH:mm') : String(v || '').trim();
  const rows = s.getDataRange().getValues().slice(1);
  return HARI.map(h => {
    const r = rows.find(x => String(x[0]).trim() === h) || [];
    return {hari: h, masuk: t(r[1]) || CFG.DEF.masuk, terlambat: t(r[2]) || CFG.DEF.terlambat, pulang: t(r[3]) || CFG.DEF.pulang};
  });
}
function cfgToday_() {   // null di Sabtu/Minggu
  const i = Number(now_('u')) - 1;
  return i < 5 ? cfg_()[i] : null;
}

function setCfg_(p) {
  if (String(p.pin) !== CFG.PIN) return {ok: false, error: 'PIN salah'};
  const re = /^\d{2}:\d{2}$/;
  if (!p.hari || p.hari.length !== 5) return {ok: false, error: 'Data jadwal tidak lengkap'};
  for (const h of p.hari) {
    if (!re.test(h.masuk) || !re.test(h.terlambat) || !re.test(h.pulang)) return {ok: false, error: 'Format jam salah'};
    if (!(h.masuk <= h.terlambat && h.terlambat < h.pulang)) return {ok: false, error: 'Urutan jam ' + h.hari + ' tidak valid'};
  }
  const s = sh_('Pengaturan', H_CFG);
  s.getRange(2, 1, 5, 4).setValues(HARI.map((n, i) => [n, p.hari[i].masuk, p.hari[i].terlambat, p.hari[i].pulang]));
  return {ok: true};
}

// Rekap absen masuk + pulang untuk diunduh. p: {pin, dari:'yyyy-MM-dd', sampai:'yyyy-MM-dd'}
function rekap_(p) {
  if (String(p.pin) !== CFG.PIN) return {ok: false, error: 'PIN salah'};
  const dari = String(p.dari), sampai = String(p.sampai);
  const rows = sh_('Absensi', H_ABSEN).getDataRange().getValues().slice(1)
    .filter(r => { const t = String(r[0]); return t >= dari && t <= sampai; })
    .slice(0, 20000)
    .map(r => ({tgl: String(r[0]), nis: String(r[1]), nama: r[2], kelas: r[3], jam: String(r[4] || ''), status: r[5], pulang: String(r[6] || '')}));
  return {ok: true, rows: rows};
}

function init_() {
  const rows = sh_('Siswa', H_SISWA).getDataRange().getValues().slice(1);
  const gallery = [];
  rows.forEach(r => {
    if (!r[3]) return;
    try { gallery.push({nis: String(r[0]), nama: r[1], kelas: r[2], d: JSON.parse(r[3])}); } catch (_) {}
  });
  return {ok: true, gallery: gallery, today: today_(), total: rows.length, cfg: {hari: cfg_()}};
}

function today_() {
  const t = now_('yyyy-MM-dd');
  return sh_('Absensi', H_ABSEN).getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === t)
    .map(r => ({nis: String(r[1]), nama: r[2], kelas: r[3], jam: String(r[4] || ''), status: r[5], pulang: String(r[6] || '')}));
}

// jenis: 'masuk' | 'pulang'
function absen_(nis, jenis) {
  const siswa = sh_('Siswa', H_SISWA).getDataRange().getValues().slice(1).find(r => String(r[0]) === nis);
  if (!siswa) return {ok: false, error: 'NIS tidak terdaftar'};
  const c = cfgToday_();
  if (!c) return {ok: false, error: 'Hari ini libur (jadwal hanya Senin–Jumat)'};
  const s = sh_('Absensi', H_ABSEN), t = now_('yyyy-MM-dd'), jam = now_('HH:mm');
  const data = s.getDataRange().getValues();
  let idx = -1;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]) === t && String(data[i][1]) === nis) { idx = i; break; }
  const row = idx > 0 ? data[idx] : null;
  const jn = jenis === 'pulang' ? 'pulang' : 'masuk';
  const base = {ok: true, jenis: jn, nama: siswa[1], kelas: siswa[2]};

  if (jn === 'masuk') {
    if (row && row[4]) return Object.assign({}, base, {dup: true, jam: String(row[4]), status: row[5], pulang: String(row[6] || ''), at: String(row[4])});
    const status = jam > c.terlambat ? 'Terlambat' : 'Hadir';
    if (row) s.getRange(idx + 1, 5, 1, 2).setValues([[jam, status]]);          // baris sudah ada (pulang tanpa masuk)
    else s.appendRow([t, nis, siswa[1], siswa[2], jam, status, '']);
    return Object.assign({}, base, {dup: false, jam: jam, status: status, pulang: String(row ? row[6] || '' : ''), at: jam});
  }

  if (row && row[6]) return Object.assign({}, base, {dup: true, jam: String(row[4] || ''), status: row[5], pulang: String(row[6]), at: String(row[6])});
  const awal = jam < c.pulang;
  if (row) s.getRange(idx + 1, 7).setValue(jam);
  else s.appendRow([t, nis, siswa[1], siswa[2], '', 'Tanpa Masuk', jam]);
  return Object.assign({}, base, {dup: false, jam: String(row ? row[4] || '' : ''), status: row ? row[5] : 'Tanpa Masuk', pulang: jam, at: jam, awal: awal});
}

function daftar_(p) {
  if (String(p.pin) !== CFG.PIN) return { ok: false, error: 'PIN salah' };
  if (!p.nis || !p.nama || !p.d || !p.d.length) return { ok: false, error: 'Data tidak lengkap' };
  const s = sh_('Siswa', H_SISWA);
  const rows = s.getDataRange().getValues();
  const wajah = JSON.stringify(p.d);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.nis)) {
      s.getRange(i + 1, 1, 1, 4).setValues([[String(p.nis), p.nama, p.kelas || '', wajah]]);
      return { ok: true, updated: true };
    }
  }
  s.appendRow([String(p.nis), p.nama, p.kelas || '', wajah]);
  return { ok: true, updated: false };
}
