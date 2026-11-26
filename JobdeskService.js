const SHEET_JOB = 'JOBDESK';
const JD_HEADERS_ = [
  'JOBDESK_ID',
  'TIMESTAMP',
  'UPDATED_AT',
  'PLANT',
  'AREA_ID',
  'AREA_NAME',
  'ZONE_ID',
  'ZONE_NAME',
  'PIC_NAME',
  'NIK',
  'GROUP',
  'CATEGORY',
  'KEGIATAN',
  'KRITERIA'
];
const JD_CATEGORIES_ = ['RINGKAS', 'RAPI', 'RESIK', 'RAWAT', 'RAJIN', 'SAFETY'];

function getJobdeskList() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_JOB);
  if (!sheet || sheet.getLastRow() < 2) return [];
  validateJobdeskHeaders_(sheet);
  const range = sheet.getRange(1, 1, sheet.getLastRow(), JD_HEADERS_.length);
  const rows = range.getValues();
  rows.shift();
  return rows.map(r => ({
    id: String(r[0] || ''),
    timestamp: jdTime_(r[1]),
    updatedAt: jdTime_(r[2]),
    plant: String(r[3] || ''),
    areaId: String(r[4] || ''),
    areaName: String(r[5] || ''),
    zoneId: String(r[6] || ''),
    zoneName: String(r[7] || ''),
    category: String(r[11] || ''),
    kegiatan: String(r[12] || ''),
    kriteria: String(r[13] || '')
  })).filter(x => x.id);
}

function getZoneJobdesks_(zoneId) {
  const key = String(zoneId || '').trim();
  if (!key) return [];
  try {
    return getJobdeskList()
      .filter(it => it.zoneId === key)
      .sort((a, b) => {
        const ca = JD_CATEGORIES_.indexOf(a.category);
        const cb = JD_CATEGORIES_.indexOf(b.category);
        const oa = ca < 0 ? 99 : ca;
        const ob = cb < 0 ? 99 : cb;
        if (oa !== ob) return oa - ob;
        return (a.updatedAt || 0) - (b.updatedAt || 0);
      })
      .map(it => ({
        category: it.category,
        kegiatan: it.kegiatan,
        kriteria: it.kriteria,
        updatedAt: it.updatedAt
      }));
  } catch (e) {
    console.warn('Jobdesk zona ' + key + ' tidak dapat dimuat: ' + (e && e.message ? e.message : e));
    return [];
  }
}

function saveJobdeskBatch(payload) {
  const head = validateJobdeskHead_(payload);
  const items = validateJobdeskItems_(payload.items);
  const sheet = ensureJobdeskSheet_();
  const now = new Date();
  const rows = items.map(it => [
    Utilities.getUuid(), now, now, head.plantName, head.areaId, head.areaName,
    head.zoneId, head.zoneName, '', '', '',
    it.category, it.kegiatan, it.kriteria
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, JD_HEADERS_.length).setValues(rows);
  return { success: true, count: rows.length, message: rows.length + ' jobdesk berhasil disimpan.' };
}

function updateJobdeskItem(payload) {
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('ID jobdesk tidak valid.');
  const items = validateJobdeskItems_([{ category: payload.category, kegiatan: payload.kegiatan, kriteria: payload.kriteria }]);
  const item = items[0];
  const sheet = ensureJobdeskSheet_();
  const row = findJobdeskRow_(sheet, id);
  sheet.getRange(row, 3).setValue(new Date());
  sheet.getRange(row, 12, 1, 3).setValues([[item.category, item.kegiatan, item.kriteria]]);
  return { success: true, message: 'Jobdesk berhasil diperbarui.' };
}

function deleteJobdeskItem(id) {
  const key = String(id || '').trim();
  if (!key) throw new Error('ID jobdesk tidak valid.');
  const sheet = ensureJobdeskSheet_();
  const row = findJobdeskRow_(sheet, key);
  sheet.deleteRow(row);
  return { success: true, message: 'Jobdesk berhasil dihapus.' };
}

function bersihkanKolomPicJobdesk() {
  const sheet = ensureJobdeskSheet_();
  const last = sheet.getLastRow();
  if (last < 2) {
    Logger.log('Tidak ada data jobdesk untuk dibersihkan.');
    return;
  }
  sheet.getRange(2, 9, last - 1, 3).clearContent();
  SpreadsheetApp.flush();
  Logger.log('Kolom PIC_NAME, NIK, dan GROUP pada JOBDESK berhasil dikosongkan.');
}

function setupJobdeskHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_JOB);
  if (!sheet) sheet = ss.insertSheet(SHEET_JOB);
  if (sheet.getLastRow() > 1) {
    throw new Error('Sheet JOBDESK sudah berisi data. Backup lalu kosongkan dulu sebelum menjalankan setupJobdeskHeaders().');
  }
  if (sheet.getLastRow() > 0) {
    const first = sheet.getRange(1, 1, 1, Math.max(sheet.getMaxColumns(), JD_HEADERS_.length)).getDisplayValues()[0];
    if (first.some(v => String(v).trim() !== '')) {
      const matches = JD_HEADERS_.every((h, i) => String(first[i] || '').trim() === h);
      if (matches) {
        Logger.log('Header JOBDESK sudah sesuai.');
        return;
      }
      throw new Error('Baris pertama sheet JOBDESK berisi data lain. Backup lalu kosongkan dulu sebelum menjalankan setupJobdeskHeaders().');
    }
  }
  writeJobdeskHeaders_(sheet);
  Logger.log('Header JOBDESK siap.');
}

function migrateJobdeskToKegiatanKriteria() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_JOB);
  if (!sheet || sheet.getLastRow() === 0) {
    Logger.log('Sheet JOBDESK kosong. Jalankan setupJobdeskHeaders() untuk membuat header baru.');
    return;
  }
  const lastCol = Math.max(sheet.getMaxColumns(), JD_HEADERS_.length);
  const header = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || '').trim().toUpperCase());
  if (header[12] === 'KEGIATAN' && header[13] === 'KRITERIA') {
    Logger.log('Sheet JOBDESK sudah termigrasi (KEGIATAN & KRITERIA). Tidak ada yang diubah.');
    return;
  }
  if (header[11] !== 'CATEGORY' || header[12] !== 'DESCRIPTION') {
    throw new Error('Struktur header JOBDESK tidak dikenali. Kolom L harus CATEGORY dan M harus DESCRIPTION.');
  }
  if (sheet.getMaxColumns() < JD_HEADERS_.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), JD_HEADERS_.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 13).setValue('KEGIATAN');
  sheet.getRange(1, 14).setValue('KRITERIA');
  const last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 14, last - 1, 1).clearContent();
  }
  writeJobdeskHeaders_(sheet);
  SpreadsheetApp.flush();
  Logger.log('Migrasi JOBDESK selesai: DESCRIPTION menjadi KEGIATAN, KRITERIA dikosongkan. Total ' + Math.max(last - 1, 0) + ' baris.');
}

function ensureJobdeskSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_JOB);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_JOB);
    writeJobdeskHeaders_(sheet);
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    writeJobdeskHeaders_(sheet);
    return sheet;
  }
  validateJobdeskHeaders_(sheet);
  return sheet;
}

function writeJobdeskHeaders_(sheet) {
  if (sheet.getMaxColumns() < JD_HEADERS_.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), JD_HEADERS_.length - sheet.getMaxColumns());
  }
  const range = sheet.getRange(1, 1, 1, JD_HEADERS_.length);
  range.setValues([JD_HEADERS_]);
  range.setBackground('#17365D').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 42);
  sheet.setColumnWidths(1, JD_HEADERS_.length, 140);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 165);
  sheet.setColumnWidth(3, 165);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 240);
  sheet.setColumnWidth(9, 190);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 90);
  sheet.setColumnWidth(12, 120);
  sheet.setColumnWidth(13, 420);
  sheet.setColumnWidth(14, 420);
  sheet.getRange(2, 10, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  SpreadsheetApp.flush();
}

function validateJobdeskHeaders_(sheet) {
  if (sheet.getMaxColumns() < JD_HEADERS_.length) {
    throw new Error('Kolom JOBDESK kurang dari A-N. Backup spreadsheet lalu jalankan setupJobdeskHeaders().');
  }
  const actual = sheet.getRange(1, 1, 1, JD_HEADERS_.length).getDisplayValues()[0];
  const errors = JD_HEADERS_.filter((wanted, i) => String(actual[i] || '').trim() !== wanted);
  if (errors.length) {
    throw new Error('Header JOBDESK belum sesuai: ' + errors.join(', ') + '. Backup spreadsheet lalu jalankan setupJobdeskHeaders().');
  }
}

function validateJobdeskHead_(payload) {
  const required = ['plantName', 'areaId', 'areaName', 'zoneId', 'zoneName'];
  required.forEach(k => {
    if (!payload[k] || String(payload[k]).trim() === '') throw new Error('Field wajib belum lengkap: ' + k);
  });
  return {
    plantName: String(payload.plantName).trim(),
    areaId: String(payload.areaId).trim(),
    areaName: String(payload.areaName).trim(),
    zoneId: String(payload.zoneId).trim(),
    zoneName: String(payload.zoneName).trim()
  };
}

function validateJobdeskItems_(items) {
  const list = Array.isArray(items) ? items : [];
  const out = [];
  list.forEach(it => {
    const category = String(it && it.category || '').toUpperCase().trim();
    const kegiatan = String(it && it.kegiatan || '').trim();
    const kriteria = String(it && it.kriteria || '').trim();
    if (!kegiatan && !kriteria) return;
    if (JD_CATEGORIES_.indexOf(category) < 0) throw new Error('Kategori 6S tidak valid: ' + category);
    if (!kegiatan) throw new Error('Kegiatan jobdesk wajib diisi.');
    if (!kriteria) throw new Error('Kriteria jobdesk wajib diisi.');
    if (kegiatan.length > 1000) throw new Error('Kegiatan maksimal 1000 karakter.');
    if (kriteria.length > 1000) throw new Error('Kriteria maksimal 1000 karakter.');
    out.push({ category: category, kegiatan: kegiatan, kriteria: kriteria });
  });
  if (!out.length) throw new Error('Isi minimal satu jobdesk.');
  return out;
}

function findJobdeskRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) throw new Error('Jobdesk tidak ditemukan.');
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === id) return i + 2;
  }
  throw new Error('Jobdesk tidak ditemukan atau sudah dihapus.');
}

function jdTime_(value) {
  if (value instanceof Date) return value.getTime();
  const t = new Date(value).getTime();
  return isNaN(t) ? 0 : t;
}
