/* ================================================================
   KONFIGURASI
   ================================================================ */
const SPREADSHEET_ID = '1azav1pZln8GUgVk964MyhR4H3yT-EIoBHMUQTvMJ5PQ';
const DRIVE_FOLDER_ID = '1bOSDRFqokgiNnjjs-PcXQAOTrpALzMzO';
const SHEET_AREA = 'MASTER_AREA';
const SHEET_ZONE = 'MASTER_ZONE';
const SHEET_SUB = 'SUBMISSIONS';
const MAX_FILE_BYTES = 200 * 1024;
const FALLBACK_FOLDER_NAME = 'Photo PIC Upload (Fallback)';

/* ================================================================
   ENTRY POINT
   ================================================================ */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('6S Smart Zone - PIC Area')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ================================================================
   MASTER DATA
   ================================================================ */
function getMasterData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetArea = ss.getSheetByName(SHEET_AREA);
  if (!sheetArea) throw new Error("Sheet MASTER_AREA tidak ditemukan. Cek nama sheet di Spreadsheet.");

  const areaRows = sheetArea.getDataRange().getValues();
  areaRows.shift();

  const areas = [];
  const uniqueGroups = new Set();

  areaRows.forEach(r => {
    const id = String(r[0] || '').trim();
    if (!id) return;
    const groupId = String(r[2] || '').trim();
    const status = String(r[3] || '').toUpperCase().trim();
    if (groupId) uniqueGroups.add(groupId);
    if (status === 'ACTIVE' || status === '') {
      areas.push({
        id: id,
        name: String(r[1] || '').trim(),
        groupId: groupId,
        layoutUrl: /^https?:\/\//.test(String(r[5] || '')) ? String(r[5]).trim() : ''
      });
    }
  });

  let plants = Array.from(uniqueGroups).sort().map(gid => {
    let plantName = gid;
    if (gid.toUpperCase().startsWith('A')) {
      const num = parseInt(gid.replace(/[^0-9]/g, ''), 10);
      if (!isNaN(num)) plantName = 'Plant ' + num;
    }
    return { id: gid, name: plantName };
  });

  if (plants.length === 0) {
    plants = [
      { id: 'A01', name: 'Plant 1' }, { id: 'A02', name: 'Plant 2' },
      { id: 'A03', name: 'Plant 3' }, { id: 'A04', name: 'Plant 4' }
    ];
  }

  const sheetZone = ss.getSheetByName(SHEET_ZONE);
  const zoneRows = sheetZone ? sheetZone.getDataRange().getValues() : [];
  if (zoneRows.length > 0) zoneRows.shift();

  const zones = zoneRows
    .filter(r => r[0] && String(r[0]).trim() !== '')
    .filter(r => String(r[4] || '').toUpperCase().trim() === '' || String(r[4] || '').toUpperCase().trim() === 'ACTIVE')
    .map(r => ({
      id: String(r[0]).trim(), name: String(r[1] || '').trim(),
      areaId: String(r[2]).trim(),
      layoutUrl: /^https?:\/\//.test(String(r[3] || '')) ? String(r[3]).trim() : ''
    }));

  return { plants, areas, zones };
}

/* ================================================================
   HELPER KOLOM BERDASARKAN HEADER (urutan kolom bebas)
   ================================================================ */
function hmap_(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const k = String(h).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (k && m[k] === undefined) m[k] = i;
  });
  return m;
}
function col_(m, names, def) {
  for (const n of names) if (m[n] !== undefined) return m[n];
  return def;
}
// Susunan kolom tetap: A-Q (I=NIK, L=PHOTO_PIC_URL, Q=PHOTO_PIC_KB)
const COLS_ = {
  ts: [['TIMESTAMP'], 0], email: [['SUBMITTEREMAIL', 'EMAIL'], 1],
  plant: [['PLANT'], 2], areaId: [['AREAID'], 3],
  areaName: [['AREANAME'], 4], zoneId: [['ZONEID'], 5],
  zoneName: [['ZONENAME'], 6], name: [['PICNAME', 'NAMAPIC'], 7],
  nik: [['NIK'], 8], dept: [['DEPARTMENT', 'DEPARTEMEN'], 9],
  grup: [['GROUP', 'GRUP', 'SHIFT'], 10],
  url: [['PHOTOPICURL', 'PHOTOURL'], 11],
  notes: [['NOTES', 'KETERANGAN'], 12],
  lat: [['GPSLAT'], 13], lng: [['GPSLNG'], 14],
  ua: [['USERAGENT'], 15], kb: [['PHOTOPICKB'], 16]
};
function idx_(m, key) { return col_(m, COLS_[key][0], COLS_[key][1]); }

/* ================================================================
   SUBMIT DATA - format A-Q:
   H PIC_NAME, I NIK, J DEPARTMENT, K GROUP, L PHOTO_PIC_URL,
   M NOTES, N GPS_LAT, O GPS_LNG, P USER_AGENT, Q PHOTO_PIC_KB
   ================================================================ */
function submitForm(payload) {
  const required = ['plantName', 'areaId', 'areaName', 'zoneId', 'zoneName', 'namaPic', 'nik', 'departemen', 'grup'];
  for (const k of required) {
    if (!payload[k] || String(payload[k]).trim() === '')
      throw new Error('Field wajib belum lengkap: ' + k);
  }
  const nik = String(payload.nik).trim();
  if (!/^[A-Za-z0-9]{4,20}$/.test(nik)) throw new Error('NIK harus 4-20 karakter huruf/angka tanpa spasi.');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_SUB);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SUB);
    sheet.appendRow(PIC_HEADERS_);
    sheet.setFrozenRows(1);
  }
  ensureHeaders_(sheet); // validasi posisi sebelum upload, tanpa mengubah header/data

  const picUpload = uploadToDrive(
    payload.fotoPic, payload.namaPic, 'PIC',
    payload.plantName, `${payload.areaId} — ${payload.areaName}`, `${payload.zoneId} — ${payload.zoneName}`
  );

  // Tanpa login: SUBMITTER_EMAIL selalu kosong.
  const email = '';

  const headers = PIC_HEADERS_;
  const m = hmap_(headers);
  const row = new Array(headers.length).fill('');
  const put = (key, val) => { const i = idx_(m, key); row[i] = val; };
  put('ts', new Date()); put('email', email); put('plant', payload.plantName);
  put('areaId', payload.areaId); put('areaName', payload.areaName);
  put('zoneId', payload.zoneId); put('zoneName', payload.zoneName);
  put('name', payload.namaPic); put('nik', nik);
  put('dept', payload.departemen); put('grup', payload.grup);
  put('url', picUpload.url); put('notes', payload.keterangan || '');
  put('lat', payload.gpsLat || ''); put('lng', payload.gpsLng || '');
  put('ua', payload.userAgent || ''); put('kb', picUpload.sizeKB);

  const r = sheet.getLastRow() + 1;
  sheet.getRange(r, 9).setNumberFormat('@'); // format teks pada I (NIK)
  sheet.getRange(r, 1, 1, row.length).setValues([row]);

  const note = picUpload.fallback
    ? ' Catatan: folder utama tidak dapat diakses, foto disimpan di folder fallback.'
    : '';
  return { success: true, message: `Data tersimpan (Grup ${payload.grup}).` + note, picUrl: picUpload.url };
}

/* ================================================================
   OUTPUT: DAFTAR ZONA & PAPAN PIC
   ================================================================ */
function stripId_(s) {            // "Z01 — Nama" -> "Nama"
  s = String(s || '');
  const i = s.indexOf(' — ');
  return i >= 0 ? s.substring(i + 3).trim() : s.trim();
}

// Validasi skema; tidak pernah menebak NIK dari isi kolom atau ukuran foto.
function ensureHeaders_(sheet) {
  if (sheet.getLastColumn() < PIC_HEADERS_.length) {
    throw new Error('Kolom SUBMISSIONS kurang dari A-Q. Jalankan setupHeaders().');
  }
  const actual = sheet.getRange(1, 1, 1, PIC_HEADERS_.length).getDisplayValues()[0];
  const errors = PIC_HEADERS_.flatMap((wanted, i) =>
    String(actual[i] || '').trim() === wanted
      ? [] : [String.fromCharCode(65 + i) + ': ' + wanted]
  );
  if (errors.length) {
    throw new Error('Header SUBMISSIONS belum sesuai (' + errors.join(', ') + '). Jalankan setupHeaders() setelah backup spreadsheet.');
  }
}

// NIK hanya boleh dibaca dari kolom I; data lama yang kosong/tidak valid tetap kosong.
function cleanNik_(value) {
  const nik = String(value == null ? '' : value).trim();
  return /^[A-Za-z0-9]{4,20}$/.test(nik) ? nik : '';
}

// Normalisasi grup: "Grup a", "GROUP B", "A" -> A/B/C/D; selain itu -> DT
function normGrup_(g) {
  const s = String(g || '').toUpperCase().replace(/GRUP|GROUP|SHIFT/g, '').trim();
  return ['A', 'B', 'C', 'D'].indexOf(s) >= 0 ? s : 'DT';
}

function getLatestPics_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SUB);
  if (!sheet || sheet.getLastRow() < 2) return [];
  ensureHeaders_(sheet);
  const range = sheet.getRange(1, 1, sheet.getLastRow(), PIC_HEADERS_.length);
  const all = range.getValues();
  const nikText = range.getDisplayValues(); // simpan nol awal pada NIK
  const m = hmap_(all.shift());
  nikText.shift();
  const I = k => idx_(m, k);
  const map = {};
  all.forEach((r, rowIndex) => {
    const zoneId = String(r[I('zoneId')] || '').trim();
    if (!zoneId) return;
    const nik = cleanNik_(nikText[rowIndex][8]); // I, bukan Q (PHOTO_PIC_KB)
    const name = String(r[I('name')] || '');
    const key = zoneId + '|' + (nik || name.toLowerCase().trim());
    const t = r[I('ts')];
    const ts = t instanceof Date ? t.getTime() : 0;
    // URL foto: dari kolom header; jika bukan URL, cari sel mana pun yang berisi link Drive
    let url = String(r[I('url')] || '');
    if (!/^https?:\/\//.test(url)) {
      url = String(r.filter(v => /drive\.google\.com/.test(String(v)))[0] || '');
    }
    if (!map[key] || ts >= map[key].ts) {
      map[key] = {
        ts, zoneId, name, dept: String(r[I('dept')] || ''),
        grup: normGrup_(r[I('grup')]), url, nik
      };
    }
  });
  return Object.keys(map).map(k => map[k]);
}

function getZoneList() {
  const m = getMasterData();
  const plantName = {}; m.plants.forEach(p => plantName[p.id] = p.name);
  const areaById = {}; m.areas.forEach(a => areaById[a.id] = a);
  const stat = {};
  getLatestPics_().forEach(p => {
    const s = stat[p.zoneId] || (stat[p.zoneId] = { count: 0, last: 0 });
    s.count++; if (p.ts > s.last) s.last = p.ts;
  });
  return m.zones.filter(z => areaById[z.areaId]).map(z => {
    const a = areaById[z.areaId], s = stat[z.id] || { count: 0, last: 0 };
    return {
      zoneId: z.id, zoneName: z.name, areaId: a.id, areaName: a.name,
      plantId: a.groupId, plantName: plantName[a.groupId] || a.groupId,
      count: s.count, last: s.last
    };
  }).sort((x, y) => (x.plantId + x.areaId + x.zoneId).localeCompare(y.plantId + y.areaId + y.zoneId));
}

function getZoneBoard(zoneId) {
  const list = getZoneList();
  const info = list.filter(z => z.zoneId === zoneId)[0];
  if (!info) throw new Error('Zona tidak ditemukan.');
  const pics = getLatestPics_().filter(p => p.zoneId === zoneId)
    .sort((a, b) => a.ts - b.ts)
    .map(p => {
      const f = photoDataUrl_(p.url);
      return { name: p.name, nik: p.nik, dept: p.dept, grup: p.grup, photo: f.data, err: f.err };
    });
  const jobdesks = typeof getZoneJobdesks_ === 'function' ? getZoneJobdesks_(zoneId) : [];
  const jobUpdated = jobdesks.reduce((max, j) => Math.max(max, j.updatedAt || 0), 0);
  return { info, pics, jobdesks, updated: Math.max(info.last || 0, jobUpdated) };
}

function photoDataUrl_(url) {
  try {
    const m = String(url).match(/\/d\/([-\w]+)/) || String(url).match(/[?&]id=([-\w]+)/);
    if (!m) return { data: '', err: 'URL foto tidak valid: ' + url };
    const blob = DriveApp.getFileById(m[1]).getBlob();
    return { data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()), err: '' };
  } catch (e) { return { data: '', err: String(e.message || e) }; }
}

/* ================================================================
   MENU FOTO: DAFTAR ORANG, EDIT DATA & GANTI FOTO
   ================================================================ */
function driveFileId_(url) {
  const s = String(url || '');
  const m = s.match(/\/d\/([-\w]+)/) || s.match(/[?&]id=([-\w]+)/);
  return m ? m[1] : '';
}

function driveThumb_(url, size) {
  const id = driveFileId_(url);
  return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + (size || 400) : '';
}

function picSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SUB);
  if (!sheet) throw new Error('Sheet SUBMISSIONS tidak ditemukan.');
  ensureHeaders_(sheet);
  return sheet;
}

function picRowUrl_(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, PIC_HEADERS_.length).getDisplayValues()[0];
  const hm = hmap_(headers);
  return String(sheet.getRange(row, idx_(hm, 'url') + 1).getValue() || '');
}

// Satu kartu per orang (submit terbaru), lengkap dengan nomor baris untuk edit.
function getPicList() {
  const sheet = picSheet_();
  if (sheet.getLastRow() < 2) return [];
  const range = sheet.getRange(1, 1, sheet.getLastRow(), PIC_HEADERS_.length);
  const all = range.getValues();
  const nikText = range.getDisplayValues();
  const m = hmap_(all.shift());
  nikText.shift();
  const I = k => idx_(m, k);
  const map = {};

  all.forEach((r, i) => {
    const zoneId = String(r[I('zoneId')] || '').trim();
    if (!zoneId) return;
    const nik = String(nikText[i][I('nik')] || '').trim();
    const name = String(r[I('name')] || '').trim();
    const key = zoneId + '|' + (nik || name.toLowerCase());
    const t = r[I('ts')];
    const ts = t instanceof Date ? t.getTime() : 0;
    let url = String(r[I('url')] || '');
    if (!/^https?:\/\//.test(url)) {
      url = String(r.filter(v => /drive\.google\.com/.test(String(v)))[0] || '');
    }
    if (!map[key] || ts >= map[key].ts) {
      map[key] = {
        row: i + 2, ts: ts,
        plant: String(r[I('plant')] || ''),
        areaId: String(r[I('areaId')] || ''),
        areaName: String(r[I('areaName')] || ''),
        zoneId: zoneId,
        zoneName: String(r[I('zoneName')] || ''),
        name: name, nik: nik,
        dept: String(r[I('dept')] || ''),
        grup: normGrup_(r[I('grup')]),
        notes: String(r[I('notes')] || ''),
        url: url, thumb: driveThumb_(url, 400),
        kb: r[I('kb')]
      };
    }
  });

  return Object.keys(map).map(k => map[k])
    .sort((a, b) => (a.plant + a.areaId + a.zoneId + a.name)
      .localeCompare(b.plant + b.areaId + b.zoneId + b.name));
}

function updatePicData(payload) {
  const sheet = picSheet_();
  const row = parseInt(payload && payload.row, 10);
  if (!row || row < 2 || row > sheet.getLastRow()) throw new Error('Baris data tidak valid.');

  const plantName = String(payload.plantName || '').trim();
  const areaId = String(payload.areaId || '').trim();
  const areaName = String(payload.areaName || '').trim();
  const zoneId = String(payload.zoneId || '').trim();
  const zoneName = String(payload.zoneName || '').trim();
  const name = String(payload.name || '').trim();
  const nik = String(payload.nik || '').trim();
  const dept = String(payload.dept || '').trim();
  const grup = String(payload.grup || '').trim();
  const notes = String(payload.notes || '').trim();

  if (!plantName || !areaId || !areaName || !zoneId || !zoneName) {
    throw new Error('Plant, Area, dan Zona wajib dipilih.');
  }
  if (!name) throw new Error('Nama PIC wajib diisi.');
  if (!/^[A-Za-z0-9]{4,20}$/.test(nik)) throw new Error('NIK harus 4-20 karakter huruf/angka tanpa spasi.');
  if (!dept) throw new Error('Departemen wajib diisi.');
  if (!grup) throw new Error('Grup wajib dipilih.');

  const md = getMasterData();
  const zone = md.zones.filter(z => z.id === zoneId)[0];
  if (!zone) throw new Error('Zona tidak ditemukan di master.');
  if (zone.areaId !== areaId) throw new Error('Zona tidak sesuai dengan Area.');
  const area = md.areas.filter(a => a.id === areaId)[0];
  if (!area) throw new Error('Area tidak ditemukan di master.');
  const plant = md.plants.filter(p => p.id === area.groupId)[0];
  if (plant && plant.name !== plantName) throw new Error('Area tidak sesuai dengan Plant.');

  const headers = sheet.getRange(1, 1, 1, PIC_HEADERS_.length).getDisplayValues()[0];
  const hm = hmap_(headers);
  const set = (key, val) => sheet.getRange(row, idx_(hm, key) + 1).setValue(val);

  set('plant', plantName);
  set('areaId', areaId);
  set('areaName', areaName);
  set('zoneId', zoneId);
  set('zoneName', zoneName);
  set('name', name);
  sheet.getRange(row, idx_(hm, 'nik') + 1).setNumberFormat('@').setValue(nik);
  set('dept', dept);
  set('grup', normGrup_(grup));
  set('notes', notes);
  SpreadsheetApp.flush();

  return { success: true, message: 'Data PIC berhasil diperbarui.' };
}

function getPicPhoto(row) {
  const sheet = picSheet_();
  const r = parseInt(row, 10);
  if (!r || r < 2 || r > sheet.getLastRow()) throw new Error('Baris data tidak valid.');
  const url = picRowUrl_(sheet, r);
  if (!/^https?:\/\//.test(url)) throw new Error('Foto PIC tidak ditemukan.');
  const f = photoDataUrl_(url);
  if (!f.data) throw new Error(f.err || 'Foto tidak dapat dimuat.');
  return { data: f.data };
}

function updatePicPhoto(payload) {
  const sheet = picSheet_();
  const row = parseInt(payload && payload.row, 10);
  if (!row || row < 2 || row > sheet.getLastRow()) throw new Error('Baris data tidak valid.');

  const headers = sheet.getRange(1, 1, 1, PIC_HEADERS_.length).getDisplayValues()[0];
  const hm = hmap_(headers);
  const vals = sheet.getRange(row, 1, 1, PIC_HEADERS_.length).getValues()[0];
  const get = key => vals[idx_(hm, key)];
  const plantName = String(get('plant') || '').trim();
  const areaId = String(get('areaId') || '').trim();
  const areaName = String(get('areaName') || '').trim();
  const zoneId = String(get('zoneId') || '').trim();
  const zoneName = String(get('zoneName') || '').trim();
  const name = String(get('name') || '').trim();
  const oldUrl = String(get('url') || '').trim();

  const upload = uploadToDrive(
    payload.fotoPic, name, 'PIC',
    plantName || 'UNKNOWN', `${areaId} — ${areaName}`, `${zoneId} — ${zoneName}`
  );

  sheet.getRange(row, idx_(hm, 'url') + 1).setValue(upload.url);
  sheet.getRange(row, idx_(hm, 'kb') + 1).setValue(upload.sizeKB);
  SpreadsheetApp.flush();

  const oldId = driveFileId_(oldUrl);
  if (oldId && oldId !== driveFileId_(upload.url)) {
    try { DriveApp.getFileById(oldId).setTrashed(true); }
    catch (e) { console.warn('Foto lama tidak dapat dihapus: ' + (e && e.message ? e.message : e)); }
  }

  const note = upload.fallback ? ' Foto disimpan di folder fallback.' : '';
  return {
    success: true,
    message: 'Foto PIC berhasil diperbarui.' + note,
    url: upload.url,
    thumb: driveThumb_(upload.url, 400)
  };
}

// Jalankan manual dari editor untuk cek data: debugZone('1BS5')
function debugZone(zoneId) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SUB);
  Logger.log('HEADER: ' + JSON.stringify(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]));
  getLatestPics_().filter(p => p.zoneId === zoneId)
    .forEach(p => Logger.log(p.name + ' | grup=' + p.grup + ' | ' + p.url));
  Logger.log('selesai');
}

/* ================================================================
   UPLOAD KE DRIVE
   ================================================================ */
function uploadToDrive(fileData, picName, prefix, plantName, areaLabel, zoneLabel) {
  if (!fileData || !fileData.data) throw new Error('Foto ' + prefix + ' kosong.');
  const base64 = fileData.data.split(',')[1] || fileData.data;
  const bytes = Utilities.base64Decode(base64);
  const sizeKB = Math.round(bytes.length / 1024);
  if (bytes.length > MAX_FILE_BYTES) throw new Error(`Foto ${prefix} ${sizeKB} KB melebihi batas 200 KB.`);

  const uploadRoot = getUploadRoot_();
  const root = uploadRoot.folder;
  const plantFolder = getOrCreateFolder(root, sanitizeFolderName(plantName));
  const areaFolder = getOrCreateFolder(plantFolder, sanitizeFolderName(areaLabel));
  const zoneFolder = getOrCreateFolder(areaFolder, sanitizeFolderName(zoneLabel));

  const ts = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
  const safe = sanitizeFolderName(picName || 'unknown');
  const filename = `${ts}_${prefix}_${safe}.jpg`;

  const blob = Utilities.newBlob(bytes, fileData.mimeType || 'image/jpeg', filename);
  const file = zoneFolder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    console.warn('setSharing gagal: ' + (e && e.message ? e.message : e));
  }
  return { url: file.getUrl(), sizeKB, fallback: uploadRoot.fallback };
}

function getUploadRoot_() {
  try {
    return { folder: DriveApp.getFolderById(DRIVE_FOLDER_ID), fallback: false };
  } catch (e) {
    const detail = e && e.message ? e.message : String(e);
    console.warn('Folder utama ' + DRIVE_FOLDER_ID + ' tidak dapat diakses: ' + detail);
    try {
      const root = DriveApp.getRootFolder();
      return { folder: getOrCreateFolder(root, FALLBACK_FOLDER_NAME), fallback: true };
    } catch (e2) {
      const detail2 = e2 && e2.message ? e2.message : String(e2);
      throw new Error(
        'Folder upload tidak dapat diakses. Folder utama (' + DRIVE_FOLDER_ID + ') gagal: ' + detail +
        '. Folder fallback juga gagal: ' + detail2 +
        '. Pastikan web app dieksekusi oleh akun yang punya izin edit folder Drive.'
      );
    }
  }
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function sanitizeFolderName(name) {
  return String(name || '').replace(/[\\/:*?"<>|]/g, '-').trim().substring(0, 120);
}


const PIC_HEADERS_ = [
  'TIMESTAMP',        // A
  'SUBMITTER_EMAIL',  // B
  'PLANT',            // C
  'AREA_ID',          // D
  'AREA_NAME',        // E
  'ZONE_ID',          // F
  'ZONE_NAME',        // G
  'PIC_NAME',         // H
  'NIK',              // I
  'DEPARTMENT',       // J
  'GROUP',            // K
  'PHOTO_PIC_URL',    // L
  'NOTES',            // M
  'GPS_LAT',          // N
  'GPS_LNG',          // O
  'USER_AGENT',       // P
  'PHOTO_PIC_KB'      // Q
];

function setupHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SUB);

  if (!sheet) {
    throw new Error('Sheet SUBMISSIONS tidak ditemukan!');
  }

  const headers = PIC_HEADERS_;

  // Pastikan jumlah kolom mencukupi
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      headers.length - sheet.getMaxColumns()
    );
  }

  // Periksa apakah baris pertama sudah memiliki header
  const firstRow = sheet.getRange(1, 1, 1, headers.length)
    .getDisplayValues()[0];

  const normalize = v => String(v || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');

  const hasHeader =
    normalize(firstRow[0]) === 'TIMESTAMP' ||
    firstRow.filter((v, i) => normalize(v) === normalize(headers[i])).length >= 3;

  const isEmpty = firstRow.every(
    v => String(v).trim() === ''
  );

  // Sisipkan baris hanya jika baris pertama berisi data
  if (!hasHeader && !isEmpty) {
    sheet.insertRowBefore(1);
  }

  // Tulis header A sampai Q
  const range = sheet.getRange(1, 1, 1, headers.length);

  range.setValues([headers]);

  // Desain header
  range
    .setBackground('#17365D')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 42);

  // Pengaturan lebar kolom
  sheet.setColumnWidths(1, 17, 130);

  sheet.setColumnWidth(1, 170);   // Timestamp
  sheet.setColumnWidth(2, 200);   // Email
  sheet.setColumnWidth(5, 270);   // Area name
  sheet.setColumnWidth(7, 270);   // Zone name
  sheet.setColumnWidth(8, 190);   // PIC name
  sheet.setColumnWidth(9, 140);   // NIK
  sheet.setColumnWidth(10, 160);  // Department
  sheet.setColumnWidth(11, 130);  // Group
  sheet.setColumnWidth(12, 350);  // Photo URL
  sheet.setColumnWidth(13, 200);  // Notes

  // Kolom I (NIK) menggunakan format teks
  sheet.getRange(2, 9, sheet.getMaxRows() - 1, 1)
    .setNumberFormat('@');

  SpreadsheetApp.flush();

  Logger.log('Header SUBMISSIONS berhasil diperbaiki!');
}
