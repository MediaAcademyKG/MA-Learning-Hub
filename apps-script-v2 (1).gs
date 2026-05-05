// ============================================================
// KG Media Learning Portal — Apps Script Web App (OPTIMIZED)
// Deploy: Execute as "Me", Access "Anyone"
// ============================================================

const SHEET_ID = '11T58NlMVg2NAiygt3IDl0nR-UZJoWZBLbui1TFZmpFE';

const CACHE = {
  access:    1800,
  searchbar: 1800,
  recap:     1800,
  part_full:  600,
  database:   600,
};

function doGet(e) {
  const params = e.parameter;
  const action = params.action || '';
  let result;

  try {
    switch (action) {

      case 'access':
        result = getCached('access', CACHE.access, () =>
          getSheetData('Access', ['NIK','nama','role','layer'])
        );
        break;

      case 'searchbar':
        const nikSB = normalizeNIK(params.nik || '');
        if (!nikSB) throw new Error('NIK diperlukan');
        result = getSearchBarByNIK(nikSB);
        break;

      case 'participation':
        const nikP = normalizeNIK(params.nik || '');
        if (!nikP) throw new Error('NIK diperlukan');
        result = getParticipationByNIK(nikP);
        break;

      case 'participation_full':
        result = getCached('part_full', CACHE.part_full, () =>
          getSheetData('Learning Participation')
        );
        break;

      case 'database_full':
        result = getCached('database', CACHE.database, () =>
          getSheetData('Database')
        );
        break;

      case 'recap':
        result = getCached('recap', CACHE.recap, () =>
          getSheetData('Recap')
        );
        break;

      // Warmup: hanya cache access + searchbar — TIDAK Database/Participation
      case 'warmup':
        getCached('access', CACHE.access, () =>
          getSheetData('Access', ['NIK','nama','role','layer'])
        );
        getCached('searchbar_all', CACHE.searchbar, () =>
          getSheetData('Search Bar')
        );
        result = { warmed: true, time: new Date().toISOString() };
        break;

      case 'ping':
        result = { status: 'ok', time: new Date().toISOString() };
        break;

      case 'clear_cache':
        clearAllCache();
        result = { cleared: true };
        break;

      default:
        throw new Error('Action tidak dikenal: ' + action);
    }

    return jsonResponse({ success: true, data: result });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function getSearchBarByNIK(inputNik) {
  const allRows = getCached('searchbar_all', CACHE.searchbar, () =>
    getSheetData('Search Bar')
  );
  return allRows.find(r =>
    normalizeNIK(r['NIK'] || r['nik'] || '') === inputNik
  ) || null;
}

// Scan per batch — jauh lebih cepat dari load seluruh tab
function getParticipationByNIK(inputNik) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Learning Participation');
  if (!sheet) throw new Error('Tab Learning Participation tidak ditemukan');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => h.toString().trim());
  const nikIdx = header.findIndex(h => h.toUpperCase() === 'NIK');
  if (nikIdx === -1) throw new Error('Kolom NIK tidak ditemukan');

  const BATCH = 500;
  const rows  = [];

  for (let startRow = 2; startRow <= lastRow; startRow += BATCH) {
    const rowCount = Math.min(BATCH, lastRow - startRow + 1);
    const data     = sheet.getRange(startRow, 1, rowCount, lastCol).getValues();
    for (const row of data) {
      if (normalizeNIK(row[nikIdx]) === inputNik) {
        const obj = {};
        header.forEach((h, j) => { obj[h] = (row[j] ?? '').toString(); });
        rows.push(obj);
      }
    }
  }
  return rows;
}

function getSheetData(tabName, colFilter) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Tab "' + tabName + '" tidak ditemukan');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const header      = data[0].map(h => h.toString().trim());
  const filterUpper = colFilter ? colFilter.map(c => c.toUpperCase()) : null;
  const rows        = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === '' || cell === null || cell === undefined)) continue;
    const obj = {};
    header.forEach((h, j) => {
      if (!filterUpper || filterUpper.includes(h.toUpperCase())) {
        obj[h] = (row[j] ?? '').toString();
      }
    });
    rows.push(obj);
  }
  return rows;
}

function getCached(key, ttl, fetchFn) {
  const cache = CacheService.getScriptCache();
  const meta  = cache.get('meta_' + key);

  if (meta) {
    try {
      const { chunks } = JSON.parse(meta);
      let full = '';
      for (let i = 0; i < chunks; i++) {
        const chunk = cache.get(key + '_c' + i);
        if (!chunk) { full = null; break; }
        full += chunk;
      }
      if (full) return JSON.parse(full);
    } catch(e) {}
  }

  const data  = fetchFn();
  try {
    const str   = JSON.stringify(data);
    const CHUNK = 90000;
    const n     = Math.ceil(str.length / CHUNK);
    const pairs = { ['meta_' + key]: JSON.stringify({ chunks: n }) };
    for (let i = 0; i < n; i++) {
      pairs[key + '_c' + i] = str.slice(i * CHUNK, (i + 1) * CHUNK);
    }
    cache.putAll(pairs, ttl);
  } catch(e) {}

  return data;
}

function clearAllCache() {
  const cache = CacheService.getScriptCache();
  ['access','searchbar_all','recap','part_full','database'].forEach(k => {
    cache.remove('meta_' + k);
    for (let i = 0; i < 20; i++) cache.remove(k + '_c' + i);
  });
}

function normalizeNIK(nik) {
  const s = (nik || '').toString().trim();
  if (/^\d+$/.test(s)) return String(parseInt(s, 10));
  return s.toUpperCase();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
