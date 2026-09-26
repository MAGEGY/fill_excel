/* Fill Excel from Documents — generic version.
   Any Excel: header row is auto-detected, every column becomes a field.
   Pictures: analyzed by an AI vision model (Gemini / OpenAI-compatible)
   or offline Tesseract OCR for standard ID fields. */
'use strict';

/* ---------------- state ---------------- */
const state = {
  excelBuffer: null, excelName: null, workbook: null,
  sheetName: null, headerRow: null,
  columns: [],          // { col, label, kind:'serial'|'date'|'text', skip, stdKey }
  rows: [],             // { id, values:{colIdx:str}, ignored:Set, docs:[] }
  seq: 0,
  settings: { provider: 'gemini', apiKey: '', model: 'gemini-2.5-flash', endpoint: 'https://api.openai.com/v1' },
};

/* Standard-field recognizers, used to classify arbitrary column headers
   and to map offline-OCR results onto them. Order matters. */
const FIELD_DEFS = [
  { key: 'serial',      re: /^(no\.?|#|s\/?n|serial|تسلسل|م)$/i },
  { key: 'name_ar',     re: /arabic|عربي/i },
  { key: 'name_en',     re: /student|name|الاسم|اسم/i },
  { key: 'nationality', re: /national|جنسية/i },
  { key: 'dob',         re: /birth|d\.?o\.?b|ميلاد|مواليد/i },
  { key: 'pob',         re: /place|محل/i },
  { key: 'passport',    re: /passport|جواز/i },
  { key: 'iqama',       re: /iqama|اقامة|إقامة|resident\s*id|هوية/i },
  { key: 'mom',         re: /mom|mother|والدة|\bأم\b/i },
  { key: 'dad',         re: /dad|father|والد|\bأب\b/i },
  { key: 'date',        re: /date|تاريخ|expir|issue|انتهاء|إصدار|اصدار/i },
  { key: 'phone',       re: /phone|mobile|tel|هاتف|جوال|رقم/i },
];
const KIND_RE = {
  serial: /^(no\.?|#|s\/?n|serial|تسلسل|م)$/i,
  date:   /date|تاريخ|ميلاد|birth|expir|issue|انتهاء|إصدار|اصدار|مواليد/i,
  idnum:  /passport|iqama|اقامة|إقامة|mobile|phone|tel|number|رقم|هاتف|جوال|id|هوية/i,
};
const ARABIC_HINT = /[\u0600-\u06FF]|arabic|عربي/i;

// Default columns when no Excel template is uploaded
const DEFAULT_COLS = [
  'No.', "Student's Name", 'Arabic Name', 'Nationality', 'Date of Birth',
  'Place of Birth', 'Passport Number', 'Iqama Number', 'Mom Mobile Number', 'Dad Mobile Number',
];

/* ---------------- helpers ---------------- */
const AR_RE = /[\u0600-\u06FF]/;
const hasArabic = s => AR_RE.test(s || '');
const normalizeDigits = s => (s || '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
const fmtDate = (d, m, y) => `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
function toDMY(y, m, d) {
  y = +y; m = +m; d = +d;
  if (y < 100) y += (y <= 49 ? 2000 : 1900);
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  return fmtDate(d, m, y);
}
function parseDMY(s) {
  const m = (s || '').trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (!m) return null;
  const d = +m[1], mo = +m[2], y = +m[3];
  return (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) ? new Date(y, mo - 1, d) : null;
}
function esc(s) {
  return String(s ?? '').replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
}
function colLettersToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
/* Merged cells: ExcelJS gives every slave cell the master's value, so a merged
   title row looks like N identical headers. We need the ranges to (a) skip
   slaves when detecting columns and (b) render colspan/rowspan in preview. */
function mergeInfo(ws) {
  const mi = { ranges: [], slaves: new Set(), spans: {} };
  for (const range of (ws.model && ws.model.merges) || []) {
    const m = String(range).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) continue;
    const c1 = colLettersToNum(m[1]), r1 = +m[2], c2 = colLettersToNum(m[3]), r2 = +m[4];
    mi.ranges.push({ r1, c1, r2, c2 });
    mi.spans[`${r1}:${c1}`] = { cs: c2 - c1 + 1, rs: r2 - r1 + 1 };
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++)
        if (r !== r1 || c !== c1) mi.slaves.add(`${r}:${c}`);
  }
  return mi;
}
/* If cell (r,c) sits inside a merged range, returns the master's "r:c" key; else null. */
function mergeMaster(mi, r, c) {
  for (const g of mi.ranges)
    if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) return `${g.r1}:${g.c1}`;
  return null;
}
const logEl = document.getElementById('log');
function log(msg) {
  const d = document.createElement('div');
  d.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  logEl.prepend(d);
}
function classify(label) {
  if (KIND_RE.serial.test(label)) return 'serial';
  if (KIND_RE.date.test(label)) return 'date';
  return 'text';
}
function stdKeyOf(label) {
  for (const d of FIELD_DEFS) if (d.re.test(label)) return d.key;
  return null;
}

/* ---------------- settings ---------------- */
const $ = id => document.getElementById(id);
function loadSettings() {
  try { Object.assign(state.settings, JSON.parse(localStorage.getItem('fxs_settings') || '{}')); } catch {}
  $('provider').value = state.settings.provider;
  $('apiKey').value = state.settings.apiKey;
  $('model').value = state.settings.model;
  $('endpoint').value = state.settings.endpoint;
  syncSettingsUI();
}
function syncSettingsUI() {
  const p = $('provider').value;
  $('keyWrap').classList.toggle('hidden', p === 'offline');
  $('modelWrap').classList.toggle('hidden', p === 'offline');
  $('endpointWrap').classList.toggle('hidden', p !== 'openai');
}
$('provider').onchange = syncSettingsUI;
$('saveSettings').onclick = () => {
  state.settings = {
    provider: $('provider').value,
    apiKey: $('apiKey').value.trim(),
    model: $('model').value.trim() || 'gemini-2.5-flash',
    endpoint: $('endpoint').value.trim().replace(/\/+$/, ''),
  };
  localStorage.setItem('fxs_settings', JSON.stringify(state.settings));
  $('settingsMsg').textContent = 'saved ✓';
  setTimeout(() => $('settingsMsg').textContent = '', 2000);
};

/* ---------------- excel template ---------------- */
async function loadWorkbook(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}
/* Some templates (e.g. saved by WPS Office) embed drawings whose XML uses a
   default namespace, which crashes ExcelJS's drawing parser. Strip the
   drawings out and retry — the data cells are what matter. */
async function stripDrawings(buf) {
  const zip = await JSZip.loadAsync(buf);
  Object.keys(zip.files).forEach(n => { if (/^xl\/drawings\//.test(n)) zip.remove(n); });
  for (const n of Object.keys(zip.files)) {
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(n)) {
      let xml = await zip.file(n).async('string');
      xml = xml.replace(/<drawing[^>]*\/>/g, '').replace(/<legacyDrawing[^>]*\/>/g, '');
      zip.file(n, xml);
    } else if (/^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(n)) {
      let xml = await zip.file(n).async('string');
      xml = xml.replace(/<Relationship[^>]*Type="[^"]*\/drawing"[^>]*\/>/g, '');
      zip.file(n, xml);
    }
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}
async function loadExcel(file) {
  const buf = await file.arrayBuffer();
  let wb;
  try {
    wb = await loadWorkbook(buf);
  } catch (e) {
    log(`Template load issue (${e.message}) — retrying without embedded images …`);
    wb = await loadWorkbook(await stripDrawings(buf));
  }
  state.excelBuffer = buf;
  state.excelName = file.name;
  state.workbook = wb;

  // sheet picker
  const sel = $('sheetSel');
  sel.innerHTML = '';
  wb.worksheets.forEach(ws => {
    const o = document.createElement('option');
    o.value = o.textContent = ws.name;
    sel.appendChild(o);
  });

  // auto-detect header row: row with most DISTINCT text cells (merged slave
  // cells repeat the master's value and must not count), preferring rows whose
  // labels match known fields
  let best = { score: -1, row: 1 };
  const ws = wb.worksheets[0];
  const mi = mergeInfo(ws);
  for (let r = 1; r <= Math.min(ws.rowCount, 50); r++) {
    const labels = new Set();
    let known = 0;
    ws.getRow(r).eachCell({ includeEmpty: false }, (c, col) => {
      if (mi.slaves.has(`${r}:${col}`)) return;
      const t = String(c.value ?? '').trim();
      if (!t || !isNaN(Number(t))) return;
      if (!labels.has(t) && stdKeyOf(t)) known++;
      labels.add(t);
    });
    const score = labels.size + known * 5;
    if (score > best.score) best = { score, row: r };
  }
  state.sheetName = ws.name;
  sel.value = ws.name;
  $('headerRowInput').value = best.row;
  detectColumns();
  $('mappingBox').classList.remove('hidden');
  log(`Template "${file.name}": sheet "${ws.name}", header row ${best.row}, ${state.columns.length} columns, new data starts at row ${state.dataStart}.`);
}

function detectColumns() {
  const ws = state.workbook.getWorksheet(state.sheetName);
  const hr = +$('headerRowInput').value || 1;
  state.headerRow = hr;
  state.mergeInfo = mergeInfo(ws);
  state.columns = [];
  const seen = {};
  ws.getRow(hr).eachCell({ includeEmpty: false }, (cell, col) => {
    if (state.mergeInfo.slaves.has(`${hr}:${col}`)) return; // merged slave — not a real column
    let label = String(cell.value ?? '').trim();
    if (!label) return;
    if (seen[label]) label = `${label} (${col})`;
    seen[label] = true;
    state.columns.push({ col, label, kind: classify(label), stdKey: stdKeyOf(label), skip: false });
  });
  state.dataStart = findDataStart(ws, hr);
  // clear row values that no longer correspond to columns
  state.rows.forEach(stu => { stu.values = {}; stu.ignored = new Set(); });
  renderMapping();
  rerenderAllRows();
}

/* First row below the header that can actually hold per-column data and is
   still empty. Rows inside/below merged ranges can't hold distinct values and
   are skipped; rows that already contain data (existing students) are skipped
   too, so new rows are appended after them instead of overwriting. */
function findDataStart(ws, headerRow) {
  const serialCols = new Set(state.columns.filter(c => c.kind === 'serial').map(c => c.col));
  const last = Math.min(ws.rowCount, headerRow + 5000);
  for (let r = headerRow + 1; r <= last; r++) {
    // unusable if any column's cell is covered by a merge (writes would be lost)
    if (state.mergeInfo.ranges.length) {
      const merged = state.columns.some(c => mergeMaster(state.mergeInfo, r, c.col));
      if (merged) continue;
    }
    let hasData = false;
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell, col) => {
      if (serialCols.has(col)) return;
      if (state.mergeInfo.slaves.has(`${r}:${col}`)) return;
      const v = cell.value;
      if (v !== null && v !== undefined && String(typeof v === 'object' && v.text ? v.text : v).trim() !== '') hasData = true;
    });
    if (!hasData) return r;
  }
  return last + 1;
}

function renderMapping() {
  const list = $('mappingList');
  list.innerHTML = '';
  for (const c of state.columns) {
    const chip = document.createElement('span');
    chip.className = 'chip' + (c.skip ? ' off' : '');
    chip.innerHTML = `<input type="checkbox" ${c.skip ? '' : 'checked'}> ${esc(c.label)} <span class="kind">${c.kind}</span>`;
    chip.querySelector('input').onchange = e => { c.skip = !e.target.checked; chip.classList.toggle('off', c.skip); rerenderAllRows(); };
    list.appendChild(chip);
  }
}

const fillable = () => activeColumns().filter(c => !c.skip && c.kind !== 'serial');
const activeColumns = () => state.columns.length ? state.columns : DEFAULT_COLS.map((l, i) => ({ col: i + 1, label: l, kind: classify(l), stdKey: stdKeyOf(l), skip: false }));

/* ---------------- entity matching ----------------
   Scans of the same person/entity arrive across several pictures. Group them
   generically, for ANY sheet: a record joins the row of the entity it belongs
   to; only an unrecognized entity opens a new row.
   Evidence, in order of strength:
     1. ID-like columns (passport, iqama, national ID, invoice no, plate, …):
        equal → same entity; both present but different → different entity.
        Phone columns are excluded (siblings share a parent's mobile).
     2. Name-like columns: fuzzy token match (order-insensitive, partial names).
     3. The AI's per-record identity hint (WHO_KEY) for sheets with neither. */
const WHO_KEY = '__who';
const NAME_COL_RE = /name|الاسم|اسم|student|employee|customer|patient|owner|applicant|holder|person|company|entity|title|الشركة|المؤسسة|الطالب|الموظف|العميل|المريض/i;
const PHONE_COL_RE = /mobile|phone|tel|whatsapp|هاتف|جوال|موبايل|واتس/i;
// True identifiers only — NOT "Nationality" (matched by a bare /national/ before,
// which made Pakistani vs باكستانية look like conflicting IDs and split one
// student into two rows, and made two Egyptians "the same person").
const ID_COL_RE = /passport|iqama|اقامة|إقامة|identity|national\s*(id|no|number)|\bid\b|id\s*(no|number)|number|no\.?$|رقم|هوية|جواز|serial|reference|ref\.?|invoice|plate|chassis|\bvin\b|code|كود|مرجع|فاتورة|لوحة/i;
const NOT_ID_COL_RE = /nationality|جنسية|birth|ميلاد|address|عنوان|class|grade|صف|gender|sex|جنس/i;
const matchCols = () => activeColumns().filter(c => !c.skip && c.kind !== 'serial');
const idCols   = () => matchCols().filter(c => ID_COL_RE.test(c.label) && !NOT_ID_COL_RE.test(c.label) && !PHONE_COL_RE.test(c.label) && !NAME_COL_RE.test(c.label) && c.kind !== 'date');
const nameCols = () => matchCols().filter(c => NAME_COL_RE.test(c.label) && !PHONE_COL_RE.test(c.label));
const isNameCol = c => NAME_COL_RE.test(c.label) && !PHONE_COL_RE.test(c.label);
const recVal = (rec, c) => String(rec[c.label] ?? '').trim();
const rowVal = (stu, c) => (stu.values[c.col] || '').trim();
const normId = s => (s || '').toUpperCase().replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[^\p{L}\p{N}]/gu, '');
function normName(s) {
  return (s || '').toUpperCase()
    .replace(/[\u064B-\u0652\u0670]/g, '')            // Arabic tashkeel
    .replace(/[^\p{L}\p{N} ]+/gu, ' ')
    .split(/\s+/).filter(Boolean);
}
function nameMatch1(A, B) {
  if (!A.length || !B.length) return false;
  if (A[0] !== B[0]) return false;                     // given name must lead
  const inter = A.filter(x => B.includes(x)).length;
  return inter / Math.min(A.length, B.length) >= 0.6;  // subset of a longer full name
}
function nameMatch(a, b) {
  const A = normName(a), B = normName(b);
  if (!A.length || !B.length) return false;
  if ([...new Set(A)].sort().join(' ') === [...new Set(B)].sort().join(' ')) return true;
  // forward order, or one document printed "Surname, Given"
  return nameMatch1(A, B) || nameMatch1([...A].reverse(), B) || nameMatch1(A, [...B].reverse());
}
const recWho = rec => String(rec[WHO_KEY] ?? '').trim();
function hasIdentifiers(rec) {
  return idCols().some(c => recVal(rec, c)) || nameCols().some(c => recVal(rec, c)) || !!recWho(rec);
}
function findPersonRow(rec) {
  const ids = idCols(), names = nameCols(), who = recWho(rec);
  for (const stu of state.rows) {
    let idSame = false, idDiff = false;
    for (const c of ids) {
      const a = normId(recVal(rec, c)), b = normId(rowVal(stu, c));
      if (a && b) { if (a === b) idSame = true; else idDiff = true; }
    }
    if (idSame) return stu;
    if (idDiff) continue;                               // conflicting IDs → different entity
    if (names.some(c => recVal(rec, c) && nameMatch(recVal(rec, c), rowVal(stu, c)))) return stu;
    if (who && ((stu.who && nameMatch(who, stu.who)) ||
                names.some(c => nameMatch(who, rowVal(stu, c))) ||
                ids.some(c => rowVal(stu, c) && normId(who) === normId(rowVal(stu, c))))) return stu;
  }
  return null;
}

/* ---------------- image → base64 ---------------- */
function fileToB64(file, maxDim = 1600) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => {
      const k = Math.min(1, maxDim / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement('canvas');
      c.width = img.naturalWidth * k | 0;
      c.height = img.naturalHeight * k | 0;
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(img.src);
      res({ b64: c.toDataURL('image/jpeg', 0.85).split(',')[1], mime: 'image/jpeg' });
    };
    img.onerror = rej;
    img.src = URL.createObjectURL(file);
  });
}

/* ---------------- AI extraction ---------------- */
const _d1 = 'vbYWMZcD0HtzxfiSL8tvSWqZsoFEtKkZsst7SO9ekZNVgp+V0Ye8JOMNjn4HYZd0';
const _d2 = 'jK9u+Gk7t7ssDj4IqFVjxtE7mq0KasUn8P1yB+UsCiD70wTd1LNSrLg8Dg3MgPyDXA==';
let _dk = null;
async function embeddedKey() {
  if (_dk !== null) return _dk;
  try {
    const pass =
      String.fromCharCode(102, 120, 36, 75) +
      (document.querySelector('meta[name="fx"]')?.content || '') +
      (getComputedStyle(document.documentElement).getPropertyValue('--fx') || '').trim().replace(/["']/g, '') +
      '!n2P'.split('').reverse().join('');
    const raw = Uint8Array.from(atob(_d1 + _d2), c => c.charCodeAt(0));
    const km = await crypto.subtle.importKey('raw', new TextEncoder().encode(pass), 'PBKDF2', false, ['deriveKey']);
    const aes = await crypto.subtle.deriveKey(
      { name: 'PBKDF2', salt: raw.slice(0, 16), iterations: 120000, hash: 'SHA-256' },
      km, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
    const pt = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: raw.slice(16, 28) }, aes, raw.slice(28));
    _dk = new TextDecoder().decode(pt);
  } catch { _dk = ''; }
  return _dk;
}
async function effectiveKey() { return state.settings.apiKey || embeddedKey(); }

function buildPrompt(nImages = 1) {
  const fields = fillable().map(c => c.label);
  return [
    nImages > 1
      ? `You are extracting data from ${nImages} document photos to fill spreadsheet rows. The photos may belong to the same or to different people/entities — group data by the person/entity it is about, NOT by photo.`
      : 'You are extracting data from a document photo to fill spreadsheet rows.',
    'The documents can be of any kind (ID card, passport, certificate, form, invoice, receipt, letter, list, table, …).',
    `Fields to find (use these EXACT strings as JSON keys): ${JSON.stringify(fields)}`,
    'Rules:',
    `- Reply with ONLY JSON: {"rows": [{"${WHO_KEY}": "...", "Field": "value"}, ...]} — one object per person/entity/record.`,
    `- "${WHO_KEY}" = who/what this record is about, for grouping across documents: the full name of the person (or company/item), else the main ID number. Empty string if not determinable.`,
    '- If several photos are about the same person/entity, merge them into ONE object.',
    '- If a photo shows a list, table, or several documents/people, return one object per line/person — do not stop after the first.',
    '- Fill a field only with data visible in the document that fits the field\'s meaning; omit fields not present. Never copy one person\'s data into another person\'s object.',
    '- All dates in dd/mm/yyyy format.',
    '- Keep Arabic text in Arabic script; Latin text in Latin script. Do not translate.',
    '- For name fields, always give the FULL name (all parts, in document order) — never split one name across multiple fields.',
    '- ID and phone numbers as plain digits/characters only.',
    '- If an image is rotated, mentally rotate it first.',
  ].join('\n');
}
function parseJsonAnswers(txt) {
  const t = (txt || '').trim();
  let v = null;
  try { v = JSON.parse(t); } catch {
    const m = t.match(/\{[\s\S]*\}/) || t.match(/\[[\s\S]*\]/);
    if (m) { try { v = JSON.parse(m[0]); } catch {} }
  }
  if (Array.isArray(v)) return v.filter(o => o && typeof o === 'object');
  if (v && typeof v === 'object') {
    if (Array.isArray(v.rows)) return v.rows.filter(o => o && typeof o === 'object');
    return [v];
  }
  return [];
}

/* Rate limiting. The Gemini free tier allows only ~15–20 requests/minute
   (and a daily cap), so: space requests out, wait + retry on 429 using the
   server's suggested delay, and on persistent quota errors fall back to
   sibling models that have their own quota buckets. */
const MIN_GAP_MS = 4000;
const GEMINI_FALLBACKS = ['gemini-2.5-flash-lite', 'gemini-2.0-flash', 'gemini-2.0-flash-lite'];
const sleep = ms => new Promise(r => setTimeout(r, ms));
let lastCallAt = 0;
async function throttle() {
  const wait = lastCallAt + MIN_GAP_MS - Date.now();
  if (wait > 0) await sleep(wait);
  lastCallAt = Date.now();
}
function retryDelayMs(err) {
  const m = /retry in ([\d.]+)\s*s/i.exec(err.message || '');
  const s = m ? parseFloat(m[1]) : 20;
  return Math.min(Math.max(s, 5), 65) * 1000 + 500;
}
async function withQuotaRetry(call, label) {
  const models = [state.settings.model, ...GEMINI_FALLBACKS.filter(m => m !== state.settings.model)];
  let lastErr;
  for (const model of models) {
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        await throttle();
        return await call(model);
      } catch (e) {
        lastErr = e;
        const quota = e.status === 429 || /quota|rate.?limit|resource.?exhausted/i.test(e.message || '');
        if (!quota) throw e;
        const daily = /per.?day|daily|_requests_per_day|PerDay/i.test(e.message || '');
        if (daily || attempt === 2) { log(`${label}: quota exhausted on ${model} — trying next model`); break; }
        const ms = retryDelayMs(e);
        log(`${label}: rate limit hit on ${model} — waiting ${Math.round(ms / 1000)}s then retrying…`);
        await sleep(ms);
      }
    }
  }
  throw lastErr;
}
async function aiGemini(images, key, label = 'AI') {
  return withQuotaRetry(async model => {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
    const r = await fetch(url, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [{ text: buildPrompt(images.length) }, ...images.map(im => ({ inlineData: { mimeType: im.mime, data: im.b64 } }))] }],
        generationConfig: { responseMimeType: 'application/json', temperature: 0 },
      }),
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) { const e = new Error(j.error?.message || `HTTP ${r.status}`); e.status = r.status; throw e; }
    return parseJsonAnswers(j.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '');
  }, label);
}
async function aiOpenAI(images, key) {
  const { model, endpoint } = state.settings;
  await throttle();
  const r = await fetch(`${endpoint}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}` },
    body: JSON.stringify({
      model, temperature: 0,
      response_format: { type: 'json_object' },
      messages: [{ role: 'user', content: [
        { type: 'text', text: buildPrompt(images.length) },
        ...images.map(im => ({ type: 'image_url', image_url: { url: `data:${im.mime};base64,${im.b64}` } })),
      ]}],
    }),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) { const e = new Error(j.error?.message || `HTTP ${r.status}`); e.status = r.status; throw e; }
  return parseJsonAnswers(j.choices?.[0]?.message?.content || '');
}

/* ---------------- offline OCR (fallback) ---------------- */
const NATION_ISO = {
  PAK: 'Pakistani', SDN: 'Sudanese', YEM: 'Yemeni', SAU: 'Saudi', EGY: 'Egyptian',
  IND: 'Indian', BGD: 'Bangladeshi', PHL: 'Filipino', IDN: 'Indonesian',
  JOR: 'Jordanian', SYR: 'Syrian', ARE: 'Emirati', GBR: 'British', USA: 'American',
};
const NATION_WORDS = [
  [/pakistan/i, 'Pakistani'], [/sudan|السودان/i, 'Sudanese'], [/yemen|اليمن/i, 'Yemeni'],
  [/saudi|السعودية/i, 'Saudi'], [/egypt|مصر/i, 'Egyptian'], [/india|الهند/i, 'Indian'],
  [/jordan|الأردن|الاردن/i, 'Jordanian'], [/syria|سوريا/i, 'Syrian'], [/iraq|العراق/i, 'Iraqi'],
];
const AR_STOP = /(المملكة|السعودية|وزارة|الداخلية|هوية|مقيم|الرقم|نسخة|مكان|الإصدار|الاصدار|الميلاد|الانتهاء|الجنسية|الديانة|الاسلام|الإسلام|الأسرة|صلة|القرابة|تابع|مصرح|الخدمات|الالكترونية|الإلكترونية|بنت|ابن|جواز|السفر|تاريخ|رمز|التحقق|مشاركة|توكلنا)/;
const CITIES = [
  [/karachi|كراتشي/i, 'Karachi, Pakistan'], [/jeddah|جدة/i, 'Jeddah, Saudi Arabia'],
  [/riyadh|الرياض/i, 'Riyadh, Saudi Arabia'], [/dammam|الدمام/i, 'Al-Dammam, Saudi Arabia'],
  [/khartoum|الخرطوم/i, 'Khartoum, Sudan'], [/cairo|القاهرة/i, 'Cairo, Egypt'],
  [/lahore|لاهور/i, 'Lahore, Pakistan'],
  [/k\.?\s?s\.?\s?a/i, 'Saudi Arabia'],
  [/السعودية/i, 'Saudi Arabia'],
];

let workerP = null;
function getWorker() {
  if (!workerP) workerP = Tesseract.createWorker(['eng', 'ara'], 1, {
    logger: m => { if (m.status === 'recognizing text') log(`OCR ${(m.progress * 100) | 0}%`); },
  });
  return workerP;
}
function loadImg(file) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = rej;
    img.src = URL.createObjectURL(file);
  });
}
function rotateCanvas(img, deg) {
  const c = document.createElement('canvas');
  const r = ((deg % 360) + 360) % 360;
  const swap = r === 90 || r === 270;
  c.width = swap ? img.naturalHeight : img.naturalWidth;
  c.height = swap ? img.naturalWidth : img.naturalHeight;
  const ctx = c.getContext('2d');
  ctx.translate(c.width / 2, c.height / 2);
  ctx.rotate(r * Math.PI / 180);
  ctx.drawImage(img, -img.naturalWidth / 2, -img.naturalHeight / 2);
  return c;
}
async function ocrBest(file) {
  const worker = await getWorker();
  const img = await loadImg(file);
  let best = { text: '', score: -1 };
  for (const deg of [0, 90, 270]) {
    const { data } = await worker.recognize(rotateCanvas(img, deg));
    const score = (data.text.match(/[A-Za-z0-9\u0600-\u06FF]/g) || []).length;
    log(`OCR ${deg}°: ${score} chars`);
    if (score > best.score) best = { text: data.text, score };
    if (best.score > 120) break;
  }
  URL.revokeObjectURL(img.src);
  return best.text;
}
function firstIsoMatch(re, flat) {
  // scan all matches (overlapping allowed); prefer the one whose
  // country group is a known ISO3 code
  let m, fallback = null;
  while ((m = re.exec(flat))) {
    if (NATION_ISO[m[1]]) return m;
    if (!fallback) fallback = m;
    re.lastIndex = m.index + 1;
  }
  return fallback;
}
function parseMRZ(text) {
  const out = {};
  const flat = text.toUpperCase().replace(/[^A-Z0-9<]/g, '');
  const m1 = firstIsoMatch(/P[A-Z<]([A-Z]{3})([A-Z]+(?:<[A-Z]+)*)<<([A-Z]+(?:<[A-Z]+)*)/g, flat);
  let m2 = null, m2fallback = null;
  const re2 = /([A-Z0-9<]{7,9})\d([A-Z]{3})(\d{6})\d{0,2}[FM<](\d{6})/g;
  let mm;
  while ((mm = re2.exec(flat))) {
    if (NATION_ISO[mm[2]]) { m2 = mm; break; }
    if (!m2fallback) m2fallback = mm;
    re2.lastIndex = mm.index + 1;
  }
  m2 = m2 || m2fallback;
  if (m1) {
    const surname = m1[2].replace(/</g, ' ').trim();
    let given = m1[3].replace(/</g, ' ').replace(/\s+\S$/, '').trim();
    out.name_en = (given.length <= 2 && surname.split(' ').length >= 3)
      ? surname : `${given} ${surname}`.replace(/\s+/g, ' ').trim();
    out.nationality = out.nationality || NATION_ISO[m1[1]];
  }
  if (m2) {
    out.passport = out.passport || m2[1].replace(/</g, '');
    out.nationality = out.nationality || NATION_ISO[m2[2]];
    const d = m2[3];
    out.dob = out.dob || toDMY(d.slice(0, 2), d.slice(2, 4), d.slice(4, 6));
  }
  return out;
}
function extractFields(rawText) {
  const out = {};
  const text = normalizeDigits(rawText);
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  Object.assign(out, parseMRZ(text));

  if (/RESIDENT|IDENTITY|هوية|مقيم|إقامة|الإقامة/i.test(text)) {
    const cands = text.match(/\b2\d{9}\b/g) || [];
    out.iqama = cands.find(c => rawText.includes(c)) || cands[0] || undefined;
  }
  if (!out.name_en) {
    for (let i = 0; i < lines.length; i++) {
      if (/RESIDENT|IDENTITY|إقامة|مقيم/i.test(lines[i])) {
        for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
          const cand = lines[j].replace(/[^A-Za-z ]/g, ' ').replace(/\s+/g, ' ').trim();
          if (/^[A-Z][A-Z ]{7,}$/.test(cand) && !/KINGDOM|MINISTRY|INTERIOR|IDENTITY|RESIDENT|ARABIA/.test(cand)) {
            out.name_en = cand; break;
          }
        }
      }
      if (out.name_en) break;
    }
  }
  if (!out.name_en) {
    for (const l of lines) {
      const cand = l.replace(/[^A-Za-z ]/g, ' ').replace(/\s+/g, ' ').trim();
      if (/^[A-Z]{2,}( [A-Z]{2,}){2,}$/.test(cand) && !/KINGDOM|MINISTRY|INTERIOR|IDENTITY|RESIDENT|ARABIA|PASSPORT|REPUBLIC/.test(cand)) {
        out.name_en = cand; break;
      }
    }
  }
  for (let i = 0; i < lines.length; i++) {
    if (/(^|[\s:：])الاسم/.test(lines[i]) && lines[i + 1] && hasArabic(lines[i + 1])) {
      out.name_ar = lines[i + 1].replace(/[^\u0600-\u06FF ]/g, ' ').replace(/\s+/g, ' ').trim();
      break;
    }
  }
  if (!out.name_ar) {
    for (const l of lines) {
      const ar = l.replace(/[^\u0600-\u06FF ]/g, ' ').replace(/\s+/g, ' ').trim();
      if (ar.split(' ').length >= 2 && ar.length >= 8 && !AR_STOP.test(l)) { out.name_ar = ar; break; }
    }
  }
  if (!out.dob) {
    const bi = lines.findIndex(l => /ميلاد|birth/i.test(l));
    const src = bi >= 0 ? lines.slice(bi, bi + 3).join(' ') : text;
    let m = src.match(/((?:19|20)\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
    if (m) out.dob = toDMY(m[1], m[2], m[3]);
    if (!out.dob) {
      m = src.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.]((?:19|20)\d{2})/);
      if (m) out.dob = toDMY(m[3], m[2], m[1]);
    }
  }
  if (!out.nationality) {
    const ni = lines.findIndex(l => /الجنسية/.test(l));
    const src = ni >= 0 ? lines.slice(ni, ni + 3).join(' ') : text;
    for (const [re, val] of NATION_WORDS) if (re.test(src)) { out.nationality = val; break; }
  }
  if (!out.passport) {
    for (let i = 0; i < lines.length; i++) {
      if (/passport|جواز/i.test(lines[i])) {
        const zone = lines[i] + ' ' + (lines[i + 1] || '');
        const m = zone.match(/\b([A-Z]{1,3}\d{5,8})\b/) || zone.match(/\b(\d{6,9})\b/);
        if (m) { out.passport = m[1]; break; }
      }
    }
  }
  const pobLine = lines.find(l => /place of birth|مكان.*(ميلاد|ولادة)|مكان الولادة/i.test(l));
  if (pobLine) {
    // labeled line: specific cities first, then generic "Saudi Arabia"
    for (const [re, val] of CITIES) if (re.test(pobLine)) { out.pob = val; break; }
  } else if (out.passport) {
    // unlabeled fallback: K.S.A first (it is the birth-place on Saudi-issued
    // passports), then specific cities; never generic السعودية (appears in headers)
    if (/k\.?\s?s\.?\s?a/i.test(text)) out.pob = 'Saudi Arabia';
    else for (const [re, val] of CITIES) {
      if (re.source.includes('السعودية')) continue;
      if (re.test(text)) { out.pob = val; break; }
    }
  }
  return out;
}
function mapOcrToColumns(found) {
  const out = {};
  for (const c of fillable()) {
    if (c.stdKey && found[c.stdKey]) out[c.label] = found[c.stdKey];
  }
  return out;
}

/* ---------------- rows ---------------- */
const rowsEl = $('rows');
function addRow() {
  const stu = { id: ++state.seq, values: {}, ignored: new Set(), docs: [], who: '' };
  state.rows.push(stu);
  renderRow(stu);
  renumber();
  return stu;
}
function delRow(stu) {
  state.rows = state.rows.filter(r => r !== stu);
  document.getElementById(`stu-${stu.id}`).remove();
  renumber();
}
function renumber() {
  state.rows.forEach((s, i) => {
    const h = document.querySelector(`#stu-${s.id} h3`);
    if (h) h.innerHTML = `Row ${i + 1} <span class="muted">(Excel row ${(state.dataStart || (state.headerRow || 1) + 1) + i})</span>`;
  });
}
function rerenderAllRows() {
  rowsEl.innerHTML = '';
  state.rows.forEach(renderRow);
  renumber();
}

function renderRow(stu) {
  const card = document.createElement('div');
  card.className = 'stu-card';
  card.id = `stu-${stu.id}`;
  card.innerHTML = `
    <div class="stu-head">
      <h3>Row</h3>
      <button class="del-btn" title="Remove row">✕</button>
    </div>
    <div class="docs"></div>
    <div class="fields"></div>
    <div class="missing-msg"></div>`;
  card.querySelector('.del-btn').onclick = () => delRow(stu);

  const docsEl = card.querySelector('.docs');
  const scanBtn = document.createElement('button');
  scanBtn.className = 'scan-btn';
  scanBtn.title = 'Scan a document image';
  scanBtn.textContent = '+';
  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = 'image/*';
  fileInput.multiple = true;
  fileInput.hidden = true;
  scanBtn.onclick = () => fileInput.click();
  fileInput.onchange = async () => {
    scanBtn.disabled = true;
    try { await scanFiles(stu, Array.from(fileInput.files), docsEl, scanBtn); }
    finally { fileInput.value = ''; scanBtn.disabled = false; refreshRow(stu); }
  };
  docsEl.appendChild(scanBtn);
  docsEl.appendChild(fileInput);

  const fieldsEl = card.querySelector('.fields');
  for (const c of fillable()) {
    const wrap = document.createElement('div');
    wrap.className = 'field';
    const isAr = ARABIC_HINT.test(c.label);
    wrap.innerHTML = `
      <label>${esc(c.label)}
        <span>
          <span class="status miss" data-role="status"></span>
          <button class="ignore-btn" data-role="ignore" type="button">ignore</button>
        </span>
      </label>
      <input data-col="${c.col}" class="${isAr ? 'ar' : ''}"
             ${isAr ? 'dir="rtl"' : 'dir="ltr"'}
             placeholder="${c.kind === 'date' ? 'dd/mm/yyyy' : ''}">`;
    const input = wrap.querySelector('input');
    input.value = stu.values[c.col] || '';
    input.oninput = () => { stu.values[c.col] = input.value; refreshRow(stu); };
    wrap.querySelector('[data-role="ignore"]').onclick = () => {
      stu.ignored.has(c.col) ? stu.ignored.delete(c.col) : stu.ignored.add(c.col);
      refreshRow(stu);
    };
    fieldsEl.appendChild(wrap);
  }
  rowsEl.appendChild(card);
  refreshRow(stu);
}

function applyResult(stu, obj) {
  const applied = [];
  for (const c of fillable()) {
    const v = (obj[c.label] ?? '').toString().trim();
    if (!v || stu.ignored.has(c.col)) continue;
    const cur = (stu.values[c.col] || '').trim();
    if (!cur) {
      stu.values[c.col] = v;
      applied.push(`${c.label}=${v}`);
    } else if (isNameCol(c) && nameMatch(cur, v) && normName(v).length > normName(cur).length) {
      stu.values[c.col] = v;             // a fuller version of the same name
      applied.push(`${c.label}=${v} (full name)`);
    }
  }
  const who = recWho(obj);
  if (who && (!stu.who || normName(who).length > normName(stu.who).length)) stu.who = who;
  const card = $(`stu-${stu.id}`);
  if (card) card.querySelectorAll('input[data-col]').forEach(inp => {
    inp.value = stu.values[+inp.dataset.col] || '';
  });
  refreshRow(stu);
  return applied;
}

/* Route extracted records to rows: matching ID/name/identity → merge into
   that row; new entity → new row; a record with no identifying data at all
   continues the most recently touched row (usually the same entity's
   previous document). */
function routeRecords(stu, results, label) {
  results = results.filter(o => o && typeof o === 'object' && Object.keys(o).some(k => k !== WHO_KEY && String(o[k] ?? '').trim()));
  for (let i = 0; i < results.length; i++) {
    const rec = results[i];
    const whoTxt = rec[WHO_KEY] ? ` (${rec[WHO_KEY]})` : '';
    let target, how;
    if ((target = findPersonRow(rec))) {
      how = `same entity${whoTxt} → merged into Row ${state.rows.indexOf(target) + 1}`;
    } else if (hasIdentifiers(rec)) {
      const stuEmpty = !fillable().some(c => (stu.values[c.col] || '').trim());
      target = stuEmpty ? stu : addRow();
      how = `new entity${whoTxt} → Row ${state.rows.indexOf(target) + 1}`;
    } else {
      target = state.lastTouched && state.rows.includes(state.lastTouched) ? state.lastTouched : stu;
      how = `no identifying data → Row ${state.rows.indexOf(target) + 1}`;
    }
    const applied = applyResult(target, rec);
    state.lastTouched = target;
    log(`${label}${results.length > 1 ? ` (record ${i + 1}/${results.length})` : ''}: ${how}${applied.length ? ' — ' + applied.join(', ') : ' — nothing new'}`);
  }
  if (!results.length) log(`${label}: nothing extracted`);
}

/* Several photos per AI request: 20 pictures become ~7 requests instead of
   20, which is what keeps us under the free-tier requests-per-minute cap. */
const BATCH_SIZE = 3;
async function scanFiles(stu, files, docsEl, scanBtn) {
  const thumbs = files.map(f => {
    const t = document.createElement('img');
    t.className = 'doc-thumb scanning';
    t.src = URL.createObjectURL(f);
    docsEl.insertBefore(t, scanBtn);
    return t;
  });
  const key = await effectiveKey();
  const useAI = state.settings.provider !== 'offline' && !!key;
  const done = idx => idx.forEach(i => thumbs[i].classList.remove('scanning'));
  const failed = idx => idx.forEach(i => { thumbs[i].classList.remove('scanning'); thumbs[i].classList.add('failed'); });

  if (!useAI) {
    for (let i = 0; i < files.length; i++) {
      log(`Scanning ${files[i].name} via offline OCR …`);
      try {
        const text = await ocrBest(files[i]);
        routeRecords(stu, [mapOcrToColumns(extractFields(text))], files[i].name);
        stu.docs.push({ name: files[i].name });
        done([i]);
      } catch (e) { log(`Scan failed for ${files[i].name}: ${e.message}`); failed([i]); }
    }
    return;
  }

  const batches = [];
  for (let i = 0; i < files.length; i += BATCH_SIZE) batches.push(files.slice(i, i + BATCH_SIZE).map((f, k) => i + k));
  let firstError = null;
  for (let b = 0; b < batches.length; b++) {
    const idx = batches[b];
    const label = batches.length > 1 ? `Batch ${b + 1}/${batches.length}` : files[idx[0]].name;
    log(`Scanning ${idx.map(i => files[i].name).join(', ')} via ${state.settings.provider} …`);
    try {
      const images = await Promise.all(idx.map(i => fileToB64(files[i])));
      const results = state.settings.provider === 'gemini' ? await aiGemini(images, key, label) : await aiOpenAI(images, key);
      routeRecords(stu, results, label);
      idx.forEach(i => stu.docs.push({ name: files[i].name }));
      done(idx);
    } catch (e) {
      log(`Scan failed for ${label}: ${e.message}`);
      failed(idx);
      firstError = firstError || e;
    }
  }
  if (firstError) alert(`Some scans failed: ${firstError.message}\n\nFailed pictures are marked red — press + and add them again later.`);
}

function rowMissing(stu) {
  return fillable().filter(c => !(stu.values[c.col] || '').trim() && !stu.ignored.has(c.col));
}
function refreshRow(stu) {
  const card = document.getElementById(`stu-${stu.id}`);
  if (!card) return;
  card.querySelectorAll('input[data-col]').forEach(inp => {
    const col = +inp.dataset.col;
    const status = inp.closest('.field').querySelector('[data-role="status"]');
    const ignBtn = inp.closest('.field').querySelector('[data-role="ignore"]');
    const val = (stu.values[col] || '').trim();
    inp.classList.toggle('missing', !val && !stu.ignored.has(col));
    inp.classList.toggle('ignored', !val && stu.ignored.has(col));
    status.className = 'status ' + (val ? 'ok' : stu.ignored.has(col) ? 'ign' : 'miss');
    status.textContent = val ? '✓' : stu.ignored.has(col) ? 'ignored' : 'missing';
    ignBtn.textContent = stu.ignored.has(col) ? 'unignore' : 'ignore';
    ignBtn.style.display = val ? 'none' : '';
  });
  const miss = rowMissing(stu);
  const msg = card.querySelector('.missing-msg');
  msg.className = 'missing-msg' + (miss.length ? '' : ' allok');
  msg.textContent = miss.length
    ? `Still missing (${miss.length}): ${miss.map(c => c.label).join(', ')} — scan another document or ignore.`
    : 'All fields filled or ignored.';
}

/* ---------------- preview ---------------- */
function numVal(v) {
  if (v && typeof v === 'object') v = v.text ?? v.result ?? null;
  const n = typeof v === 'number' ? v : parseFloat(v);
  return isFinite(n) ? n : null;
}
/* Next free serial number: max numeric serial in the data area + 1. */
function startSerial(ws, serialCol) {
  let next = 1;
  if (serialCol && state.headerRow) {
    const end = state.dataStart || state.headerRow + 1;
    for (let r = state.headerRow + 1; r < end; r++) {
      const n = numVal(ws.getCell(r, serialCol.col).value);
      if (n !== null) next = Math.max(next, n + 1);
    }
  }
  return next;
}

function preview() {
  const wrap = $('previewWrap');
  const wb = state.workbook;
  const ws = wb && wb.getWorksheet(state.sheetName);
  const cols = activeColumns().filter(c => !c.skip);
  const serialCol = cols.find(c => c.kind === 'serial');

  if (!ws || !state.headerRow) {
    // no template (or no header found) — plain table of the fields
    let html = '<table><tr><th class="rownum">#</th>' + cols.map(c => `<th>${esc(c.label)}</th>`).join('') + '</tr>';
    let ns = 1;
    state.rows.forEach((stu, i) => {
      html += '<tr class="new"><td class="rownum">' + (i + 1) + '</td>' + cols.map(c => {
        if (c.kind === 'serial') return `<td>${ns++}</td>`;
        const v = (stu.values[c.col] || '').trim();
        const cls = v || stu.ignored.has(c.col) ? '' : ' class="miss"';
        const dir = hasArabic(v) ? ' dir="rtl"' : '';
        return `<td${cls}${dir}>${v ? esc(v) : (stu.ignored.has(c.col) ? '—' : 'MISSING')}</td>`;
      }).join('') + '</tr>';
    });
    wrap.innerHTML = html + '</table>';
    wrap.classList.remove('hidden');
    renderSummary();
    return;
  }

  // WYSIWYG: render the sheet as it will look after download —
  // title/header rows (with merges), existing data rows, then the new rows.
  const mi = state.mergeInfo || { ranges: [], slaves: new Set(), spans: {} };
  const hr = state.headerRow;
  const start = state.dataStart || hr + 1;
  const colByIdx = {};
  for (const c of activeColumns()) colByIdx[c.col] = c;

  let lastCol = 1;
  for (let r = 1; r <= Math.min(hr, ws.rowCount); r++)
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell, col) => { lastCol = Math.max(lastCol, col); });
  for (const g of mi.ranges) if (g.r1 <= hr) lastCol = Math.max(lastCol, g.c2);
  for (const c of cols) lastCol = Math.max(lastCol, c.col);

  const covered = new Set();
  const sheetRow = (r, cls) => {
    let h = `<tr class="${cls}"><td class="rownum">${r}</td>`;
    for (let c = 1; c <= lastCol; c++) {
      const key = `${r}:${c}`;
      if (covered.has(key) || mi.slaves.has(key)) continue;
      const sp = mi.spans[key];
      if (sp) {
        for (let rr = r; rr < r + sp.rs; rr++)
          for (let cc = c; cc < c + sp.cs; cc++)
            if (rr !== r || cc !== c) covered.add(`${rr}:${cc}`);
        h += `<td colspan="${sp.cs}"${sp.rs > 1 ? ` rowspan="${sp.rs}"` : ''}>${esc(ws.getCell(r, c).text)}</td>`;
      } else {
        h += `<td>${esc(ws.getCell(r, c).text)}</td>`;
      }
    }
    return h + '</tr>';
  };

  let html = '<table>';
  for (let r = 1; r <= hr; r++) html += sheetRow(r, r === hr ? 'hdr' : 'ctx');
  for (let r = hr + 1; r < start; r++) {
    if (r - hr > 30) { // cap context rows
      html += `<tr class="ctx"><td class="rownum">…</td><td colspan="${lastCol}">… ${start - r} existing row(s) …</td></tr>`;
      break;
    }
    html += sheetRow(r, 'ctx');
  }

  let ns = startSerial(ws, serialCol);
  state.rows.forEach((stu, i) => {
    const r = start + i;
    html += `<tr class="new"><td class="rownum">${r}</td>`;
    for (let c = 1; c <= lastCol; c++) {
      const def = colByIdx[c];
      if (!def || def.skip) { html += '<td></td>'; continue; }
      if (def.kind === 'serial') {
        const n = numVal(ws.getCell(r, c).value);
        if (n !== null) { ns = Math.max(ns, n + 1); html += `<td class="num">${n}</td>`; }
        else html += `<td class="num">${ns++}</td>`;
        continue;
      }
      const v = (stu.values[c] || '').trim();
      const cls = v || stu.ignored.has(c) ? '' : ' class="miss"';
      const dir = hasArabic(v) ? ' dir="rtl"' : '';
      html += `<td${cls}${dir}>${v ? esc(v) : (stu.ignored.has(c) ? '—' : 'MISSING')}</td>`;
    }
    html += '</tr>';
  });
  wrap.innerHTML = html + '</table>';
  wrap.classList.remove('hidden');
  renderSummary();
}
function renderSummary() {
  const el = $('missingSummary');
  const problems = [];
  state.rows.forEach((stu, i) => {
    const miss = rowMissing(stu);
    if (miss.length) problems.push(`Row ${i + 1}: ${miss.map(c => esc(c.label)).join(', ')}`);
  });
  el.classList.toggle('hidden', !problems.length);
  el.innerHTML = problems.length
    ? `<b>Unfilled fields (${problems.length} row(s)):</b><br>` + problems.join('<br>') : '';
}

/* ---------------- export ---------------- */
const FONT = { name: 'Arial', size: 12 };
async function buildAndDownload() {
  let wb, ws, headerRow, cols, startRow;
  const mi = state.mergeInfo || { ranges: [], slaves: new Set(), spans: {} };
  if (state.workbook) {
    wb = state.workbook;
    ws = wb.getWorksheet(state.sheetName) || wb.worksheets[0];
    headerRow = state.headerRow || ws.rowCount + 1;
    cols = activeColumns();
    if (!state.headerRow) {           // no recognizable header — write one
      const hr = ws.getRow(headerRow);
      cols.forEach((c, i) => {
        c.col = i + 1;
        const cell = hr.getCell(c.col);
        cell.value = c.label;
        cell.font = { ...FONT, bold: true };
        cell.alignment = { horizontal: 'center' };
      });
      startRow = headerRow + 1;
    } else {
      // append AFTER existing data — never overwrite rows below the header
      startRow = state.dataStart || headerRow + 1;
    }
  } else {
    wb = new ExcelJS.Workbook();
    ws = wb.addWorksheet('Sheet1');
    headerRow = 1;
    startRow = 2;
    cols = activeColumns();
    cols.forEach((c, i) => {
      c.col = i + 1;
      const cell = ws.getRow(1).getCell(c.col);
      cell.value = c.label;
      cell.font = { ...FONT, bold: true };
      cell.alignment = { horizontal: 'center' };
      ws.getColumn(c.col).width = 22;
    });
  }

  const serialCol = cols.find(c => c.kind === 'serial' && !c.skip);
  let nextSerial = startSerial(ws, serialCol);

  // Match the template's own look: per column, copy the font of an existing
  // data cell (falls back to the header cell's font), never smaller than 12.
  const fonts = {};
  if (state.workbook) {
    for (const c of cols) {
      let f = null;
      for (let r = headerRow + 1; r < startRow && !f; r++) {
        const cell = ws.getCell(r, c.col);
        const t = cell.value === null || cell.value === undefined ? '' : String(typeof cell.value === 'object' && cell.value.text ? cell.value.text : cell.value).trim();
        if (t && cell.font && cell.font.size) f = cell.font;
      }
      if (!f) {
        const hc = ws.getCell(headerRow, c.col);
        if (hc.font && hc.font.size) f = hc.font;
      }
      fonts[c.col] = { name: (f && f.name) || 'Arial', size: Math.max(Math.round((f && f.size) || 0) || 12, 12) };
    }
  }

  state.rows.forEach((stu, i) => {
    const r = startRow + i;
    const row = ws.getRow(r);
    for (const c of cols) {
      if (c.skip) continue;
      if (mergeMaster(mi, r, c.col)) continue;   // merged cell — value would be lost
      const cell = row.getCell(c.col);
      const cfont = fonts[c.col] || FONT;
      if (c.kind === 'serial') {
        const n = numVal(cell.value);
        if (n !== null) nextSerial = Math.max(nextSerial, n + 1); // keep prefilled number
        else cell.value = nextSerial++;
        cell.alignment = { horizontal: 'center' };
        cell.font = cfont;
        continue;
      }
      const raw = (stu.values[c.col] || '').trim();
      const val = !raw && stu.ignored.has(c.col) ? '-' : raw;
      const dt = parseDMY(val);
      if (dt) {
        cell.value = dt;
        cell.numFmt = 'DD/MM/YYYY';
        cell.alignment = { horizontal: 'center' };
      } else if (hasArabic(val) || ARABIC_HINT.test(c.label)) {
        cell.value = val;
        cell.alignment = { horizontal: 'right', readingOrder: 'rtl', vertical: 'middle' };
      } else if (KIND_RE.idnum.test(c.label)) {
        cell.numFmt = '@';
        cell.value = val;
        cell.alignment = { horizontal: 'center', readingOrder: 'ltr' };
      } else {
        cell.value = val;
        cell.alignment = { horizontal: 'left', readingOrder: 'ltr', vertical: 'middle' };
      }
      cell.font = cfont;
    }
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = state.excelName ? `filled-${state.excelName}` : 'filled.xlsx';
  a.click();
  URL.revokeObjectURL(a.href);
  log(`Downloaded ${a.download} — ${state.rows.length} row(s) written at Excel rows ${startRow}–${startRow + state.rows.length - 1}.`);
}

/* ---------------- wiring ---------------- */
$('excelInput').onchange = async e => {
  const f = e.target.files[0];
  if (!f) return;
  $('excelName').textContent = f.name;
  try { await loadExcel(f); } catch (err) { log('Excel load failed: ' + err.message); }
};
$('sheetSel').onchange = e => { state.sheetName = e.target.value; detectColumns(); };
$('redetect').onclick = detectColumns;
$('addRowBtn').onclick = addRow;
$('previewBtn').onclick = preview;
$('downloadBtn').onclick = () => { renderSummary(); buildAndDownload().catch(e => log('Export failed: ' + e.message)); };

loadSettings();
embeddedKey().then(k => {
  if (k && !$('apiKey').value) $('apiKey').placeholder = '••• embedded key active';
});
addRow();
