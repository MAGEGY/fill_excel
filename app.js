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

const fillable = () => state.columns.filter(c => !c.skip && c.kind !== 'serial');
const activeColumns = () => state.columns.length ? state.columns : DEFAULT_COLS.map((l, i) => ({ col: i + 1, label: l, kind: classify(l), stdKey: stdKeyOf(l), skip: false }));

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

function buildPrompt() {
  const fields = fillable().map(c => c.label);
  return [
    'You are extracting data from a document photo to fill spreadsheet rows.',
    `Fields to find (use these EXACT strings as JSON keys): ${JSON.stringify(fields)}`,
    'Rules:',
    '- Reply with ONLY JSON: {"rows": [{"Field": "value"}, ...]} — one object per person/record.',
    '- If the image shows a single document/person, "rows" has ONE object.',
    '- If the image shows a list, table, or several documents/people, return one object per line/person — do not stop after the first.',
    '- Omit fields not visible in the document.',
    '- All dates in dd/mm/yyyy format.',
    '- Keep Arabic text in Arabic script; Latin names in Latin script. Do not translate.',
    '- For name fields, always give the person\'s FULL name (all parts, in document order) — never split one name across multiple fields.',
    '- ID and phone numbers as plain digits only.',
    '- If the image is rotated, mentally rotate it first.',
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
async function aiGemini(b64, mime, key) {
  const { model } = state.settings;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
  const r = await fetch(url, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: buildPrompt() }, { inlineData: { mimeType: mime, data: b64 } }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0 },
    }),
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error?.message || r.status);
  return parseJsonAnswers(j.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '');
}
async function aiOpenAI(b64, mime, key) {
  const { model, endpoint } = state.settings;
  const r = await fetch(`${endpoint}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}` },
    body: JSON.stringify({
      model, temperature: 0,
      response_format: { type: 'json_object' },
      messages: [{ role: 'user', content: [
        { type: 'text', text: buildPrompt() },
        { type: 'image_url', image_url: { url: `data:${mime};base64,${b64}` } },
      ]}],
    }),
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error?.message || r.status);
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
  const stu = { id: ++state.seq, values: {}, ignored: new Set(), docs: [] };
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
    for (const f of fileInput.files) await scanDoc(stu, f, docsEl, scanBtn);
    fileInput.value = '';
    scanBtn.disabled = false;
    refreshRow(stu);
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
    if (v && !stu.values[c.col] && !stu.ignored.has(c.col)) {
      stu.values[c.col] = v;
      applied.push(`${c.label}=${v}`);
    }
  }
  const card = $(`stu-${stu.id}`);
  if (card) card.querySelectorAll('input[data-col]').forEach(inp => {
    inp.value = stu.values[+inp.dataset.col] || '';
  });
  refreshRow(stu);
  return applied;
}

async function scanDoc(stu, file, docsEl, scanBtn) {
  const thumb = document.createElement('img');
  thumb.className = 'doc-thumb scanning';
  thumb.src = URL.createObjectURL(file);
  docsEl.insertBefore(thumb, scanBtn);
  const key = await effectiveKey();
  const useAI = state.settings.provider !== 'offline' && !!key;
  log(`Scanning ${file.name} via ${useAI ? state.settings.provider : 'offline OCR'} …`);
  try {
    let results;
    if (useAI) {
      const { b64, mime } = await fileToB64(file);
      results = state.settings.provider === 'gemini' ? await aiGemini(b64, mime, key) : await aiOpenAI(b64, mime, key);
    } else {
      const text = await ocrBest(file);
      results = [mapOcrToColumns(extractFields(text))];
    }
    // First record goes to this row; extra people/records create new rows.
    results = results.filter(o => o && typeof o === 'object' && Object.keys(o).length);
    for (let i = 0; i < results.length; i++) {
      const target = i === 0 ? stu : addRow();
      const applied = applyResult(target, results[i]);
      log(`${file.name}${i ? ` (record ${i + 1} → Row ${state.rows.indexOf(target) + 1})` : ''}: ${applied.length ? 'filled → ' + applied.join(', ') : 'nothing new extracted'}`);
    }
    if (!results.length) log(`${file.name}: nothing extracted`);
    stu.docs.push({ name: file.name });
  } catch (e) {
    log(`Scan failed for ${file.name}: ${e.message}`);
    alert(`Scan failed: ${e.message}`);
  } finally {
    thumb.classList.remove('scanning');
  }
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
const FONT = { name: 'Arial', size: 11 };
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

  state.rows.forEach((stu, i) => {
    const r = startRow + i;
    const row = ws.getRow(r);
    for (const c of cols) {
      if (c.skip) continue;
      if (mergeMaster(mi, r, c.col)) continue;   // merged cell — value would be lost
      const cell = row.getCell(c.col);
      if (c.kind === 'serial') {
        const n = numVal(cell.value);
        if (n !== null) nextSerial = Math.max(nextSerial, n + 1); // keep prefilled number
        else cell.value = nextSerial++;
        cell.alignment = { horizontal: 'center' };
        cell.font = FONT;
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
      cell.font = FONT;
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
