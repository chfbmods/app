/**
 * CHApp – Google Apps Script backend
 * Sheets used (tab names are configurable below):
 *   "Main Responses"  – source data (row 1 = headers; needs Area, Timestamp, MLBB Server, MLBB UID)
 *   "My Memberlist"   – approved members (created/filled automatically)
 *   "Settings"        – app settings with the columns Name | Value | Notes (e.g. Logo, Favicon). Admins edit the Value column in Settings > Admin.
 *   "Users"           – logins: Username | Password | Role | Area | Nickname
 *                       Area = the Area this user manages (matched against the "Area" column in the data sheets).
 *                       Full Name | Position | Org Level | Photo Link feed the About Us page (Org Level like 1, 2, 3, 4-A, 4-A-1).
 *                       Nickname = the name shown in the upper right (falls back to Username if blank).
 *                       Role = "admin" gives access to ALL Areas. Leave Role blank for regular users.
 */
const CFG = {
  SPREADSHEET_ID: '16dS2wvXIrsQqOdX6eEfuuW0vUVEVa0OryJvHG5pnrwc', // the ID in the Google Sheet URL
  MEMBER_SPREADSHEET_ID: '',                        // leave blank to use the same file
  MAIN_SHEET: 'Main Responses',
  USERS_SHEET: 'Users',
  SETTINGS_SPREADSHEET_ID: '',                      // only if the Settings tab lives in a different spreadsheet (blank = same file as SPREADSHEET_ID)
  SETTINGS_SHEET: 'Settings',                        // app settings (logo / favicon links ...), editable by admins on the Admin page
  MEMBER_SHEET: 'My Memberlist',
  AREA_COL: 'Area',
  SERVER_COL: 'MLBB Server',                       // MLBB UID + MLBB Server identify a player; if that pair is Approved in My Memberlist, Approve is hidden
  UID_COL: 'MLBB UID',
  TIMESTAMP_COL: 'Timestamp',                       // both tables are sorted by this, newest first
  TIMESTAMP_FORMAT: 'M/d/yyyy H:mm:ss',             // date + time format used for Timestamp in My Memberlist
  PAGE_SIZE: 15,                                    // rows sent per request (the "Show more" button loads the next page)
  STATUS_COL: 'Status',
  SOURCE_COL: 'Source Row',                         // legacy column: no longer written; hidden in the UI if it still exists
  NOTE_COL: 'Note',                                 // editable in Member Details (My Memberlist)
  UPDATED_BY_COL: 'Last Updated By',                // filled with the username on approve / status change
  SEARCH_COLS: ['Email Address', 'MLBB Server', 'MLBB UID'], // searchable even when not shown in the table
  // Table columns after Status, in order. (Text in {braces} would make a virtual column, e.g. '{MLBB UID} ({MLBB Server})'.)
  // Current layout: Status | Complete Name | MLBB UID | MLBB Server | Area
  // (A name that doesn't exist in the sheet is replaced by the next unused column, so the positions stay the same.)
  LIST_COLS: ['Complete Name', 'MLBB UID', 'MLBB Server', 'Area'],
  DATA_CACHE_SECONDS: 300,                          // server-side cache of both tables (cleared automatically on approve / status save)
  USER_CACHE_SECONDS: 120,                          // how long user roles/areas are cached
  PHOTO_FOLDER_ID: '',                              // Drive folder (link or ID) for uploaded profile photos. Blank = a folder "CHApp Photos" is created automatically
  PHOTO_FOLDER_NAME: 'CHApp Photos',
  DEFAULT_EMAIL_DOMAIN: '@gmail.com',               // users may sign in with their email WITHOUT this part (e.g. "juan" for juan@gmail.com)
  SESSION_HOURS: 0,                                 // 0 = sessions never expire (users stay signed in until they sign out). Set e.g. 24 for a daily limit.
  STATUSES: ['Approved', 'Rejected', 'Duplicate', 'Transferred'],
  ADMIN_STATUSES: ['Blacklisted']                   // can only be set (or changed again) by admins; a blacklisted player cannot be approved
};

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index'), br = brand_();
  t.brand = JSON.stringify({ logo: br.logo }).replace(/</g, '\\u003c');   // logo + favicon links come from the Settings tab
  const out = t.evaluate()
    .setTitle('CHApp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (br.favicon) { try { out.setFaviconUrl(br.favicon); } catch (e) { /* unsupported image: the app still loads */ } }
  return out;
}

/* ---------- helpers ---------- */
const norm_ = s => String(s == null ? '' : s).trim().toLowerCase();
const books_ = {};                                   // open each spreadsheet only once per request
function book_(id) {
  id = id || CFG.SPREADSHEET_ID;
  return books_[id] || (books_[id] = SpreadsheetApp.openById(id));
}
/** Finds a tab by name (ignores capitals and extra spaces) and explains clearly if it doesn't exist. */
function tab_(book, name) {
  const sh = book.getSheetByName(name) || book.getSheets().find(s => norm_(s.getName()) === norm_(name));
  if (!sh) throw new Error('Tab "' + name + '" was not found. Tabs in this spreadsheet: ' +
    book.getSheets().map(s => '"' + s.getName() + '"').join(', ') + '. Rename the tab or change the tab name in CFG at the top of Code.gs.');
  return sh;
}
const mainSheet_ = () => tab_(book_(CFG.SPREADSHEET_ID), CFG.MAIN_SHEET);
const usersSheet_ = () => tab_(book_(), CFG.USERS_SHEET);
function memberSheet_() {
  const b = book_(CFG.MEMBER_SPREADSHEET_ID);
  return b.getSheetByName(CFG.MEMBER_SHEET) || b.getSheets().find(s => norm_(s.getName()) === norm_(CFG.MEMBER_SHEET)) || b.insertSheet(CFG.MEMBER_SHEET);
}
function hash_(p) {
  return 'sha256:' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
/** Run once from the editor to get a hash you can paste into the Users sheet instead of a plain password. */
function hashPassword() { Logger.log(hash_('type-password-here')); }

function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
const sign_ = t => Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(t, secret_()));
function makeToken_(user) {
  const exp = CFG.SESSION_HOURS > 0 ? Date.now() + CFG.SESSION_HOURS * 3600 * 1000 : 0;   // 0 = never expires
  const body = Utilities.base64EncodeWebSafe(user, Utilities.Charset.UTF_8) + '.' + exp;
  return body + '.' + sign_(body);
}
function tokenUser_(token) {
  const p = String(token || '').split('.');
  const exp = Number(p[1]);
  if (p.length !== 3 || sign_(p[0] + '.' + p[1]) !== p[2] || (exp > 0 && Date.now() > exp)) throw new Error('SESSION');
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString();
}
const col_ = (headers, name) => headers.findIndex(h => norm_(h) === norm_(name));
const headers_ = sh => sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim()) : [];
/** Admins match every row; others only when their (non-empty) Users.Area equals the row's Area. */
const match_ = (rowArea, a) => a.admin || (a.area !== '' && norm_(rowArea) === norm_(a.area));

/* ---------- caching (CacheService, chunked because each entry is limited to 100 KB) ---------- */
const CHUNK = 30000;
function cget_(name) {
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get(name + ':n'));
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(name + ':' + i);
    const got = c.getAll(keys); let s = '';
    for (let i = 0; i < n; i++) { const p = got[name + ':' + i]; if (p == null) return null; s += p; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function cput_(name, obj) {
  try {
    const s = JSON.stringify(obj), n = Math.max(1, Math.ceil(s.length / CHUNK)), all = {};
    for (let i = 0; i < n; i++) all[name + ':' + i] = s.substr(i * CHUNK, CHUNK);
    all[name + ':n'] = String(n);
    CacheService.getScriptCache().putAll(all, CFG.DATA_CACHE_SECONDS);
  } catch (e) { /* too big or cache unavailable: just skip caching */ }
}
const cdel_ = name => { try { CacheService.getScriptCache().remove(name + ':n'); } catch (e) {} };
function base_(name, builder, force) {
  if (!force) { const c = cget_(name); if (c) return c; }
  const b = builder(); cput_(name, b); return b;
}

/** Users are cached briefly (no passwords stored) so the Users sheet isn't re-read on every click. */
function usersMap_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('users');
  if (hit) return JSON.parse(hit);
  const sh = usersSheet_(), h = headers_(sh), map = {};
  const ui = col_(h, 'Username'), ri = col_(h, 'Role'), ai = col_(h, CFG.AREA_COL), ni = col_(h, 'Nickname');
  const fu = col_(h, 'Full Name'), po = col_(h, 'Position'), oi = col_(h, 'Org Level'), fi = col_(h, 'Photo Link');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getDisplayValues().forEach(v => {
    const user = String(v[ui]).trim(); if (!user) return;
    map[norm_(user)] = { user, full: fu >= 0 ? String(v[fu]).trim() : '', pos: po >= 0 ? String(v[po]).trim() : '', org: oi >= 0 ? String(v[oi]).trim() : '', photo: fi >= 0 ? imgUrl_(v[fi], 240) : '', nick: ni >= 0 ? String(v[ni]).trim() : '', admin: ri >= 0 && norm_(v[ri]) === 'admin', area: ai >= 0 ? String(v[ai]).trim() : user };
  });
  cache.put('users', JSON.stringify(map), CFG.USER_CACHE_SECONDS);
  return map;
}
function auth_(token) {
  const u = usersMap_()[norm_(tokenUser_(token))];
  if (!u) throw new Error('SESSION');
  return u;
}

/* ---------- table builders (only run on a cache miss) ---------- */
/** Player identity = MLBB UID + MLBB Server (empty when there is no UID, so blank rows never match each other). */
const idKey_ = (server, uid) => norm_(uid) ? norm_(server) + '|' + norm_(uid) : '';
const searchIdx_ = headers => CFG.SEARCH_COLS.map(n => col_(headers, n)).filter(i => i >= 0);
const searchText_ = (v, idx) => idx.map(i => v[i]).join(' ').toLowerCase();
/**
 * Columns shown in a table, following CFG.LIST_COLS (Status goes first on My Memberlist).
 * Each column is {h: header label, refs: [sheet column indexes it needs], get: row values -> cell text}.
 * "{MLBB UID} ({MLBB Server})" is a virtual column: values are joined, and empty "()" is dropped.
 */
function listCols_(headers, vis, statusFirst) {
  const si = statusFirst ? col_(headers, CFG.STATUS_COL) : -1;
  const pool = vis.filter(i => i !== si && norm_(headers[i]) !== norm_(CFG.STATUS_COL));
  const isTpl = n => /\{[^}]+\}/.test(n);
  const refsOf = n => (n.match(/\{[^}]+\}/g) || []).map(t => col_(headers, t.slice(1, -1))).filter(i => i >= 0);
  const named = CFG.LIST_COLS.map(n => isTpl(n) ? -2 : col_(headers, n));
  const reserved = [].concat.apply([], CFG.LIST_COLS.filter(isTpl).map(refsOf));  // used by virtual columns: never auto-picked as fillers
  const used = {}, out = [];
  CFG.LIST_COLS.forEach((n, k) => {
    if (isTpl(n)) {
      out.push({
        h: n.replace(/[{}]/g, ''), refs: refsOf(n),
        get: v => n.replace(/\{([^}]+)\}/g, (m, c) => { const i = col_(headers, c); return i >= 0 ? v[i] : ''; })
                   .replace(/\s*\(\s*\)/g, '').trim()
      });
      return;
    }
    let i = named[k];
    if (i < 0 || used[i] || pool.indexOf(i) < 0) i = pool.find(j => !used[j] && named.indexOf(j) < 0 && reserved.indexOf(j) < 0);
    if (i != null && i >= 0) { used[i] = 1; out.push({ h: headers[i], refs: [i], get: v => v[i] }); }
  });
  return si >= 0 ? [{ h: headers[si], refs: [si], get: v => v[si] }].concat(out) : out;
}
const refsMax_ = (cols, extra) => Math.max.apply(null, [].concat.apply([], cols.map(c => c.refs)).concat(extra)) + 1;
/** Row number -> timestamp (ms), read from the raw cell values so dates sort correctly. */
function tsKeys_(sh, ti) {
  const keys = {}, last = sh.getLastRow();
  if (ti < 0 || last < 2) return keys;
  sh.getRange(2, ti + 1, last - 1, 1).getValues().forEach((x, i) => {
    const t = x[0] instanceof Date ? x[0].getTime() : Date.parse(x[0]);
    keys[i + 2] = isNaN(t) ? 0 : t;
  });
  return keys;
}
const byNewest_ = keys => (p, q) => ((keys[q.r] || 0) - (keys[p.r] || 0)) || (q.r - p.r);

/** Main Responses columns and positions (read fresh every time: Main Responses is never cached). */
function mainMeta_() {
  const sh = mainSheet_(), headers = headers_(sh);
  const ai = col_(headers, CFG.AREA_COL), sv = col_(headers, CFG.SERVER_COL), ui = col_(headers, CFG.UID_COL);
  if (ai < 0) throw new Error('"' + CFG.AREA_COL + '" column not found in ' + CFG.MAIN_SHEET);
  if (sv < 0 || ui < 0) throw new Error('"' + CFG.SERVER_COL + '" and "' + CFG.UID_COL + '" columns are required in ' + CFG.MAIN_SHEET);
  const cols = listCols_(headers, headers.map((h, i) => i), false), sIdx = searchIdx_(headers);
  return { sh, headers, ai, sv, ui, cols, sIdx, need: refsMax_(cols, sIdx.concat([ai, sv, ui])), ti: col_(headers, CFG.TIMESTAMP_COL) };
}
const mainRow_ = (m, r, v) => ({ r, v: m.cols.map(c => c.get(v)), k: searchText_(v, m.sIdx), id: idKey_(v[m.sv], v[m.ui]), a: v[m.ai] });

/** Every row, newest first (used for searching). */
function buildMain_() {
  const m = mainMeta_(), last = m.sh.getLastRow(), rows = [];
  const vals = last > 1 ? m.sh.getRange(2, 1, last - 1, m.need).getDisplayValues() : [];
  vals.forEach((v, i) => { if (v.some(x => x !== '')) rows.push(mainRow_(m, i + 2, v)); });
  rows.sort(byNewest_(tsKeys_(m.sh, m.ti)));
  return { headers: m.headers, lh: m.cols.map(c => c.h), rows };
}
/**
 * Fast start: only the newest rows [offset, offset+limit). It reads just the Timestamp column to find them,
 * then reads those rows only (one block when they sit close together, which is normal for a form sheet).
 */
function mainLatest_(offset, limit) {
  const m = mainMeta_(), last = m.sh.getLastRow(), lh = m.cols.map(c => c.h);
  if (last < 2) return { headers: m.headers, lh, total: 0, rows: [] };
  const keys = tsKeys_(m.sh, m.ti), order = [];
  for (let r = 2; r <= last; r++) order.push(r);
  order.sort((p, q) => ((keys[q] || 0) - (keys[p] || 0)) || (q - p));
  const pick = order.slice(offset, offset + limit), got = {};
  if (pick.length) {
    const lo = Math.min.apply(null, pick), hi = Math.max.apply(null, pick);
    if (hi - lo < 500) m.sh.getRange(lo, 1, hi - lo + 1, m.need).getDisplayValues().forEach((v, i) => got[lo + i] = v);
    else pick.forEach(r => got[r] = m.sh.getRange(r, 1, 1, m.need).getDisplayValues()[0]);
  }
  return { headers: m.headers, lh, total: last - 1, rows: pick.filter(r => got[r].some(x => x !== '')).map(r => mainRow_(m, r, got[r])) };
}

function buildMem_() {
  const sh = memberSheet_(), all = headers_(sh);
  if (!all.length) return { headers: [], lh: [], rows: [] };
  const ai = col_(all, CFG.AREA_COL), si = col_(all, CFG.STATUS_COL);
  const sv = col_(all, CFG.SERVER_COL), ui = col_(all, CFG.UID_COL);
  const vis = all.map((h, i) => i).filter(i => norm_(all[i]) !== norm_(CFG.SOURCE_COL));
  const cols = listCols_(all, vis, true), sIdx = searchIdx_(all);
  const need = refsMax_(cols, sIdx.concat([ai, si, sv, ui]));
  const last = sh.getLastRow(), rows = [];
  const vals = last > 1 ? sh.getRange(2, 1, last - 1, need).getDisplayValues() : [];
  vals.forEach((v, i) => {
    if (!v.some(x => x !== '')) return;
    rows.push({ r: i + 2, v: cols.map(c => c.get(v)), k: searchText_(v, sIdx), id: idKey_(v[sv], v[ui]),
                a: ai >= 0 ? v[ai] : '', s: si >= 0 ? String(v[si]).trim() : '' });
  });
  rows.sort(byNewest_(tsKeys_(sh, col_(all, CFG.TIMESTAMP_COL))));
  return { headers: vis.map(i => all[i]), lh: cols.map(c => c.h), rows };
}

const allStatuses_ = () => CFG.STATUSES.concat(CFG.ADMIN_STATUSES);
const isAdminStatus_ = s => CFG.ADMIN_STATUSES.some(x => norm_(x) === norm_(s));
/** Identity key -> {s: status, r: row}. If a player has several entries, "Approved" wins; otherwise the latest entry. */
function statusOf_(memRows) {
  const map = {};
  memRows.forEach(m => {
    if (!m.id) return;
    const c = map[m.id];
    // Blacklisted beats Approved, Approved beats everything else; otherwise the latest entry wins
    const rk = x => isAdminStatus_(x) ? 2 : x === 'Approved' ? 1 : 0;
    if (!c || rk(m.s) > rk(c.s) || (rk(m.s) === rk(c.s) && rk(m.s) === 0 && m.r > c.r)) map[m.id] = { s: m.s, r: m.r, a: m.a };
  });
  return map;
}

/* ---------- Settings tab & branding ---------- */
const driveId_ = v => {
  v = String(v || '').trim();
  const m = v.match(/drive\.google\.com\/file\/d\/([-\w]{15,})/) || v.match(/[?&]id=([-\w]{15,})/) || v.match(/^([-\w]{25,})$/);
  return m ? m[1] : '';
};
/** Turns a Google Drive link (or bare file ID) into a displayable image URL; other https links are used as they are. */
function imgUrl_(v, w) {
  v = String(v || '').trim();
  if (!v) return '';
  const id = driveId_(v);
  if (id) return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + w;
  return /^https?:\/\//i.test(v) ? v : '';
}
/** Address variants to try for the browser-tab icon (Apps Script only accepts PNG / ICO / GIF files). */
function faviconCandidates_(v) {
  v = String(v || '').trim();
  const id = driveId_(v);
  if (id) return ['https://drive.google.com/uc?export=download&id=' + id, 'https://drive.google.com/thumbnail?id=' + id + '&sz=w64', 'https://lh3.googleusercontent.com/d/' + id + '=w64'];
  return /^https?:\/\//i.test(v) ? [v] : [];
}
/** Returns the first candidate that Apps Script accepts as a favicon, or '' (and favErr) if none is accepted. */
function pickFavicon_(v) {
  const list = faviconCandidates_(v);
  for (let i = 0; i < list.length; i++) {
    try { HtmlService.createHtmlOutput('').setFaviconUrl(list[i]); return { favicon: list[i], favErr: '' }; } catch (e) {}
  }
  return { favicon: '', favErr: list.length ? 'unsupported' : '' };
}
const kindOf_ = key => /favicon/i.test(key) ? 'favicon' : /logo/i.test(key) ? 'logo' : '';

const looksVal_ = v => /^https?:\/\//i.test(v) || /^[-\w]{25,}$/.test(v);

/**
 * Finds the existing Settings tab. It NEVER creates a sheet. Candidates are tabs named like "Settings"
 * (exact, then "Setting"/"Settings", then anything containing "setting"); the one whose header row has
 * Name | Value | Notes wins, so a stray blank tab can't be picked by mistake.
 */
function settingsSheet_() {
  const all = book_(CFG.SETTINGS_SPREADSHEET_ID).getSheets(), want = norm_(CFG.SETTINGS_SHEET), cands = [];
  const add = f => all.filter(f).forEach(x => { if (cands.indexOf(x) < 0) cands.push(x); });
  add(x => x.getName() === CFG.SETTINGS_SHEET);
  add(x => norm_(x.getName()) === want);
  add(x => /^settings?$/.test(norm_(x.getName())));
  add(x => norm_(x.getName()).includes('setting'));
  const score = x => {
    if (x.getLastColumn() < 1 || x.getLastRow() < 1) return 0;
    const h = x.getRange(1, 1, 1, x.getLastColumn()).getDisplayValues()[0].map(norm_);
    return (h.includes('name') ? 4 : 0) + (h.includes('value') ? 2 : 0) + (h.includes('notes') ? 1 : 0);
  };
  let best = null, bs = -1;
  cands.forEach(x => { const sc = score(x); if (sc > bs) { best = x; bs = sc; } });
  return best;
}

/**
 * Reads the Settings tab into [{key, value, notes, r, c}] where key = the Name column, r/c = the Value cell.
 * (If the Name | Value header is missing it falls back to plain "name, value" rows.) A labelled hyperlink is read as its link.
 */
function settingsModel_() {
  const sh = settingsSheet_();
  if (!sh) return { sh: null, items: [] };
  const last = sh.getLastRow(), cols = sh.getLastColumn(), items = [];
  if (last < 1 || cols < 1) return { sh, items };
  const rng = sh.getRange(1, 1, last, cols), disp = rng.getDisplayValues();
  let rich = null; try { rich = rng.getRichTextValues(); } catch (e) {}
  const cell = (r, c) => {
    const t = String(disp[r] && disp[r][c] != null ? disp[r][c] : '').trim();
    let u = ''; try { u = rich && rich[r] && rich[r][c] ? (rich[r][c].getLinkUrl() || '') : ''; } catch (e) {}
    return u && !looksVal_(t) ? u : t;
  };
  const first = disp[0].map(norm_);
  const nc = first.indexOf('name') >= 0 ? first.indexOf('name') : first.findIndex(c => c === 'setting' || c === 'key');
  const vc = first.indexOf('value'), tc = first.findIndex(c => c === 'notes' || c === 'note' || c === 'description');
  if (nc >= 0 && vc >= 0) {                                              // Name | Value | Notes
    for (let r = 1; r < last; r++) {
      const key = cell(r, nc); if (!key) continue;
      items.push({ key, value: cell(r, vc), notes: tc >= 0 ? String(disp[r][tc] || '').trim() : '', r: r + 1, c: vc + 1 });
    }
  } else {                                                               // no header: name in column A, value in column B
    for (let r = 0; r < last; r++) { const key = cell(r, 0); if (key) items.push({ key, value: cell(r, 1), notes: '', r: r + 1, c: 2 }); }
  }
  return { sh, items };
}
/** Resolved logo / favicon (the favicon is checked once, then cached; the cache is cleared when an admin saves settings). */
function brand_() {
  const hit = cget_('cfg3'); if (hit) return hit;
  try {
    const items = settingsModel_().items;
    const lg = items.find(i => kindOf_(i.key) === 'logo'), fv = items.find(i => kindOf_(i.key) === 'favicon');
    const pf = pickFavicon_((fv && fv.value) || (lg && lg.value));
    const out = { logo: imgUrl_(lg && lg.value, 640), favicon: pf.favicon, favErr: fv && fv.value ? pf.favErr : '' };
    cput_('cfg3', out); return out;
  } catch (e) { return { logo: '', favicon: '', favErr: '' }; }
}
function settingsPayload_() {
  const m = settingsModel_();
  if (!m.sh) throw new Error('The "' + CFG.SETTINGS_SHEET + '" tab was not found. Tabs in this spreadsheet: ' +
    book_(CFG.SETTINGS_SPREADSHEET_ID).getSheets().map(x => '"' + x.getName() + '"').join(', ') + '. Rename it, or set SETTINGS_SHEET / SETTINGS_SPREADSHEET_ID in CFG.');
  return {
    items: m.items.map(i => { const kind = kindOf_(i.key); return { key: i.key, value: i.value, notes: i.notes, kind, img: kind ? imgUrl_(i.value, 120) : '' }; }),
    brand: brand_()
  };
}

/* ---------- API ---------- */
function login(username, password) {
  const sh = usersSheet_(), h = headers_(sh);
  const ui = col_(h, 'Username'), pi = col_(h, 'Password');
  if (ui < 0 || pi < 0) throw new Error('Users sheet needs "Username" and "Password" columns.');
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getDisplayValues() : [];
  const typed = norm_(username);                     // "juan" is tried as "juan@gmail.com" first; a full email is used as typed
  const cand = typed.indexOf('@') >= 0 ? [typed] : [typed + norm_(CFG.DEFAULT_EMAIL_DOMAIN), typed];
  let row = null;
  for (let i = 0; i < cand.length && !row; i++) row = typed ? vals.find(v => norm_(v[ui]) === cand[i]) : null;
  const stored = row ? String(row[pi]) : null;
  const ok = stored !== null && (stored.indexOf('sha256:') === 0 ? stored === hash_(String(password)) : stored === String(password));
  if (!ok) { Utilities.sleep(600); throw new Error('Wrong username or password.'); }
  CacheService.getScriptCache().remove('users');     // pick up any edits to the Users sheet
  const user = String(row[ui]).trim();
  return { token: makeToken_(user), user };
}

function slice_(list, offset, limit) {
  offset = Math.max(0, Number(offset) || 0);
  limit = Math.min(500, Math.max(1, Number(limit) || CFG.PAGE_SIZE));
  return list.slice(offset, offset + limit);
}

/** Main Responses: one page of rows (newest first) with each row's status from My Memberlist. `q` searches ALL rows. */
/** Who is signed in (light call used to show the nickname in the header; touches no data sheets). */
function whoami(token) {
  const a = auth_(token);
  return { user: a.user, nick: a.nick || a.user, area: a.area, admin: a.admin, full: a.full || '', pos: a.pos || '', org: a.org || '', photo: a.photo || '' };
}

/**
 * Main Responses is NOT cached and nothing is returned until the user searches (min. 2 characters).
 * Every search reads the sheet fresh, so new responses show up immediately.
 */
function getMain(token, force, offset, limit, q) {
  const a = auth_(token), ql = norm_(q);
  const mem = base_('mem3', buildMem_, force), map = statusOf_(mem.rows);      // statuses come from My Memberlist
  const stOf = r => r.id && map[r.id] ? map[r.id].s : '';
  let page;
  if (ql.length < 2) {                                                           // no search: the newest responses, read fresh
    page = mainLatest_(Math.max(0, Number(offset) || 0), Math.min(500, Math.max(1, Number(limit) || 10)));
  } else {                                                                       // search across every response, read fresh
    const mb = buildMain_();
    const list = mb.rows.filter(r => r.v.join(' ').toLowerCase().includes(ql) || r.k.includes(ql) || stOf(r).toLowerCase().includes(ql));
    page = { headers: mb.headers, lh: mb.lh, total: list.length, rows: slice_(list, offset, limit) };
  }
  return {
    user: a.user, nick: a.nick || a.user, area: a.area, admin: a.admin, adminStatuses: CFG.ADMIN_STATUSES, headers: page.headers, lh: page.lh, total: page.total,
    rows: page.rows.map(r => {
      const st = stOf(r), m = r.id && map[r.id] ? map[r.id] : null;
      // Area where it was approved, only when it differs from this response's Area
      const apIn = st === 'Approved' && m && m.a && norm_(m.a) !== norm_(r.a) ? m.a : '';
      // Approve is shown when the user may approve this Area AND (MLBB UID + MLBB Server + Approved) is not in My Memberlist
      return { r: r.r, v: r.v, st, apIn, can: match_(r.a, a) && st !== 'Approved' && !isAdminStatus_(st) };
    })
  };
}

/** My Memberlist: one page of the user's Area (all Areas for admins), newest first. `q` searches ALL their rows. */
function getMembers(token, force, offset, limit, q, status) {
  const a = auth_(token);
  const mem = base_('mem3', buildMem_, force), ql = norm_(q);
  const mine = mem.rows.filter(r => match_(r.a, a));
  // Status counters cover everything this user can see (not just the loaded page, and not affected by the search box)
  const counts = { Total: mine.length }; allStatuses_().forEach(x => counts[x] = 0); counts.Other = 0;
  mine.forEach(r => { const x = allStatuses_().find(y => norm_(y) === norm_(r.s)); counts[x || 'Other']++; });
  // optional status filter (the clickable counters): a status name, or "Other" for anything not in the status lists
  const isKnown = r => allStatuses_().some(y => norm_(y) === norm_(r.s));
  const bySt = !status ? mine : mine.filter(r => status === 'Other' ? !isKnown(r) : norm_(r.s) === norm_(status));
  const list = ql ? bySt.filter(r => r.v.join(' ').toLowerCase().includes(ql) || r.k.includes(ql)) : bySt;
  return {
    user: a.user, nick: a.nick || a.user, area: a.area, admin: a.admin, statuses: allStatuses_(), settable: a.admin ? allStatuses_() : CFG.STATUSES, adminStatuses: CFG.ADMIN_STATUSES, counts, headers: mem.headers, lh: mem.lh, total: list.length,
    rows: slice_(list, offset, limit).map(r => ({ r: r.r, v: r.v }))
  };
}

/** One full row, loaded when a row is clicked. Values line up with `headers` from getMain/getMembers. */
function getDetail(token, page, row) {
  const a = auth_(token); row = Number(row);
  const sh = page === 'main' ? mainSheet_() : memberSheet_(), hdr = headers_(sh);
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Row not found.');
  const v = sh.getRange(row, 1, 1, hdr.length).getDisplayValues()[0];
  if (page !== 'main' && !match_(v[col_(hdr, CFG.AREA_COL)], a)) throw new Error('Not allowed.');
  return v.filter((x, i) => norm_(hdr[i]) !== norm_(CFG.SOURCE_COL));
}

function approve(token, row) {
  const a = auth_(token); row = Number(row);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ms = mainSheet_(), headers = headers_(ms), rng = ms.getRange(row, 1, 1, headers.length);
    const vals = rng.getDisplayValues()[0], raw = rng.getValues()[0];
    if (!match_(vals[col_(headers, CFG.AREA_COL)], a)) throw new Error('This row belongs to a different Area.');

    const sh = memberSheet_();
    let mh = headers_(sh);
    if (!mh.length) mh = headers.slice();
    [CFG.STATUS_COL, CFG.UPDATED_BY_COL].forEach(n => { if (col_(mh, n) < 0) mh.push(n); });
    sh.getRange(1, 1, 1, mh.length).setValues([mh]).setFontWeight('bold');

    // Same MLBB UID + MLBB Server already in My Memberlist with status Approved? Then there is nothing to approve.
    const key = idKey_(vals[col_(headers, CFG.SERVER_COL)], vals[col_(headers, CFG.UID_COL)]);
    const existing = key ? statusOf_(buildMem_().rows)[key] : null;
    if (existing && isAdminStatus_(existing.s)) throw new Error('This player is blacklisted. An admin has to change the status first.');
    if (existing && existing.s === 'Approved') {
      const rowArea = vals[col_(headers, CFG.AREA_COL)];
      throw new Error(existing.a && norm_(existing.a) !== norm_(rowArea) ? 'Already approved in ' + existing.a : 'Already approved.');
    }

    // Otherwise (new, or Rejected / Duplicate / Transferred) always add a NEW row.
    const out = [], fm = [];
    mh.forEach(h => {
      const n = norm_(h), si = col_(headers, h);
      if (n === norm_(CFG.STATUS_COL)) { out.push('Approved'); fm.push('@'); }
      else if (n === norm_(CFG.UPDATED_BY_COL)) { out.push(a.user); fm.push('@'); }
      else if (si >= 0 && n === norm_(CFG.TIMESTAMP_COL) && raw[si] instanceof Date) { out.push(raw[si]); fm.push(CFG.TIMESTAMP_FORMAT); } // keep a real date+time so it sorts
      else { out.push(si >= 0 ? vals[si] : ''); fm.push('@'); }
    });
    sh.getRange(sh.getLastRow() + 1, 1, 1, mh.length).setNumberFormats([fm]).setValues([out]);
    const tc = col_(mh, CFG.TIMESTAMP_COL);                    // make the whole Timestamp column show date AND time
    if (tc >= 0) sh.getRange(2, tc + 1, sh.getLastRow() - 1, 1).setNumberFormat(CFG.TIMESTAMP_FORMAT);
    cdel_('mem3');
    return true;
  } finally { lock.releaseLock(); }
}

/**
 * Saves the Status and/or Note of a member row. Only real changes are written; any change also stamps "Last Updated By".
 * Returns what is now stored so the page can update itself.
 */
function updateMember(token, row, status, note) {
  const a = auth_(token); row = Number(row);
  if (status != null && allStatuses_().indexOf(status) < 0) throw new Error('Invalid status.');
  const sh = memberSheet_(), headers = headers_(sh), cols = headers.length;
  const ai = col_(headers, CFG.AREA_COL), si = col_(headers, CFG.STATUS_COL), ni = col_(headers, CFG.NOTE_COL);
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Row not found.');
  const cur = sh.getRange(row, 1, 1, cols).getDisplayValues()[0];
  if (!match_(cur[ai], a)) throw new Error('Not allowed.');
  if (!a.admin && status != null && si >= 0) {                      // admin-only statuses (e.g. Blacklisted): only admins may set or change them
    const was = String(cur[si]).trim();
    if (status !== was && (isAdminStatus_(status) || isAdminStatus_(was))) throw new Error('Only an admin can set or change the "' + (isAdminStatus_(status) ? status : was) + '" status.');
  }

  let changed = false;
  if (status != null && si >= 0 && String(cur[si]).trim() !== status) { sh.getRange(row, si + 1).setValue(status); changed = true; }
  if (note != null && ni >= 0) {
    note = String(note).slice(0, 5000);
    if (String(cur[ni]) !== note) { sh.getRange(row, ni + 1).setNumberFormat('@').setValue(note); changed = true; }   // text format: never treated as a formula
  }
  if (changed) {
    let ui = col_(headers, CFG.UPDATED_BY_COL);
    if (ui < 0) { ui = cols; sh.getRange(1, ui + 1).setValue(CFG.UPDATED_BY_COL).setFontWeight('bold'); }
    sh.getRange(row, ui + 1).setValue(a.user);
    cdel_('mem3');
  }
  return {
    changed, by: a.user,
    status: si >= 0 ? (status != null ? status : cur[si]) : '',
    note: ni >= 0 ? (note != null ? note : cur[ni]) : ''
  };
}

/* ---------- About Us (organization hierarchy) & profile photos ---------- */
/** About Us: everyone who has an Org Level in the Users tab. Only name, position, photo and level are sent (never emails or roles). */
function getAbout(token, force) {
  auth_(token);
  if (!force) { const hit = cget_('about1'); if (hit) return hit; }
  const sh = usersSheet_(), h = headers_(sh);
  const oi = col_(h, 'Org Level'), fu = col_(h, 'Full Name'), po = col_(h, 'Position'), fi = col_(h, 'Photo Link'), ni = col_(h, 'Nickname');
  if (oi < 0) throw new Error('The Users tab has no "Org Level" column.');
  const last = sh.getLastRow(), vals = last > 1 ? sh.getRange(2, 1, last - 1, h.length).getDisplayValues() : [];
  const people = vals.filter(v => String(v[oi]).trim()).map(v => ({
    code: String(v[oi]).trim().toUpperCase(),
    name: (fu >= 0 && String(v[fu]).trim()) || (ni >= 0 && String(v[ni]).trim()) || 'Member',
    position: po >= 0 ? String(v[po]).trim() : '',
    photo: fi >= 0 ? imgUrl_(v[fi], 240) : ''
  }));
  const out = { people };
  cput_('about1', out);
  return out;
}

function photoFolder_() {
  const props = PropertiesService.getScriptProperties(), cfg = String(CFG.PHOTO_FOLDER_ID || '');
  const m = cfg.match(/folders\/([-\w]+)/), id = m ? m[1] : (driveId_(cfg) || props.getProperty('PHOTO_FOLDER'));
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(CFG.PHOTO_FOLDER_NAME), f = it.hasNext() ? it.next() : DriveApp.createFolder(CFG.PHOTO_FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}
/** Moves a replaced / removed photo to the trash, but only if the app uploaded it (it sits in the photo folder). */
function trashOwnPhoto_(link, folder) {
  const id = driveId_(link); if (!id) return;
  try {
    const f = DriveApp.getFileById(id), ps = f.getParents();
    while (ps.hasNext()) if (ps.next().getId() === folder.getId()) { f.setTrashed(true); return; }
  } catch (e) {}
}
function userRow_(sh, h, user) {
  const ui = col_(h, 'Username'), last = sh.getLastRow(), names = last > 1 ? sh.getRange(2, ui + 1, last - 1, 1).getDisplayValues() : [];
  const i = names.findIndex(x => norm_(x[0]) === norm_(user));
  if (i < 0) throw new Error('SESSION');
  return i + 2;
}
/** Settings page: saves the user's photo (a square JPEG made in the browser) to Drive and puts its link in Users > Photo Link. */
function uploadPhoto(token, b64) {
  const a = auth_(token);
  b64 = String(b64 || '');
  if (b64.length < 200 || b64.length > 1500000) throw new Error('The photo is empty or too large.');
  const bytes = Utilities.base64Decode(b64);
  if ((bytes[0] & 255) !== 255 || (bytes[1] & 255) !== 216 || (bytes[2] & 255) !== 255) throw new Error('Please choose a JPG or PNG picture.');
  const sh = usersSheet_(), h = headers_(sh), fi = col_(h, 'Photo Link');
  if (fi < 0) throw new Error('The Users tab has no "Photo Link" column.');
  const row = userRow_(sh, h, a.user), old = sh.getRange(row, fi + 1).getDisplayValue();
  const folder = photoFolder_();
  const file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', 'photo-' + a.user.replace(/[^\w.@-]/g, '_') + '-' + Date.now() + '.jpg'));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { file.setTrashed(true); throw new Error('Drive would not make the photo visible to others (link sharing may be turned off for your organization).'); }
  const link = 'https://drive.google.com/file/d/' + file.getId() + '/view';
  sh.getRange(row, fi + 1).setNumberFormat('@').setValue(link);
  trashOwnPhoto_(old, folder);
  CacheService.getScriptCache().remove('users'); cdel_('about1');
  return { photo: imgUrl_(link, 240) };
}
function removePhoto(token) {
  const a = auth_(token);
  const sh = usersSheet_(), h = headers_(sh), fi = col_(h, 'Photo Link');
  if (fi < 0) throw new Error('The Users tab has no "Photo Link" column.');
  const row = userRow_(sh, h, a.user), old = sh.getRange(row, fi + 1).getDisplayValue();
  sh.getRange(row, fi + 1).setValue('');
  trashOwnPhoto_(old, photoFolder_());
  CacheService.getScriptCache().remove('users'); cdel_('about1');
  return { photo: '' };
}
/** Run this once from the Apps Script editor so the app is allowed to store uploaded photos in Drive. */
function authorizeDrive() { DriveApp.getRootFolder(); Logger.log('Drive access is authorised.'); }

/** Settings page: a user changes their own nickname (stored in the Users tab, "Nickname" column). */
function updateNickname(token, nickname) {
  const a = auth_(token);
  const nick = String(nickname == null ? '' : nickname).trim().replace(/\s+/g, ' ').slice(0, 40);
  const sh = usersSheet_(), h = headers_(sh), ui = col_(h, 'Username'), ni = col_(h, 'Nickname');
  if (ni < 0) throw new Error('The Users tab has no "Nickname" column.');
  const last = sh.getLastRow(), names = last > 1 ? sh.getRange(2, ui + 1, last - 1, 1).getDisplayValues() : [];
  const i = names.findIndex(x => norm_(x[0]) === norm_(a.user));
  if (i < 0) throw new Error('SESSION');
  sh.getRange(i + 2, ni + 1).setNumberFormat('@').setValue(nick);               // text format: never treated as a formula
  CacheService.getScriptCache().remove('users');
  return { nick: nick || a.user };
}

/** Admin: the Settings tab (Name | Value | Notes). Admins only. */
function getSettings(token) {
  if (!auth_(token).admin) throw new Error('Not allowed.');
  return settingsPayload_();
}
/**
 * Admin: save changed values ([{key, value}]). Each row is found by its Name and only its Value cell is updated.
 * Nothing is created: unknown names are ignored, and everything is validated first so nothing is half-saved.
 */
function saveSettings(token, changes) {
  if (!auth_(token).admin) throw new Error('Not allowed.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const m = settingsModel_();
    if (!m.sh) settingsPayload_();                                   // throws the "tab not found" message
    const byKey = {}; m.items.forEach(i => byKey[i.key] = i);
    const todo = [];
    (changes || []).forEach(c => {
      const it = byKey[c && c.key]; if (!it) return;
      const v = String(c.value == null ? '' : c.value).trim().slice(0, 2000);
      if (kindOf_(it.key) && v && !imgUrl_(v, 64)) throw new Error('"' + it.key + '" must be a Google Drive link or an image link starting with https://');
      if (v !== it.value.trim()) todo.push({ it, v });
    });
    todo.forEach(t => m.sh.getRange(t.it.r, t.it.c).setNumberFormat('@').setValue(t.v));   // text format: never treated as a formula
    cdel_('cfg3');
    return settingsPayload_();
  } finally { lock.releaseLock(); }
}

/** Optional: run once from the editor to show date + time for every existing Timestamp in My Memberlist. */
function fixTimestampFormat() {
  const sh = memberSheet_(), tc = col_(headers_(sh), CFG.TIMESTAMP_COL), last = sh.getLastRow();
  if (tc < 0 || last < 2) return;
  sh.getRange(2, tc + 1, last - 1, 1).setNumberFormat(CFG.TIMESTAMP_FORMAT);
  cdel_('mem3');
}
