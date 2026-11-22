/**
 * >>> لوحة التحكم / Control panel password <<<
 * اكتب هنا كلمة سر خاصة بيك (بدل change-me) قبل ما تستخدم لوحة التحكم، وبعدها اعمل Deploy جديد.
 * Put your own secret password here (instead of change-me), then deploy a new version.
 */
const ADMIN_PASSWORD = 'change-me';
const SPREADSHEET_ID = '1JCg-HsmH4_Xzz86lZ45Va-KMa4wpiFQwtsbZdvF4ymA';

/**
 * Ahmed Abd Elhamid - 3D Portfolio Web App (Container-Bound Edition)
 * Bound directly to Google Spreadsheet: 1JCg-HsmH4_Xzz86lZ45Va-KMa4wpiFQwtsbZdvF4ymA
 * Root Drive Folder: 1AoQOxSiPPQkPqJWV6BPGBE-Be0eKlDB0
 *
 * The images on the site come straight from the sub-folders of the root Drive folder:
 *   Trophies  -> Trophies tab          Gifts          -> Gifts tab
 *   Boxes     -> Boxes tab             Booths & Events-> Booths & Events tab
 *   Heritage Gifts -> Heritage tab     Ramadan Gifts  -> Ramadan Gifts tab
 *   Video Animation -> Animation Videos tab
 * Partner logos: folder "Saudi Logos" with the sub-folders  Banks / Company / Ministry.
 * Add a picture to a folder = it appears on the site. Delete it = it disappears.
 */

const ROOT_FOLDER_ID = '1AoQOxSiPPQkPqJWV6BPGBE-Be0eKlDB0';
const CACHE_SECONDS = 60;   // how long the list is reused (lower = faster updates, higher = faster site)

function doGet(e) {
  // JSON feed for sites hosted outside Google (e.g. GitHub Pages):  <web app url>?action=images
  if (e && e.parameter && e.parameter.action === 'images') {
    return ContentService
      .createTextOutput(JSON.stringify(getDriveImages()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // JSON feed for the YouTube video names:  <web app url>?action=titles&ids=ID1,ID2
  if (e && e.parameter && e.parameter.action === 'titles') {
    const ids = String(e.parameter.ids || '').split(',').filter(Boolean).slice(0, 60);
    return ContentService
      .createTextOutput(JSON.stringify(getVideoTitles(ids)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // JSON feed for the partner logos:  <web app url>?action=logos
  if (e && e.parameter && e.parameter.action === 'logos') {
    return ContentService
      .createTextOutput(JSON.stringify(getDriveLogos()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Ahmed Abd Elhamid | 3D Product Visualization & Motion Craft")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Folder name -> site category. Order matters (e.g. "Heritage Gifts" must be checked before "Gifts"). */
function categoryForFolder_(name) {
  const n = String(name || '').toLowerCase();
  if (n.indexOf('ramadan') > -1) return 'ramadan';
  if (n.indexOf('heritage') > -1) return 'heritage';
  if (n.indexOf('booth') > -1 || n.indexOf('event') > -1) return 'booths';
  if (n.indexOf('trophy') > -1 || n.indexOf('trophies') > -1 || n.indexOf('shield') > -1) return 'trophy';
  if (n.indexOf('box') > -1) return 'boxes';
  if (n.indexOf('video') > -1 || n.indexOf('animation') > -1 || n.indexOf('motion') > -1) return 'motion';
  if (n.indexOf('gift') > -1) return 'gifts';
  return null;   // unknown folder: ignored
}

/** Called by the page (google.script.run) and by doGet?action=images. */
function getDriveImages() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('drive_images_v1');
  if (cached) return JSON.parse(cached);

  const out = [];
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const subs = root.getFolders();
  while (subs.hasNext()) {
    const folder = subs.next();
    if (/logo/i.test(folder.getName())) continue;   // the logos folder is handled by getDriveLogos()
    const category = categoryForFolder_(folder.getName());
    if (category) collectImages_(folder, category, out, folder.getId());
  }

  try { cache.put('drive_images_v1', JSON.stringify(out), CACHE_SECONDS); } catch (err) { /* list too big for cache: fine */ }
  return out;
}

/** Collects every image in a folder AND in its nested sub-folders. */
function collectImages_(folder, category, out, topId) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getMimeType()).indexOf('image/') === 0) {
      out.push({ id: f.getId(), name: f.getName(), category: category, folder: topId });
    }
  }
  const inner = folder.getFolders();
  while (inner.hasNext()) collectImages_(inner.next(), category, out, topId);
}

/**
 * Run this ONCE from the Apps Script editor (select setupPublicSharing, press Run).
 * It lets visitors see the pictures: "Anyone with the link can view". Files you add later inherit it.
 */
function setupPublicSharing() {
  DriveApp.getFolderById(ROOT_FOLDER_ID).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

/** Optional: run after changing things if you want the site to refresh immediately. */
function clearImageCache() {
  CacheService.getScriptCache().removeAll(['drive_images_v1', 'drive_logos_v1']);
}

/* ================= Partner logos: "Saudi Logos" folder ================= */
const LOGOS_FOLDER_NAME = 'Saudi Logos';

/** Sub-folder name -> logo row on the site. */
function logoGroupForFolder_(name) {
  const n = String(name || '').toLowerCase();
  if (n.indexOf('bank') > -1 || n.indexOf('financ') > -1) return 'banks';                 // Banks & Financial Institutions
  if (n.indexOf('minist') > -1 || n.indexOf('gov') > -1 || n.indexOf('authorit') > -1) return 'gov';   // Ministries & Government
  if (n.indexOf('compan') > -1 || n.indexOf('corporat') > -1 || n.indexOf('major') > -1) return 'major';  // Major Companies
  return null;
}

/** Called by the page (google.script.run) and by doGet?action=logos. */
function getDriveLogos() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('drive_logos_v1');
  if (cached) return JSON.parse(cached);

  const found = [];
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const subs = root.getFolders();
  while (subs.hasNext()) {
    const logos = subs.next();
    if (logos.getName().toLowerCase() !== LOGOS_FOLDER_NAME.toLowerCase()) continue;
    const groups = logos.getFolders();
    while (groups.hasNext()) {
      const gf = groups.next();
      const group = logoGroupForFolder_(gf.getName());
      if (group) collectImages_(gf, group, found, gf.getId());
    }
  }
  // file name without extension becomes the logo name (used as hover text)
  const out = found.map(function (f) {
    return { id: f.id, name: String(f.name).replace(/\.[^.]+$/, '').replace(/[_-]+/g, ' ').trim(), group: f.category };
  });

  try { cache.put('drive_logos_v1', JSON.stringify(out), CACHE_SECONDS); } catch (err) { }
  return out;
}

/* ================= YouTube video names (same title as on YouTube) ================= */
/** Called by the page (google.script.run) and by doGet?action=titles. Returns {videoId: title}. */
function getVideoTitles(ids) {
  const cache = CacheService.getScriptCache();
  const out = {};
  (ids || []).forEach(function (id) {
    if (!/^[\w-]{6,20}$/.test(id)) return;
    const key = 'yt_title_' + id;
    let title = cache.get(key);
    if (!title) {
      try {
        const r = UrlFetchApp.fetch(
          'https://www.youtube.com/oembed?format=json&url=' + encodeURIComponent('https://www.youtube.com/watch?v=' + id),
          { muteHttpExceptions: true });
        if (r.getResponseCode() === 200) {
          title = JSON.parse(r.getContentText()).title || '';
          if (title) cache.put(key, title, 300);   // keep for 5 minutes, so a renamed video is renamed on the site quickly
        }
      } catch (err) { /* ignore: the page will keep the old name */ }
    }
    if (title) out[id] = title;
  });
  return out;
}

/* ================= Control panel API (called from the secret panel inside the site) ================= */
function adminSS_() {
  try { const a = SpreadsheetApp.getActiveSpreadsheet(); if (a) return a; } catch (e) { }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function adminSheet_(name) {
  const sh = adminSS_().getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}
function adminCols_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const m = {};
  head.forEach(function (h, i) { m[String(h).trim().toLowerCase()] = i; });
  return m;
}
function adminYt_(u) {
  const s = String(u || '').trim();
  const m = s.match(/(?:v=|youtu\.be\/|embed\/|shorts\/|live\/)([\w-]{11})/);
  if (m) return m[1];
  return /^[\w-]{11}$/.test(s) ? s : '';
}
function adminIsVideoRow_(r, c) {
  return String(r[c['type']] || '').toLowerCase() === 'video' || !!adminYt_(r[c['media_url']]);
}

/** One entry point for the panel. Every call needs the password. */
function adminApi(password, fn, p) {
  if (!ADMIN_PASSWORD || ADMIN_PASSWORD === 'change-me') return { ok: false, error: 'SET_PASSWORD' };
  if (String(password) !== String(ADMIN_PASSWORD)) { Utilities.sleep(700); return { ok: false, error: 'BAD_PASSWORD' }; }
  p = p || {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    switch (fn) {
      case 'ping': return { ok: true };
      case 'listVideos': return adminListVideos_();
      case 'addVideo': return adminAddVideo_(p);
      case 'updateVideo': return adminUpdateVideo_(p);
      case 'moveVideo': return adminMoveVideo_(p);
      case 'deleteVideo': return adminDeleteVideo_(p);
      case 'setHeroFirst': return adminSetHeroFirst_(p);
      case 'diagnostics': return adminDiagnostics_();
      default: return { ok: false, error: 'UNKNOWN_FUNCTION' };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function adminListVideos_() {
  const sh = adminSheet_('Projects'), c = adminCols_(sh), last = sh.getLastRow();
  const out = [];
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    vals.forEach(function (r, i) {
      if (!adminIsVideoRow_(r, c)) return;
      out.push({ row: i + 2, id: String(r[c['id']]), category: r[c['category']], tags: r[c['tags']], status: r[c['status']] || 'Active', yt: adminYt_(r[c['media_url']]) });
    });
  }
  return { ok: true, videos: out };
}

function adminAddVideo_(p) {
  const yt = adminYt_(p.url);
  if (!yt) return { ok: false, error: 'BAD_URL' };
  const sh = adminSheet_('Projects'), c = adminCols_(sh), last = sh.getLastRow(), width = sh.getLastColumn();
  let maxId = 0;
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, width).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (adminYt_(vals[i][c['media_url']]) === yt) return { ok: false, error: 'DUPLICATE' };
      const n = Number(vals[i][c['id']]); if (n > maxId) maxId = n;
    }
  }
  const row = []; for (let i = 0; i < width; i++) row.push('');
  const set = function (k, v) { if (c[k] !== undefined) row[c[k]] = v; };
  set('id', maxId + 1); set('category', String(p.category || 'motion')); set('type', 'video');
  set('title_en', '3D Animation'); set('title_ar', '3D Animation');
  set('media_url', 'https://www.youtube.com/watch?v=' + yt);
  set('thumbnail_url', 'https://img.youtube.com/vi/' + yt + '/maxresdefault.jpg');
  set('tags', String(p.tags || '')); set('status', 'Active');
  sh.insertRowBefore(2);                       // newest video goes first
  sh.getRange(2, 1, 1, width).setValues([row]);
  return { ok: true, id: maxId + 1 };
}

function adminCheckRow_(sh, c, p) {
  const row = Number(p.row);
  if (!row || row < 2 || row > sh.getLastRow()) return null;
  const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  if (String(vals[c['id']]) !== String(p.id)) return null;      // the sheet changed in the meantime
  return row;
}

function adminUpdateVideo_(p) {
  const sh = adminSheet_('Projects'), c = adminCols_(sh), row = adminCheckRow_(sh, c, p);
  if (!row) return { ok: false, error: 'STALE' };
  if (p.tags !== undefined && c['tags'] !== undefined) sh.getRange(row, c['tags'] + 1).setValue(String(p.tags));
  if (p.status !== undefined && c['status'] !== undefined) sh.getRange(row, c['status'] + 1).setValue(String(p.status) === 'Active' ? 'Active' : 'Hidden');
  return { ok: true };
}

function adminMoveVideo_(p) {
  const sh = adminSheet_('Projects'), c = adminCols_(sh), row = adminCheckRow_(sh, c, p);
  if (!row) return { ok: false, error: 'STALE' };
  const dir = Number(p.dir) < 0 ? -1 : 1, last = sh.getLastRow(), width = sh.getLastColumn();
  let other = row + dir;
  while (other >= 2 && other <= last) {
    const r = sh.getRange(other, 1, 1, width).getValues()[0];
    if (adminIsVideoRow_(r, c)) break;
    other += dir;
  }
  if (other < 2 || other > last) return { ok: true };           // already first / last
  const a = sh.getRange(row, 1, 1, width), b = sh.getRange(other, 1, 1, width);
  const av = a.getValues(), bv = b.getValues();
  a.setValues(bv); b.setValues(av);
  return { ok: true };
}

function adminDeleteVideo_(p) {
  const sh = adminSheet_('Projects'), c = adminCols_(sh), row = adminCheckRow_(sh, c, p);
  if (!row) return { ok: false, error: 'STALE' };
  sh.deleteRow(row);
  return { ok: true };
}

/** Which video opens the home page (Settings sheet, key hero_first). Empty = keep the normal order. */
function adminSetHeroFirst_(p) {
  const sh = adminSheet_('Settings'), yt = adminYt_(p.yt) || '';
  const last = sh.getLastRow();
  const keys = last >= 1 ? sh.getRange(1, 1, last, 1).getValues() : [];
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === 'hero_first') { sh.getRange(i + 1, 2, 1, 2).setValues([[yt, yt]]); return { ok: true }; }
  }
  sh.appendRow(['hero_first', yt, yt]);
  return { ok: true };
}

/** What the script sees in Drive: every folder, the category it maps to, and how many pictures it holds. */
function adminDiagnostics_() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID), out = [], subs = root.getFolders();
  while (subs.hasNext()) {
    const f = subs.next(), name = f.getName(), isLogos = /logo/i.test(name);
    const cat = isLogos ? '(partner logos)' : categoryForFolder_(name);
    const cnt = { images: 0, other: 0, sub: 0 };
    adminCount_(f, cnt);
    out.push({ folder: name, category: cat || '— ignored (name not recognised) —', images: cnt.images, other: cnt.other, subfolders: cnt.sub });
  }
  return { ok: true, folders: out };
}
function adminCount_(folder, cnt) {
  const files = folder.getFiles();
  while (files.hasNext()) { if (String(files.next().getMimeType()).indexOf('image/') === 0) cnt.images++; else cnt.other++; }
  const inner = folder.getFolders();
  while (inner.hasNext()) { cnt.sub++; adminCount_(inner.next(), cnt); }
}
