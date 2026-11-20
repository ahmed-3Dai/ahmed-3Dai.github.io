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
    if (category) collectImages_(folder, category, out);
  }

  try { cache.put('drive_images_v1', JSON.stringify(out), CACHE_SECONDS); } catch (err) { /* list too big for cache: fine */ }
  return out;
}

/** Collects every image in a folder AND in its nested sub-folders. */
function collectImages_(folder, category, out) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getMimeType()).indexOf('image/') === 0) {
      out.push({ id: f.getId(), name: f.getName(), category: category });
    }
  }
  const inner = folder.getFolders();
  while (inner.hasNext()) collectImages_(inner.next(), category, out);
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
      if (group) collectImages_(gf, group, found);
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
          if (title) cache.put(key, title, 21600);   // keep for 6 hours
        }
      } catch (err) { /* ignore: the page will keep the old name */ }
    }
    if (title) out[id] = title;
  });
  return out;
}
