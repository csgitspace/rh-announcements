/**
 * Announcement Hub — Configuration & one-time setup
 * Royalton-Hartland CSD
 *
 * Run setup() ONCE from the editor. It creates the spreadsheet, the Drive
 * folder structure, seeds the routing table, and stores every ID in Script
 * Properties so re-pasting this code never loses your configuration.
 */

var APP = {
  TITLE: 'Announcement Hub',
  ROOT_FOLDER: 'Announcement Hub',
  MAX_FILE_MB: 20,           // per file; bigger video belongs in the links field
  CHUNK_BYTES: 3145728,      // 3 MB upload chunks
  ALLOWED_EXT: ['jpg', 'jpeg', 'png', 'gif', 'pdf'],
  DEFAULT_LEAD_DAYS: 2,      // "two full school days" guidance
  LOGO: 'https://files.smartsites.parentsquare.com/4898/img_pd_102121_mxvilf.png',
  BRAND: {
    purple: '#63419A',
    dark: '#4A2F78',
    lavender: '#AC9AC9',
    tint: '#F3F0F8',
    grey: '#636263'
  }
};

var SHEET = {
  SUBMISSIONS: 'Submissions',
  TASKS: 'Tasks',
  ROUTING: 'Routing',
  LOG: 'Log'
};

var HEADERS = {
  SUBMISSIONS: ['ID', 'Submitted', 'Name', 'Email', 'Content', 'Locations',
                'Other Location', 'Start Date', 'Expire Date', 'Media Links',
                'Files', 'Folder URL', 'PDF URL', 'Status'],
  TASKS: ['Task ID', 'ID', 'Location', 'Assignees', 'Status', 'Posted By',
          'Posted At', 'Notes'],
  ROUTING: ['Location', 'Recipients (comma separated)', 'Posting notes for this location',
            'Lead time (days)', 'Active'],
  LOG: ['When', 'Who', 'Action', 'Detail']
};

/** Destinations seeded on first run — edit them later on the Routing tab. */
var SEED_ROUTING = [
  ['HS Morning Announcements', '', 'Read during first period. Keep to ~40 words.', 0, 'Yes'],
  ['HS Announcement TVs', '', 'Landscape image required. Slides rotate every 12 seconds.', 2, 'Yes'],
  ['MS Morning Announcements', '', 'Read during homeroom.', 0, 'Yes'],
  ['MS Cafeteria TV', '', 'Landscape image required.', 2, 'Yes'],
  ['District Social Media (Facebook/Instagram/Twitter)', '', 'Square or landscape image preferred. No student names without a release on file.', 2, 'Yes'],
  ['Other', '', 'Check the submitter note for where this needs to go.', 2, 'Yes']
];

/* ---------------------------------------------------------------- properties */

function props_() { return PropertiesService.getScriptProperties(); }
function prop_(key, fallback) {
  var v = props_().getProperty(key);
  return (v === null || v === '') ? (fallback === undefined ? '' : fallback) : v;
}

function adminEmails_() {
  return prop_('ADMIN_EMAILS', '').split(',')
    .map(function (s) { return s.trim().toLowerCase(); })
    .filter(String);
}

function isAdmin_(email) {
  return adminEmails_().indexOf(String(email).toLowerCase()) > -1;
}

/* -------------------------------------------------------------------- setup */

/**
 * ONE-TIME SETUP. Run this from the editor, approve the permissions prompt,
 * then open the Execution log for the links it prints.
 */
function setup() {
  var p = props_();

  // 1. Spreadsheet
  var ssId = p.getProperty('SHEET_ID');
  var ss;
  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create(APP.TITLE + ' — Data');
    p.setProperty('SHEET_ID', ss.getId());
  }

  ensureSheet_(ss, SHEET.SUBMISSIONS, HEADERS.SUBMISSIONS);
  ensureSheet_(ss, SHEET.TASKS, HEADERS.TASKS);
  ensureSheet_(ss, SHEET.LOG, HEADERS.LOG);

  var routing = ensureSheet_(ss, SHEET.ROUTING, HEADERS.ROUTING);
  if (routing.getLastRow() < 2) {
    routing.getRange(2, 1, SEED_ROUTING.length, SEED_ROUTING[0].length).setValues(SEED_ROUTING);
    routing.setColumnWidth(1, 260).setColumnWidth(2, 280).setColumnWidth(3, 340);
  }

  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  // 2. Drive folders
  var rootId = p.getProperty('ROOT_FOLDER_ID');
  var root = rootId ? DriveApp.getFolderById(rootId) : DriveApp.createFolder(APP.ROOT_FOLDER);
  p.setProperty('ROOT_FOLDER_ID', root.getId());
  p.setProperty('DRAFTS_FOLDER_ID', childFolder_(root, '_drafts').getId());

  // 3. Admin defaults
  if (!p.getProperty('ADMIN_EMAILS')) {
    p.setProperty('ADMIN_EMAILS', Session.getEffectiveUser().getEmail());
  }
  if (!p.getProperty('REPLY_TO')) {
    p.setProperty('REPLY_TO', Session.getEffectiveUser().getEmail());
  }

  installTriggers();

  Logger.log('Setup complete.');
  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Drive folder: ' + root.getUrl());
  Logger.log('Next: fill in Recipients on the Routing tab, deploy the web app, ' +
             'then run saveWebAppUrl() with the deployment URL.');
}

/**
 * After deploying, paste your /exec URL here and run this once so that emails
 * can link back to the dashboard.
 */
function saveWebAppUrl() {
  var url = 'PASTE_YOUR_EXEC_URL_HERE';
  if (url.indexOf('http') !== 0) throw new Error('Edit saveWebAppUrl() and paste your /exec URL first.');
  props_().setProperty('WEB_APP_URL', url.trim());
  Logger.log('Saved: ' + url);
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['sendDailyDigest', 'expireAnnouncements', 'cleanUpDrafts'].indexOf(t.getHandlerFunction()) > -1) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendDailyDigest').timeBased().atHour(6).everyDays(1).create();
  ScriptApp.newTrigger('expireAnnouncements').timeBased().atHour(1).everyDays(1).create();
  ScriptApp.newTrigger('cleanUpDrafts').timeBased().atHour(2).everyDays(1).create();
}

/* ------------------------------------------------------------------ helpers */

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground(APP.BRAND.purple).setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}

function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function ss_() { return SpreadsheetApp.openById(prop_('SHEET_ID')); }
function sheet_(name) { return ss_().getSheetByName(name); }
function rootFolder_() { return DriveApp.getFolderById(prop_('ROOT_FOLDER_ID')); }
function draftsFolder_() { return DriveApp.getFolderById(prop_('DRAFTS_FOLDER_ID')); }

function schoolYear_(d) {
  d = d || new Date();
  var y = d.getFullYear();
  return (d.getMonth() >= 6) ? y + '-' + (y + 1) : (y - 1) + '-' + y;
}

function rows_(name) {
  var sh = sheet_(name);
  if (sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return values.map(function (r, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, c) { if (h) o[h] = r[c]; });
    return o;
  });
}

function colIndex_(name, header) {
  var head = sheet_(name).getRange(1, 1, 1, sheet_(name).getLastColumn()).getValues()[0];
  var i = head.indexOf(header);
  if (i < 0) throw new Error('Column "' + header + '" not found on ' + name);
  return i + 1;
}

function fmtDate_(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'EEE, MMM d, yyyy');
}

/**
 * Drive throws transient "Service error: Drive" under load. Retrying with
 * backoff clears almost all of them. Only retries errors that look transient —
 * a genuine permission or not-found error still fails fast.
 */
function withRetry_(fn, label, tries) {
  tries = tries || 4;
  var wait = 700;
  for (var i = 0; i < tries; i++) {
    try {
      return fn();
    } catch (err) {
      var msg = String((err && err.message) || err);
      var transient = /service error|internal error|try again|timed out|timeout|rate limit|too many|unavailable|backend/i.test(msg);
      if (!transient || i === tries - 1) {
        if (transient) log_('Drive gave up', (label || '') + ': ' + msg);
        throw err;
      }
      Utilities.sleep(wait + Math.floor(Math.random() * 500));
      wait *= 2;
    }
  }
}

function log_(action, detail) {
  try {
    sheet_(SHEET.LOG).appendRow([new Date(), userEmail_(), action, detail]);
  } catch (err) { /* logging must never break the app */ }
}

function userEmail_() {
  return (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();
}

/** Active routing rows, as plain objects. */
function locations_() {
  return rows_(SHEET.ROUTING)
    .filter(function (r) { return r.Location && String(r.Active).toLowerCase() !== 'no'; })
    .map(function (r) {
      return {
        name: String(r.Location),
        notes: String(r['Posting notes for this location'] || ''),
        leadDays: Number(r['Lead time (days)'] || 0),
        recipients: String(r['Recipients (comma separated)'] || '')
          .split(',').map(function (s) { return s.trim(); }).filter(String)
      };
    });
}
