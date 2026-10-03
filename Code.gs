/**
 * Announcement Hub — web app entry points and server API
 */

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'form';
  var id = (e && e.parameter && e.parameter.id) || '';

  if (page === 'print') {
    var p = HtmlService.createTemplateFromFile('Print');
    p.data = JSON.stringify(printRecord_(id));
    return p.evaluate()
      .setTitle(id + ' — Announcement')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var file = (page === 'dashboard') ? 'Dashboard' : 'Form';
  var t = HtmlService.createTemplateFromFile(file);
  t.userEmail = userEmail_();
  t.isAdmin = isAdmin_(t.userEmail);
  t.focusId = id;
  return t.evaluate()
    .setTitle(APP.TITLE + ' — Royalton-Hartland')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function webAppUrl_() {
  return prop_('WEB_APP_URL', ScriptApp.getService().getUrl());
}

/* ------------------------------------------------------------ form bootstrap */

function getFormBootstrap() {
  var email = userEmail_();
  return {
    email: email,
    guessName: guessName_(email),
    locations: locations_().map(function (l) {
      return { name: l.name, notes: l.notes, leadDays: l.leadDays };
    }),
    maxFileMb: APP.MAX_FILE_MB,
    allowedExt: APP.ALLOWED_EXT,
    defaultLeadDays: APP.DEFAULT_LEAD_DAYS,
    dashboardUrl: webAppUrl_() + '?page=dashboard',
    isAdmin: isAdmin_(email)
  };
}

function guessName_(email) {
  var local = String(email).split('@')[0].replace(/[._]+/g, ' ');
  return local.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

/* ---------------------------------------------------------------- uploading */

/**
 * Just mints an id. The folder is not created until the first chunk actually
 * arrives, so people who attach nothing never touch Drive at all.
 */
function startDraft() {
  return 'draft-' + Utilities.getUuid().slice(0, 8);
}

function findDraftFolder_(draftId) {
  if (!/^draft-[0-9a-f]{8}$/i.test(String(draftId))) return null;
  var it = draftsFolder_().getFoldersByName(draftId);
  return it.hasNext() ? it.next() : null;
}

function ensureDraftFolder_(draftId) {
  if (!/^draft-[0-9a-f]{8}$/i.test(String(draftId))) {
    throw new Error('That upload session is not valid. Refresh the page and try again.');
  }
  return withRetry_(function () {
    return findDraftFolder_(draftId) || draftsFolder_().createFolder(draftId);
  }, 'ensureDraftFolder');
}

function draftFolder_(draftId) {
  var f = findDraftFolder_(draftId);
  if (!f) throw new Error('This upload session expired. Refresh the page and try again.');
  return f;
}

function chunkFolder_(draftId) {
  return withRetry_(function () {
    return childFolder_(ensureDraftFolder_(draftId), '_chunks');
  }, 'chunkFolder');
}

function uploadChunk(draftId, key, index, base64) {
  var bytes = Utilities.base64Decode(base64);
  var folder = chunkFolder_(draftId);
  withRetry_(function () {
    folder.createFile(Utilities.newBlob(bytes, 'application/octet-stream', key + '.' + index));
  }, 'uploadChunk');
  return true;
}

function finishFile(draftId, key, name, mime, totalChunks) {
  var ext = String(name).split('.').pop().toLowerCase();
  if (APP.ALLOWED_EXT.indexOf(ext) < 0) {
    throw new Error(name + ' is a .' + ext + ' file. Allowed types: ' + APP.ALLOWED_EXT.join(', ') + '.');
  }

  var chunks = chunkFolder_(draftId);
  var bytes = [];
  for (var i = 0; i < totalChunks; i++) {
    var part = withRetry_(function () {
      var it = chunks.getFilesByName(key + '.' + i);
      return it.hasNext() ? it.next() : null;
    }, 'findChunk');
    if (!part) throw new Error('Part of ' + name + ' did not arrive. Remove it and upload again.');
    bytes = bytes.concat(part.getBlob().getBytes());
    try { part.setTrashed(true); } catch (e) {}
  }

  var file = withRetry_(function () {
    return ensureDraftFolder_(draftId)
      .createFile(Utilities.newBlob(bytes, mime || 'application/octet-stream', name));
  }, 'finishFile');

  return { id: file.getId(), name: file.getName(), size: file.getSize(), url: file.getUrl() };
}

function removeDraftFile(draftId, fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
  return true;
}

/* --------------------------------------------------------------- submitting */

function submitAnnouncement(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var email = userEmail_();
    var name = String(payload.name || '').trim();
    var content = String(payload.content || '').trim();
    var picked = (payload.locations || []).filter(String);

    if (!name) throw new Error('Add your name so the poster knows who to ask about this.');
    if (!content) throw new Error('Add the announcement text.');
    if (!picked.length) throw new Error('Pick at least one place for this to be posted.');
    if (picked.indexOf('Other') > -1 && !String(payload.otherLocation || '').trim()) {
      throw new Error('You chose "Other" — tell us where it should go.');
    }

    var id = nextId_();
    var folder = moveDraftToFolder_(payload.draftId, id, name);
    var files = folder ? listFiles_(folder) : [];

    var record = {
      id: id,
      submitted: new Date(),
      name: name,
      email: email,
      content: content,
      locations: picked,
      otherLocation: String(payload.otherLocation || '').trim(),
      start: payload.start || '',
      expire: payload.expire || '',
      links: String(payload.links || '').trim(),
      files: files,
      folderUrl: folder ? folder.getUrl() : ''
    };

    // Summary PDF, saved beside the media and attached to poster emails.
    var pdfUrl = '';
    var pdfBlob = null;
    try {
      pdfBlob = buildSummaryPdf_(record);
      var pdfFile = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);
      shareInDomain_(pdfFile);
      pdfUrl = pdfFile.getUrl();
    } catch (err) {
      log_('PDF failed', id + ': ' + err.message);
    }
    record.pdfUrl = pdfUrl;

    sheet_(SHEET.SUBMISSIONS).appendRow([
      id, record.submitted, name, email, content, picked.join(', '),
      record.otherLocation, record.start, record.expire, record.links,
      files.map(function (f) { return f.name; }).join(', '),
      record.folderUrl, pdfUrl, 'Open'
    ]);

    var taskSheet = sheet_(SHEET.TASKS);
    var routing = locations_();
    picked.forEach(function (loc, i) {
      var r = routing.filter(function (x) { return x.name === loc; })[0] || { recipients: [] };
      taskSheet.appendRow([
        id + '-' + (i + 1), id, loc, r.recipients.join(', '), 'To post', '', '', ''
      ]);
    });

    notifyPosters_(record, pdfBlob);
    confirmToSubmitter_(record);
    log_('Submitted', id);

    return {
      id: id,
      folderUrl: record.folderUrl,
      pdfUrl: pdfUrl,
      notified: routedRecipientCount_(picked)
    };
  } finally {
    lock.releaseLock();
  }
}

function nextId_() {
  var p = props_();
  var year = schoolYear_().split('-')[0];
  var key = 'COUNTER_' + year;
  var n = Number(p.getProperty(key) || 0) + 1;
  p.setProperty(key, String(n));
  return 'ANN-' + year + '-' + ('000' + n).slice(-4);
}

function moveDraftToFolder_(draftId, id, name) {
  var target = yearFolder_();
  var draft = draftId ? findDraftFolder_(draftId) : null;

  if (!draft) {
    return withRetry_(function () { return target.createFolder(id + ' — ' + name); }, 'newFolder');
  }

  return withRetry_(function () {
    var chunks = draft.getFoldersByName('_chunks');
    while (chunks.hasNext()) {
      try { chunks.next().setTrashed(true); } catch (e) {}
    }
    draft.setName(id + ' — ' + name);
    if (typeof draft.moveTo === 'function') {
      draft.moveTo(target);
    } else {
      target.addFolder(draft);
      draftsFolder_().removeFolder(draft);
    }
    shareFolderInDomain_(draft);
    return draft;
  }, 'moveDraft');
}

function yearFolder_() {
  return childFolder_(rootFolder_(), schoolYear_());
}

function listFiles_(folder) {
  var out = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().slice(-4).toLowerCase() === '.pdf' && f.getName().indexOf('Summary') > -1) continue;
    shareInDomain_(f);
    out.push({
      id: f.getId(),
      name: f.getName(),
      size: f.getSize(),
      mime: f.getMimeType(),
      url: f.getUrl(),
      download: 'https://drive.google.com/uc?export=download&id=' + f.getId(),
      preview: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w400'
    });
  }
  return out;
}

function shareInDomain_(file) {
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
}
function shareFolderInDomain_(folder) {
  try { folder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
}

function routedRecipientCount_(picked) {
  var set = {};
  locations_().forEach(function (l) {
    if (picked.indexOf(l.name) > -1) l.recipients.forEach(function (r) { set[r.toLowerCase()] = 1; });
  });
  return Object.keys(set).length;
}

/* ---------------------------------------------------------------- dashboard */

function getDashboard(opts) {
  opts = opts || {};
  var me = userEmail_();
  var admin = isAdmin_(me);
  var scopeAll = admin && opts.scope === 'all';

  var subs = {};
  rows_(SHEET.SUBMISSIONS).forEach(function (s) { subs[s.ID] = s; });

  var items = rows_(SHEET.TASKS).filter(function (t) {
    if (!t['Task ID']) return false;
    if (scopeAll) return true;
    var mine = String(t.Assignees || '').toLowerCase().indexOf(me) > -1;
    return mine || (admin && !String(t.Assignees || '').trim());
  }).map(function (t) {
    var s = subs[t.ID] || {};
    return {
      taskId: t['Task ID'],
      id: t.ID,
      location: t.Location,
      status: t.Status || 'To post',
      postedBy: t['Posted By'] || '',
      postedAt: fmtDate_(t['Posted At']),
      notes: t.Notes || '',
      name: s.Name || '',
      email: s.Email || '',
      content: s.Content || '',
      otherLocation: s['Other Location'] || '',
      start: fmtDate_(s['Start Date']),
      expire: fmtDate_(s['Expire Date']),
      startRaw: s['Start Date'] ? new Date(s['Start Date']).getTime() : 0,
      submitted: fmtDate_(s.Submitted),
      submittedRaw: s.Submitted ? new Date(s.Submitted).getTime() : 0,
      links: s['Media Links'] || '',
      folderUrl: s['Folder URL'] || '',
      pdfUrl: s['PDF URL'] || '',
      fileNames: s.Files || ''
    };
  });

  items.sort(function (a, b) { return (b.startRaw || b.submittedRaw) - (a.startRaw || a.submittedRaw); });

  return {
    me: me,
    isAdmin: admin,
    scope: scopeAll ? 'all' : 'mine',
    items: items,
    formUrl: webAppUrl_()
  };
}

/** The essentials only — what a student at the microphone needs to read. */
function printRecord_(id) {
  var s = rows_(SHEET.SUBMISSIONS).filter(function (r) { return r.ID === id; })[0];
  if (!s) return { error: 'Announcement ' + (id || '(none given)') + ' was not found.' };
  return {
    id: s.ID,
    name: s.Name || '',
    email: s.Email || '',
    content: s.Content || '',
    locations: String(s.Locations || ''),
    otherLocation: s['Other Location'] || '',
    start: fmtDate_(s['Start Date']),
    expire: fmtDate_(s['Expire Date']),
    submitted: fmtDate_(s.Submitted)
  };
}

function getMedia(submissionId) {
  var s = rows_(SHEET.SUBMISSIONS).filter(function (r) { return r.ID === submissionId; })[0];
  if (!s || !s['Folder URL']) return [];
  var folderId = String(s['Folder URL']).match(/folders\/([^/?]+)/);
  if (!folderId) return [];
  return listFiles_(DriveApp.getFolderById(folderId[1]));
}

function markPosted(taskId, note) {
  return setTaskStatus_(taskId, 'Posted', note);
}

function reopenTask(taskId) {
  return setTaskStatus_(taskId, 'To post', '');
}

function setTaskStatus_(taskId, status, note) {
  var sh = sheet_(SHEET.TASKS);
  var all = rows_(SHEET.TASKS);
  var t = all.filter(function (r) { return r['Task ID'] === taskId; })[0];
  if (!t) throw new Error('That item is no longer on the board.');

  sh.getRange(t._row, colIndex_(SHEET.TASKS, 'Status')).setValue(status);
  sh.getRange(t._row, colIndex_(SHEET.TASKS, 'Posted By')).setValue(status === 'Posted' ? userEmail_() : '');
  sh.getRange(t._row, colIndex_(SHEET.TASKS, 'Posted At')).setValue(status === 'Posted' ? new Date() : '');
  if (note !== undefined) sh.getRange(t._row, colIndex_(SHEET.TASKS, 'Notes')).setValue(note || '');

  rollUpSubmissionStatus_(t.ID);
  log_(status, taskId);
  return true;
}

function rollUpSubmissionStatus_(id) {
  var tasks = rows_(SHEET.TASKS).filter(function (t) { return t.ID === id; });
  var done = tasks.every(function (t) { return t.Status === 'Posted'; });
  var some = tasks.some(function (t) { return t.Status === 'Posted'; });
  var status = done ? 'Posted' : (some ? 'Partly posted' : 'Open');

  var sh = sheet_(SHEET.SUBMISSIONS);
  var s = rows_(SHEET.SUBMISSIONS).filter(function (r) { return r.ID === id; })[0];
  if (s) sh.getRange(s._row, colIndex_(SHEET.SUBMISSIONS, 'Status')).setValue(status);
}

/** Zips every media file for one submission and returns a download link. */
function zipMedia(submissionId) {
  var files = getMedia(submissionId);
  if (!files.length) throw new Error('This announcement has no uploaded media.');

  var blobs = files.map(function (f) { return DriveApp.getFileById(f.id).getBlob(); });
  var zip = Utilities.zip(blobs, submissionId + ' media.zip');
  var folder = childFolder_(rootFolder_(), '_zips');

  var existing = folder.getFilesByName(zip.getName());
  while (existing.hasNext()) existing.next().setTrashed(true);

  var file = folder.createFile(zip);
  shareInDomain_(file);
  return 'https://drive.google.com/uc?export=download&id=' + file.getId();
}
