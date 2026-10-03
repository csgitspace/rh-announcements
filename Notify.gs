/**
 * Announcement Hub — email
 * Table layouts, fully inline styles, no <style> blocks (Gmail/Outlook safe).
 */

/**
 * One email per person, not one per location. If Wendy covers two of the
 * places the teacher checked, she gets a single email listing both.
 */
function notifyPosters_(rec, pdfBlob) {
  var routing = locations_();
  var byPerson = {};

  rec.locations.forEach(function (locName) {
    var loc = routing.filter(function (l) { return l.name === locName; })[0];
    if (!loc || !loc.recipients.length) return;
    loc.recipients.forEach(function (addr) {
      var k = addr.toLowerCase();
      byPerson[k] = byPerson[k] || { to: addr, locs: [] };
      byPerson[k].locs.push(loc);
    });
  });

  var keys = Object.keys(byPerson);
  if (!keys.length) {
    MailApp.sendEmail({
      to: adminEmails_().join(','),
      subject: '[' + rec.id + '] No recipients set for: ' + rec.locations.join(', '),
      htmlBody: emailShell_('Nobody is assigned',
        '<p style="margin:0 0 12px;">' + esc_(rec.id) + ' was submitted for <b>' + esc_(rec.locations.join(', ')) +
        '</b>, but the Routing tab has no email addresses for those locations, so no one was notified.</p>')
    });
    return;
  }

  keys.forEach(function (k) {
    var person = byPerson[k];
    var attachments = pdfBlob ? [pdfBlob] : [];
    MailApp.sendEmail({
      to: person.to,
      replyTo: rec.email || prop_('REPLY_TO'),
      subject: 'Announcement to post — ' + firstLine_(rec.content, 60) + ' [' + rec.id + ']',
      htmlBody: posterBody_(rec, person.locs),
      attachments: attachments,
      name: 'Roy-Hart Announcements'
    });
  });
}

function posterBody_(rec, locs) {
  var b = APP.BRAND;

  var locNames = locs.map(function (l) {
    return (l.name === 'Other' && rec.otherLocation) ? l.name + ' (' + rec.otherLocation + ')' : l.name;
  });

  var notes = locs.filter(function (l) { return l.notes; }).map(function (l) {
    return '<div style="margin-bottom:4px;">' + esc_(l.name) + ' — ' + esc_(l.notes) + '</div>';
  }).join('');

  var media = rec.files.length
    ? mediaBlock_(rec)
    : '<p style="margin:0 0 16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:' + b.grey + ';">No media was attached.</p>';

  var dash = webAppUrl_() + '?page=dashboard&id=' + encodeURIComponent(rec.id);
  var printUrl = webAppUrl_() + '?page=print&id=' + encodeURIComponent(rec.id);

  var body =
    '<p style="margin:0 0 18px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:' + b.grey + ';">' +
      esc_(rec.name) + ' submitted an announcement for you to post to <b style="color:#2B2B2B;">' +
      esc_(locNames.join(' and ')) + '</b>.</p>' +

    runDatesBlock_(rec) +

    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" ' +
      'style="background:' + b.tint + ';border-radius:14px;margin:0 0 18px;">' +
      '<tr><td style="padding:18px 20px;font-family:Georgia,\'Times New Roman\',serif;font-size:16px;' +
      'line-height:1.55;color:#231C2E;white-space:pre-wrap;">' + esc_(rec.content) + '</td></tr>' +
    '</table>' +

    (notes ? '<p style="margin:0 0 18px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:' +
      b.grey + ';">' + notes + '</p>' : '') +

    media +

    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:6px 0 18px;">' +
      '<tr>' +
        '<td style="background:' + b.purple + ';border-radius:999px;">' +
          '<a href="' + printUrl + '" style="display:inline-block;padding:12px 26px;' +
          'font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;color:#FFFFFF;' +
          'text-decoration:none;">Print this announcement</a></td>' +
        '<td width="10">&nbsp;</td>' +
        '<td style="border:1px solid ' + b.lavender + ';border-radius:999px;">' +
          '<a href="' + dash + '" style="display:inline-block;padding:11px 24px;' +
          'font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;color:' + b.purple + ';' +
          'text-decoration:none;">Open it in the dashboard</a></td>' +
      '</tr>' +
    '</table>' +

    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:' + b.grey + ';">' +
      'The print page is sized for a reader at the microphone — large type, run dates at the top. ' +
      'Mark it posted in the dashboard so ' + esc_(rec.name) + ' can see it went up, or just reply to this ' +
      'email to reach ' + esc_(String(rec.name).split(' ')[0]) + ' directly.</p>';

  return emailShell_(rec.id, body);
}

/** The dates, given their own hard-to-miss block. */
function runDatesBlock_(rec) {
  var b = APP.BRAND;
  var start = fmtDate_(rec.start);
  var expire = fmtDate_(rec.expire);

  function cell(label, value, muted) {
    return '<td width="50%" style="padding:14px 18px;vertical-align:top;">' +
      '<div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;letter-spacing:.09em;color:' +
        b.purple + ';margin-bottom:5px;">' + label + '</div>' +
      '<div style="font-family:Arial,Helvetica,sans-serif;font-size:19px;font-weight:bold;line-height:1.25;color:' +
        (muted ? b.grey : '#231C2E') + ';">' + esc_(value) + '</div></td>';
  }

  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" ' +
    'style="border:2px solid ' + b.purple + ';border-radius:14px;margin:0 0 18px;background:#FFFFFF;">' +
    '<tr>' +
      cell('STARTS', start || 'Right away', !start) +
      cell('ENDS', expire || 'No end date given', !expire) +
    '</tr></table>';
}

function mediaBlock_(rec) {
  var b = APP.BRAND;
  var items = rec.files.map(function (f) {
    var thumb = /image/.test(f.mime)
      ? '<img src="' + f.preview + '" width="120" alt="" style="display:block;border-radius:10px;border:1px solid #E7E1F0;">'
      : '<div style="width:120px;height:80px;background:' + b.tint + ';border-radius:10px;"></div>';
    return '<tr>' +
      '<td width="132" style="padding:8px 12px 8px 0;vertical-align:top;">' + thumb + '</td>' +
      '<td style="padding:8px 0;vertical-align:top;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#2B2B2B;">' +
        esc_(f.name) + '<div style="margin-top:2px;font-size:12px;color:' + b.grey + ';">' + kb_(f.size) + '</div>' +
        '<div style="margin-top:6px;"><a href="' + f.download + '" style="color:' + b.purple + ';font-size:13px;">Download</a>' +
        ' &nbsp;·&nbsp; <a href="' + f.url + '" style="color:' + b.purple + ';font-size:13px;">Preview</a></div>' +
      '</td></tr>';
  }).join('');

  var links = rec.links
    ? '<p style="margin:10px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:' + b.grey + ';">' +
      'Links from the submitter: ' + esc_(rec.links) + '</p>' : '';

  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:0 0 18px;">' +
    '<tr><td style="padding-bottom:6px;font-family:Arial,Helvetica,sans-serif;font-size:12px;letter-spacing:.08em;color:' +
    b.grey + ';">MEDIA</td></tr>' + items + '</table>' + links;
}

function confirmToSubmitter_(rec) {
  var b = APP.BRAND;
  var body =
    '<p style="margin:0 0 14px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#2B2B2B;">' +
      'Thanks — your announcement is with the people who post to ' + esc_(rec.locations.join(', ')) + '.</p>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" ' +
      'style="background:' + b.tint + ';border-radius:14px;margin:0 0 16px;">' +
      '<tr><td style="padding:16px 18px;font-family:Georgia,\'Times New Roman\',serif;font-size:15px;line-height:1.55;' +
      'color:#231C2E;white-space:pre-wrap;">' + esc_(rec.content) + '</td></tr></table>' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:' + b.grey + ';">' +
      'Reference ' + esc_(rec.id) + '. If something needs to change, reply to this email and mention that number.</p>';

  MailApp.sendEmail({
    to: rec.email,
    subject: 'Got it — ' + firstLine_(rec.content, 50) + ' [' + rec.id + ']',
    htmlBody: emailShell_('Submitted', body),
    name: 'Roy-Hart Announcements'
  });
}

/** 6 a.m. nudge: anything still unposted that starts today or has passed. */
function sendDailyDigest() {
  var subs = {};
  rows_(SHEET.SUBMISSIONS).forEach(function (s) { subs[s.ID] = s; });

  var today = new Date(); today.setHours(23, 59, 59);
  var byPerson = {};

  rows_(SHEET.TASKS).forEach(function (t) {
    if (!t['Task ID'] || t.Status === 'Posted') return;
    var s = subs[t.ID];
    if (!s) return;
    var start = s['Start Date'] ? new Date(s['Start Date']) : new Date(s.Submitted);
    if (start > today) return;
    if (s['Expire Date'] && new Date(s['Expire Date']) < new Date()) return;

    String(t.Assignees || '').split(',').map(function (x) { return x.trim(); }).filter(String)
      .forEach(function (addr) {
        var k = addr.toLowerCase();
        byPerson[k] = byPerson[k] || { to: addr, rows: [] };
        byPerson[k].rows.push({ id: t.ID, loc: t.Location, who: s.Name, start: fmtDate_(start), text: s.Content });
      });
  });

  var b = APP.BRAND;
  Object.keys(byPerson).forEach(function (k) {
    var p = byPerson[k];
    var list = p.rows.map(function (r) {
      return '<tr><td style="padding:10px 0;border-bottom:1px solid #E7E1F0;font-family:Arial,Helvetica,sans-serif;' +
        'font-size:14px;color:#2B2B2B;"><b>' + esc_(r.loc) + '</b> — ' + esc_(firstLine_(r.text, 70)) +
        '<div style="margin-top:3px;font-size:12px;color:' + b.grey + ';">' + esc_(r.who) + ' · starts ' +
        esc_(r.start) + ' · ' + esc_(r.id) + '</div></td></tr>';
    }).join('');

    MailApp.sendEmail({
      to: p.to,
      subject: p.rows.length + ' announcement' + (p.rows.length === 1 ? '' : 's') + ' waiting to be posted',
      name: 'Roy-Hart Announcements',
      htmlBody: emailShell_('Still waiting',
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' + list + '</table>' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:18px 0 0;">' +
        '<tr><td style="background:' + b.purple + ';border-radius:999px;"><a href="' + webAppUrl_() +
        '?page=dashboard" style="display:inline-block;padding:12px 26px;font-family:Arial,Helvetica,sans-serif;' +
        'font-size:14px;font-weight:bold;color:#FFFFFF;text-decoration:none;">Open the dashboard</a></td></tr></table>')
    });
  });
}

/** Closes out anything past its expire date so the board stays honest. */
function expireAnnouncements() {
  var sh = sheet_(SHEET.SUBMISSIONS);
  var now = new Date();
  rows_(SHEET.SUBMISSIONS).forEach(function (s) {
    if (!s['Expire Date'] || s.Status === 'Expired') return;
    if (new Date(s['Expire Date']) < now) {
      sh.getRange(s._row, colIndex_(SHEET.SUBMISSIONS, 'Status')).setValue('Expired');
    }
  });
}

/** Sweeps abandoned upload folders older than 3 days. */
function cleanUpDrafts() {
  var cutoff = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
  var it = draftsFolder_().getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getDateCreated() < cutoff) f.setTrashed(true);
  }
}

/* ------------------------------------------------------------------ helpers */

function emailShell_(eyebrow, inner) {
  var b = APP.BRAND;
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" ' +
    'style="background:#F6F5F8;padding:24px 0;"><tr><td align="center">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" ' +
    'style="max-width:600px;background:#FFFFFF;border-radius:16px;overflow:hidden;' +
    'box-shadow:0 2px 10px rgba(99,65,154,.12);">' +
      '<tr><td style="background:' + b.purple + ';padding:18px 24px;">' +
        '<img src="' + APP.LOGO + '" height="34" alt="Royalton-Hartland" style="display:block;border:0;">' +
      '</td></tr>' +
      '<tr><td style="padding:10px 24px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;' +
        'letter-spacing:.08em;color:' + b.lavender + ';">' + esc_(eyebrow).toUpperCase() + '</td></tr>' +
      '<tr><td style="padding:12px 24px 26px;">' + inner + '</td></tr>' +
      '<tr><td style="padding:14px 24px;background:' + b.tint + ';font-family:Arial,Helvetica,sans-serif;' +
        'font-size:11px;color:' + b.grey + ';">Royalton-Hartland Central School District · sent by Announcement Hub</td></tr>' +
    '</table></td></tr></table>';
}

function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function firstLine_(text, max) {
  var t = String(text || '').replace(/\s+/g, ' ').trim();
  return t.length > max ? t.slice(0, max - 1) + '…' : (t || 'Announcement');
}

function kb_(bytes) {
  var n = Number(bytes || 0);
  return n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
}
