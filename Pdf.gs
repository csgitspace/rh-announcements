/**
 * Announcement Hub — one-page PDF summary
 *
 * Google's HTML-to-PDF converter is a simple renderer: tables and inline
 * styles only, no flexbox or grid. Images are embedded as data URIs when they
 * are small enough; if that trips the converter we retry without them.
 */

function buildSummaryPdf_(rec) {
  var name = rec.id + ' — Announcement Summary.pdf';
  try {
    return htmlToPdf_(summaryHtml_(rec, true), name);
  } catch (err) {
    return htmlToPdf_(summaryHtml_(rec, false), name);
  }
}

function htmlToPdf_(html, name) {
  return Utilities.newBlob(html, MimeType.HTML, name.replace(/\.pdf$/, '.html'))
    .getAs(MimeType.PDF).setName(name);
}

function summaryHtml_(rec, withImages) {
  var b = APP.BRAND;
  var routing = locations_();

  var locRows = rec.locations.map(function (n) {
    var l = routing.filter(function (x) { return x.name === n; })[0] || { notes: '' };
    var extra = (n === 'Other' && rec.otherLocation) ? ' — ' + esc_(rec.otherLocation) : '';
    return '<tr>' +
      '<td style="padding:7px 12px;border-bottom:1px solid #E7E1F0;font-size:12px;width:45%;">' +
        '<b>' + esc_(n) + '</b>' + extra + '</td>' +
      '<td style="padding:7px 12px;border-bottom:1px solid #E7E1F0;font-size:11px;color:' + b.grey + ';">' +
        esc_(l.notes) + '</td></tr>';
  }).join('');

  var imgHtml = '';
  if (withImages) {
    var shots = rec.files.filter(function (f) { return /image/.test(f.mime) && f.size < 2500000; }).slice(0, 2);
    imgHtml = shots.map(function (f) {
      var blob = DriveApp.getFileById(f.id).getBlob();
      var uri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
      return '<div style="margin:0 0 10px;"><img src="' + uri + '" style="max-width:420px;border:1px solid #E7E1F0;">' +
        '<div style="font-size:10px;color:' + b.grey + ';margin-top:3px;">' + esc_(f.name) + '</div></div>';
    }).join('');
  }

  var fileList = rec.files.length
    ? rec.files.map(function (f) {
        return '<div style="font-size:11px;margin-bottom:3px;">' + esc_(f.name) +
          ' <span style="color:' + b.grey + ';">(' + kb_(f.size) + ')</span></div>';
      }).join('')
    : '<div style="font-size:11px;color:' + b.grey + ';">No media attached.</div>';

  return '<html><body style="margin:0;font-family:Arial,Helvetica,sans-serif;color:#231C2E;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:' + b.purple + ';">' +
      '<tr><td style="padding:22px 28px;">' +
        '<div style="color:#FFFFFF;font-size:22px;font-weight:bold;">Announcement summary</div>' +
        '<div style="color:' + b.lavender + ';font-size:12px;margin-top:4px;">' + esc_(rec.id) +
        ' · Royalton-Hartland Central School District</div>' +
      '</td></tr></table>' +

    '<table width="100%" cellpadding="0" cellspacing="0" style="margin:22px 0 0;">' +
      '<tr><td style="padding:0 28px;">' +

        '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px;">' +
          row_('Submitted by', esc_(rec.name) + ' &lt;' + esc_(rec.email) + '&gt;') +
          row_('Submitted', esc_(fmtDate_(rec.submitted))) +
          row_('Runs', esc_(fmtDate_(rec.start) || 'As soon as possible') +
               (rec.expire ? ' through ' + esc_(fmtDate_(rec.expire)) : '')) +
        '</table>' +

        '<div style="font-size:11px;letter-spacing:.08em;color:' + b.grey + ';margin-bottom:6px;">ANNOUNCEMENT</div>' +
        '<table width="100%" cellpadding="0" cellspacing="0" style="background:' + b.tint + ';margin-bottom:20px;">' +
          '<tr><td style="padding:16px 18px;font-family:Georgia,serif;font-size:13px;line-height:1.6;">' +
            esc_(rec.content).replace(/\n/g, '<br>') + '</td></tr></table>' +

        '<div style="font-size:11px;letter-spacing:.08em;color:' + b.grey + ';margin-bottom:6px;">WHERE IT GOES</div>' +
        '<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #E7E1F0;margin-bottom:20px;">' +
          locRows + '</table>' +

        '<div style="font-size:11px;letter-spacing:.08em;color:' + b.grey + ';margin-bottom:6px;">MEDIA</div>' +
        imgHtml + fileList +
        (rec.links ? '<div style="font-size:11px;margin-top:6px;">Links: ' + esc_(rec.links) + '</div>' : '') +
        (rec.folderUrl ? '<div style="font-size:11px;margin-top:6px;">Files: ' + esc_(rec.folderUrl) + '</div>' : '') +

        '<div style="margin-top:26px;padding-top:10px;border-top:2px solid ' + b.purple + ';font-size:10px;color:' +
          b.grey + ';">Questions about this announcement go to ' + esc_(rec.email) + '.</div>' +
      '</td></tr></table></body></html>';
}

function row_(label, value) {
  return '<tr>' +
    '<td width="130" style="padding:4px 0;font-size:11px;color:' + APP.BRAND.grey + ';vertical-align:top;">' +
      label + '</td>' +
    '<td style="padding:4px 0;font-size:12px;">' + value + '</td></tr>';
}
