function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return sh;
  }
  // Legacy sheets whose first row is data (not a header) get a header inserted
  var first = sh.getRange(1, 1).getValue();
  if (String(first) !== headers[0]) {
    sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return sh;
  }
  // Append any header columns added after the sheet was created
  var existing = sh.getRange(1, 1, 1, sh.getLastColumn() || 1).getValues()[0];
  var add = [];
  for (var i = 0; i < headers.length; i++) {
    if (existing.indexOf(headers[i]) === -1) add.push(headers[i]);
  }
  if (add.length) sh.getRange(1, (sh.getLastColumn() || 0) + 1, 1, add.length).setValues([add]);
  return sh;
}

function prepend_(sh, values) {
  // Insert newest rows at the top, just under the header
  sh.insertRowBefore(2);
  sh.getRange(2, 1, 1, values.length).setValues([values]);
}

function formatTs_(ts) {
  var d = ts ? new Date(ts) : new Date();
  if (isNaN(d.getTime())) d = new Date();
  var parts = {};
  new Intl.DateTimeFormat('en-IN', {
    timeZone: 'Asia/Kolkata',
    day: '2-digit', month: '2-digit', year: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    hour12: true
  }).formatToParts(d).forEach(function(p) {
    if (p.type !== 'literal') parts[p.type] = p.value;
  });
  return parts.day + '/' + parts.month + '/' + parts.year +
    ' ' + parts.hour + ':' + parts.minute + ':' + parts.second + ' ' + parts.dayPeriod.toUpperCase();
}

function doGet(e) {
  var p = e.parameter;

  var sid = p.s || '';
  var did = p.d || '';
  var evt = p.evt || '';
  var name = p.name || '';
  var roll = p.roll || '';
  var extra = p.extra || '';
  var ts = formatTs_(p.ts);
  var ua = p.ua || '';
  var ref = p.ref || '';

  // Page open events
  if (evt === 'pageview' || name === '__page_view__') {
    logActivity_(sid, did, ts, 'pageview', '', '', '', ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Button-click easter egg events
  if (evt === 'click' || name === '__do_not_click__') {
    logActivity_(sid, did, ts, 'click', '', '', extra, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Name search events
  if (evt === 'search') {
    logActivity_(sid, did, ts, 'search', '', '', extra, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // PWA install events
  if (evt === 'install' || evt === 'installed' || evt === 'install_dismissed' ||
      evt === 'install_guide' || evt === 'installed_open') {
    logActivity_(sid, did, ts, evt, '', '', extra, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Suggestion box opened (may not lead to a submission)
  if (evt === 'suggest_open') {
    logSuggestion_(sid, ts, name, roll, '[opened]', ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Suggestion box closed/cancelled without submitting
  if (evt === 'suggest_cancel') {
    logSuggestion_(sid, ts, name, roll, '[opened, closed without submitting]', ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Suggestion box closed while blocked by the send cooldown
  if (evt === 'suggest_cancel_cooldown') {
    logSuggestion_(sid, ts, name, roll, '[closed during send cooldown]', ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Suggestion box submissions
  if (evt === 'suggest') {
    logSuggestion_(sid, ts, name, roll, extra, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Custom events (e.g. fac_week) keep their real type instead of collapsing into 'select'
  if (evt && evt !== 'select') {
    logActivity_(sid, did, ts, evt, name, roll, extra, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Student selection / access logging
  logActivity_(sid, did, ts, 'select', name, roll, extra, ua, ref);
  return ContentService.createTextOutput('ok');
}

function logActivity_(sid, did, ts, evt, name, roll, extra, ua, ref) {
  var log = ensureSheet_('Activity Log', ['Session ID', 'Device', 'Timestamp', 'Event', 'Name', 'Roll', 'Extra', 'User Agent', 'Referrer']);
  prepend_(log, [sid, did, ts, evt, name, roll, extra, ua, ref]);
}

function logSuggestion_(sid, ts, name, roll, text, ua, ref) {
  var sheet = ensureSheet_('Suggestions', ['Session ID', 'Timestamp', 'Name', 'Roll', 'Suggestion', 'User Agent', 'Referrer']);
  prepend_(sheet, [sid, ts, name, roll, text, ua, ref]);
}
