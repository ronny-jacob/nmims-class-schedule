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
  }
  return sh;
}

function prepend_(sh, values) {
  // Insert newest rows at the top, just under the header
  sh.insertRowBefore(2);
  sh.getRange(2, 1, 1, values.length).setValues([values]);
}

function now_() {
  return new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
}

function doGet(e) {
  var p = e.parameter;

  if (p.debug === 'sheets') {
    var out = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var last = sh.getLastRow();
      var rows = last > 1 ? sh.getRange(1, 1, Math.min(last, 6), sh.getLastColumn()).getValues() : [];
      out.push('SHEET: ' + sh.getName() + ' (rows=' + last + ', cols=' + sh.getLastColumn() + ')');
      for (var r = 0; r < rows.length; r++) out.push('  ' + rows[r].join(' | '));
    }
    return ContentService.createTextOutput(out.join('\n'));
  }

  var sid = p.sid || '';
  var evt = p.evt || '';
  var name = p.name || '';
  var roll = p.roll || '';
  var extra = p.extra || '';
  var ts = p.ts || now_();
  var ua = p.ua || '';
  var ref = p.ref || '';

  // Page open events
  if (evt === 'pageview' || name === '__page_view__') {
    logActivity_(sid, ts, 'pageview', '', '', '', ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Button-click easter egg events
  if (evt === 'click' || name === '__do_not_click__') {
    logActivity_(sid, ts, 'click', '', '', extra, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Name search events
  if (evt === 'search') {
    logActivity_(sid, ts, 'search', '', '', extra, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // Student selection / access logging
  logActivity_(sid, ts, 'select', name, roll, extra, ua, ref);
  return ContentService.createTextOutput('ok');
}

function logActivity_(sid, ts, evt, name, roll, extra, ua, ref) {
  var log = ensureSheet_('Activity Log', ['Session ID', 'Timestamp', 'Event', 'Name', 'Roll', 'Extra', 'User Agent', 'Referrer']);
  prepend_(log, [sid, ts, evt, name, roll, extra, ua, ref]);
}

function debugListSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var last = sh.getLastRow();
    out.push('SHEET: ' + sh.getName() + ' (rows=' + last + ', cols=' + sh.getLastColumn() + ')');
  }
  return out.join('\n');
}
