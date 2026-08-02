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
  // TEMP DEBUG: list sheet names
  if (e.parameter.debug === 'sheets') {
    var dbg = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (sh) { return sh.getName(); });
    return ContentService.createTextOutput(JSON.stringify(dbg));
  }

  // Page open events
  if (e.parameter.evt === 'pageview' || e.parameter.name === '__page_view__') {    var views = ensureSheet_('Page Views', ['Timestamp', 'User Agent', 'Referrer']);
    prepend_(views, [e.parameter.ts || now_(), e.parameter.ua || '', e.parameter.ref || '']);
    return ContentService.createTextOutput('ok');
  }

  // Button-click easter egg events
  if (e.parameter.evt === 'click' || e.parameter.name === '__do_not_click__') {
    var clicks = ensureSheet_('Button Clicks', ['Timestamp', 'Message']);
    prepend_(clicks, [e.parameter.ts || now_(), e.parameter.extra || '']);
    return ContentService.createTextOutput('ok');
  }

  // Student access logging -> fixed "Access Log" sheet (no more getActiveSheet())
  var access = ensureSheet_('Access Log', ['Name', 'Roll', 'Timestamp', 'User Agent', 'Referrer']);
  prepend_(access, [
    e.parameter.name || '',
    e.parameter.roll || '',
    e.parameter.ts || now_(),
    e.parameter.ua || '',
    e.parameter.ref || ''
  ]);
  return ContentService.createTextOutput('ok');
}
