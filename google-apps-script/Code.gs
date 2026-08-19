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

  // One-time repair: fix rows written with shifted columns after the Device column was added
  if (p.repair === '1') {
    return ContentService.createTextOutput('repaired: ' + repairActivityLog_());
  }

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

  // Free Together: friend added / removed
  if (evt === 'free' || evt === 'free_remove') {
    logFreeTogether_(sid, did, ts, evt, extra, name, roll, ua, ref);
    return ContentService.createTextOutput('ok');
  }

  // End Term Exam Timetable events go to their own sheet
  if (evt === 'exam_open' || evt === 'exam_close' || evt === 'exam_tab' || evt === 'exam_blocked_noname') {
    logExam_(sid, did, ts, evt, extra, ua, ref);
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
  var log = ensureSheet_('Activity Log', ['Session ID', 'Timestamp', 'Event', 'Name', 'Roll', 'Extra', 'User Agent', 'Referrer', 'Device']);
  prepend_(log, [sid, ts, evt, name, roll, extra, ua, ref, did]);
}

function logFreeTogether_(sid, did, ts, evt, you, friend, roll, ua, ref) {
  var sheet = ensureSheet_('Free Together', ['Session ID', 'Timestamp', 'Event', 'You', 'Friend', 'Roll', 'Device']);
  prepend_(sheet, [sid, ts, evt, you, friend, roll, did]);
}

function logExam_(sid, did, ts, evt, extra, ua, ref) {
  var student = '';
  var tab = '';
  var parts = String(extra || '').split(' :: ');
  for (var i = 0; i < parts.length; i++) {
    if (parts[i].indexOf('tab:') === 0) {
      tab = parts[i].replace('tab:', '');
    } else if (parts[i] !== '') {
      student = parts[i];
    }
  }
  var sheet = ensureSheet_('Exam Timetable', ['Session ID', 'Timestamp', 'Event', 'Student', 'Tab', 'Device']);
  prepend_(sheet, [sid, ts, evt, student, tab, did]);
}

function logSuggestion_(sid, ts, name, roll, text, ua, ref) {
  var sheet = ensureSheet_('Suggestions', ['Session ID', 'Timestamp', 'Name', 'Roll', 'Suggestion', 'User Agent', 'Referrer']);
  prepend_(sheet, [sid, ts, name, roll, text, ua, ref]);
}

function repairActivityLog_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activity Log');
  if (!sh) return 0;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var colCount = sh.getLastColumn();
  var range = sh.getRange(1, 1, lastRow, colCount);
  var rows = range.getValues();
  // Header layout: Session ID | Timestamp | Event | Name | Roll | Extra | User Agent | Referrer | Device
  // Misplaced rows (written with Device 2nd): [sid, did, ts, evt, name, roll, extra, ua, ref]
  // Correct layout:                                   [sid, ts, evt, name, roll, extra, ua, ref, did]
  var fixed = [rows[0]];
  var repaired = 0;
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    var hasTs = /^\d{2}\/\d{2}\/\d{2}/.test(String(row[1] || ''));
    var hasTs2 = /^\d{2}\/\d{2}\/\d{2}/.test(String(row[2] || ''));
    if (!hasTs && hasTs2) {
      fixed.push([
        row[0], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[1]
      ]);
      repaired++;
    } else {
      fixed.push(row);
    }
  }
  if (repaired) {
    sh.getRange(1, 1, fixed.length, fixed[0].length).setValues(fixed);
  }
  return repaired;
}

function doRepair(e) {
  return ContentService.createTextOutput('repaired: ' + repairActivityLog_());
}
