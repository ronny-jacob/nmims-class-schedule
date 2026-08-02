function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['Timestamp', 'Message']);
  }
  return sh;
}

function doGet(e) {
  // Button-click easter egg events go to a dedicated sheet
  if (e.parameter.evt === 'click' || e.parameter.name === '__do_not_click__') {
    var clicks = getOrCreateSheet_('Button Clicks');
    clicks.appendRow([
      e.parameter.ts || new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      e.parameter.extra || ''
    ]);
    return ContentService.createTextOutput('ok');
  }

  // Existing access logging (unchanged)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([
    e.parameter.name,
    e.parameter.roll,
    e.parameter.ts,
    e.parameter.ua,
    e.parameter.ref
  ]);
  return ContentService.createTextOutput('ok');
}
