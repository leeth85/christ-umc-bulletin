/**
 * Christ United Methodist Church — Digital Bulletin Google Sheets Integration Script
 * Automatically creates & populates tabs for:
 *   1_VisitorLog      — Anonymous page-open counter
 *   2_ConnectCards    — Guest & Member Sunday Check-Ins
 *   3_PrayerRequests  — Joys, Concerns & Sermon Reflections
 *   4_CareAndRSVP     — Event RSVPs, Volunteer Sign-Ups & Pastoral Visits
 *   [Campaign Title]  — Seasonal Gratitude / Devotional Journals (supports live edit & delete)
 */

function doPost(e) {
  if (e && e.parameter && e.parameter.type === 'special') {
    return saveSpecialRecord_(e);
  }

  var output = ContentService.createTextOutput();

  try {
    var param = (e && e.parameter) ? e.parameter : {};
    var type = param.type || '';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = '';
    var headers = [];
    var rowData = [];

    if (type === 'visitor') {
      sheetName = '1_VisitorLog';
      headers = ['Timestamp', 'Bulletin Date', 'Activity'];
      rowData = [param.date || new Date(), param.bulletinDate || '', param.content || 'Opened Bulletin'];

    } else if (type === 'connect') {
      sheetName = '2_ConnectCards';
      headers = ['Submitted At', 'Bulletin Date', 'Full Name', 'Service Attended', 'Visitor Status', 'Email', 'Phone', 'Interests / Next Steps', 'Notes'];
      rowData = [
        param.date || new Date(),
        param.bulletinDate || '',
        param.name || '',
        param.service || '',
        param.status || '',
        param.email || '',
        param.phone || '',
        param.interests || '',
        param.notes || ''
      ];

    } else if (type === 'prayer') {
      sheetName = '3_PrayerRequests';
      headers = ['Submitted At', 'Bulletin Date', 'Name', 'Sharing Preference', 'Prayer Request (Joys & Concerns)', 'Sermon Reflection', 'Scripture', 'Sermon Title'];
      rowData = [
        param.date || new Date(),
        param.bulletinDate || '',
        param.author || '',
        param.sharing || '',
        param.prayer || '',
        param.grace || '',
        param.biblePassage || '',
        param.sermonTitle || ''
      ];

    } else if (type === 'rsvp') {
      sheetName = '4_CareAndRSVP';
      headers = ['Submitted At', 'Bulletin Date', 'Request Category', 'Full Name', 'Email / Phone', 'Preferred Date & Time', 'Details'];
      rowData = [
        param.date || new Date(),
        param.bulletinDate || '',
        param.requestType || '',
        param.author || '',
        param.contact || '',
        param.schedule || '',
        param.text || ''
      ];

    } else {
      return output
        .setContent(JSON.stringify({ status: 'SUCCESS', message: 'Unknown type bypassed' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(sheetName);
      targetSheet.appendRow(headers);
      targetSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E2E8F0');
      targetSheet.setFrozenRows(1);
    }

    targetSheet.appendRow(rowData);
    var lastRow = targetSheet.getLastRow();
    targetSheet.getRange(lastRow, 1, 1, headers.length).setWrap(true).setVerticalAlignment('top');

    return output
      .setContent(JSON.stringify({ status: 'SUCCESS', sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return output
      .setContent(JSON.stringify({ status: 'ERROR', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Christ UMC Digital Bulletin Cloud Sheet Engine is active.');
}

function saveSpecialRecord_(e) {
  var p = e.parameter || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name = String(p.recordTitle || 'SeasonalJournal').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90) || 'SeasonalJournal';
    var sheet = ss.getSheetByName(name) || ss.insertSheet(name);

    var fields = [];
    try { fields = JSON.parse(p.fieldsJson || '[]'); } catch (err) { fields = []; }

    var fixed = ['Submitted At', 'Name', 'Service', 'Submission Type', 'Completed Prompts'];
    var header = sheet.getLastRow() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
      : [];
    if (header.length === 0) header = fixed.slice();

    fields.forEach(function(f) {
      if (f && f.label && header.indexOf(f.label) === -1) header.push(f.label);
    });
    if (header.indexOf('Full Summary') === -1) header.push('Full Summary');
    if (header.indexOf('Entry ID') === -1) header.push('Entry ID');

    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#FEF3C7');
    sheet.setFrozenRows(1);

    var idCol = header.indexOf('Entry ID') + 1;
    var foundRow = 0;
    var entryId = String(p.entryId || '');
    if (entryId && sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === entryId) { foundRow = i + 2; break; }
      }
    }

    if (p.action === 'delete') {
      if (foundRow) sheet.deleteRow(foundRow);
      return ContentService.createTextOutput('deleted');
    }

    var row = header.map(function() { return ''; });
    var put = function(col, val) { var k = header.indexOf(col); if (k !== -1) row[k] = val; };
    put('Submitted At', p.date || new Date());
    put('Name', p.author || '');
    put('Service', p.service || '');
    put('Submission Type', foundRow ? 'Updated' : (p.submitKind || 'Initial'));
    put('Completed Prompts', p.filled || '');
    fields.forEach(function(f) { if (f && f.label) put(f.label, f.text || ''); });
    put('Full Summary', p.content || '');
    put('Entry ID', entryId);

    if (foundRow) {
      sheet.getRange(foundRow, 1, 1, header.length).setValues([row]);
    } else {
      sheet.appendRow(row);
      foundRow = sheet.getLastRow();
    }
    sheet.getRange(foundRow, 1, 1, header.length).setWrap(true).setVerticalAlignment('top');
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
