/**
 * Christ United Methodist Church — Digital Bulletin & Live Attendance Google Sheets Engine
 * Automatically creates & populates tabs for:
 *   0_MemberRoster    — Saved member/regular names for 1-tap paper attendance entry
 *   1_Attendance      — Unified Sunday Check-Ins (One-Tap Phone App + Paper Pew Sheet, deduplicated!)
 *   1_Headcounts      — Optional Usher Total Seated Headcounts per Service
 *   2_ConnectCards    — Full Visitor / Ministry Connect Cards
 *   3_PrayerRequests  — Joys, Concerns & Sermon Reflections
 *   4_CareAndRSVP     — Event RSVPs, Volunteer Sign-Ups & Pastoral Visits
 *   5_VisitorLog      — Anonymous bulletin open log
 *   [Campaign Title]  — Seasonal Gratitude / Devotional Journals
 */

function doPost(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (p.type === 'special') return saveSpecialRecord_(e);
  if (p.type === 'attendance') return saveAttendance_(p);
  if (p.type === 'headcount') return saveHeadcount_(p);

  var output = ContentService.createTextOutput();

  try {
    var type = p.type || '';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = '';
    var headers = [];
    var rowData = [];

    if (type === 'visitor') {
      sheetName = '5_VisitorLog';
      headers = ['Timestamp', 'Bulletin Date', 'Activity'];
      rowData = [p.date || new Date(), p.bulletinDate || '', p.content || 'Opened Bulletin'];

    } else if (type === 'connect') {
      sheetName = '2_ConnectCards';
      headers = ['Submitted At', 'Bulletin Date', 'Full Name', 'Service Attended', 'Visitor Status', 'Email', 'Phone', 'Interests / Next Steps', 'Notes'];
      rowData = [
        p.date || new Date(), p.bulletinDate || '', p.name || '', p.service || '',
        p.status || '', p.email || '', p.phone || '', p.interests || '', p.notes || ''
      ];

    } else if (type === 'prayer') {
      sheetName = '3_PrayerRequests';
      headers = ['Submitted At', 'Bulletin Date', 'Name', 'Sharing Preference', 'Prayer Request (Joys & Concerns)', 'Sermon Reflection', 'Scripture', 'Sermon Title'];
      rowData = [
        p.date || new Date(), p.bulletinDate || '', p.author || '', p.sharing || '',
        p.prayer || '', p.grace || '', p.biblePassage || '', p.sermonTitle || ''
      ];

    } else if (type === 'rsvp') {
      sheetName = '4_CareAndRSVP';
      headers = ['Submitted At', 'Bulletin Date', 'Request Category', 'Full Name', 'Email / Phone', 'Preferred Date & Time', 'Details'];
      rowData = [
        p.date || new Date(), p.bulletinDate || '', p.requestType || '', p.author || '',
        p.contact || '', p.schedule || '', p.text || ''
      ];

    } else {
      return output
        .setContent(JSON.stringify({ status: 'SUCCESS', message: 'Unknown type bypassed' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var targetSheet = ensureSheetWithHeaders_(ss, sheetName, headers, '#E2E8F0');
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

/**
 * Two-way GET endpoint for Admin Attendance Dashboard
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  if (params.action === 'stats') {
    var targetDate = String(params.date || '').trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var records = [];
    var attSheet = ss.getSheetByName('1_Attendance');
    if (attSheet && attSheet.getLastRow() > 1) {
      var rows = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, 12).getDisplayValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (!targetDate || r[0] === targetDate) {
          records.push({
            bulletinDate: r[0],
            service: r[1],
            primaryName: r[2],
            attendees: r[3],
            count: Number(r[4]) || 1,
            status: r[5],
            source: r[6],
            phone: r[7],
            email: r[8],
            address: r[9],
            submittedAt: r[10],
            checkInId: r[11]
          });
        }
      }
    }

    var roster = [];
    var rosterSheet = ss.getSheetByName('0_MemberRoster');
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      var rVals = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 1).getDisplayValues();
      for (var j = 0; j < rVals.length; j++) {
        if (rVals[j][0]) roster.push(rVals[j][0]);
      }
    }

    var headcounts = {};
    var hcSheet = ss.getSheetByName('1_Headcounts');
    if (hcSheet && hcSheet.getLastRow() > 1) {
      var hVals = hcSheet.getRange(2, 1, hcSheet.getLastRow() - 1, 3).getDisplayValues();
      for (var k = 0; k < hVals.length; k++) {
        if (hVals[k][0] === targetDate) {
          try { headcounts = JSON.parse(hVals[k][2] || '{}'); } catch (err) {}
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'SUCCESS', date: targetDate, records: records, roster: roster, headcounts: headcounts }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('Christ UMC Digital Bulletin & Live Attendance Engine is active.');
}

function ensureSheetWithHeaders_(ss, name, headers, bg) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(bg || '#E2E8F0');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveAttendance_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var headers = ['Bulletin Date', 'Service', 'Primary Name', 'Attendees Present', 'Count', 'Status', 'Source', 'Phone', 'Email', 'Address', 'Submitted At', 'CheckIn ID'];
    var sheet = ensureSheetWithHeaders_(ss, '1_Attendance', headers, '#DCFCE7');
    var checkInId = String(p.checkInId || '');
    var foundRow = 0;

    if (checkInId && sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, 12, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (ids[i][0] === checkInId) { foundRow = i + 2; break; }
      }
    }

    if (p.action === 'delete') {
      if (foundRow) sheet.deleteRow(foundRow);
      return ContentService.createTextOutput('deleted');
    }

    var cnt = Math.max(1, parseInt(p.count, 10) || 1);
    var row = [
      p.bulletinDate || '',
      p.service || '',
      p.primaryName || '',
      p.attendees || p.primaryName || '',
      cnt,
      p.status || 'Member',
      p.source || 'App',
      p.phone || '',
      p.email || '',
      p.address || '',
      p.date || new Date(),
      checkInId
    ];

    if (foundRow) {
      sheet.getRange(foundRow, 1, 1, headers.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    if (p.saveRoster === '1' && p.primaryName) {
      var rSheet = ensureSheetWithHeaders_(ss, '0_MemberRoster', ['Household / Member Name', 'Usual Service', 'Last Updated'], '#F1F5F9');
      var label = cnt > 1 && p.source === 'Paper' ? (p.primaryName + ' (' + cnt + ')') : p.primaryName;
      var exists = false;
      if (rSheet.getLastRow() > 1) {
        var existingNames = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 1).getDisplayValues();
        for (var m = 0; m < existingNames.length; m++) {
          if (existingNames[m][0].toLowerCase() === label.toLowerCase()) { exists = true; break; }
        }
      }
      if (!exists) rSheet.appendRow([label, p.service || '', p.date || new Date()]);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

function saveHeadcount_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ensureSheetWithHeaders_(ss, '1_Headcounts', ['Bulletin Date', 'Updated At', 'Headcounts JSON'], '#E0F2FE');
    var bDate = String(p.bulletinDate || '');
    var foundRow = 0;
    if (bDate && sheet.getLastRow() > 1) {
      var dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (var i = 0; i < dates.length; i++) {
        if (dates[i][0] === bDate) { foundRow = i + 2; break; }
      }
    }
    var row = [bDate, p.date || new Date(), p.headcountsJson || '{}'];
    if (foundRow) sheet.getRange(foundRow, 1, 1, 3).setValues([row]);
    else sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
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
