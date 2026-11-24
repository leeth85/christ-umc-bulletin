/**
 * Christ United Methodist Church — Digital Bulletin, Attendance, Welcoming & Prayer Auto-Email Engine
 * Automatically creates & populates tabs for:
 *   0_MemberRoster    — Saved member/regular names for 1-tap paper attendance entry
 *   1_Attendance      — Unified Sunday Check-Ins (Adults, Young Adults, Youth, Children, App + Paper)
 *   1_Headcounts      — Optional Usher Total Seated Headcounts per Service
 *   2_ConnectCards    — Full Visitor / Ministry Connect Cards
 *   3_PrayerRequests  — Joys, Concerns & Sermon Reflections
 *   4_CareAndRSVP     — Event RSVPs, Volunteer Sign-Ups & Pastoral Visits
 *   5_VisitorLog      — Anonymous bulletin open log
 *   [Campaign Title]  — Seasonal Gratitude / Devotional Journals
 *
 * Built-in Email Automation (via MailApp + ScriptApp Triggers):
 *   1. Instant Welcome Email sent automatically to 1st-Time / Returning Guests who share their email
 *   2. Instant Pastor / Welcome Team Email Alert on new Connect Cards, 1st-Time Guest Check-Ins & Care/Event RSVPs
 *   3. Instant Prayer Request Routing (Confidential -> Pastors Only; Shared -> Pastors + Prayer Team)
 *   4. Scheduled Weekly Sunday Post-Worship Attendance Summary Email
 */

var ATT_HEADERS = [
  'Bulletin Date', 'Service', 'Primary Name', 'Attendees Present',
  'Total Count', 'Adults', 'Young Adults', 'Youth', 'Children',
  'Status', 'Source', 'Phone', 'Email', 'Address', 'Submitted At', 'CheckIn ID'
];

var DEFAULT_WELCOME_BODY =
  'Thank you so much for worshiping with us at Christ United Methodist Church! It was truly a blessing to have you with us.\n\n' +
  'At Christ UMC, our mission is simple: Gather · Grow · Go. Whether you are exploring faith, looking for a church home, or new to the Rockford area, there is a place for you and your family here.\n\n' +
  'If you have any questions or would enjoy grabbing a cup of coffee with one of our pastors, simply reply to this email or reach our church office at (815) 399-5910.\n\n' +
  'Grace and peace,\nPastoral Team · Christ United Methodist Church\nwww.christumc.cc';

function doPost(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (p.type === 'special') return saveSpecialRecord_(e);
  if (p.type === 'attendance') return saveAttendance_(p);
  if (p.type === 'headcount') return saveHeadcount_(p);
  if (p.type === 'email_config') return configureWeeklyEmail_(p);
  if (p.type === 'automation_config') return configureAutomationSettings_(p);
  if (p.type === 'send_report_now') return sendReportNow_(p);

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

    // Trigger instant email notifications & welcoming automation (non-blocking if email fails)
    try {
      if (type === 'connect') handleConnectNotifications_(ss, p);
      else if (type === 'prayer') handlePrayerNotifications_(ss, p);
      else if (type === 'rsvp') handleRsvpNotifications_(ss, p);
    } catch (mailErr) {
      Logger.log('Notification email error: ' + mailErr);
    }

    return output
      .setContent(JSON.stringify({ status: 'SUCCESS', sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return output
      .setContent(JSON.stringify({ status: 'ERROR', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  if (params.action === 'stats') {
    var targetDate = String(params.date || '').trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var records = readAttendanceRecords_(ss, targetDate);

    var roster = [];
    var rosterSheet = ss.getSheetByName('0_MemberRoster');
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      var rVals = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 1).getDisplayValues();
      for (var j = 0; j < rVals.length; j++) {
        if (rVals[j][0]) roster.push(rVals[j][0]);
      }
    }

    var headcounts = readHeadcountsForDate_(ss, targetDate);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'SUCCESS', date: targetDate, records: records, roster: roster, headcounts: headcounts }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('Christ UMC Digital Bulletin, Attendance & Auto-Email Engine is active.');
}

function escHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function resolveSetting_(paramVal, propKey, fallback) {
  var v = String(paramVal !== undefined && paramVal !== null ? paramVal : '').trim();
  if (v !== '') return v;
  var saved = PropertiesService.getScriptProperties().getProperty(propKey);
  if (saved !== null && String(saved).trim() !== '') return String(saved).trim();
  return fallback || '';
}

function mergeEmailLists_(listA, listB) {
  var raw = (String(listA || '') + ',' + String(listB || '')).split(/[,;]+/);
  var seen = {};
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var em = raw[i].trim();
    if (em && em.indexOf('@') > 0 && !seen[em.toLowerCase()]) {
      seen[em.toLowerCase()] = true;
      out.push(em);
    }
  }
  return out.join(', ');
}

function firstValidEmail_(listStr) {
  var parts = String(listStr || '').split(/[,;]+/);
  for (var i = 0; i < parts.length; i++) {
    var em = parts[i].trim();
    if (em && em.indexOf('@') > 0) return em;
  }
  return '';
}

function configureAutomationSettings_(p) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('CHURCH_NAME', String(p.churchName || 'Christ United Methodist Church').trim());
  props.setProperty('WELCOME_ALERT_EMAILS', String(p.welcomeAlertEmails || '').trim());
  props.setProperty('AUTO_WELCOME_GUEST', String(p.autoWelcomeGuest || '1').trim());
  props.setProperty('WELCOME_EMAIL_SUBJECT', String(p.welcomeEmailSubject || '').trim());
  props.setProperty('WELCOME_EMAIL_BODY', String(p.welcomeEmailBody || '').trim());
  props.setProperty('PASTOR_PRAYER_EMAILS', String(p.pastorPrayerEmails || '').trim());
  props.setProperty('PRAYER_TEAM_EMAILS', String(p.prayerTeamEmails || '').trim());
  return ContentService.createTextOutput('saved');
}

/**
 * 1. Connect Card: Alert Pastors/Welcoming Team + Auto-Welcome Email to New Guest
 */
function handleConnectNotifications_(ss, p) {
  var churchName = resolveSetting_(p.churchName, 'CHURCH_NAME', 'Christ United Methodist Church');
  var alertEmails = mergeEmailLists_(resolveSetting_(p.welcomeAlertEmails, 'WELCOME_ALERT_EMAILS', ''), '');
  var status = String(p.status || 'Visitor').trim();
  var isGuest = /guest/i.test(status);

  if (alertEmails) {
    var badgeBg = isGuest ? '#FFF1F2' : '#EFF6FF';
    var badgeCol = isGuest ? '#BE123C' : '#1D4ED8';
    var subjPrefix = isGuest ? '🌟 [New Guest Connect Card] ' : '🪪 [Connect Card] ';
    var subject = subjPrefix + (p.name || 'Visitor') + ' — ' + (p.service || churchName);

    var html =
      '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:580px;margin:0 auto;color:#1E293B;line-height:1.6;">' +
        '<div style="background:#1E1F4B;color:#fff;padding:18px 22px;border-radius:14px 14px 0 0;">' +
          '<div style="font-size:11px;letter-spacing:1.3px;text-transform:uppercase;opacity:0.85;">' + escHtml_(churchName) + ' · Welcoming &amp; Connect</div>' +
          '<h2 style="margin:4px 0 0;font-size:20px;">' + escHtml_(p.name || 'New Submission') + '</h2>' +
        '</div>' +
        '<div style="border:1px solid #E2E8F0;border-top:none;padding:20px 22px;border-radius:0 0 14px 14px;background:#FFFFFF;">' +
          '<span style="display:inline-block;background:' + badgeBg + ';color:' + badgeCol + ';font-weight:700;font-size:12px;padding:4px 10px;border-radius:999px;margin-bottom:14px;">' +
            escHtml_(status) + ' · ' + escHtml_(p.service || 'Sunday Worship') +
          '</span>' +
          '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:14px;">' +
            '<tr><td style="padding:6px 0;color:#64748B;width:130px;"><b>Bulletin Date:</b></td><td style="padding:6px 0;">' + escHtml_(p.bulletinDate || p.date || '') + '</td></tr>' +
            '<tr><td style="padding:6px 0;color:#64748B;"><b>Email:</b></td><td style="padding:6px 0;">' + (p.email ? '<a href="mailto:' + escHtml_(p.email) + '">' + escHtml_(p.email) + '</a>' : '—') + '</td></tr>' +
            '<tr><td style="padding:6px 0;color:#64748B;"><b>Phone:</b></td><td style="padding:6px 0;">' + (p.phone ? '<a href="tel:' + escHtml_(p.phone) + '">' + escHtml_(p.phone) + '</a>' : '—') + '</td></tr>' +
            '<tr><td style="padding:6px 0;color:#64748B;vertical-align:top;"><b>Next Steps:</b></td><td style="padding:6px 0;">' + escHtml_(p.interests || 'None selected') + '</td></tr>' +
            '<tr><td style="padding:6px 0;color:#64748B;vertical-align:top;"><b>Comments:</b></td><td style="padding:6px 0;white-space:pre-wrap;">' + escHtml_(p.notes || '—') + '</td></tr>' +
          '</table>' +
          '<a href="' + ss.getUrl() + '" style="display:inline-block;background:#1E1F4B;color:#fff;text-decoration:none;padding:10px 16px;border-radius:9px;font-weight:700;font-size:13px;">Open 2_ConnectCards in Google Sheets</a>' +
        '</div>' +
      '</div>';

    var mailOpts = { to: alertEmails, subject: subject, htmlBody: html };
    if (p.email && p.email.indexOf('@') > 0) mailOpts.replyTo = String(p.email).trim();
    MailApp.sendEmail(mailOpts);
  }

  // Send Automated Warm Welcome Email to Guest if enabled & guest shared an email
  var autoWelcome = resolveSetting_(p.autoWelcomeGuest, 'AUTO_WELCOME_GUEST', '1');
  var guestEmail = String(p.email || '').trim();
  if (isGuest && autoWelcome === '1' && guestEmail && guestEmail.indexOf('@') > 0) {
    sendGuestWelcomeEmail_(guestEmail, p.name, churchName, alertEmails, p);
  }
}

function sendGuestWelcomeEmail_(guestEmail, guestName, churchName, replyToList, p) {
  var subj = resolveSetting_(p.welcomeEmailSubject, 'WELCOME_EMAIL_SUBJECT', 'Welcome to ' + churchName + '!');
  var bodyRaw = resolveSetting_(p.welcomeEmailBody, 'WELCOME_EMAIL_BODY', DEFAULT_WELCOME_BODY);
  var firstName = String(guestName || '').trim().split(/\s+/)[0] || 'Friend';
  var bodyHtml = escHtml_(bodyRaw).replace(/\r?\n/g, '<br>');

  var html =
    '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:580px;margin:0 auto;color:#1E293B;line-height:1.7;">' +
      '<div style="background:#1E1F4B;color:#fff;padding:24px;border-radius:16px 16px 0 0;text-align:center;">' +
        '<div style="font-size:11px;letter-spacing:1.8px;text-transform:uppercase;opacity:0.85;">Gather · Grow · Go</div>' +
        '<h2 style="margin:6px 0 0;font-size:22px;">' + escHtml_(churchName) + '</h2>' +
      '</div>' +
      '<div style="border:1px solid #E2E8F0;border-top:none;padding:24px;border-radius:0 0 16px 16px;background:#FFFFFF;font-size:15px;">' +
        '<p style="margin-top:0;">Dear <b>' + escHtml_(firstName) + '</b>,</p>' +
        '<div>' + bodyHtml + '</div>' +
      '</div>' +
    '</div>';

  var opts = {
    to: guestEmail,
    name: churchName,
    subject: subj,
    htmlBody: html
  };
  var replyEm = firstValidEmail_(replyToList);
  if (replyEm) opts.replyTo = replyEm;
  MailApp.sendEmail(opts);
}

/**
 * 2. Prayer Request: Respect Confidential (Pastors Only) vs Pastors & Prayer Team
 */
function handlePrayerNotifications_(ss, p) {
  var churchName = resolveSetting_(p.churchName, 'CHURCH_NAME', 'Christ United Methodist Church');
  var pastorEmails = resolveSetting_(p.pastorPrayerEmails, 'PASTOR_PRAYER_EMAILS', '');
  var teamEmails = resolveSetting_(p.prayerTeamEmails, 'PRAYER_TEAM_EMAILS', '');
  var sharing = String(p.sharing || 'Pastors & Prayer Team').trim();
  var isConfidential = /confidential/i.test(sharing);
  var isPraise = /joy|praise/i.test(sharing);

  var recipients = isConfidential
    ? mergeEmailLists_(pastorEmails, '')
    : mergeEmailLists_(pastorEmails, teamEmails);
  if (!recipients) return;

  var headerBg = isConfidential ? '#881337' : (isPraise ? '#15803D' : '#1E1F4B');
  var tagPrefix = isConfidential
    ? '🔒 [CONFIDENTIAL — Pastors Only] '
    : (isPraise ? '🙌 [Joy & Praise Report] ' : '🙏 [Prayer Request] ');
  var subject = tagPrefix + (p.author || 'Anonymous') + ' — ' + churchName;

  var html =
    '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:580px;margin:0 auto;color:#1E293B;line-height:1.6;">' +
      '<div style="background:' + headerBg + ';color:#fff;padding:18px 22px;border-radius:14px 14px 0 0;">' +
        '<div style="font-size:11px;letter-spacing:1.3px;text-transform:uppercase;opacity:0.9;">' + escHtml_(sharing) + '</div>' +
        '<h2 style="margin:4px 0 0;font-size:20px;">Prayer from ' + escHtml_(p.author || 'Anonymous') + '</h2>' +
      '</div>' +
      '<div style="border:1px solid #E2E8F0;border-top:none;padding:20px 22px;border-radius:0 0 14px 14px;background:#FFFFFF;">' +
        (isConfidential ? '<p style="background:#FFF1F2;border:1px solid #FECDD3;color:#9F1239;padding:10px 12px;border-radius:10px;font-size:12.5px;font-weight:700;margin-top:0;">🔒 Confidential Pastoral Prayer — Sent strictly to pastors only (not shared with the Prayer Team).</p>' : '') +
        (p.prayer ? '<div style="margin-bottom:14px;"><div style="font-size:12px;font-weight:700;color:#64748B;text-transform:uppercase;margin-bottom:4px;">Prayer Request / Joy or Concern</div><div style="background:#F8FAFC;padding:13px 15px;border-radius:11px;font-size:15px;white-space:pre-wrap;">' + escHtml_(p.prayer) + '</div></div>' : '') +
        (p.grace ? '<div style="margin-bottom:14px;"><div style="font-size:12px;font-weight:700;color:#64748B;text-transform:uppercase;margin-bottom:4px;">Personal Sermon Reflection</div><div style="background:#FFFBEB;padding:12px 15px;border-radius:11px;font-size:14px;white-space:pre-wrap;">' + escHtml_(p.grace) + '</div></div>' : '') +
        '<p style="font-size:12.5px;color:#64748B;margin:0 0 14px;">Date: ' + escHtml_(p.bulletinDate || p.date || '') + (p.sermonTitle ? ' · Message: ' + escHtml_(p.sermonTitle) : '') + '</p>' +
        '<a href="' + ss.getUrl() + '" style="display:inline-block;background:#1E1F4B;color:#fff;text-decoration:none;padding:10px 16px;border-radius:9px;font-weight:700;font-size:13px;">Open 3_PrayerRequests in Google Sheets</a>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({ to: recipients, subject: subject, htmlBody: html });
}

/**
 * 3. Pastoral Care & Event RSVP Alert
 */
function handleRsvpNotifications_(ss, p) {
  var churchName = resolveSetting_(p.churchName, 'CHURCH_NAME', 'Christ United Methodist Church');
  var recipients = mergeEmailLists_(
    resolveSetting_(p.welcomeAlertEmails, 'WELCOME_ALERT_EMAILS', ''),
    resolveSetting_(p.pastorPrayerEmails, 'PASTOR_PRAYER_EMAILS', '')
  );
  if (!recipients) return;

  var subject = '📅 [' + (p.requestType || 'Care / Event RSVP') + '] ' + (p.author || '') + ' — ' + churchName;
  var html =
    '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:580px;margin:0 auto;color:#1E293B;line-height:1.6;">' +
      '<div style="background:#0369A1;color:#fff;padding:18px 22px;border-radius:14px 14px 0 0;">' +
        '<div style="font-size:11px;letter-spacing:1.3px;text-transform:uppercase;opacity:0.9;">Pastoral Care &amp; Event RSVP</div>' +
        '<h2 style="margin:4px 0 0;font-size:20px;">' + escHtml_(p.requestType || 'RSVP') + '</h2>' +
      '</div>' +
      '<div style="border:1px solid #E2E8F0;border-top:none;padding:20px 22px;border-radius:0 0 14px 14px;background:#FFFFFF;font-size:14px;">' +
        '<p style="margin:0 0 8px;"><b>Name:</b> ' + escHtml_(p.author || '') + '</p>' +
        '<p style="margin:0 0 8px;"><b>Contact:</b> ' + escHtml_(p.contact || '') + '</p>' +
        '<p style="margin:0 0 8px;"><b>Preferred Time:</b> ' + escHtml_(p.schedule || 'Flexible') + '</p>' +
        '<p style="margin:0 0 16px;white-space:pre-wrap;"><b>Details:</b> ' + escHtml_(p.text || '—') + '</p>' +
        '<a href="' + ss.getUrl() + '" style="display:inline-block;background:#0369A1;color:#fff;text-decoration:none;padding:10px 16px;border-radius:9px;font-weight:700;font-size:13px;">Open 4_CareAndRSVP in Google Sheets</a>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({ to: recipients, subject: subject, htmlBody: html });
}

function readAttendanceRecords_(ss, targetDate) {
  var records = [];
  var attSheet = ss.getSheetByName('1_Attendance');
  if (!attSheet || attSheet.getLastRow() <= 1) return records;
  var cols = Math.max(12, attSheet.getLastColumn());
  var rows = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, cols).getDisplayValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!targetDate || r[0] === targetDate) {
      if (cols >= 16) {
        records.push({
          bulletinDate: r[0], service: r[1], primaryName: r[2], attendees: r[3],
          count: Number(r[4]) || 1, adults: Number(r[5]) || 0, youngAdults: Number(r[6]) || 0,
          youth: Number(r[7]) || 0, children: Number(r[8]) || 0,
          status: r[9], source: r[10], phone: r[11], email: r[12], address: r[13],
          submittedAt: r[14], checkInId: r[15]
        });
      } else {
        var c = Number(r[4]) || 1;
        records.push({
          bulletinDate: r[0], service: r[1], primaryName: r[2], attendees: r[3],
          count: c, adults: c, youngAdults: 0, youth: 0, children: 0,
          status: r[5], source: r[6], phone: r[7], email: r[8], address: r[9],
          submittedAt: r[10], checkInId: r[11]
        });
      }
    }
  }
  return records;
}

function readHeadcountsForDate_(ss, targetDate) {
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
  return headcounts;
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
  var isNewFirstTimeGuest = false;
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ensureSheetWithHeaders_(ss, '1_Attendance', ATT_HEADERS, '#DCFCE7');
    var checkInId = String(p.checkInId || '');
    var idCol = 16;
    var foundRow = 0;

    if (checkInId && sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (ids[i][0] === checkInId) { foundRow = i + 2; break; }
      }
    }

    if (p.action === 'delete') {
      if (foundRow) sheet.deleteRow(foundRow);
      return ContentService.createTextOutput('deleted');
    }

    var cnt = Math.max(1, parseInt(p.count, 10) || 1);
    var ya = parseInt(p.youngAdults, 10) || 0;
    var yth = parseInt(p.youth, 10) || 0;
    var chd = parseInt(p.children, 10) || 0;
    var adt = (p.adults !== undefined && p.adults !== '') ? (parseInt(p.adults, 10) || 0) : Math.max(0, cnt - ya - yth - chd);

    var row = [
      p.bulletinDate || '',
      p.service || '',
      p.primaryName || '',
      p.attendees || p.primaryName || '',
      cnt, adt, ya, yth, chd,
      p.status || 'Member',
      p.source || 'App',
      p.phone || '',
      p.email || '',
      p.address || '',
      p.date || new Date(),
      checkInId
    ];

    if (foundRow) {
      sheet.getRange(foundRow, 1, 1, ATT_HEADERS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
      if (p.status === 'First-Time Guest' && p.source !== 'Paper') {
        isNewFirstTimeGuest = true;
      }
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

  // If a First-Time Guest checked in on their phone for the first time today, trigger Welcome Alert + Auto-Welcome Email!
  if (isNewFirstTimeGuest && ss) {
    try {
      handleConnectNotifications_(ss, {
        name: p.attendees || p.primaryName || 'Guest',
        service: p.service || '',
        status: 'First-Time Guest (1-Tap Check-In)',
        bulletinDate: p.bulletinDate || '',
        email: p.email || '',
        phone: p.phone || '',
        interests: 'Checked in via Home 1-Tap Attendance',
        notes: p.address ? ('Address: ' + p.address) : '',
        churchName: p.churchName,
        welcomeAlertEmails: p.welcomeAlertEmails,
        autoWelcomeGuest: p.autoWelcomeGuest,
        welcomeEmailSubject: p.welcomeEmailSubject,
        welcomeEmailBody: p.welcomeEmailBody
      });
    } catch (guestMailErr) {
      Logger.log('1st-time guest check-in email error: ' + guestMailErr);
    }
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

/**
 * Configure Weekly Sunday Automated Email Trigger
 */
function configureWeeklyEmail_(p) {
  var props = PropertiesService.getScriptProperties();
  var emails = String(p.emails || '').trim();
  var hourStr = String(p.hour || '13').trim();
  var churchName = String(p.churchName || 'Christ United Methodist Church').trim();

  props.setProperty('REPORT_EMAILS', emails);
  props.setProperty('REPORT_HOUR', hourStr);
  props.setProperty('CHURCH_NAME', churchName);

  // Clear existing triggers for sendWeeklySundayAttendanceEmail
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendWeeklySundayAttendanceEmail') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  if (hourStr !== 'off' && emails) {
    var hr = Math.max(0, Math.min(23, parseInt(hourStr, 10) || 13));
    ScriptApp.newTrigger('sendWeeklySundayAttendanceEmail')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(hr)
      .create();
  }

  return ContentService.createTextOutput('scheduled');
}

function sendReportNow_(p) {
  var emails = String(p.emails || '').trim();
  var bDate = String(p.bulletinDate || '').trim();
  var churchName = String(p.churchName || 'Christ United Methodist Church').trim();
  if (emails) {
    sendAttendanceSummaryEmailForDate_(emails, bDate, churchName);
  }
  return ContentService.createTextOutput('sent');
}

/**
 * Triggered automatically every Sunday afternoon by Google Apps Script
 */
function sendWeeklySundayAttendanceEmail() {
  var props = PropertiesService.getScriptProperties();
  var emails = props.getProperty('REPORT_EMAILS') || '';
  var churchName = props.getProperty('CHURCH_NAME') || 'Christ United Methodist Church';
  if (!emails) return;

  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/Chicago';
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  sendAttendanceSummaryEmailForDate_(emails, todayStr, churchName);
}

function sendAttendanceSummaryEmailForDate_(emails, dateStr, churchName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetDate = dateStr;
  var records = readAttendanceRecords_(ss, targetDate);
  if (records.length === 0) {
    var allRecs = readAttendanceRecords_(ss, '');
    if (allRecs.length > 0) {
      targetDate = allRecs[allRecs.length - 1].bulletinDate;
      records = readAttendanceRecords_(ss, targetDate);
    }
  }

  var total = 0, members = 0, guests = 0, appCnt = 0, paperCnt = 0;
  var adults = 0, ya = 0, youth = 0, children = 0;
  var byService = {};
  var firstTimeGuests = [];

  for (var i = 0; i < records.length; i++) {
    var r = records[i];
    var c = Math.max(1, Number(r.count) || 1);
    total += c;
    adults += Number(r.adults) || 0;
    ya += Number(r.youngAdults) || 0;
    youth += Number(r.youth) || 0;
    children += Number(r.children) || 0;
    if (r.status === 'First-Time Guest') {
      guests += c;
      firstTimeGuests.push(r);
    } else {
      members += c;
    }
    if (r.source === 'Paper') paperCnt += c; else appCnt += c;
    byService[r.service] = (byService[r.service] || 0) + c;
  }

  var hc = readHeadcountsForDate_(ss, targetDate);
  var svcHtml = Object.keys(byService).map(function(s) {
    var hcNote = hc[s] ? ' <span style="color:#64748B;">(Usher Room Count: ' + escHtml_(hc[s]) + ')</span>' : '';
    return '<li style="margin-bottom:5px;"><b>' + escHtml_(s) + ':</b> ' + byService[s] + ' checked in' + hcNote + '</li>';
  }).join('') || '<li>No check-ins recorded yet.</li>';

  var guestHtml = firstTimeGuests.length ? firstTimeGuests.map(function(g) {
    var contact = [g.phone, g.email, g.address].filter(Boolean).join(' · ') || 'No contact info provided';
    return '<li style="margin-bottom:6px;"><b>' + escHtml_(g.attendees) + '</b> (' + escHtml_(g.service) + ') — ' + escHtml_(contact) + '</li>';
  }).join('') : '<li style="color:#64748B;">No first-time guests recorded for this date.</li>';

  var sheetUrl = ss.getUrl();
  var subject = '📊 ' + churchName + ' Attendance Summary — ' + (targetDate || 'Sunday') + ' (Total: ' + total + ')';

  var htmlBody =
    '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:600px;margin:0 auto;color:#1E293B;line-height:1.6;">' +
      '<div style="background:#1E1F4B;color:#fff;padding:22px 24px;border-radius:16px 16px 0 0;">' +
        '<div style="font-size:11px;letter-spacing:1.5px;text-transform:uppercase;opacity:0.85;">Weekly Worship Attendance Report</div>' +
        '<h2 style="margin:6px 0 2px;font-size:22px;">' + escHtml_(churchName) + '</h2>' +
        '<div style="font-size:14px;opacity:0.9;">Service Date: <b>' + escHtml_(targetDate || 'N/A') + '</b></div>' +
      '</div>' +
      '<div style="border:1px solid #E2E8F0;border-top:none;padding:22px 24px;border-radius:0 0 16px 16px;background:#FFFFFF;">' +
        '<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px;">' +
          '<div style="flex:1;min-width:120px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:12px;padding:12px;text-align:center;">' +
            '<div style="font-size:26px;font-weight:800;color:#15803D;">' + total + '</div>' +
            '<div style="font-size:12px;color:#166534;font-weight:700;">Total Checked In</div>' +
          '</div>' +
          '<div style="flex:1;min-width:120px;background:#F8FAFC;border:1px solid #E2E8F0;border-radius:12px;padding:12px;text-align:center;">' +
            '<div style="font-size:26px;font-weight:800;color:#1E1F4B;">' + members + '</div>' +
            '<div style="font-size:12px;color:#475569;font-weight:700;">Members &amp; Regular</div>' +
          '</div>' +
          '<div style="flex:1;min-width:120px;background:#FFF1F2;border:1px solid #FECDD3;border-radius:12px;padding:12px;text-align:center;">' +
            '<div style="font-size:26px;font-weight:800;color:#BE123C;">' + guests + '</div>' +
            '<div style="font-size:12px;color:#9F1239;font-weight:700;">1st-Time Guests</div>' +
          '</div>' +
        '</div>' +

        '<h3 style="font-size:15px;color:#1E1F4B;margin:16px 0 8px;">Age Group Breakdown</h3>' +
        '<p style="background:#F8FAFC;padding:10px 14px;border-radius:10px;font-size:14px;margin:0 0 16px;">' +
          '<b>Adults:</b> ' + adults + ' &nbsp;·&nbsp; ' +
          '<b>Young Adults:</b> ' + ya + ' &nbsp;·&nbsp; ' +
          '<b>Youth:</b> ' + youth + ' &nbsp;·&nbsp; ' +
          '<b>Children:</b> ' + children +
          '<br><span style="font-size:12px;color:#64748B;">Check-in source: ' + appCnt + ' Phone App · ' + paperCnt + ' Paper Sheet</span>' +
        '</p>' +

        '<h3 style="font-size:15px;color:#1E1F4B;margin:16px 0 8px;">By Worship Service</h3>' +
        '<ul style="padding-left:20px;margin:0 0 16px;font-size:14px;">' + svcHtml + '</ul>' +

        '<h3 style="font-size:15px;color:#BE123C;margin:16px 0 8px;">🌟 First-Time Guests (Pastoral Follow-Up)</h3>' +
        '<ul style="padding-left:20px;margin:0 0 20px;font-size:14px;">' + guestHtml + '</ul>' +

        '<a href="' + sheetUrl + '" style="display:inline-block;background:#1E1F4B;color:#fff;text-decoration:none;padding:11px 18px;border-radius:10px;font-weight:700;font-size:13.5px;">Open Full Google Spreadsheet</a>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: emails,
    subject: subject,
    htmlBody: htmlBody
  });
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
        if (String(ids[i][0]) === entryId) foundRow = i + 2;
        if (foundRow) break;
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
