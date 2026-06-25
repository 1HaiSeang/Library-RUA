// ============================================================
// CONFIGURATION
// ============================================================
var SHEET_NAME   = 'Booking';
var COL_DATE     = 1;   // A
var COL_TIME     = 3;   // C
var COL_ROOM     = 9;   // I
var COL_PHONE    = 7;   // G
var COL_EMAIL    = 8;   // H
var COL_STATUS   = 10;  // J
var COL_REMARK   = 11;  // K
var LIBRARY_NAME = 'បណ្ណាល័យ';
// ============================================================
// TELEGRAM CONFIG — add these near the top
// ============================================================
var TELEGRAM_TOKEN   = 'Telegram_Bot_API_Token';        // from BotFather
var TELEGRAM_PRIVATE_ID = '1795526925';    // your chat ID
var TELEGRAM_CHAT_ID    = '-1003887735287';
var TELEGRAM_API     = 'https://api.telegram.org/bot' + TELEGRAM_TOKEN;

// ============================================================
// GET
// ============================================================
function doGet(e) {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1)
    .filter(function(row) { return row[0] !== ''; })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return {
        date:    obj['Date']      || '',
        day:     obj['Day']       || '',
        time:    obj['Time slot'] || '',
        topic:   obj['Topic']     || '',
        group:   obj['Group']     || '',
        members: obj['Members']   || '',
        phone:   obj['Phone']     || '',
        email:   obj['Email']     || '',
        room:    obj['Room']      || '',
        status:  obj['Status']    || 'Pending',
        remark:  obj['Remark']    || ''
      };
    });
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// POST
// ============================================================
function doPost(e) {
  var contents;
  
  try {
    contents = JSON.parse(e.postData.contents);
  } catch(err) {
    // Cannot parse JSON at all
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: 'Invalid JSON'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── Telegram webhook call ──────────────────────────────
  if (contents.update_id) {
    // This is from Telegram
    if (contents.callback_query) {
      handleTelegramCallback(contents.callback_query);
    } else if (contents.message) {
      handleTelegramMessage(contents.message);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── Booking form POST ──────────────────────────────────
  if (!contents.date && !contents.room && !contents.phone) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: 'Unknown request'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.appendRow([
    contents.date    || '',
    contents.day     || '',
    contents.time    || '',
    contents.topic   || '—',
    contents.group   || '—',
    contents.members || '—',
    contents.phone   || '',
    contents.email   || '—',
    contents.room    || '',
    'Pending',
    ''
  ]);

  addStatusDropdown(sheet);
  colorStatusCell(sheet, sheet.getLastRow(), 'Pending');

  // Notify admin via Telegram
  var lastRow = sheet.getLastRow();
  var row     = sheet.getRange(lastRow, 1, 1, 11).getValues()[0];
  try {
    sendTelegramBooking(
      lastRow,
      row[0], row[1], row[2],
      row[8], row[6], row[7],
      row[5], row[3], row[4]
    );
  } catch(tgErr) {
    Logger.log('Telegram notify error: ' + tgErr.message);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Dropdown
// ============================================================
function addStatusDropdown(sheet) {
  var lastRow = sheet.getLastRow();
  var cell    = sheet.getRange(lastRow, COL_STATUS);
  var rule    = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending','Approved','Rejected'], true)
    .build();
  cell.setDataValidation(rule);
  cell.setValue('Pending');
}

// ============================================================
// onEdit — silent, logs to Remark column
// ============================================================
function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var editedCol = e.range.getColumn();
  var editedRow = e.range.getRow();

  // Only watch Status column J, skip header
  if (editedCol !== COL_STATUS || editedRow <= 1) return;

  // ── IMPORTANT: ignore edits to Remark column triggered by setValue ──
  // If the value is not one of our 3 status values, ignore completely
  var newStatus = e.value;
  if (newStatus !== 'Approved' && newStatus !== 'Rejected') return;

  // Use lock to prevent concurrent/double execution
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000); // wait max 3 seconds
  } catch(lockErr) {
    Logger.log('Could not get lock: ' + lockErr);
    return;
  }

  try {
    colorStatusCell(sheet, editedRow, newStatus);

    var rowData   = sheet.getRange(editedRow, 1, 1, COL_STATUS + 1).getValues()[0];
    var userEmail = rowData[COL_EMAIL - 1];
    var date      = rowData[COL_DATE  - 1];
    var time      = rowData[COL_TIME  - 1];
    var room      = rowData[COL_ROOM  - 1];

    // Write to Remark BEFORE sending — avoids re-trigger issue
    var remarkCell = sheet.getRange(editedRow, COL_REMARK);
    var timestamp  = Utilities.formatDate(new Date(), 'Asia/Phnom_Penh', 'dd/MM/yyyy HH:mm');

    if (!userEmail || userEmail === '—' || userEmail === '') {
      remarkCell.setValue('No Email — not sent ' + timestamp);
      return;
    }

    // Mark as "sending" first
    remarkCell.setValue('Sending...');

    // Send email
    sendStatusEmail(userEmail, newStatus, date, time, room);

    // Update remark to success AFTER send
    remarkCell.setValue('Sent ' + timestamp + ' > ' + userEmail);
    Logger.log('Email sent to: ' + userEmail);

  } catch(err) {
    // Only write error if it's NOT a permission false-alarm after successful send
    var errMsg = err.message || '';
    if (errMsg.indexOf('mail.google.com') === -1 &&
        errMsg.indexOf('gmail.send') === -1 &&
        errMsg.indexOf('gmail.compose') === -1) {
      // Real error — log it
      sheet.getRange(editedRow, COL_REMARK)
           .setValue('Failed: ' + errMsg.substring(0, 100));
      Logger.log('Real error: ' + errMsg);
    } else {
      // False-alarm permission error after successful send — ignore silently
      Logger.log('False-alarm ignored (email was sent): ' + errMsg);
    }
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// Send Email
// ============================================================
function sendStatusEmail(toEmail, status, date, time, room) {
  var isApproved = (status === 'Approved');

  var subject = isApproved
    ? 'បណ្ណាល័យ៖ ការស្នើរសុំបន្ទប់របស់លោកអ្នកត្រូវបានអនុម័ត - APPROVED'
    : 'បណ្ណាល័យ៖ ការស្នើរសុំបន្ទប់របស់លោកអ្នកត្រូវបានបដិសេធ - REJECTED';

  var body = isApproved
    ? 'ជំរាបសួរ,\n\nការស្នើរសុំបន្ទប់ពិភាក្សារបស់លោកអ្នកត្រូវបានអនុម័ត។\n\n' +
      'ថ្ងៃ: ' + date + '\nម៉ោង: ' + time + '\nបន្ទប់: ' + room + '\n\n' +
      'សូមទៅដល់បន្ទប់ ' + room + ' ក្នុងម៉ោងដែលបានកំណត់។\n\nសូមអរគុណ,\nបណ្ណាល័យ'
    : 'ជំរាបសួរ,\n\nជាអកុសល ការស្នើរសុំបន្ទប់ពិភាក្សារបស់លោកអ្នកត្រូវបានបដិសេធ។\n\n' +
      'ថ្ងៃ: ' + date + '\nម៉ោង: ' + time + '\nបន្ទប់: ' + room + '\n\n' +
      'សូមមេត្តាទំនាក់ទំនងបណ្ណាល័យ ឬស្នើរសុំម្ដងទៀត។\n\nសូមអរគុណ,\nបណ្ណាល័យ';

  var htmlBody = buildEmailHtml(
    isApproved ? 'APPROVED — ការស្នើរសុំត្រូវបានអនុម័ត' : 'REJECTED — ការស្នើរសុំត្រូវបានបដិសេធ',
    isApproved ? '#1E6B3C' : '#C00000',
    isApproved ? '#D4EDDA' : '#FFCCCC',
    date, time, room,
    isApproved
      ? 'ការស្នើរសុំបន្ទប់របស់លោកអ្នកត្រូវបានអនុម័ត។ សូមទៅដល់បន្ទប់ ' + room + ' ក្នុងម៉ោងដែលបានកំណត់។'
      : 'ការស្នើរសុំរបស់លោកអ្នកត្រូវបានបដិសេធ។ សូមមេត្តាទំនាក់ទំនងបណ្ណាល័យ ឬស្នើរសុំម្ដងទៀត។'
  );

  GmailApp.sendEmail(toEmail, subject, body, {
    htmlBody: htmlBody,
    name:     'បណ្ណាល័យ RUA',
    // Anti-spam headers
    replyTo:  'cheanlyheng@rua.edu.kh',   // ← your school email
    cc:       '',
    bcc:      ''
  });
}

// ============================================================
// Build HTML Email with SVG Icons
// ============================================================
function buildEmailHtml(title, color, bgColor, date, time, room, message) {

  // SVG icon — library/book for header
  var iconLibrary =
    '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" ' +
    'stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" ' +
    'style="display:block;">' +
      '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/>' +
      '<path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>' +
    '</svg>';

  // SVG icon — checkmark for approved
  var iconApproved =
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" ' +
    'stroke="' + color + '" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" ' +
    'style="vertical-align:middle;margin-right:8px;">' +
      '<circle cx="12" cy="12" r="10"/>' +
      '<polyline points="9 12 11 14 15 10"/>' +
    '</svg>';

  // SVG icon — X for rejected
  var iconRejected =
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" ' +
    'stroke="' + color + '" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" ' +
    'style="vertical-align:middle;margin-right:8px;">' +
      '<circle cx="12" cy="12" r="10"/>' +
      '<line x1="15" y1="9" x2="9" y2="15"/>' +
      '<line x1="9" y1="9" x2="15" y2="15"/>' +
    '</svg>';

  var statusIcon = (bgColor === '#D4EDDA') ? iconApproved : iconRejected;

  // SVG icon — calendar
  // Replace these 3 SVG icon variables:

  // Calendar icon → colored square + text
  var iconCalendar =
    '<span style="display:inline-block;width:10px;height:10px;' +
    'background:#2E75B6;border-radius:2px;margin-right:8px;' +
    'vertical-align:middle;"></span>';

  // Clock icon → colored circle
  var iconClock =
    '<span style="display:inline-block;width:10px;height:10px;' +
    'background:#2E75B6;border-radius:50%;margin-right:8px;' +
    'vertical-align:middle;"></span>';

  // Room icon → colored diamond shape
  var iconRoom =
    '<span style="display:inline-block;width:10px;height:10px;' +
    'background:#2E75B6;border-radius:2px;margin-right:8px;' +
    'transform:rotate(45deg);vertical-align:middle;"></span>';

  return (
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;' +
    'border:1px solid #ddd;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);">' +

    // ── Header ──────────────────────────────────────────
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1F4E79;">' +
      '<tr>' +
        '<td style="padding:20px 24px;">' +
          '<table cellpadding="0" cellspacing="0"><tr>' +
            '<td style="padding-right:14px;vertical-align:middle;">' +
              '<div style="background:rgba(255,255,255,0.18);border-radius:10px;' +
              'width:44px;height:44px;text-align:center;line-height:44px;display:inline-block;' +
              'padding:8px;box-sizing:border-box;">' +
                iconLibrary +
              '</div>' +
            '</td>' +
            '<td style="vertical-align:middle;">' +
              '<div style="color:white;font-size:18px;font-weight:700;' +
              'font-family:Arial,sans-serif;line-height:1.2;">បណ្ណាល័យ RUA</div>' +
              '<div style="color:rgba(255,255,255,0.65);font-size:12px;' +
              'font-family:Arial,sans-serif;margin-top:2px;">ប្រព័ន្ធស្នើរសុំបន្ទប់ពិភាក្សា</div>' +
            '</td>' +
          '</tr></table>' +
        '</td>' +
      '</tr>' +
    '</table>' +

    // ── Status Banner ────────────────────────────────────
    '<div style="background:' + bgColor + ';padding:16px 24px;' +
    'border-bottom:3px solid ' + color + ';">' +
      '<div style="font-size:15px;font-weight:700;color:' + color + ';' +
      'font-family:Arial,sans-serif;">' +
        statusIcon + title +
      '</div>' +
    '</div>' +

    // ── Message ──────────────────────────────────────────
    '<div style="padding:22px 24px 10px;background:#ffffff;">' +
      '<p style="color:#444;font-size:14px;line-height:1.8;margin:0;' +
      'font-family:Arial,sans-serif;">' + message + '</p>' +
    '</div>' +

    // ── Details Table ────────────────────────────────────
    '<div style="padding:4px 24px 24px;background:#ffffff;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" ' +
      'style="border-collapse:collapse;font-size:14px;border-radius:8px;' +
      'overflow:hidden;border:1px solid #E8F0F9;">' +

        // Date row
        '<tr style="background:#F0F6FF;">' +
          '<td style="padding:12px 16px;color:#5a6a7a;width:36%;' +
          'border-bottom:1px solid #E8F0F9;font-family:Arial,sans-serif;">' +
            iconCalendar + 'ថ្ងៃខែឆ្នាំ' +
          '</td>' +
          '<td style="padding:12px 16px;font-weight:700;color:#1a2a3a;' +
          'border-bottom:1px solid #E8F0F9;font-family:Arial,sans-serif;">' +
            date +
          '</td>' +
        '</tr>' +

        // Time row
        '<tr style="background:#ffffff;">' +
          '<td style="padding:12px 16px;color:#5a6a7a;' +
          'border-bottom:1px solid #E8F0F9;font-family:Arial,sans-serif;">' +
            iconClock + 'ម៉ោង' +
          '</td>' +
          '<td style="padding:12px 16px;font-weight:700;color:#1a2a3a;' +
          'border-bottom:1px solid #E8F0F9;font-family:Arial,sans-serif;">' +
            time +
          '</td>' +
        '</tr>' +

        // Room row
        '<tr style="background:#F0F6FF;">' +
          '<td style="padding:12px 16px;color:#5a6a7a;font-family:Arial,sans-serif;">' +
            iconRoom + 'បន្ទប់' +
          '</td>' +
          '<td style="padding:12px 16px;font-weight:700;color:#1a2a3a;' +
          'font-family:Arial,sans-serif;">' +
            room +
          '</td>' +
        '</tr>' +

      '</table>' +
    '</div>' +

    // ── Footer ───────────────────────────────────────────
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1F4E79;">' +
      '<tr>' +
        '<td style="padding:14px 24px;text-align:center;">' +
          '<p style="color:rgba(255,255,255,0.7);font-size:12px;margin:0;' +
          'font-family:Arial,sans-serif;">' +
            'ប្រព័ន្ធស្នើរសុំបន្ទប់ &nbsp;|&nbsp; បណ្ណាល័យ' +
          '</p>' +
        '</td>' +
      '</tr>' +
    '</table>' +

    '</div>'
  );
}

// ============================================================
// Color cell
// ============================================================
function colorStatusCell(sheet, row, status) {
  var cell = sheet.getRange(row, COL_STATUS);
  if (status === 'Approved') {
    cell.setBackground('#D4EDDA').setFontColor('#1E6B3C').setFontWeight('bold');
  } else if (status === 'Rejected') {
    cell.setBackground('#FFCCCC').setFontColor('#C00000').setFontWeight('bold');
  } else {
    cell.setBackground('#FFF3CD').setFontColor('#7D5A00').setFontWeight('bold');
  }
}

// ============================================================
// Setup existing rows
// ============================================================
function setupExistingRows() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    var dateVal = sheet.getRange(i, COL_DATE).getValue();
    if (!dateVal) continue;
    var statusCell = sheet.getRange(i, COL_STATUS);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Approved', 'Rejected'], true)
      .build();
    statusCell.setDataValidation(rule);
    var current = statusCell.getValue();
    if (!current || current === 'កក់ហើយ') statusCell.setValue('Pending');
    colorStatusCell(sheet, i, statusCell.getValue());
  }
  SpreadsheetApp.getUi().alert('Done! ' + (lastRow-1) + ' rows updated.');
}

// ============================================================
// Send Telegram notification when new booking arrives
// ============================================================
function sendTelegramBooking(rowIndex, date, day, time, room, phone, email, members, topic, group) {
  var text =
    '📋 *ការស្នើរសុំបន្ទប់ថ្មី* \\#' + rowIndex + '\n' +
    '━━━━━━━━━━━━━━━\n' +
    '📅 *ថ្ងៃ:* ' + date + ' (' + day + ')\n' +
    '🕐 *ម៉ោង:* ' + time + '\n' +
    '🚪 *បន្ទប់:* ' + room + '\n' +
    '👥 *ក្រុម:* ' + (group   || '—') + '\n' +
    '👤 *សមាជិក:* ' + (members || '—') + '\n' +
    '📚 *ប្រធានបទ:* ' + (topic  || '—') + '\n' +
    '📞 *ទូរសព្ទ:* ' + (phone  || '—') + '\n' +
    '📧 *Email:* '   + (email  || '—') + '\n' +
    '━━━━━━━━━━━━━━━\n' +
    '⏳ *ស្ថានភាព:* រង់ចាំការអនុម័ត';

  var keyboard = {
    inline_keyboard: [[
      { text: '✅ Approve', callback_data: 'APPROVE_' + rowIndex },
      { text: '❌ Reject',  callback_data: 'REJECT_'  + rowIndex }
    ]]
  };

  // Send to BOTH group and private
  var targets = [TELEGRAM_CHAT_ID, TELEGRAM_PRIVATE_ID];
  targets.forEach(function(chatId) {
    if (!chatId || chatId === '') return; // skip if empty
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id:      chatId,
        text:         text,
        parse_mode:   'Markdown',
        reply_markup: JSON.stringify(keyboard)
      }),
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', options);
  });
}

// ============================================================
// Handle Telegram button press (webhook)
// ============================================================
function doPost(e) {
  // Check if this is a Telegram webhook call
  try {
    var contents = JSON.parse(e.postData.contents);

    // Handle Telegram callback (button press)
    if (contents.callback_query) {
      handleTelegramCallback(contents.callback_query);
      return ContentService
        .createTextOutput(JSON.stringify({ok: true}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Handle Telegram message (e.g. /start command)
    if (contents.message) {
      handleTelegramMessage(contents.message);
      return ContentService
        .createTextOutput(JSON.stringify({ok: true}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    // Not a Telegram call — handle as normal booking form POST
  }

  // Normal booking form POST
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var payload = JSON.parse(e.postData.contents);
  sheet.appendRow([
    payload.date    || '',
    payload.day     || '',
    payload.time    || '',
    payload.topic   || '—',
    payload.group   || '—',
    payload.members || '—',
    payload.phone   || '',
    payload.email   || '—',
    payload.room    || '',
    'Pending',
    ''
  ]);

  addStatusDropdown(sheet);
  colorStatusCell(sheet, sheet.getLastRow(), 'Pending');

  // Notify admin via Telegram
  var lastRow = sheet.getLastRow();
  var row     = sheet.getRange(lastRow, 1, 1, 11).getValues()[0];
  try {
    sendTelegramBooking(
      lastRow,
      row[0], row[1], row[2],  // date, day, time
      row[8],                   // room (col I = index 8)
      row[6],                   // phone (col G = index 6)
      row[7],                   // email (col H = index 7)
      row[5],                   // members (col F = index 5)
      row[3],                   // topic (col D = index 3)
      row[4]                    // group (col E = index 4)
    );
  } catch(tgErr) {
    Logger.log('Telegram error: ' + tgErr.message);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Handle button press from Telegram
// ============================================================
function handleTelegramCallback(callback) {
  var callbackId = callback.id;
  var data       = callback.data;
  var messageId  = callback.message.message_id;
  var chatId     = callback.message.chat.id.toString();

  var parts     = data.split('_');
  var action    = parts[0];
  var rowIndex  = parseInt(parts[1]);
  var newStatus = (action === 'APPROVE') ? 'Approved' : 'Rejected';

  var sheet      = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var statusCell = sheet.getRange(rowIndex, COL_STATUS);

  // Already processed
  var currentStatus = statusCell.getValue();
  if (currentStatus === 'Approved' || currentStatus === 'Rejected') {
    answerCallbackQuery(callbackId,
      '⚠️ ការស្នើរសុំនេះបានដំណើរការរួចហើយ: ' + currentStatus
    );
    return;
  }

  // Update sheet
  statusCell.setValue(newStatus);
  colorStatusCell(sheet, rowIndex, newStatus);

  // Get row data
  var rowData   = sheet.getRange(rowIndex, 1, 1, COL_STATUS + 1).getValues()[0];
  var userEmail = rowData[COL_EMAIL - 1];
  var date      = rowData[COL_DATE  - 1];
  var time      = rowData[COL_TIME  - 1];
  var room      = rowData[COL_ROOM  - 1];
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Phnom_Penh', 'dd/MM/yyyy HH:mm');
  var remarkCell = sheet.getRange(rowIndex, COL_REMARK);

  // Send Gmail
  if (userEmail && userEmail !== '—' && userEmail !== '') {
    try {
      sendStatusEmail(userEmail, newStatus, date, time, room);
      remarkCell.setValue(newStatus + ' via Telegram | Sent ' + timestamp + ' > ' + userEmail);
    } catch(err) {
      remarkCell.setValue(newStatus + ' via Telegram | Email failed: ' + err.message.substring(0, 60));
    }
  } else {
    remarkCell.setValue(newStatus + ' via Telegram | No email ' + timestamp);
  }

  // ── Step 1: Remove buttons from original message (keep text, remove keyboard) ──
  removeInlineKeyboard(chatId, messageId);

  // ── Step 2: Send NEW result message below the original ──
  var isApproved  = (newStatus === 'Approved');
  var resultEmoji = isApproved ? '✅' : '❌';
  var resultLabel = isApproved ? 'APPROVED' : 'REJECTED';
  var emailStatus = (userEmail && userEmail !== '—' && userEmail !== '')
    ? '📧 Email បានផ្ញើទៅ: ' + userEmail
    : '📧 គ្មាន Email — មិនបានផ្ញើ';

  var resultText =
    resultEmoji + ' *' + resultLabel + '* — Row ' + rowIndex + '\n' +
    '━━━━━━━━━━━━━━━\n' +
    '📅 ' + date + '\n' +
    '🕐 ' + time + '\n' +
    '🚪 ' + room + '\n' +
    '━━━━━━━━━━━━━━━\n' +
    emailStatus + '\n' +
    '🕓 _' + timestamp + '_';

  // Send result to the chat where button was pressed
  sendTelegramText(chatId, resultText);

  // Notify the OTHER chat too
  var otherChatId = (chatId === TELEGRAM_CHAT_ID.toString())
    ? TELEGRAM_PRIVATE_ID
    : TELEGRAM_CHAT_ID;

  if (otherChatId && otherChatId !== '' && otherChatId !== chatId) {
    var notifyText =
      resultEmoji + ' *Row ' + rowIndex + '* ត្រូវបាន *' + resultLabel + '*\n' +
      '📅 ' + date + ' | 🕐 ' + time + ' | 🚪 ' + room + '\n' +
      '🕓 _' + timestamp + '_';
    sendTelegramText(otherChatId, notifyText);
  }

  // ── Step 3: Answer callback — shows popup alert to admin ──
  var alertMsg = isApproved
    ? '✅ បានអនុម័តរួចហើយ!'
    : '❌ បានបដិសេធរួចហើយ!';
  answerCallbackQuery(callbackId, alertMsg);
}

// ============================================================
// Handle Telegram text commands
// ============================================================
function handleTelegramMessage(message) {
  // Ignore old messages
  var msgTime = message.date;
  var now     = Math.floor(Date.now() / 1000);
  if (now - msgTime > 30) return;

  var text   = message.text || '';
  var chatId = message.chat.id.toString();

  // ── Accept BOTH group and private chat ──
  var isGroup   = (chatId === TELEGRAM_CHAT_ID.toString());
  var isPrivate = (chatId === TELEGRAM_PRIVATE_ID.toString());
  if (!isGroup && !isPrivate) return; // ignore unknown chats

  if (text.indexOf('/start') === 0 || text.indexOf('/help') === 0) {
    sendTelegramText(chatId,
      '🏛️ *RUA Library Admin Bot*\n\n' +
      'Commands:\n' +
      '/pending — show all pending bookings\n' +
      '/today — show today bookings\n' +
      '/stats — show booking statistics\n\n' +
      'New bookings appear here automatically\\!'
    );
  } else if (text.indexOf('/pending') === 0) {
    showPendingBookings(chatId);
  } else if (text.indexOf('/today') === 0) {
    showTodayBookings(chatId);
  } else if (text.indexOf('/stats') === 0) {
    showStats(chatId);
  }
}

// ============================================================
// /pending command — show all pending rows with buttons
// ============================================================
function showPendingBookings(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows  = sheet.getDataRange().getValues();
  var pending = [];

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][COL_STATUS - 1] === 'Pending' && rows[i][0] !== '') {
      pending.push({ rowIndex: i + 1, row: rows[i] });
    }
  }

  if (pending.length === 0) {
    sendTelegramText(chatId, '✅ មិនមានការស្នើរសុំរង់ចាំទេ!');
    return;
  }

  // Send count summary first
  sendTelegramText(chatId, '⏳ *Pending: ' + pending.length + ' ការស្នើរសុំ*');

  // Send each pending booking with buttons
  pending.forEach(function(item) {
    var r = item.row;
    var text =
      '📋 *Pending \\#' + item.rowIndex + '*\n' +
      '📅 ' + r[0] + ' (' + r[1] + ')\n' +
      '🕐 ' + r[2] + '\n' +
      '🚪 ' + r[8] + '\n' +
      '📞 ' + (r[6] || '—') + '\n' +
      '📧 ' + (r[7] || '—');

    var keyboard = {
      inline_keyboard: [[
        { text: '✅ Approve', callback_data: 'APPROVE_' + item.rowIndex },
        { text: '❌ Reject',  callback_data: 'REJECT_'  + item.rowIndex }
      ]]
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id:      chatId,
        text:         text,
        parse_mode:   'Markdown',
        reply_markup: JSON.stringify(keyboard)
      }),
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', options);
  });
}

// ============================================================
// /today command
// ============================================================
function showTodayBookings(chatId) {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows    = sheet.getDataRange().getValues();
  var today   = Utilities.formatDate(new Date(), 'Asia/Phnom_Penh', 'dd/MM/yyyy');
  var found   = [];

  for (var i = 1; i < rows.length; i++) {
    var rowDate = rows[i][0] ? rows[i][0].toString() : '';
    // Handle both Date objects and strings
    if (rowDate.indexOf(today) !== -1 || rowDate === today) {
      found.push(rows[i]);
    }
  }

  if (found.length === 0) {
    sendTelegramText(chatId, '📭 មិនមានការកក់សម្រាប់ថ្ងៃនេះទេ។');
    return;
  }

  var text = '📅 *ការកក់ថ្ងៃនេះ (' + today + ')*\n━━━━━━━━━━━━━━━\n';
  found.forEach(function(r) {
    var statusIcon = r[COL_STATUS-1] === 'Approved' ? '✅' :
                     r[COL_STATUS-1] === 'Rejected'  ? '❌' : '⏳';
    text += statusIcon + ' ' + r[2] + ' — ' + r[8] + '\n';
  });
  sendTelegramText(chatId, text);
}

// ============================================================
// /stats command
// ============================================================
function showStats(chatId) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows   = sheet.getDataRange().getValues();
  var total = 0, approved = 0, rejected = 0, pending = 0;

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    total++;
    var s = rows[i][COL_STATUS - 1];
    if (s === 'Approved') approved++;
    else if (s === 'Rejected') rejected++;
    else pending++;
  }

  sendTelegramText(chatId,
    '📊 *ស្ថិតិការស្នើរសុំបន្ទប់*\n' +
    '━━━━━━━━━━━━━━━\n' +
    '📋 សរុប: *' + total + '*\n' +
    '✅ អនុម័ត: *' + approved + '*\n' +
    '❌ បដិសេធ: *' + rejected + '*\n' +
    '⏳ រង់ចាំ: *' + pending + '*'
  );
}

// ============================================================
// Helper: send plain text to Telegram
// ============================================================
function sendTelegramText(chatId, text) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id:    chatId,
      text:       text,
      parse_mode: 'Markdown'
    }),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', options);
}

// ============================================================
// Helper: answer callback query (dismiss loading on button)
// ============================================================
function answerCallbackQuery(callbackId, text) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      callback_query_id: callbackId,
      text:              text,
      show_alert:        true
    }),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(TELEGRAM_API + '/answerCallbackQuery', options);
}

// ============================================================
// Helper: edit existing Telegram message after action
// ============================================================
function editTelegramMessage(chatId, messageId, newText) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id:    chatId,
      message_id: messageId,
      text:       newText,
      parse_mode: 'Markdown'
    }),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(TELEGRAM_API + '/editMessageText', options);
}

function removeInlineKeyboard(chatId, messageId) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id:      chatId,
      message_id:   messageId,
      reply_markup: JSON.stringify({ inline_keyboard: [] }) // empty = remove buttons
    }),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(TELEGRAM_API + '/editMessageReplyMarkup', options);
}

function setTelegramWebhook() {
  var webAppUrl = 'https://script.google.com/macros/s/AKfycbxHui_TGSFcvADbwsZsq4ubJ1TAVRQvhdAcd3IXEdvggR_OTHlcKXRL4EqyGfERLVb2/exec'; // your deployed Apps Script URL
  var url = TELEGRAM_API + '/setWebhook?url=' + encodeURIComponent(webAppUrl);
  var response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
  SpreadsheetApp.getUi().alert('Webhook set! ' + response.getContentText());
}

function clearTelegramUpdates() {
  // Delete webhook first
  UrlFetchApp.fetch(TELEGRAM_API + '/deleteWebhook');
  
  // Get all pending updates
  var response = UrlFetchApp.fetch(TELEGRAM_API + '/getUpdates');
  var data = JSON.parse(response.getContentText());
  
  if (data.result && data.result.length > 0) {
    // Clear all updates by setting offset to last_update_id + 1
    var lastId = data.result[data.result.length - 1].update_id;
    UrlFetchApp.fetch(TELEGRAM_API + '/getUpdates?offset=' + (lastId + 1));
  }
  
  SpreadsheetApp.getUi().alert('✅ Webhook deleted and updates cleared!');
}