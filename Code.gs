// ============================================================
// CONFIGURATION
// ============================================================
var SHEET_NAME   = 'Booking';       // ← check your tab name at the bottom of the sheet
var COL_DATE     = 1;              // A — Date
var COL_TIME     = 3;              // C — Time slot
var COL_ROOM     = 9;              // I — Room
var COL_PHONE    = 7;              // G — Phone
var COL_EMAIL    = 8;              // H — Email ✅
var COL_STATUS   = 10;             // J — Status
var COL_REMARK   = 11;             // K — Remark (not used yet, but ready)
var LIBRARY_NAME = 'បណ្ណាល័យ RUA';

// ============================================================
// GET — return all rows as JSON
// ============================================================
function doGet(e) {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1)
    .filter(function(row) { return row[0] !== ''; }) // skip empty rows
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return {
        date:    obj['Date']      || '',
        day:     obj['Day']       || '',
        time:    obj['Time slot'] || '',   // ← exact header name
        topic:   obj['Topic']     || '',
        group:   obj['Group']     || '',
        members: obj['Members']   || '',
        phone:   obj['Phone']     || '',
        email:   obj['Email']     || '',
        room:    obj['Room']      || '',   // ← exact header name
        status:  obj['Status']    || 'Pending',
        remark:  obj['Remark']    || ''
      };
    });
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// POST — append a new booking row
// ============================================================
function doPost(e) {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var payload = JSON.parse(e.postData.contents);

  // A       B            C              D                    E                   F                    G                H                I              J         K
  sheet.appendRow([
    payload.date    || '',   // A — Date
    payload.day     || '',   // B — Day
    payload.time    || '',   // C — Time slot
    payload.topic   || '—',  // D — Topic
    payload.group   || '—',  // E — Group
    payload.members || '—',  // F — Members
    payload.phone   || '',   // G — Phone
    payload.email   || '—',  // H — Email
    payload.room    || '',   // I — Room
    'Pending',               // J — Status (default)
    ''                       // K — Remark (blank)
  ]);

  addStatusDropdown(sheet);
  colorStatusCell(sheet, sheet.getLastRow(), 'Pending');

  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}
// ============================================================
// Add data-validation dropdown to Status column
// ============================================================
function addStatusDropdown(sheet) {
  var lastRow = sheet.getLastRow();
  var cell    = sheet.getRange(lastRow, COL_STATUS); // Status column
  var rule    = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending','Approved','Rejected'], true)
    .build();
  cell.setDataValidation(rule);
  cell.setValue('Pending');
}

// ============================================================
// onEdit trigger — fires when admin changes Status cell
// ============================================================
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var editedCol = e.range.getColumn();
  var editedRow = e.range.getRow();
  if (editedCol !== COL_STATUS || editedRow === 1) return; // skip header

  var newStatus = e.value;
  if (newStatus !== 'Approved' && newStatus !== 'Rejected') return;

  var row      = sheet.getRange(editedRow, 1, 1, COL_STATUS).getValues()[0];
  var userEmail= row[COL_EMAIL - 1];
  var date     = row[COL_DATE  - 1];
  var time     = row[COL_TIME  - 1];
  var room     = row[COL_ROOM  - 1];

  if (!userEmail || userEmail === '—') return; // no email provided

  sendStatusEmail(userEmail, newStatus, date, time, room);
  colorStatusCell(sheet, editedRow, newStatus);
}

// ============================================================
// Send Gmail notification
// ============================================================
function sendStatusEmail(toEmail, status, date, time, room) {
  var isApproved = (status === 'Approved');
  var subject    = isApproved
    ? '[' + LIBRARY_NAME + '] ✅ ការស្នើរសុំបន្ទប់របស់លោកអ្នកត្រូវបានអនុម័ត'
    : '[' + LIBRARY_NAME + '] ❌ ការស្នើរសុំបន្ទប់របស់លោកអ្នកត្រូវបានបដិសេធ';

  var body = isApproved
    ? 'ជំរាបសួរ,\n\n' +
      'ការស្នើរសុំបន្ទប់ពិភាក្សារបស់លោកអ្នកត្រូវបាន **អនុម័ត** ។\n\n' +
      '📅 ថ្ងៃ: ' + date + '\n' +
      '🕐 ម៉ោង: ' + time + '\n' +
      '🚪 បន្ទប់: ' + room + '\n\n' +
      'សូមទៅដល់បន្ទប់ ' + room + ' ក្នុងម៉ោងដែលបានកំណត់។\n\n' +
      'សូមអរគុណ,\n' + LIBRARY_NAME
    : 'ជំរាបសួរ,\n\n' +
      'ជាអកុសល ការស្នើរសុំបន្ទប់ពិភាក្សារបស់លោកអ្នកត្រូវបាន **បដិសេធ** ។\n\n' +
      '📅 ថ្ងៃ: ' + date + '\n' +
      '🕐 ម៉ោង: ' + time + '\n' +
      '🚪 បន្ទប់: ' + room + '\n\n' +
      'សូមមេត្តាទំនាក់ទំនងបណ្ណាល័យ ឬព្យាយាមស្នើរសុំម្ដងទៀតចំពោះម៉ោងផ្សេង។\n\n' +
      'សូមអរគុណ,\n' + LIBRARY_NAME;

  var htmlBody = isApproved
    ? buildEmailHtml('✅ ការស្នើរសុំត្រូវបានអនុម័ត', '#1E6B3C', '#D4EDDA',
        date, time, room,
        'ការស្នើរសុំបន្ទប់ពិភាក្សារបស់លោកអ្នកត្រូវបានអនុម័ត។ សូមទៅដល់បន្ទប់ ' + room + ' ក្នុងម៉ោងដែលបានកំណត់។')
    : buildEmailHtml('❌ ការស្នើរសុំត្រូវបានបដិសេធ', '#C00000', '#FFCCCC',
        date, time, room,
        'ជាអកុសល ការស្នើរសុំរបស់លោកអ្នកត្រូវបានបដិសេធ។ សូមមេត្តាទំនាក់ទំនងបណ្ណាល័យ ឬស្នើរសុំម្ដងទៀតចំពោះម៉ោងផ្សេង។');

  GmailApp.sendEmail(toEmail, subject, body, { htmlBody: htmlBody });
}

// ============================================================
// Build a clean HTML email
// ============================================================
function buildEmailHtml(title, color, bgColor, date, time, room, message) {
  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #ddd;border-radius:10px;overflow:hidden;">' +
    '<div style="background:#1F4E79;padding:20px 24px;">' +
      '<h2 style="color:white;margin:0;font-size:18px;">🏛️ បណ្ណាល័យ RUA</h2>' +
    '</div>' +
    '<div style="background:' + bgColor + ';padding:16px 24px;border-bottom:2px solid ' + color + ';">' +
      '<h3 style="color:' + color + ';margin:0;font-size:16px;">' + title + '</h3>' +
    '</div>' +
    '<div style="padding:24px;background:#ffffff;">' +
      '<p style="color:#333;font-size:14px;">' + message + '</p>' +
      '<table style="width:100%;margin-top:16px;border-collapse:collapse;font-size:14px;">' +
        '<tr style="background:#F4F7FB;"><td style="padding:10px 12px;color:#6b7c93;width:40%;">📅 ថ្ងៃ</td><td style="padding:10px 12px;font-weight:bold;color:#1a2a3a;">' + date + '</td></tr>' +
        '<tr><td style="padding:10px 12px;color:#6b7c93;">🕐 ម៉ោង</td><td style="padding:10px 12px;font-weight:bold;color:#1a2a3a;">' + time + '</td></tr>' +
        '<tr style="background:#F4F7FB;"><td style="padding:10px 12px;color:#6b7c93;">🚪 បន្ទប់</td><td style="padding:10px 12px;font-weight:bold;color:#1a2a3a;">' + room + '</td></tr>' +
      '</table>' +
    '</div>' +
    '<div style="padding:14px 24px;background:#F4F7FB;text-align:center;font-size:12px;color:#6b7c93;">' +
      'ប្រព័ន្ធស្នើរសុំបន្ទប់ — បណ្ណាល័យ RUA &nbsp;|&nbsp; admin@rua.edu.kh' +
    '</div>' +
  '</div>';
}

// ============================================================
// Color-code the Status cell automatically
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

function setupExistingRows() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var dateVal = sheet.getRange(i, COL_DATE).getValue();
    if (!dateVal) continue; // skip blank rows

    // Add dropdown to column J (Status)
    var statusCell = sheet.getRange(i, COL_STATUS);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Approved', 'Rejected'], true)
      .build();
    statusCell.setDataValidation(rule);

    // Replace old "កក់ហើយ" values with "Pending"
    var current = statusCell.getValue();
    if (!current || current === 'កក់ហើយ') {
      statusCell.setValue('Pending');
    }

    // Apply color
    colorStatusCell(sheet, i, statusCell.getValue());
  }

  SpreadsheetApp.getUi().alert('✅ រួចរាល់! Rows ទាំង ' + (lastRow - 1) + ' ត្រូវបានដំឡើង dropdown Status។');
}
function findSheetName() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var names = sheets.map(function(s) { return s.getName(); });
  SpreadsheetApp.getUi().alert('Sheet tabs found:\n' + names.join('\n'));
}