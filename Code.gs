const SPREADSHEET_NAME = '2026_公司旅遊';
const SHEET_NAME = '報名資料';

function doPost(event) {
  try {
    const fields = event && event.parameter ? event.parameter : {};
    const requiredFields = ['name', 'email', 'phone'];
    if (requiredFields.some(function (field) { return !String(fields[field] || '').trim(); })) {
      return jsonResponse({ ok: false, message: '缺少必填欄位' });
    }

    const sheet = getRegistrationSheet();
    const nextRow = sheet.getLastRow() + 1;
    const values = [
      safeText(fields.name),
      safeText(fields.email),
      safeText(fields.unit),
      safeText(fields.phone),
      safeText(fields.title)
    ];
    sheet.getRange(nextRow, 2, 1, values.length).setNumberFormat('@');
    sheet.getRange(nextRow, 1, 1, 1).setValue(new Date());
    sheet.getRange(nextRow, 2, 1, values.length).setValues([values]);

    return jsonResponse({ ok: true });
  } catch (error) {
    return jsonResponse({ ok: false, message: String(error) });
  }
}

function initialize() {
  getRegistrationSheet();
}

function getRegistrationSheet() {
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  let spreadsheet;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      spreadsheet = SpreadsheetApp.openById(file.getId());
      break;
    }
  }
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME);
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([['提交時間', '姓名', 'Email', '單位', '電話', '職稱']]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e4ede3');
    sheet.setFrozenRows(1);
  }
  sheet.getRange('E:E').setNumberFormat('@');
  return sheet;
}

function safeText(value) {
  const text = String(value || '').trim();
  return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}