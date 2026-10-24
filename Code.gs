const SHEET_NAME = '表單回覆 1';
const REQUEST_ID_COLUMN = 10;

function getSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) throw new Error('SPREADSHEET_ID_NOT_CONFIGURED');

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (err) {
    console.error(err);
    throw new Error('SPREADSHEET_ACCESS_ERROR');
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('SHEET_NOT_FOUND');
  return sheet;
}

function hasId_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return false;
  return !!sheet.getRange(2, REQUEST_ID_COLUMN, last - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
}

function safe_(value) {
  const str = String(value == null ? '' : value);
  return /^\s*[=+@-]/.test(str) ? "'" + str : str;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getErrorCode_(err, fallback) {
  const code = String(err && err.message || '');
  return /^[A-Z][A-Z0-9_]{1,50}$/.test(code) ? code : fallback;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    try {
      lock.waitLock(20000);
    } catch (err) {
      throw new Error('LOCK_TIMEOUT');
    }
    if (!e || !e.postData || !e.postData.contents) throw new Error('EMPTY_REQUEST');

    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      throw new Error('INVALID_JSON');
    }

    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('INVALID_FORM_DATA');
    const id = String(data.requestId || '');
    if (!/^[0-9a-f-]{36}$/i.test(id)) throw new Error('INVALID_REQUEST_ID');

    const guests = Number(data.guests);
    const vegetarian = Number(data.vegetarian);
    const babyChair = Number(data.babyChair);
    const name = String(data.name || '').trim();
    const email = String(data.email || '').trim();
    const relationship = String(data.relationship || '').trim();
    if (!name || !email || !relationship ||
        !Number.isInteger(guests) || guests < 0 || guests > 100 ||
        !Number.isInteger(vegetarian) || vegetarian < 0 || vegetarian > guests ||
        !Number.isInteger(babyChair) || babyChair < 0 || babyChair > guests) {
      throw new Error('INVALID_FORM_DATA');
    }

    const sheet = getSheet_();
    if (!hasId_(sheet, id)) {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, 10).setValues([[
        new Date(), safe_(name), safe_(email), safe_(data.address),
        safe_(relationship), guests, vegetarian, babyChair,
        safe_(data.message), id
      ]]);
      SpreadsheetApp.flush();
    }
    return json_({success: true});
  } catch (err) {
    console.error(err);
    return json_({success: false, error: getErrorCode_(err, 'SUBMISSION_FAILED')});
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function doGet(e) {
  const parameters = e && e.parameter || {};
  const callback = String(parameters.callback || '');
  const id = String(parameters.id || '');
  if (!/^rsvpCallback_[A-Za-z0-9_]{1,50}$/.test(callback)) {
    return ContentService.createTextOutput('invalid callback');
  }

  let result;
  if (!/^[0-9a-f-]{36}$/i.test(id)) {
    result = {saved: false, error: 'INVALID_REQUEST_ID'};
  } else {
    try {
      result = {saved: hasId_(getSheet_(), id)};
    } catch (err) {
      console.error(err);
      result = {saved: false, error: getErrorCode_(err, 'STATUS_LOOKUP_FAILED')};
    }
  }

  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
