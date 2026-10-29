const SHEETS = {
  students: 'Students',
  logs: 'SitUpLogs'
};

function setupSpreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(spreadsheet, SHEETS.students, ['name', 'studentClass', 'studentId']);
  ensureSheet_(spreadsheet, SHEETS.logs, ['time', 'studentClass', 'studentId', 'name', 'count', 'duration']);
}

function doGet() {
  return json_({ status: 'success', message: '仰臥起坐 Google Apps Script API 正常運作' });
}

function doPost(event) {
  try {
    const body = parseRequest_(event);
    switch (body.action) {
      case 'searchStudent':
        return searchStudent_(body.studentId);
      case 'saveSitUpBatch':
        return saveSitUpBatch_(body.students, body.duration);
      case 'getSitUpLogs':
        return getSitUpLogs_();
      default:
        return json_({ status: 'error', message: 'Unknown action' });
    }
  } catch (error) {
    console.error(error.stack || error.message);
    return json_({ status: 'error', message: error.message || '伺服器要求失敗' });
  }
}

function searchStudent_(studentId) {
  const query = String(studentId || '').trim().toLowerCase();
  if (!query) return json_({ status: 'not_found' });

  const rows = readObjects_(SHEETS.students);
  const row = rows.find(item =>
    String(item.studentId || item['學號'] || '').trim().toLowerCase() === query
  );
  if (!row) return json_({ status: 'not_found' });

  return json_({
    status: 'found',
    student: {
      name: row.name || row['姓名'] || '',
      studentClass: row.studentClass || row['班別'] || '',
      studentId: row.studentId || row['學號'] || query
    }
  });
}

function saveSitUpBatch_(students, duration) {
  if (!Array.isArray(students) || students.length < 1 || students.length > 10) {
    throw new Error('學生數量格式不正確');
  }
  const seconds = Number(duration);
  if (!Number.isInteger(seconds) || seconds < 1 || seconds > 3600) {
    throw new Error('測驗時間格式不正確');
  }

  const rows = students.map(student => {
    const count = Number(student && student.count);
    const name = String(student && student.name || '').trim();
    if (!name || !Number.isInteger(count) || count < 0) {
      throw new Error('學生資料或次數格式不正確');
    }
    return [
      new Date(),
      String(student.studentClass || ''),
      String(student.studentId || ''),
      name,
      count,
      seconds
    ];
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_(SHEETS.logs);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return json_({ status: 'success' });
}

function getSitUpLogs_() {
  const rows = readObjects_(SHEETS.logs);
  const logs = rows.slice(-100).reverse().map(row => ({
    time: formatDate_(row.time || row['時間']),
    studentClass: row.studentClass || row['班別'] || '',
    name: row.name || row['姓名'] || '',
    count: row.count || row['次數'] || '',
    duration: row.duration || row['秒數'] || ''
  }));
  return json_({ status: 'success', logs: logs });
}

function parseRequest_(event) {
  if (!event || !event.postData || !event.postData.contents) {
    throw new Error('缺少要求內容');
  }
  const body = JSON.parse(event.postData.contents);
  if (!body || typeof body !== 'object') throw new Error('要求格式不正確');
  return body;
}

function readObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(row => Object.fromEntries(
    headers.map((header, index) => [header, row[index] === undefined ? '' : row[index]])
  ));
}

function getSheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`找不到工作表：${sheetName}，請先執行 setupSpreadsheet`);
  return sheet;
}

function ensureSheet_(spreadsheet, sheetName, headers) {
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function formatDate_(value) {
  if (!(value instanceof Date)) return String(value || '');
  return Utilities.formatDate(
    value,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
