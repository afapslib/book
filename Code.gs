const CONFIG = {
  SPREADSHEET_ID: '1Y6f-TD8uxHGIwAT41p72u60KLLeKSoqVHAtqfbT5e9Q',
  SHEET_NAME: 'book'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบเสนอความต้องการหนังสือ | ห้องสมุด กสป.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function submitBookRequests(payload) {
  if (!payload || !payload.unit || !Array.isArray(payload.books) || payload.books.length === 0) {
    throw new Error('กรุณาเลือกหน่วยและกรอกข้อมูลหนังสืออย่างน้อย 1 รายการ');
  }

  const unit = String(payload.unit).trim();
  const allowedUnits = [
    'สกศ.ฯ', 'กรม นร.ฯ', 'กพล.ฯ', 'กสน.ฯ', 'กพท.ฯ',
    'กวท.ฯ', 'กทส.ฯ', 'สมท.บก.ฯ', 'บก.รร.ตท.สปท.', 'ร้อย.บร.ฯ'
  ];
  if (!allowedUnits.includes(unit)) throw new Error('หน่วยที่เลือกไม่ถูกต้อง');

  const cleanBooks = payload.books.map((book, index) => ({
    title: String(book.title || '').trim(),
    author: String(book.author || '').trim(),
    row: index + 1
  })).filter(book => book.title || book.author);

  if (!cleanBooks.length) throw new Error('กรุณากรอกชื่อหนังสือหรือชื่อผู้แต่งอย่างน้อย 1 รายการ');
  const incomplete = cleanBooks.find(book => !book.title || !book.author);
  if (incomplete) throw new Error('รายการที่ ' + incomplete.row + ' กรุณากรอกชื่อหนังสือและชื่อผู้แต่งให้ครบ');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

    const headers = ['วันที่เสนอ', 'รหัสคำขอ', 'หน่วย', 'ชื่อหนังสือ', 'ชื่อผู้แต่ง', 'ลำดับรายการ', 'สถานะ'];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    } else {
      const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
      if (existing[0] !== headers[0] || existing[1] !== headers[1]) {
        throw new Error('ชีท book มีรูปแบบหัวตารางไม่ตรงกับระบบ กรุณาตรวจสอบก่อนใช้งาน');
      }
    }

    const now = new Date();
    const requestId = 'BK-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') +
      '-' + Math.floor(100 + Math.random() * 900);
    const rows = cleanBooks.map(book => [
      now, requestId, unit, book.title, book.author, book.row, 'เสนอความต้องการแล้ว'
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(sheet.getLastRow() - rows.length + 1, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    return { success: true, requestId: requestId, count: rows.length };
  } finally {
    lock.releaseLock();
  }
}
