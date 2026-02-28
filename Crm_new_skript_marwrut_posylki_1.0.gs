// ============================================
// ЮРА ТРАНСПОРТЕЙШН — МАРШРУТИ ПОСИЛКИ v1.0
// Apps Script API для таблиці "Маршрути посилки"
// ID: 1Pd3nv3fbwZ_0YSzdG4cda-q52BQT57E0hDe7eQej6z8
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Маршрути посилки" → Розширення → Apps Script
// 2. Видали весь старий код і встав цей файл
// 3. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Скопіюй URL деплоя
// 5. Встав URL в CRM HTML файл та BOTI Driver
// ============================================

// ============================================
// КОНФІГУРАЦІЯ
// ============================================

var SPREADSHEET_ID = '1Pd3nv3fbwZ_0YSzdG4cda-q52BQT57E0hDe7eQej6z8';

// URL архівного скрипта (Crm_Arhiv_1.0)
var ARCHIVE_API_URL = 'https://script.google.com/macros/s/AKfycbwJLGZgYT333VdMW-nM5kPjYs2WIGGjfqkZnDJYjJxUt8nzE8GDGCPm7EzMHhcxNDOn/exec';

// Аркуші
var SHEET_LOGS = 'Маршрути водіїв';
var SHEET_MAILING = 'Провірка розсилки';

// Кольори статусів для водіїв
var STATUS_COLORS = {
  'pending':     { bg: '#fffbf0', border: '#ffc107', font: '#ffc107' },
  'in-progress': { bg: '#e3f2fd', border: '#2196F3', font: '#2196F3' },
  'completed':   { bg: '#e8f5e9', border: '#4CAF50', font: '#4CAF50' },
  'cancelled':   { bg: '#ffebee', border: '#dc3545', font: '#dc3545' }
};

// Колонки (A-Z = 26, індекс 0-25)
// A:ВО  B:Номер№  C:Номер ТТН  D:Вага  E:Адреса Отримувача
// F:Напрямок  G:Телефон Отримувача  H:Сума Є  I:Статус оплати
// J:Оплата  K:Телефон Реєстратора  L:Примітка  M:Статус посилки
// N:ІД  O:ПіБ  P:дата оформлення  Q:Таймінг  R:Примітка смс
// S:Дата отримання  T:фото  U:Статус
// V:DATE_ARCHIVE  W:ARCHIVED_BY  X:ARCHIVE_REASON
// Y:SOURCE_SHEET  Z:ARCHIVE_ID
var COL = {
  VO: 0,              // A — ВО (відправник-отримувач)
  NUMBER: 1,          // B — Номер№
  TTN: 2,             // C — Номер ТТН
  WEIGHT: 3,          // D — Вага
  ADDRESS: 4,         // E — Адреса Отримувача
  DIRECTION: 5,       // F — Напрямок
  PHONE: 6,           // G — Телефон Отримувача
  AMOUNT: 7,          // H — Сума Є
  PAY_STATUS: 8,      // I — Статус оплати
  PAYMENT: 9,         // J — Оплата
  PHONE_REG: 10,      // K — Телефон Реєстратора
  NOTE: 11,           // L — Примітка
  PARCEL_STATUS: 12,  // M — Статус посилки (pending/in-progress/completed/cancelled)
  ID: 13,             // N — ІД
  NAME: 14,           // O — ПіБ
  DATE_REG: 15,       // P — дата оформлення
  TIMING: 16,         // Q — Таймінг
  SMS_NOTE: 17,       // R — Примітка смс
  DATE_RECEIVE: 18,   // S — Дата отримання
  PHOTO: 19,          // T — фото
  STATUS: 20,         // U — Статус (CRM: new/work/archived/refused/deleted)
  DATE_ARCHIVE: 21,   // V — DATE_ARCHIVE
  ARCHIVED_BY: 22,    // W — ARCHIVED_BY
  ARCHIVE_REASON: 23, // X — ARCHIVE_REASON
  SOURCE_SHEET: 24,   // Y — SOURCE_SHEET
  ARCHIVE_ID: 25      // Z — ARCHIVE_ID
};
var TOTAL_COLS = 26;

// Заголовки для нового аркуша
var HEADERS = [
  'ВО', 'Номер№', 'Номер ТТН', 'Вага', 'Адреса Отримувача',
  'Напрямок', 'Телефон Отримувача', 'Сума Є', 'Статус оплати', 'Оплата',
  'Телефон Реєстратора', 'Примітка', 'Статус посилки', 'ІД', 'ПіБ',
  'дата оформлення', 'Таймінг', 'Примітка смс', 'Дата отримання', 'фото',
  'Статус', 'DATE_ARCHIVE', 'ARCHIVED_BY', 'ARCHIVE_REASON',
  'SOURCE_SHEET', 'ARCHIVE_ID'
];

// Статуси для архівації
var ARCHIVE_STATUSES = ['archived', 'refused', 'deleted', 'transferred'];

// Маппінг полів API → індексів колонок
var FIELD_MAP = {
  vo: COL.VO,
  number: COL.NUMBER,
  ttn: COL.TTN,
  weight: COL.WEIGHT,
  address: COL.ADDRESS,
  direction: COL.DIRECTION,
  phone: COL.PHONE,
  amount: COL.AMOUNT,
  payStatus: COL.PAY_STATUS,
  payment: COL.PAYMENT,
  phoneReg: COL.PHONE_REG,
  note: COL.NOTE,
  parcelStatus: COL.PARCEL_STATUS,
  id: COL.ID,
  name: COL.NAME,
  dateReg: COL.DATE_REG,
  timing: COL.TIMING,
  smsNote: COL.SMS_NOTE,
  dateReceive: COL.DATE_RECEIVE,
  photo: COL.PHOTO,
  status: COL.STATUS,
  dateArchive: COL.DATE_ARCHIVE,
  archivedBy: COL.ARCHIVED_BY,
  archiveReason: COL.ARCHIVE_REASON,
  sourceSheet: COL.SOURCE_SHEET,
  archiveId: COL.ARCHIVE_ID
};

// ============================================
// doGet — ВОДІЇ (BOTI Driver) + Health Check
// ============================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'health';
    var sheetParam = (e && e.parameter) ? (e.parameter.sheet || '') : '';

    switch (action) {
      case 'health':
        return respond({
          success: true,
          version: '1.0',
          service: 'Маршрути Посилки — ЮРА ТРАНСПОРТЕЙШН',
          totalCols: TOTAL_COLS,
          timestamp: new Date().toISOString()
        });

      case 'getDeliveries':
        if (!sheetParam) return respond({ success: false, error: 'Не вказано маршрут (sheet)' });
        return respond(getDeliveries(sheetParam));

      case 'getAvailableRoutes':
        return respond(getAvailableRoutes());

      default:
        return respond({ success: false, error: 'Невідома GET дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// doPost — CRM + ВОДІЇ
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var payload = data.payload || data;

    switch (action) {
      // --- CRM: МАРШРУТИ ---
      case 'copyToRoute':
        return respond(copyToRoute(payload));

      case 'checkRouteSheets':
        return respond(checkRouteSheets(payload));

      case 'getRoutePassengers':
        return respond(getRoutePackages(payload));

      case 'getRoutePackages':
        return respond(getRoutePackages(payload));

      case 'getAvailableRoutes':
        return respond(getAvailableRoutes());

      case 'deleteRouteSheet':
        return respond(deleteRouteSheet(payload));

      case 'createRouteSheet':
        return respond(createRouteSheet(payload));

      // --- CRM: ОНОВЛЕННЯ ---
      case 'updateField':
        return respond(updateField(payload));

      case 'updateStatus':
        return respond(updateStatus(payload));

      case 'updateMultiple':
        return respond(updateMultiple(payload));

      // --- CRM: АРХІВАЦІЯ ---
      case 'archivePackages':
        return respond(changeStatus(payload, 'archived'));

      case 'restorePackages':
        return respond(changeStatus(payload, 'work'));

      case 'deletePackages':
        return respond(changeStatus(payload, 'deleted'));

      case 'archiveToExternal':
        return respond(archiveToExternal(payload));

      // --- ВОДІЙ: СТАТУС ---
      case 'updateDriverStatus':
      case 'updateStatus_driver':
        return respond(handleDriverStatusUpdate(data));

      // --- РОЗСИЛКА ---
      case 'getMailingStatus':
        return respond(getMailingStatus());

      case 'addMailingRecord':
        return respond(addMailingRecord(payload));

      case 'checkMailing':
        return respond(checkMailingByIds(payload));

      // --- ДЕБАГ ---
      case 'getStructure':
        return respond(getStructure());

      default:
        return respond({ success: false, error: 'Невідома дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// getDeliveries — Читання посилок для водіїв
// ============================================
function getDeliveries(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, deliveries: [], count: 0, sheetName: sheetName };
    }

    var readCols = Math.min(sheet.getLastColumn(), TOTAL_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();
    var deliveries = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var internalNumber = str(row[COL.NUMBER]);
      if (!internalNumber && !str(row[COL.VO])) continue;

      // Пропускаємо архівовані
      var crmStatus = str(row[COL.STATUS]).toLowerCase();
      if (ARCHIVE_STATUSES.indexOf(crmStatus) !== -1) continue;

      deliveries.push({
        rowNum: i + 2,
        internalNumber: internalNumber,
        vo: str(row[COL.VO]),
        ttn: str(row[COL.TTN]),
        weight: str(row[COL.WEIGHT]),
        address: str(row[COL.ADDRESS]),
        direction: str(row[COL.DIRECTION]),
        phone: str(row[COL.PHONE]),
        price: str(row[COL.AMOUNT]),
        paymentStatus: str(row[COL.PAY_STATUS]),
        payment: str(row[COL.PAYMENT]),
        registrarPhone: str(row[COL.PHONE_REG]),
        note: str(row[COL.NOTE]),
        parcelStatus: str(row[COL.PARCEL_STATUS]) || 'pending',
        id: str(row[COL.ID]),
        name: str(row[COL.NAME]),
        createdAt: str(row[COL.DATE_REG]),
        timing: str(row[COL.TIMING]),
        smsNote: str(row[COL.SMS_NOTE]),
        receiveDate: str(row[COL.DATE_RECEIVE]),
        photo: str(row[COL.PHOTO]),
        status: crmStatus || 'new',
        archiveId: str(row[COL.ARCHIVE_ID]),
        sheet: sheetName
      });
    }

    // Статистика
    var stats = { total: deliveries.length, pending: 0, inProgress: 0, completed: 0, cancelled: 0 };
    for (var j = 0; j < deliveries.length; j++) {
      var ps = (deliveries[j].parcelStatus || 'pending').toLowerCase();
      if (ps === 'pending') stats.pending++;
      else if (ps === 'in-progress') stats.inProgress++;
      else if (ps === 'completed') stats.completed++;
      else if (ps === 'cancelled') stats.cancelled++;
    }

    return {
      success: true,
      deliveries: deliveries,
      count: deliveries.length,
      sheetName: sheetName,
      stats: stats
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// getRoutePackages — Читання маршруту (для CRM)
// ============================================
function getRoutePackages(payload) {
  try {
    var vehicleName = payload.vehicleName || '';
    var sheetName = payload.sheetName || '';

    if (vehicleName && !sheetName) {
      sheetName = vehicleName + ' марш.';
    }
    if (!sheetName) {
      return { success: false, error: 'Не вказано маршрут' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: true, packages: [], passengers: [], count: 0, sheetName: sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, packages: [], passengers: [], count: 0, sheetName: sheetName, stats: { total: 0 } };
    }

    var readCols = Math.min(sheet.getLastColumn(), TOTAL_COLS);
    var dataRange = sheet.getRange(2, 1, lastRow - 1, readCols);
    var data = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();
    var packages = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!str(row[COL.NUMBER]) && !str(row[COL.VO])) continue;

      packages.push({
        rowNum: i + 2,
        vo: str(row[COL.VO]),
        number: str(row[COL.NUMBER]),
        ttn: str(row[COL.TTN]),
        weight: str(row[COL.WEIGHT]),
        address: str(row[COL.ADDRESS]),
        direction: str(row[COL.DIRECTION]),
        phone: str(row[COL.PHONE]),
        amount: str(row[COL.AMOUNT]),
        payStatus: str(row[COL.PAY_STATUS]),
        payment: str(row[COL.PAYMENT]),
        phoneReg: str(row[COL.PHONE_REG]),
        note: str(row[COL.NOTE]),
        parcelStatus: str(row[COL.PARCEL_STATUS]) || 'pending',
        id: str(row[COL.ID]),
        name: str(row[COL.NAME]),
        dateReg: str(row[COL.DATE_REG]),
        timing: str(row[COL.TIMING]),
        smsNote: str(row[COL.SMS_NOTE]),
        dateReceive: str(row[COL.DATE_RECEIVE]),
        photo: str(row[COL.PHOTO]),
        status: str(row[COL.STATUS]),
        archiveId: str(row[COL.ARCHIVE_ID]),
        rowColor: backgrounds[i][0],
        sheet: sheetName
      });
    }

    // Статистика
    var stats = { total: packages.length, pending: 0, inProgress: 0, completed: 0, cancelled: 0 };
    for (var j = 0; j < packages.length; j++) {
      var ps = (packages[j].parcelStatus || 'pending').toLowerCase();
      if (ps === 'pending') stats.pending++;
      else if (ps === 'in-progress') stats.inProgress++;
      else if (ps === 'completed') stats.completed++;
      else if (ps === 'cancelled') stats.cancelled++;
    }

    return {
      success: true,
      packages: packages,
      passengers: packages,  // alias для CRM сумісності
      count: packages.length,
      sheetName: sheetName,
      vehicleName: vehicleName,
      stats: stats
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// getAvailableRoutes — Список маршрутних аркушів
// ============================================
function getAvailableRoutes() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var routes = [];

    // Службові аркуші які НЕ є маршрутами
    var excludePatterns = ['логи', 'logs', 'водіїв', 'розсилк', 'провірка', 'перевірка', 'template', 'шаблон', 'тест', 'test', 'архів'];

    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      var nameLower = name.toLowerCase();

      // Маршрутні аркуші закінчуються на " марш."
      var isRoute = nameLower.indexOf('марш') !== -1;

      var isExcluded = false;
      for (var e = 0; e < excludePatterns.length; e++) {
        if (nameLower.indexOf(excludePatterns[e]) !== -1 && !isRoute) {
          isExcluded = true;
          break;
        }
      }
      if (isExcluded) continue;

      // Для маршрутних аркушів обов'язково " марш."
      if (!isRoute) continue;

      var count = Math.max(0, sheets[i].getLastRow() - 1);
      routes.push({
        name: name,
        vehicle: name.replace(' марш.', ''),
        count: count,
        sheetId: sheets[i].getSheetId()
      });
    }

    return { success: true, routes: routes, count: routes.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// copyToRoute — CRM → маршрутний аркуш
// ============================================
function copyToRoute(payload) {
  try {
    var packagesByVehicle = payload.packagesByVehicle;
    var conflictAction = payload.conflictAction || 'add';

    if (!packagesByVehicle) {
      return { success: false, error: 'Немає посилок' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var totalCopied = 0;
    var totalArchived = 0;
    var totalCleared = 0;
    var results = [];

    for (var vehicleName in packagesByVehicle) {
      if (!packagesByVehicle.hasOwnProperty(vehicleName)) continue;
      var packages = packagesByVehicle[vehicleName];
      if (!packages || !packages.length) continue;

      var sheetName = vehicleName + ' марш.';
      var sheet = ss.getSheetByName(sheetName);

      // Створюємо аркуш якщо не існує
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
        sheet.getRange(1, 1, 1, HEADERS.length)
          .setBackground('#1a1a2e')
          .setFontColor('#ffffff')
          .setFontWeight('bold');
        sheet.setFrozenRows(1);
      }

      // Обробка конфліктів
      var lastRow = sheet.getLastRow();
      if (lastRow > 1 && conflictAction !== 'add') {
        if (conflictAction === 'clear') {
          totalCleared += lastRow - 1;
          sheet.deleteRows(2, lastRow - 1);
        } else if (conflictAction === 'archive') {
          totalArchived += lastRow - 1;
          // Архівуємо старі рядки (ставимо статус)
          var archiveRange = sheet.getRange(2, COL.STATUS + 1, lastRow - 1, 1);
          var archiveValues = [];
          for (var a = 0; a < lastRow - 1; a++) archiveValues.push(['archived']);
          archiveRange.setValues(archiveValues);

          var dateRange = sheet.getRange(2, COL.DATE_ARCHIVE + 1, lastRow - 1, 1);
          var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
          var dateValues = [];
          for (var d = 0; d < lastRow - 1; d++) dateValues.push([dateNow]);
          dateRange.setValues(dateValues);
        }
      }

      // Записуємо посилки
      var rows = [];
      for (var p = 0; p < packages.length; p++) {
        var pkg = packages[p];
        var newRow = new Array(TOTAL_COLS);
        for (var c = 0; c < TOTAL_COLS; c++) newRow[c] = '';

        newRow[COL.VO] = pkg.vo || '';
        newRow[COL.NUMBER] = pkg.number || '';
        newRow[COL.TTN] = pkg.ttn || '';
        newRow[COL.WEIGHT] = pkg.weight || '';
        newRow[COL.ADDRESS] = pkg.address || '';
        newRow[COL.DIRECTION] = pkg.direction || '';
        newRow[COL.PHONE] = pkg.phone || '';
        newRow[COL.AMOUNT] = pkg.amount || '';
        newRow[COL.PAY_STATUS] = pkg.payStatus || '';
        newRow[COL.PAYMENT] = pkg.payment || '';
        newRow[COL.PHONE_REG] = pkg.phoneReg || '';
        newRow[COL.NOTE] = pkg.note || '';
        newRow[COL.PARCEL_STATUS] = 'pending';
        newRow[COL.ID] = pkg.id || '';
        newRow[COL.NAME] = pkg.name || '';
        newRow[COL.DATE_REG] = pkg.dateReg || '';
        newRow[COL.TIMING] = pkg.timing || '';
        newRow[COL.SMS_NOTE] = pkg.smsNote || '';
        newRow[COL.DATE_RECEIVE] = pkg.dateReceive || '';
        newRow[COL.PHOTO] = pkg.photo || '';
        newRow[COL.STATUS] = 'new';
        newRow[COL.SOURCE_SHEET] = pkg.sourceSheet || '';

        rows.push(newRow);
      }

      if (rows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, TOTAL_COLS).setValues(rows);

        // Фарбуємо pending рядки
        var pendingColors = STATUS_COLORS['pending'];
        if (pendingColors) {
          sheet.getRange(startRow, 1, rows.length, TOTAL_COLS).setBackground(pendingColors.bg);
        }

        totalCopied += rows.length;
      }

      results.push({ vehicle: vehicleName, sheet: sheetName, copied: packages.length });

      // Авторозмір
      try { sheet.autoResizeColumns(1, Math.min(20, TOTAL_COLS)); } catch (e) {}
    }

    writeLog('copyToRoute', 'bulk', 0, 'copied: ' + totalCopied,
      'archived: ' + totalArchived + ' cleared: ' + totalCleared);

    return {
      success: true,
      copied: totalCopied,
      archived: totalArchived,
      cleared: totalCleared,
      details: results
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// checkRouteSheets — Перевірка чи є дані
// ============================================
function checkRouteSheets(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var vehicleNames = payload.vehicleNames || [];
    var existing = [];

    for (var i = 0; i < vehicleNames.length; i++) {
      var sheetName = vehicleNames[i] + ' марш.';
      var sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        existing.push({
          vehicle: vehicleNames[i],
          sheet: sheetName,
          count: sheet.getLastRow() - 1
        });
      }
    }

    return { success: true, existing: existing };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// createRouteSheet — Створити маршрутний аркуш
// ============================================
function createRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву авто' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = vehicleName + ' марш.';

    var existing = ss.getSheetByName(sheetName);
    if (existing) {
      return { success: true, sheetName: sheetName, existed: true };
    }

    var sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);

    writeLog('createRouteSheet', sheetName, 0, 'created', '');

    return { success: true, sheetName: sheetName, vehicleName: vehicleName };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// deleteRouteSheet — Видалити маршрутний аркуш
// ============================================
function deleteRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву авто' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = vehicleName + ' марш.';
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: true, message: 'Аркуш не існує', deleted: false };
    }

    var rowCount = sheet.getLastRow() - 1;
    if (rowCount > 0 && !payload.force) {
      return {
        success: false,
        error: 'Аркуш містить ' + rowCount + ' записів. Використайте force=true',
        recordsCount: rowCount
      };
    }

    ss.deleteSheet(sheet);
    writeLog('deleteRouteSheet', sheetName, 0, 'deleted', '');

    return { success: true, sheetName: sheetName, deleted: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// updateField — Оновити одне поле
// ============================================
function updateField(payload) {
  var sheetName = payload.sheet;
  var rowNum = parseInt(payload.rowNum);
  var field = payload.field;
  var value = payload.value;

  if (!sheetName || !rowNum || !field) {
    return { success: false, error: 'Відсутні sheet, rowNum або field' };
  }
  if (!FIELD_MAP.hasOwnProperty(field)) {
    return { success: false, error: 'Невідоме поле: ' + field };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Аркуш не знайдено' };
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Рядок не існує' };

  // Верифікація
  if (payload.expectedId) {
    var currentId = str(sheet.getRange(rowNum, COL.ID + 1).getValue());
    if (currentId !== String(payload.expectedId).trim()) {
      return {
        success: false,
        error: 'conflict',
        message: 'Рядок змінився. Очікувався ІД: ' + payload.expectedId + ', фактичний: ' + currentId
      };
    }
  }

  sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(value);
  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}

// ============================================
// updateStatus — Змінити CRM статус
// ============================================
function updateStatus(payload) {
  var sheetName = payload.sheet;
  var rowNum = parseInt(payload.rowNum);
  var newStatus = payload.status;

  if (!sheetName || !rowNum || !newStatus) {
    return { success: false, error: 'Відсутні sheet, rowNum або status' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Аркуш не знайдено' };
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Рядок не існує' };

  var oldStatus = str(sheet.getRange(rowNum, COL.STATUS + 1).getValue());
  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  // При архівації — записуємо дату
  if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
    sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
      Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
    );
  }

  writeLog('updateStatus', sheetName, rowNum, oldStatus + ' → ' + newStatus, '');

  return { success: true, sheet: sheetName, rowNum: rowNum, status: newStatus };
}

// ============================================
// updateMultiple — Масове оновлення полів
// ============================================
function updateMultiple(payload) {
  var items = payload.items || payload.packages || [];
  var updated = 0;
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (!item.sheet || !item.rowNum) {
      errors.push('Елемент ' + i + ': відсутні sheet/rowNum');
      continue;
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(item.sheet);
    if (!sheet) { errors.push(item.sheet + ': не знайдено'); continue; }

    var rowNum = parseInt(item.rowNum);
    if (rowNum > sheet.getLastRow()) { errors.push('Рядок ' + rowNum + ': не існує'); continue; }

    for (var field in item) {
      if (item.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
        if (item[field] !== undefined) {
          sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(item[field]);
        }
      }
    }
    updated++;
  }

  return { success: true, updated: updated, errors: errors.length > 0 ? errors : undefined };
}

// ============================================
// changeStatus — Масова зміна CRM статусу
// ============================================
function changeStatus(payload, newStatus) {
  try {
    var items = payload.packages || payload.items || payload.passengers || [];
    var note = payload.note || '';

    if (items.length === 0) {
      return { success: false, error: 'Немає записів' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var user = payload.user || 'crm';
    var changed = 0;
    var errors = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var sheet = ss.getSheetByName(item.sheet);
      if (!sheet) { errors.push(item.sheet + ': не знайдено'); continue; }

      var rowNum = parseInt(item.rowNum);
      if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
        errors.push('Рядок ' + rowNum + ': не існує');
        continue;
      }

      sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

      if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
        sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
        sheet.getRange(rowNum, COL.ARCHIVED_BY + 1).setValue(user);
        if (note) {
          sheet.getRange(rowNum, COL.ARCHIVE_REASON + 1).setValue(note);
        }
      }

      changed++;
    }

    writeLog('changeStatus:' + newStatus, 'bulk', 0, changed + '/' + items.length, note || '');

    return {
      success: true,
      changed: changed,
      total: items.length,
      status: newStatus,
      errors: errors.length > 0 ? errors : undefined
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// archiveToExternal — Відправити в Crm_Arhiv_1.0
// ============================================
function archiveToExternal(payload) {
  try {
    var items = payload.items || payload.packages || [];
    var user = payload.user || 'crm';
    var reason = payload.reason || 'route_archive';

    if (items.length === 0) {
      return { success: false, error: 'Немає записів' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var archivedIds = [];
    var count = 0;
    var errors = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var sheet = ss.getSheetByName(item.sheet);
      if (!sheet) { errors.push(item.sheet + ': не знайдено'); continue; }

      var rowNum = parseInt(item.rowNum);
      if (rowNum > sheet.getLastRow()) { errors.push('Рядок ' + rowNum); continue; }

      // Перевірка: вже архівовано?
      var existingArchiveId = str(sheet.getRange(rowNum, COL.ARCHIVE_ID + 1).getValue());
      if (existingArchiveId) {
        errors.push('Рядок ' + rowNum + ': вже архівовано (' + existingArchiveId + ')');
        continue;
      }

      var recordId = str(sheet.getRange(rowNum, COL.ID + 1).getValue());

      // Ставимо статус
      sheet.getRange(rowNum, COL.STATUS + 1).setValue('archived');
      sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
      sheet.getRange(rowNum, COL.ARCHIVED_BY + 1).setValue(user);
      sheet.getRange(rowNum, COL.ARCHIVE_REASON + 1).setValue(reason);

      if (recordId) archivedIds.push(recordId);
      count++;
    }

    // Відправляємо в зовнішній архів
    var archiveResult = { success: false, error: 'Немає ІД' };
    if (archivedIds.length > 0) {
      archiveResult = sendToArchive({
        action: 'archiveByIds',
        source: 'ROUTE_POSYLKY',
        ids: archivedIds,
        user: user,
        reason: reason,
        deleteFromSource: false
      });
    }

    writeLog('archiveToExternal', 'bulk', 0, 'archived: ' + count,
      'Archive API: ' + (archiveResult.success ? 'OK' : archiveResult.error));

    return {
      success: true,
      count: count,
      total: items.length,
      errors: errors.length > 0 ? errors : undefined,
      archiveResult: archiveResult
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// handleDriverStatusUpdate — Оновлення від водія
// ============================================
function handleDriverStatusUpdate(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Логуємо в аркуш водіїв
    var logSheet = ss.getSheetByName(SHEET_LOGS);
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOGS);
      logSheet.getRange(1, 1, 1, 10).setValues([[
        'Дата', 'Час', 'Водій', 'Маршрут', 'Номер посилки',
        'Адреса', 'Статус', 'Причина скасування', 'Телефон', 'Сума'
      ]]);
      logSheet.getRange(1, 1, 1, 10)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    var now = new Date();
    logSheet.appendRow([
      Utilities.formatDate(now, 'Europe/Kiev', 'yyyy-MM-dd'),
      Utilities.formatDate(now, 'Europe/Kiev', 'HH:mm:ss'),
      data.driverId || '',
      data.routeName || '',
      data.deliveryNumber || '',
      data.address || '',
      data.status || '',
      data.cancelReason || '',
      data.phone || '',
      data.price || ''
    ]);

    // 2. Оновлюємо статус у маршрутному аркуші
    var routeSheet = ss.getSheetByName(data.routeName);
    if (!routeSheet) {
      return { success: true, message: 'Логовано (маршрут не знайдено)' };
    }

    var allData = routeSheet.getDataRange().getValues();
    var rowsUpdated = 0;
    var deliveryNumber = str(data.deliveryNumber);

    for (var i = 1; i < allData.length; i++) {
      var num = str(allData[i][COL.NUMBER]);
      if (num === deliveryNumber) {
        var rowNum = i + 1;

        // Записуємо статус посилки (M — parcelStatus)
        routeSheet.getRange(rowNum, COL.PARCEL_STATUS + 1).setValue(data.status);

        // Якщо completed — записуємо дату отримання
        if (data.status === 'completed') {
          routeSheet.getRange(rowNum, COL.DATE_RECEIVE + 1).setValue(
            Utilities.formatDate(now, 'Europe/Kiev', 'yyyy-MM-dd HH:mm')
          );
        }

        // Причина скасування → в Примітку
        if (data.status === 'cancelled' && data.cancelReason) {
          var currentNote = str(routeSheet.getRange(rowNum, COL.NOTE + 1).getValue());
          var newNote = 'Скасовано: ' + data.cancelReason + (currentNote ? ' | ' + currentNote : '');
          routeSheet.getRange(rowNum, COL.NOTE + 1).setValue(newNote);
        }

        // Кольори
        var colors = STATUS_COLORS[data.status];
        if (colors) {
          var readCols = Math.min(routeSheet.getLastColumn(), TOTAL_COLS);
          var rangeToColor = routeSheet.getRange(rowNum, 1, 1, readCols);
          rangeToColor.setBackground(colors.bg);
          rangeToColor.setBorder(true, true, true, true, true, true,
            colors.border, SpreadsheetApp.BorderStyle.SOLID);

          var statusCell = routeSheet.getRange(rowNum, COL.PARCEL_STATUS + 1);
          statusCell.setFontColor(colors.font);
          statusCell.setFontWeight('bold');
        }

        rowsUpdated++;
      }
    }

    if (rowsUpdated === 0) {
      return { success: true, message: 'Логовано (посилку не знайдено в маршруті)' };
    }

    return {
      success: true,
      message: 'Статус записано',
      updatedRows: rowsUpdated
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// РОЗСИЛКА — "Провірка розсилки"
// ============================================

// getMailingStatus — Всі ІД з розсилки
function getMailingStatus() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, mailingIds: [], count: 0 };
    }

    var lastRow = sheet.getLastRow();
    var readCols = Math.min(sheet.getLastColumn(), 4);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();

    var mailingData = [];
    var mailingIds = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      // Шукаємо ІД (може бути в різних колонках)
      var id = '';
      var status = '';

      // Колонки A, B, C — дані, D — Статус
      for (var c = 0; c < Math.min(3, row.length); c++) {
        var val = str(row[c]);
        if (val && val.indexOf('dd.mm') === -1 && val.length > 1) {
          if (!id) id = val;
        }
      }
      if (row.length > 3) status = str(row[3]);

      if (id) {
        mailingIds.push(id);
        mailingData.push({ id: id, status: status });
      }
    }

    // Унікальні
    var uniqueIds = [];
    var seen = {};
    for (var u = 0; u < mailingIds.length; u++) {
      if (!seen[mailingIds[u]]) {
        seen[mailingIds[u]] = true;
        uniqueIds.push(mailingIds[u]);
      }
    }

    return {
      success: true,
      mailingIds: uniqueIds,
      mailingData: mailingData,
      count: uniqueIds.length
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// checkMailingByIds — Перевірити конкретні ІД (чи зроблена розсилка)
function checkMailingByIds(payload) {
  try {
    var idsToCheck = payload.ids || [];
    if (idsToCheck.length === 0) {
      return { success: true, results: {}, count: 0 };
    }

    var statusResult = getMailingStatus();
    if (!statusResult.success) return statusResult;

    var mailingSet = {};
    for (var m = 0; m < statusResult.mailingIds.length; m++) {
      mailingSet[statusResult.mailingIds[m]] = true;
    }

    var results = {};
    var mailedCount = 0;
    for (var i = 0; i < idsToCheck.length; i++) {
      var id = String(idsToCheck[i]).trim();
      var isMailed = mailingSet[id] === true;
      results[id] = isMailed;
      if (isMailed) mailedCount++;
    }

    return {
      success: true,
      results: results,
      total: idsToCheck.length,
      mailed: mailedCount,
      notMailed: idsToCheck.length - mailedCount
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// addMailingRecord — Записати що розсилка зроблена
function addMailingRecord(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_MAILING);
      sheet.getRange(1, 1, 1, 4).setValues([['Стовпець 1', 'Стовпець 2', 'Стовпець 3', 'Статус']]);
    }

    var records = payload.records || [];
    var userName = payload.userName || 'CRM';

    if (records.length === 0) {
      return { success: false, error: 'Немає записів' };
    }

    var now = new Date();
    var date = Utilities.formatDate(now, 'Europe/Kiev', 'dd.MM.yyyy');
    var rows = [];

    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      rows.push([
        r.id || r.packageId || '',
        userName,
        date,
        r.status || 'sent'
      ]);
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);

    return { success: true, added: rows.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// getStructure — Дебаг
// ============================================
function getStructure() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var result = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

    result.push({
      sheet: sheet.getName(),
      rows: lastRow,
      cols: lastCol,
      headers: headers
    });
  }

  return { success: true, sheets: result };
}

// ============================================
// ВІДПРАВКА В АРХІВ (HTTP до Crm_Arhiv_1.0)
// ============================================
function sendToArchive(payload) {
  try {
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(ARCHIVE_API_URL, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code === 200) {
      try { return JSON.parse(body); }
      catch (e) { return { success: false, error: 'Невалідна відповідь' }; }
    }
    return { success: false, error: 'HTTP ' + code };
  } catch (e) {
    Logger.log('Archive API error: ' + e.toString());
    return { success: false, error: 'Архів недоступний: ' + e.toString() };
  }
}

// ============================================
// ЛОГУВАННЯ
// ============================================
function writeLog(action, sheetName, rowNum, detail, extra) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logName = 'Логи';
    var logSheet = ss.getSheetByName(logName);

    if (!logSheet) {
      logSheet = ss.insertSheet(logName);
      logSheet.appendRow(['Дата/Час', 'Дія', 'Аркуш', 'Рядок', 'Деталі', 'Дані']);
      logSheet.getRange(1, 1, 1, 6)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }

    var timestamp = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([timestamp, action, sheetName, rowNum, detail, extra || '']);
  } catch (e) {
    Logger.log('Log error: ' + e.toString());
  }
}

// ============================================
// ДОПОМІЖНІ
// ============================================

// Безпечне перетворення в string
function str(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

// JSON відповідь
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// МЕНЮ В GOOGLE SHEETS
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Маршрути Посилки')
    .addItem('Список маршрутів', 'menuAvailableRoutes')
    .addItem('Структура таблиці', 'menuStructure')
    .addItem('Тест архів зв\'язок', 'menuTestArchive')
    .addToUi();
}

function menuAvailableRoutes() {
  var result = getAvailableRoutes();
  var msg = 'Маршрутів: ' + result.count + '\n\n';
  for (var i = 0; i < result.routes.length; i++) {
    var r = result.routes[i];
    msg += r.name + ' — ' + r.count + ' записів\n';
  }
  SpreadsheetApp.getUi().alert('Маршрути', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuStructure() {
  var result = getStructure();
  for (var i = 0; i < result.sheets.length; i++) {
    var s = result.sheets[i];
    Logger.log('[' + s.sheet + '] ' + s.rows + ' рядків, ' + s.cols + ' колонок');
    Logger.log('  Headers: ' + s.headers.join(' | '));
  }
  SpreadsheetApp.getUi().alert('Дивись Logger (Ctrl+Enter)');
}

function menuTestArchive() {
  var result = sendToArchive({ action: 'getStats' });
  SpreadsheetApp.getUi().alert(
    'Архів',
    result.success ? 'OK — зв\'язок працює' : 'ПОМИЛКА: ' + result.error,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
