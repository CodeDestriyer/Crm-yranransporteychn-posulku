// ============================================
// ЮРА ТРАНСПОРТЕЙШН — МАРШРУТИ ПАСАЖИРИ v1.0
// Apps Script API для таблиці "Маршрут Пасажири"
// ID: 1iKlD0Bj-5qB3Gc1d5ZBHscbRipcSe5xU7svqBfpB77Y
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Маршрут Пасажири" → Розширення → Apps Script
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

var SPREADSHEET_ID = '1iKlD0Bj-5qB3Gc1d5ZBHscbRipcSe5xU7svqBfpB77Y';

// URL архівного скрипта (Crm_Arhiv_1.0)
var ARCHIVE_API_URL = 'https://script.google.com/macros/s/AKfycbwJLGZgYT333VdMW-nM5kPjYs2WIGGjfqkZnDJYjJxUt8nzE8GDGCPm7EzMHhcxNDOn/exec';

// Google Maps API ключ (для оптимізації маршрутів)
var API_KEY = 'AIzaSyCthPzhD6zDM9zR-re0R2ceohyhCRdawNc';

// Точка старту за замовчуванням
var DEFAULT_START = { name: 'Ужгород', lat: 48.6209, lng: 22.2879 };
var MAX_POINTS_PER_MAP = 25;

// Службові аркуші (не маршрути)
var SHEET_LOGS = 'Логи';
var SHEET_MAILING = 'Провірка розсилки';

// Кольори статусів для водіїв
var STATUS_COLORS = {
  'pending':     { bg: '#fffbf0', border: '#ffc107', font: '#ffc107' },
  'in-progress': { bg: '#e3f2fd', border: '#2196F3', font: '#2196F3' },
  'completed':   { bg: '#e8f5e9', border: '#4CAF50', font: '#4CAF50' },
  'cancelled':   { bg: '#ffebee', border: '#dc3545', font: '#dc3545' }
};

// Колонки маршрутного аркуша пасажирів (A-W = 23, індекс 0-22)
// A:Дата виїзду  B:Адреса Відправки  C:Адреса прибуття  D:Кількість місць
// E:ПіБ  F:Телефон Пасажира  G:Відмітка  H:Оплата  I:Відсоток
// J:Диспечер  K:ІД  L:Телефон Реєстратора  M:Вага  N:Автомобіль
// O:Таймінг  P:дата оформлення  Q:Примітка
// R:Статус  S:DATE_ARCHIVE  T:ARCHIVED_BY  U:ARCHIVE_REASON
// V:SOURCE_SHEET  W:ARCHIVE_ID
var COL = {
  DATE: 0,            // A — Дата виїзду
  FROM: 1,            // B — Адреса Відправки
  TO: 2,              // C — Адреса прибуття
  SEATS: 3,           // D — Кількість місць
  NAME: 4,            // E — ПіБ
  PHONE: 5,           // F — Телефон Пасажира
  MARK: 6,            // G — Відмітка (driver status)
  PAYMENT: 7,         // H — Оплата
  PERCENT: 8,         // I — Відсоток
  DISPATCHER: 9,      // J — Диспечер
  ID: 10,             // K — ІД
  PHONE_REG: 11,      // L — Телефон Реєстратора
  WEIGHT: 12,         // M — Вага
  VEHICLE: 13,        // N — Автомобіль
  TIMING: 14,         // O — Таймінг
  DATE_REG: 15,       // P — дата оформлення
  NOTE: 16,           // Q — Примітка
  STATUS: 17,         // R — Статус (CRM: new/work/archived/refused/deleted)
  DATE_ARCHIVE: 18,   // S — DATE_ARCHIVE
  ARCHIVED_BY: 19,    // T — ARCHIVED_BY
  ARCHIVE_REASON: 20, // U — ARCHIVE_REASON
  SOURCE_SHEET: 21,   // V — SOURCE_SHEET
  ARCHIVE_ID: 22      // W — ARCHIVE_ID
};
var TOTAL_COLS = 23;

// Заголовки для нового аркуша
var HEADERS = [
  'Дата виїзду', 'Адреса Відправки', 'Адреса прибуття', 'Кількість місць',
  'ПіБ', 'Телефон Пасажира', 'Відмітка', 'Оплата', 'Відсоток',
  'Диспечер', 'ІД', 'Телефон Реєстратора', 'Вага', 'Автомобіль',
  'Таймінг', 'дата оформлення', 'Примітка',
  'Статус', 'DATE_ARCHIVE', 'ARCHIVED_BY', 'ARCHIVE_REASON',
  'SOURCE_SHEET', 'ARCHIVE_ID'
];

// Статуси для архівації
var ARCHIVE_STATUSES = ['archived', 'refused', 'deleted', 'transferred'];

// Службові аркуші — НЕ маршрути
var EXCLUDE_SHEETS = ['Логи', 'Провірка розсилки'];

// Маппінг полів API → індексів колонок
var FIELD_MAP = {
  date: COL.DATE, from: COL.FROM, to: COL.TO, seats: COL.SEATS,
  name: COL.NAME, phone: COL.PHONE, mark: COL.MARK,
  payment: COL.PAYMENT, percent: COL.PERCENT, dispatcher: COL.DISPATCHER,
  id: COL.ID, phoneReg: COL.PHONE_REG, weight: COL.WEIGHT,
  vehicle: COL.VEHICLE, timing: COL.TIMING, dateReg: COL.DATE_REG,
  note: COL.NOTE, status: COL.STATUS,
  dateArchive: COL.DATE_ARCHIVE, archivedBy: COL.ARCHIVED_BY,
  archiveReason: COL.ARCHIVE_REASON, sourceSheet: COL.SOURCE_SHEET,
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
          service: 'Маршрути Пасажири — ЮРА ТРАНСПОРТЕЙШН',
          totalCols: TOTAL_COLS,
          timestamp: new Date().toISOString()
        });

      case 'getPassengers':
        if (!sheetParam) return respond({ success: false, error: 'Не вказано маршрут (sheet)' });
        return respond(getRoutePassengers({ sheetName: sheetParam }));

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
      // --- ЧИТАННЯ ---
      case 'getRoutePassengers':
        return respond(getRoutePassengers(payload));

      case 'getAvailableRoutes':
        return respond(getAvailableRoutes());

      case 'checkRouteSheets':
        return respond(checkRouteSheets(payload));

      // --- МАРШРУТИ: CRUD ---
      case 'copyToRoute':
        return respond(copyToRoute(payload));

      case 'createRouteSheet':
        return respond(createRouteSheet(payload));

      case 'deleteRouteSheet':
        return respond(deleteRouteSheet(payload));

      case 'deleteRoutePassenger':
        return respond(deleteRoutePassenger(payload));

      // --- ОНОВЛЕННЯ ---
      case 'updateField':
        return respond(updateField(payload));

      case 'updatePassenger':
        return respond(updatePassenger(payload));

      case 'updateMultiple':
        return respond(updateMultiple(payload));

      case 'updateStatus':
        return respond(updateStatus(payload));

      // --- ВОДІЙ: СТАТУС ---
      case 'updateDriverStatus':
      case 'updateStatus_driver':
        return respond(handleDriverStatusUpdate(data));

      // --- АРХІВАЦІЯ ---
      case 'archivePassengers':
        return respond(changeStatus(payload, 'archived'));

      case 'restorePassengers':
        return respond(changeStatus(payload, 'work'));

      case 'refusePassengers':
        return respond(changeStatus(payload, 'refused'));

      case 'deletePassengers':
        return respond(changeStatus(payload, 'deleted'));

      case 'archiveToExternal':
        return respond(archiveToExternal(payload));

      case 'deletePassengersPermanently':
        return respond(deletePassengersPermanently(payload));

      // --- ОПТИМІЗАЦІЯ ---
      case 'optimize':
        return respond(optimizeRoute(payload));

      // --- РОЗСИЛКА ---
      case 'getMailingStatus':
        return respond(getMailingStatus());

      case 'addMailingRecord':
        return respond(addMailingRecord(payload));

      case 'checkMailing':
        return respond(checkMailingByIds(payload));

      case 'clearMailing':
        return respond(clearMailing(payload));

      case 'clearOldMailing':
        return respond(clearOldMailing(payload));

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
// getRoutePassengers — Читання пасажирів маршруту
// + визначення driverStatus з Відмітка/кольору
// ============================================
function getRoutePassengers(payload) {
  try {
    var vehicleName = payload.vehicleName || '';
    var sheetName = payload.sheetName || vehicleName;
    if (!sheetName) {
      return { success: false, error: 'Не вказано аркуш маршруту' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, passengers: [], count: 0, sheetName: sheetName, stats: { total: 0 } };
    }

    var readCols = Math.min(sheet.getLastColumn(), TOTAL_COLS);
    var dataRange = sheet.getRange(2, 1, lastRow - 1, readCols);
    var data = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();
    var passengers = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!str(row[COL.NAME]) && !str(row[COL.PHONE])) continue;

      // Статус водія (з Відмітка + колір рядка)
      var driverStatus = resolveDriverStatus(row, backgrounds[i]);

      // CRM статус
      var crmStatus = str(row[COL.STATUS]).toLowerCase();

      // Пропускаємо архівовані якщо не запитано
      if (!payload.includeArchived && ARCHIVE_STATUSES.indexOf(crmStatus) !== -1) continue;

      passengers.push({
        rowNum: i + 2,
        date: formatDate(row[COL.DATE]),
        from: str(row[COL.FROM]),
        to: str(row[COL.TO]),
        seats: parseInt(row[COL.SEATS]) || 1,
        name: str(row[COL.NAME]),
        phone: str(row[COL.PHONE]),
        mark: str(row[COL.MARK]),
        payment: str(row[COL.PAYMENT]),
        percent: str(row[COL.PERCENT]),
        dispatcher: str(row[COL.DISPATCHER]),
        id: str(row[COL.ID]),
        phoneReg: str(row[COL.PHONE_REG]),
        weight: str(row[COL.WEIGHT]),
        vehicle: str(row[COL.VEHICLE]),
        timing: str(row[COL.TIMING]),
        dateReg: str(row[COL.DATE_REG]),
        note: str(row[COL.NOTE]),
        status: crmStatus || 'new',
        archiveId: readCols > COL.ARCHIVE_ID ? str(row[COL.ARCHIVE_ID]) : '',
        driverStatus: driverStatus,
        rowColor: backgrounds[i][0],
        sheet: sheetName
      });
    }

    // Статистика
    var stats = { total: passengers.length, pending: 0, inProgress: 0, completed: 0, cancelled: 0, archived: 0 };
    for (var j = 0; j < passengers.length; j++) {
      var ds = passengers[j].driverStatus;
      if (ds === 'pending') stats.pending++;
      else if (ds === 'in-progress') stats.inProgress++;
      else if (ds === 'completed') stats.completed++;
      else if (ds === 'cancelled') stats.cancelled++;
      else if (ds === 'archived') stats.archived++;
    }

    return {
      success: true,
      passengers: passengers,
      count: passengers.length,
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

    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();

      // Пропускаємо службові
      var isExcluded = false;
      for (var e = 0; e < EXCLUDE_SHEETS.length; e++) {
        if (name === EXCLUDE_SHEETS[e]) { isExcluded = true; break; }
      }
      if (isExcluded) continue;

      var count = Math.max(0, sheets[i].getLastRow() - 1);
      routes.push({
        name: name,
        type: 'passenger',
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
// checkRouteSheets — Перевірка чи є дані
// ============================================
function checkRouteSheets(payload) {
  try {
    var vehicleNames = payload.vehicleNames || [];
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var existing = [];

    for (var i = 0; i < vehicleNames.length; i++) {
      var sheetName = vehicleNames[i];
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
// copyToRoute — CRM → маршрутний аркуш
// ============================================
function copyToRoute(payload) {
  try {
    var passengersByVehicle = payload.passengersByVehicle;
    var conflictAction = payload.conflictAction || 'add';

    if (!passengersByVehicle) {
      return { success: false, error: 'Немає пасажирів' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var totalCopied = 0;
    var totalArchived = 0;
    var totalCleared = 0;
    var results = [];

    for (var vehicleName in passengersByVehicle) {
      if (!passengersByVehicle.hasOwnProperty(vehicleName)) continue;
      var passengers = passengersByVehicle[vehicleName];
      if (!passengers || !passengers.length) continue;

      // Знаходимо аркуш (точна назва = vehicleName)
      var sheetName = vehicleName;
      var sheet = ss.getSheetByName(sheetName);

      // Створюємо якщо не існує
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
          // Ставимо статус "archived" на старі рядки
          var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
          var archiveRange = sheet.getRange(2, COL.STATUS + 1, lastRow - 1, 1);
          var archVals = [];
          for (var a = 0; a < lastRow - 1; a++) archVals.push(['archived']);
          archiveRange.setValues(archVals);
          var dateRange = sheet.getRange(2, COL.DATE_ARCHIVE + 1, lastRow - 1, 1);
          var dateVals = [];
          for (var d = 0; d < lastRow - 1; d++) dateVals.push([dateNow]);
          dateRange.setValues(dateVals);
        }
      }

      // Записуємо пасажирів
      var rows = [];
      for (var p = 0; p < passengers.length; p++) {
        var pass = passengers[p];
        var newRow = new Array(TOTAL_COLS);
        for (var c = 0; c < TOTAL_COLS; c++) newRow[c] = '';

        newRow[COL.DATE] = pass.date || '';
        newRow[COL.FROM] = pass.from || '';
        newRow[COL.TO] = pass.to || '';
        newRow[COL.SEATS] = pass.seats || 1;
        newRow[COL.NAME] = pass.name || '';
        newRow[COL.PHONE] = pass.phone || '';
        newRow[COL.MARK] = pass.mark || '';
        newRow[COL.PAYMENT] = pass.payment || '';
        newRow[COL.PERCENT] = pass.percent || '';
        newRow[COL.DISPATCHER] = pass.dispatcher || '';
        newRow[COL.ID] = pass.id || '';
        newRow[COL.PHONE_REG] = pass.phoneReg || '';
        newRow[COL.WEIGHT] = pass.weight || '';
        newRow[COL.VEHICLE] = vehicleName;
        newRow[COL.TIMING] = pass.timing || '';
        newRow[COL.DATE_REG] = pass.dateReg || '';
        newRow[COL.NOTE] = pass.note || '';
        newRow[COL.STATUS] = 'new';
        newRow[COL.SOURCE_SHEET] = pass.sourceSheet || pass.sheet || '';

        rows.push(newRow);
      }

      if (rows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, TOTAL_COLS).setValues(rows);

        // Фарбуємо pending
        var pendingColors = STATUS_COLORS['pending'];
        if (pendingColors) {
          sheet.getRange(startRow, 1, rows.length, TOTAL_COLS).setBackground(pendingColors.bg);
        }
        totalCopied += rows.length;
      }

      results.push({ vehicle: vehicleName, sheet: sheetName, copied: passengers.length });
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
// createRouteSheet — Створити маршрутний аркуш
// ============================================
function createRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var existing = ss.getSheetByName(vehicleName);

    if (existing) {
      return { success: true, sheetName: vehicleName, existed: true };
    }

    var sheet = ss.insertSheet(vehicleName);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);

    writeLog('createRouteSheet', vehicleName, 0, 'created', '');

    return { success: true, sheetName: vehicleName, vehicleName: vehicleName };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// deleteRouteSheet — Видалити маршрутний аркуш
// ============================================
function deleteRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName || payload.sheetName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(vehicleName);

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
    writeLog('deleteRouteSheet', vehicleName, 0, 'deleted', '');

    return { success: true, sheetName: vehicleName, deleted: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// deleteRoutePassenger — Видалити одного пасажира
// ============================================
function deleteRoutePassenger(payload) {
  try {
    var sheetName = payload.sheetName;
    var rowNum = parseInt(payload.rowNum);

    if (!sheetName || !rowNum || rowNum < 2) {
      return { success: false, error: 'Відсутні sheetName або rowNum' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Аркуш не знайдено' };
    if (rowNum > sheet.getLastRow()) return { success: false, error: 'Рядок не існує' };

    // Верифікація
    if (payload.expectedId) {
      var currentId = str(sheet.getRange(rowNum, COL.ID + 1).getValue());
      if (currentId !== String(payload.expectedId).trim()) {
        return { success: false, error: 'conflict', message: 'ІД не збігається' };
      }
    }

    sheet.deleteRow(rowNum);
    writeLog('deleteRoutePassenger', sheetName, rowNum, 'deleted', '');

    return { success: true, deleted: true, rowNum: rowNum, sheetName: sheetName };
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

  if (payload.expectedId) {
    var currentId = str(sheet.getRange(rowNum, COL.ID + 1).getValue());
    if (currentId !== String(payload.expectedId).trim()) {
      return { success: false, error: 'conflict', message: 'ІД не збігається' };
    }
  }

  sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(value);
  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}

// ============================================
// updatePassenger — Оновити кілька полів
// ============================================
function updatePassenger(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(payload.sheet);
    if (!sheet) return { success: false, error: 'Аркуш не знайдено' };

    var rowNum = parseInt(payload.rowNum);
    if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
      return { success: false, error: 'Невірний rowNum' };
    }

    if (payload.expectedId) {
      var currentId = str(sheet.getRange(rowNum, COL.ID + 1).getValue());
      if (currentId !== String(payload.expectedId).trim()) {
        return { success: false, error: 'conflict', message: 'ІД не збігається' };
      }
    }

    var updated = [];
    for (var field in payload) {
      if (payload.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
        if (payload[field] !== undefined) {
          sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(payload[field]);
          updated.push(field);
        }
      }
    }

    return { success: true, updated: updated, sheet: payload.sheet, rowNum: rowNum };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// updateMultiple — Масове оновлення
// ============================================
function updateMultiple(payload) {
  var items = payload.passengers || payload.items || [];
  var updated = 0;
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var result = updatePassenger(items[i]);
    if (result.success) updated++;
    else errors.push((items[i].id || 'unknown') + ': ' + result.error);
  }

  return { success: true, updated: updated, errors: errors.length > 0 ? errors : undefined };
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
  if (!sheet || rowNum > sheet.getLastRow()) return { success: false, error: 'Аркуш/рядок не знайдено' };

  var oldStatus = str(sheet.getRange(rowNum, COL.STATUS + 1).getValue());
  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
    sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
      Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
    );
  }

  return { success: true, sheet: sheetName, rowNum: rowNum, status: newStatus, oldStatus: oldStatus };
}

// ============================================
// handleDriverStatusUpdate — Водій оновлює статус
// ============================================
function handleDriverStatusUpdate(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Логуємо
    var logSheet = ss.getSheetByName(SHEET_LOGS);
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOGS);
      logSheet.getRange(1, 1, 1, 8).setValues([[
        'Дата', 'Час', 'Водій', 'Маршрут', 'ІД пасажира', 'Адреса', 'Статус', 'Причина'
      ]]);
      logSheet.getRange(1, 1, 1, 8)
        .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
    }

    var now = new Date();
    logSheet.appendRow([
      Utilities.formatDate(now, 'Europe/Kiev', 'yyyy-MM-dd'),
      Utilities.formatDate(now, 'Europe/Kiev', 'HH:mm:ss'),
      data.driverId || '',
      data.routeName || '',
      data.passengerId || data.deliveryNumber || '',
      data.address || '',
      data.status || '',
      data.cancelReason || ''
    ]);

    // Оновлюємо в маршруті
    var routeSheet = ss.getSheetByName(data.routeName);
    if (!routeSheet) {
      return { success: true, message: 'Логовано (маршрут не знайдено)' };
    }

    var allData = routeSheet.getDataRange().getValues();
    var searchId = str(data.passengerId || data.deliveryNumber);
    var rowsUpdated = 0;

    for (var i = 1; i < allData.length; i++) {
      var rowId = str(allData[i][COL.ID]);
      var rowPhone = str(allData[i][COL.PHONE]);

      // Пошук по ІД або телефону
      if ((searchId && rowId === searchId) || (data.phone && rowPhone === str(data.phone))) {
        var rowNum = i + 1;

        // Записуємо статус у Відмітка (G)
        routeSheet.getRange(rowNum, COL.MARK + 1).setValue(data.status);

        // Причина скасування → Примітка
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

          var markCell = routeSheet.getRange(rowNum, COL.MARK + 1);
          markCell.setFontColor(colors.font);
          markCell.setFontWeight('bold');
        }

        rowsUpdated++;
      }
    }

    return {
      success: true,
      message: rowsUpdated > 0 ? 'Статус записано' : 'Пасажира не знайдено в маршруті',
      updatedRows: rowsUpdated
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// changeStatus — Масова зміна CRM статусу
// ============================================
function changeStatus(payload, newStatus) {
  try {
    var items = payload.passengers || payload.items || [];
    var note = payload.note || '';
    if (items.length === 0) return { success: false, error: 'Немає пасажирів' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var user = payload.user || 'crm';
    var changed = 0;
    var errors = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var sheetName = item.sheet || item.sheetName;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) { errors.push(sheetName + ': не знайдено'); continue; }

      var rowNum = parseInt(item.rowNum);
      if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
        errors.push('Рядок ' + rowNum + ': не існує'); continue;
      }

      sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

      if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
        sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
        sheet.getRange(rowNum, COL.ARCHIVED_BY + 1).setValue(user);
        if (note) {
          sheet.getRange(rowNum, COL.ARCHIVE_REASON + 1).setValue(note);
        }
      }

      if (note) {
        var currentNote = str(sheet.getRange(rowNum, COL.NOTE + 1).getValue());
        var updatedNote = note + (currentNote ? ' | ' + currentNote : '');
        sheet.getRange(rowNum, COL.NOTE + 1).setValue(updatedNote);
      }

      changed++;
    }

    writeLog('changeStatus:' + newStatus, 'bulk', 0, changed + '/' + items.length, note);

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
    var items = payload.items || payload.passengers || [];
    var user = payload.user || 'crm';
    var reason = payload.reason || 'route_archive';
    if (items.length === 0) return { success: false, error: 'Немає записів' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var archivedIds = [];
    var count = 0;
    var errors = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var sheetName = item.sheet || item.sheetName;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) { errors.push(sheetName + ': не знайдено'); continue; }

      var rowNum = parseInt(item.rowNum);
      if (rowNum > sheet.getLastRow()) { errors.push('Рядок ' + rowNum); continue; }

      var existingArchiveId = str(sheet.getRange(rowNum, COL.ARCHIVE_ID + 1).getValue());
      if (existingArchiveId) {
        errors.push('Рядок ' + rowNum + ': вже архівовано'); continue;
      }

      var recordId = str(sheet.getRange(rowNum, COL.ID + 1).getValue());

      sheet.getRange(rowNum, COL.STATUS + 1).setValue('archived');
      sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
      sheet.getRange(rowNum, COL.ARCHIVED_BY + 1).setValue(user);
      sheet.getRange(rowNum, COL.ARCHIVE_REASON + 1).setValue(reason);

      if (recordId) archivedIds.push(recordId);
      count++;
    }

    var archiveResult = { success: false, error: 'Немає ІД' };
    if (archivedIds.length > 0) {
      archiveResult = sendToArchive({
        action: 'archiveByIds',
        source: 'ROUTE_PASAZHYRY',
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
// deletePassengersPermanently — Фізичне видалення
// ============================================
function deletePassengersPermanently(payload) {
  try {
    var items = payload.passengers || payload.items || [];
    if (items.length === 0) return { success: false, error: 'Немає записів' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var deleted = 0;

    var bySheet = {};
    for (var i = 0; i < items.length; i++) {
      var sheetName = items[i].sheet || items[i].sheetName;
      if (!bySheet[sheetName]) bySheet[sheetName] = [];
      bySheet[sheetName].push(parseInt(items[i].rowNum));
    }

    for (var sn in bySheet) {
      if (!bySheet.hasOwnProperty(sn)) continue;
      var sheet = ss.getSheetByName(sn);
      if (!sheet) continue;
      var rows = bySheet[sn].sort(function(a, b) { return b - a; });
      for (var d = 0; d < rows.length; d++) {
        if (rows[d] >= 2 && rows[d] <= sheet.getLastRow()) {
          sheet.deleteRow(rows[d]);
          deleted++;
        }
      }
    }

    writeLog('deletePermanently', 'bulk', 0, deleted + ' видалено', '');
    return { success: true, deleted: deleted };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// ОПТИМІЗАЦІЯ МАРШРУТУ (Google Maps)
// ============================================
function optimizeRoute(payload) {
  try {
    var passengersData = payload.passengers;
    var optimizeBy = payload.optimizeBy || 'from';
    var startAddress = payload.startAddress || '';
    var endAddress = payload.endAddress || '';

    if (!passengersData || passengersData.length === 0) {
      return { success: false, error: 'Немає пасажирів' };
    }

    var addressField = (optimizeBy === 'to') ? 'to' : 'from';
    var passengers = [];

    for (var i = 0; i < passengersData.length; i++) {
      var p = passengersData[i];
      var rawAddress = p[addressField] || '';
      if (rawAddress && rawAddress.trim().length > 0) {
        passengers.push({
          index: i,
          originalData: p,
          rawAddress: rawAddress,
          cleanAddress: rawAddress.replace(/\n/g, ' ').replace(/\r/g, ' ').replace(/\s+/g, ' ').trim(),
          coords: null,
          id: p.id || '',
          name: p.name || '',
          uid: p._uid || null
        });
      }
    }

    if (passengers.length === 0) {
      return { success: false, error: 'Немає адрес' };
    }

    // Геокодування
    var notGeocodedList = [];
    var geocodedCount = 0;
    for (var g = 0; g < passengers.length; g++) {
      try {
        var coords = geocodeAddress(passengers[g].cleanAddress);
        if (coords) { passengers[g].coords = coords; geocodedCount++; }
        else { notGeocodedList.push({ id: passengers[g].id, name: passengers[g].name, address: passengers[g].rawAddress, uid: passengers[g].uid }); }
      } catch (e) {
        notGeocodedList.push({ id: passengers[g].id, name: passengers[g].name, address: passengers[g].rawAddress, uid: passengers[g].uid });
      }
      Utilities.sleep(150);
    }

    var startCoords = { lat: DEFAULT_START.lat, lng: DEFAULT_START.lng, name: DEFAULT_START.name };
    if (startAddress) {
      var sc = geocodeAddress(startAddress);
      if (sc) startCoords = { lat: sc.lat, lng: sc.lng, name: startAddress };
    }

    var endCoords = null;
    if (endAddress) {
      var ec = geocodeAddress(endAddress);
      if (ec) endCoords = { lat: ec.lat, lng: ec.lng, name: endAddress };
    }

    var validPassengers = [];
    var invalidPassengers = [];
    for (var v = 0; v < passengers.length; v++) {
      if (passengers[v].coords) validPassengers.push(passengers[v]);
      else invalidPassengers.push(passengers[v]);
    }

    if (validPassengers.length === 0) {
      return { success: false, error: 'Жодну адресу не вдалось геокодувати' };
    }

    var optimizedOrder = null;
    var method = 'Google Directions API';

    if (validPassengers.length <= 23) {
      optimizedOrder = optimizeWithDirectionsAPI(validPassengers, startCoords, endCoords);
    }
    if (!optimizedOrder || optimizedOrder.length === 0) {
      optimizedOrder = optimizeNearestNeighbor(validPassengers, startCoords);
      method = 'Nearest Neighbor';
    }

    var orderedPassengers = [];
    var orderedForMap = [];
    for (var o = 0; o < optimizedOrder.length; o++) {
      orderedPassengers.push(validPassengers[optimizedOrder[o]].originalData);
      orderedForMap.push(validPassengers[optimizedOrder[o]]);
    }
    for (var inv = 0; inv < invalidPassengers.length; inv++) {
      var dd = invalidPassengers[inv].originalData;
      dd._notGeocoded = true;
      orderedPassengers.push(dd);
    }

    var mapLinks = generateMapLinks(orderedForMap, startCoords, endCoords);

    return {
      success: true,
      stats: { total: passengersData.length, geocoded: geocodedCount, optimized: optimizedOrder.length, notGeocoded: notGeocodedList.length },
      optimizeBy: (optimizeBy === 'to') ? 'Адреса ПРИБУТТЯ' : 'Адреса ВІДПРАВКИ',
      start: startCoords.name,
      end: endCoords ? endCoords.name : 'остання точка',
      method: method,
      orderedPassengers: orderedPassengers,
      notGeocodedList: notGeocodedList,
      mapLinks: mapLinks
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// РОЗСИЛКА — "Провірка розсилки"
// Та сама логіка що і в маршрутах посилок:
// activeIds (без ARCHIVE_ID) + clear/clearOld
// ============================================

var MAILING_COL = {
  ID: 0, USER: 1, INFO: 2, STATUS: 3,
  DATE_ARCHIVE: 4, ARCHIVED_BY: 5, ARCHIVE_REASON: 6,
  SOURCE_SHEET: 7, ARCHIVE_ID: 8
};
var MAILING_COLS = 9;

function getMailingStatus() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, mailingIds: [], activeIds: [], count: 0 };
    }

    var lastRow = sheet.getLastRow();
    var readCols = Math.min(sheet.getLastColumn(), MAILING_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();

    var activeIds = [];
    var archivedIds = [];
    var allData = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var id = str(row[MAILING_COL.ID]);
      if (!id || id.indexOf('dd.mm') !== -1) continue;

      var archiveId = readCols > MAILING_COL.ARCHIVE_ID ? str(row[MAILING_COL.ARCHIVE_ID]) : '';
      var status = readCols > MAILING_COL.STATUS ? str(row[MAILING_COL.STATUS]) : '';
      var isArchived = archiveId !== '' || status === 'archived';

      allData.push({ rowNum: i + 2, id: id, isArchived: isArchived });

      if (isArchived) archivedIds.push(id);
      else activeIds.push(id);
    }

    var uniqueActive = unique(activeIds);
    var allIds = unique(allData.map(function(d) { return d.id; }));

    return {
      success: true,
      mailingIds: allIds,
      activeIds: uniqueActive,
      archivedIds: archivedIds,
      mailingData: allData,
      count: allIds.length,
      activeCount: uniqueActive.length
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function checkMailingByIds(payload) {
  try {
    var idsToCheck = payload.ids || [];
    if (idsToCheck.length === 0) return { success: true, results: {}, count: 0 };

    var statusResult = getMailingStatus();
    if (!statusResult.success) return statusResult;

    var activeSet = toSet(statusResult.activeIds);
    var allSet = toSet(statusResult.mailingIds);

    var results = {};
    var mailedCount = 0;
    for (var i = 0; i < idsToCheck.length; i++) {
      var id = String(idsToCheck[i]).trim();
      var isActive = activeSet[id] === true;
      results[id] = {
        mailed: isActive,
        wasMailedBefore: allSet[id] === true,
        archivedMailing: allSet[id] === true && !isActive
      };
      if (isActive) mailedCount++;
    }

    return { success: true, results: results, total: idsToCheck.length, mailed: mailedCount };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function addMailingRecord(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_MAILING);
      sheet.getRange(1, 1, 1, MAILING_COLS).setValues([[
        'ІД', 'Користувач', 'Дата', 'Статус',
        'DATE_ARCHIVE', 'ARCHIVED_BY', 'ARCHIVE_REASON', 'SOURCE_SHEET', 'ARCHIVE_ID'
      ]]);
    }

    var records = payload.records || [];
    var userName = payload.userName || 'CRM';
    if (records.length === 0) return { success: false, error: 'Немає записів' };

    var date = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');
    var rows = [];
    for (var i = 0; i < records.length; i++) {
      var newRow = new Array(MAILING_COLS);
      for (var c = 0; c < MAILING_COLS; c++) newRow[c] = '';
      newRow[MAILING_COL.ID] = records[i].id || records[i].passengerId || '';
      newRow[MAILING_COL.USER] = userName;
      newRow[MAILING_COL.INFO] = date;
      newRow[MAILING_COL.STATUS] = records[i].status || 'sent';
      rows.push(newRow);
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, MAILING_COLS).setValues(rows);
    return { success: true, added: rows.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function clearMailing(payload) {
  try {
    var mode = payload.mode || 'archived';
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, deleted: 0 };

    if (mode === 'all') {
      var total = sheet.getLastRow() - 1;
      sheet.deleteRows(2, total);
      writeLog('clearMailing', SHEET_MAILING, 0, 'all', total + ' рядків');
      return { success: true, deleted: total, mode: 'all' };
    }

    var lastRow = sheet.getLastRow();
    var readCols = Math.min(sheet.getLastColumn(), MAILING_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();
    var rowsToDelete = [];

    for (var i = data.length - 1; i >= 0; i--) {
      var archiveId = readCols > MAILING_COL.ARCHIVE_ID ? str(data[i][MAILING_COL.ARCHIVE_ID]) : '';
      var status = readCols > MAILING_COL.STATUS ? str(data[i][MAILING_COL.STATUS]) : '';
      if (archiveId || status === 'archived') rowsToDelete.push(i + 2);
    }

    for (var d = 0; d < rowsToDelete.length; d++) sheet.deleteRow(rowsToDelete[d]);
    writeLog('clearMailing', SHEET_MAILING, 0, 'archived', rowsToDelete.length + ' рядків');

    return { success: true, deleted: rowsToDelete.length, remaining: (lastRow - 1) - rowsToDelete.length, mode: 'archived' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function clearOldMailing(payload) {
  try {
    var days = parseInt(payload.days) || 30;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, deleted: 0 };

    var lastRow = sheet.getLastRow();
    var readCols = Math.min(sheet.getLastColumn(), MAILING_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);

    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 0; i--) {
      var dateStr = str(data[i][MAILING_COL.INFO]);
      var recordDate = parseDate(dateStr);
      if (recordDate && recordDate < cutoff) rowsToDelete.push(i + 2);
    }

    for (var d = 0; d < rowsToDelete.length; d++) sheet.deleteRow(rowsToDelete[d]);
    writeLog('clearOldMailing', SHEET_MAILING, 0, '>' + days + 'д', rowsToDelete.length + ' рядків');

    return { success: true, deleted: rowsToDelete.length, remaining: (lastRow - 1) - rowsToDelete.length, days: days };
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
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    result.push({ sheet: sheet.getName(), rows: sheet.getLastRow(), cols: lastCol, headers: headers });
  }
  return { success: true, sheets: result };
}

// ============================================
// АРХІВ API (HTTP до Crm_Arhiv_1.0)
// ============================================
function sendToArchive(payload) {
  try {
    var response = UrlFetchApp.fetch(ARCHIVE_API_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      try { return JSON.parse(response.getContentText()); }
      catch (e) { return { success: false, error: 'Невалідна відповідь' }; }
    }
    return { success: false, error: 'HTTP ' + response.getResponseCode() };
  } catch (e) {
    return { success: false, error: 'Архів недоступний: ' + e.toString() };
  }
}

// ============================================
// ЛОГУВАННЯ
// ============================================
function writeLog(action, sheetName, rowNum, detail, extra) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(SHEET_LOGS);
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOGS);
      logSheet.appendRow(['Дата/Час', 'Дія', 'Аркуш', 'Рядок', 'Деталі', 'Дані']);
      logSheet.getRange(1, 1, 1, 6)
        .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }
    var timestamp = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([timestamp, action, sheetName, rowNum, detail, extra || '']);
  } catch (e) {
    Logger.log('Log error: ' + e.toString());
  }
}

// ============================================
// ДОПОМІЖНІ ФУНКЦІЇ
// ============================================

// Визначення статусу водія (з Відмітка + колір)
function resolveDriverStatus(row, bgColors) {
  var markValue = str(row[COL.MARK]).toLowerCase();

  if (markValue === 'completed' || markValue === 'готово') return 'completed';
  if (markValue === 'in-progress' || markValue === 'в процесі') return 'in-progress';
  if (markValue === 'cancelled' || markValue === 'відмова' || markValue === 'скасовано') return 'cancelled';
  if (markValue === 'archived' || markValue === 'архів') return 'archived';

  if (bgColors && bgColors[0]) {
    var color = bgColors[0].toLowerCase();
    if (color === '#00ff00' || color === '#b6d7a8' || color === '#93c47d') return 'completed';
    if (color === '#6fa8dc' || color === '#a4c2f4' || color === '#3d85c6') return 'in-progress';
    if (color === '#e06666' || color === '#ea9999' || color === '#cc0000') return 'cancelled';
  }

  return 'pending';
}

// Безпечне перетворення в string
function str(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

// Форматування дати
function formatDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(s)) {
    var p = s.split('.');
    return p[2] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[0]).slice(-2);
  }
  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd');
  } catch (e) {}
  return '';
}

// Парсинг дати (dd.MM.yyyy або yyyy-MM-dd)
function parseDate(dateStr) {
  if (!dateStr) return null;
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(dateStr)) {
    var p = dateStr.split('.');
    return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) return new Date(dateStr.substring(0, 10));
  try { var d = new Date(dateStr); if (!isNaN(d.getTime())) return d; } catch (e) {}
  return null;
}

// Унікальні значення масиву
function unique(arr) {
  var seen = {};
  var result = [];
  for (var i = 0; i < arr.length; i++) {
    if (!seen[arr[i]]) { seen[arr[i]] = true; result.push(arr[i]); }
  }
  return result;
}

// Масив → Set (об'єкт)
function toSet(arr) {
  var s = {};
  for (var i = 0; i < arr.length; i++) s[arr[i]] = true;
  return s;
}

// Геокодування
function geocodeAddress(address) {
  try {
    var url = 'https://maps.googleapis.com/maps/api/geocode/json'
      + '?address=' + encodeURIComponent(address) + '&key=' + API_KEY + '&language=uk';
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());
    if (json.status === 'OK' && json.results && json.results.length > 0) {
      var loc = json.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
    return null;
  } catch (e) { return null; }
}

// Directions API оптимізація
function optimizeWithDirectionsAPI(passengers, startCoords, endCoords) {
  try {
    var allCoords = [];
    for (var i = 0; i < passengers.length; i++)
      allCoords.push(passengers[i].coords.lat + ',' + passengers[i].coords.lng);
    if (allCoords.length === 0) return [];
    if (allCoords.length === 1) return [0];

    var origin = startCoords.lat + ',' + startCoords.lng;
    var destination, waypoints;
    if (endCoords) {
      destination = endCoords.lat + ',' + endCoords.lng;
      waypoints = allCoords.slice();
    } else {
      waypoints = allCoords.slice(0, allCoords.length - 1);
      destination = allCoords[allCoords.length - 1];
    }

    var url = 'https://maps.googleapis.com/maps/api/directions/json'
      + '?origin=' + encodeURIComponent(origin) + '&destination=' + encodeURIComponent(destination)
      + '&key=' + API_KEY + '&language=uk';
    if (waypoints.length > 0) url += '&waypoints=optimize:true|' + waypoints.join('|');

    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());
    if (json.status !== 'OK') return null;

    var waypointOrder = json.routes[0].waypoint_order;
    if (endCoords) return waypointOrder;
    var result = waypointOrder.slice();
    result.push(passengers.length - 1);
    return result;
  } catch (e) { return null; }
}

// Nearest Neighbor fallback
function optimizeNearestNeighbor(passengers, startCoords) {
  var n = passengers.length;
  if (n === 0) return [];
  if (n === 1) return [0];

  var currentIdx = 0;
  var minDist = Infinity;
  for (var i = 0; i < n; i++) {
    var dist = haversine(startCoords, passengers[i].coords);
    if (dist < minDist) { minDist = dist; currentIdx = i; }
  }

  var visited = [];
  for (var v = 0; v < n; v++) visited.push(false);
  var tour = [currentIdx];
  visited[currentIdx] = true;

  for (var step = 1; step < n; step++) {
    var nearest = -1;
    var nearestDist = Infinity;
    for (var j = 0; j < n; j++) {
      if (!visited[j]) {
        var d = haversine(passengers[currentIdx].coords, passengers[j].coords);
        if (d < nearestDist) { nearestDist = d; nearest = j; }
      }
    }
    if (nearest === -1) break;
    tour.push(nearest);
    visited[nearest] = true;
    currentIdx = nearest;
  }
  return tour;
}

// Google Maps посилання
function generateMapLinks(orderedPassengers, startCoords, endCoords) {
  var links = [];
  if (orderedPassengers.length === 0) return links;

  var chunkStart = 0;
  while (chunkStart < orderedPassengers.length) {
    var chunkEnd = Math.min(chunkStart + MAX_POINTS_PER_MAP - 1, orderedPassengers.length);
    var chunkItems = orderedPassengers.slice(chunkStart, chunkEnd);

    var origin = chunkStart === 0 ? startCoords.name : orderedPassengers[chunkStart - 1].cleanAddress;
    var destination, waypointItems;
    if (chunkEnd >= orderedPassengers.length && endCoords) {
      destination = endCoords.name;
      waypointItems = chunkItems;
    } else {
      destination = chunkItems[chunkItems.length - 1].cleanAddress;
      waypointItems = chunkItems.slice(0, chunkItems.length - 1);
    }

    var url = 'https://www.google.com/maps/dir/' + encodeURIComponent(origin);
    for (var w = 0; w < waypointItems.length; w++)
      url += '/' + encodeURIComponent(waypointItems[w].cleanAddress);
    url += '/' + encodeURIComponent(destination);

    links.push({ url: url, from: chunkStart + 1, to: chunkEnd, total: orderedPassengers.length });
    chunkStart = chunkEnd;
  }
  return links;
}

// Haversine
function haversine(coord1, coord2) {
  var R = 6371;
  var dLat = (coord2.lat - coord1.lat) * Math.PI / 180;
  var dLng = (coord2.lng - coord1.lng) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(coord1.lat * Math.PI / 180) * Math.cos(coord2.lat * Math.PI / 180) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// JSON відповідь
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// МЕНЮ
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Маршрути Пасажири')
    .addItem('Список маршрутів', 'menuRoutes')
    .addItem('Структура', 'menuStructure')
    .addItem('Тест архів', 'menuTestArchive')
    .addToUi();
}

function menuRoutes() {
  var result = getAvailableRoutes();
  var msg = 'Маршрутів: ' + result.count + '\n\n';
  for (var i = 0; i < result.routes.length; i++) {
    msg += result.routes[i].name + ' — ' + result.routes[i].count + ' пас.\n';
  }
  SpreadsheetApp.getUi().alert('Маршрути', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuStructure() {
  var result = getStructure();
  for (var i = 0; i < result.sheets.length; i++) {
    Logger.log('[' + result.sheets[i].sheet + '] ' + result.sheets[i].rows + 'r, ' + result.sheets[i].cols + 'c');
  }
  SpreadsheetApp.getUi().alert('Дивись Logger');
}

function menuTestArchive() {
  var result = sendToArchive({ action: 'getStats' });
  SpreadsheetApp.getUi().alert('Архів',
    result.success ? 'OK' : 'ПОМИЛКА: ' + result.error,
    SpreadsheetApp.getUi().ButtonSet.OK);
}
