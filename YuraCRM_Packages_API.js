// ============================================
// YURA TRANSPORTEYSHN — CRM POSYLKY
// Apps Script API для таблиці "Бот Посилки"
// ID: 1RyWJ-ZQ-OQbeD65fZXR-WEwP_kwuNllikiA3Q1rjtlo
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Бот Посилки" → Розширення → Apps Script
// 2. Видали весь старий код і встав цей файл
// 3. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Скопіюй URL деплоя
// 5. Встав URL в HTML файл замість YOUR_PACKAGES_API_URL_HERE
// ============================================

// ============================================
// КОНФІГУРАЦІЯ
// ============================================

// Назви аркушів — ТОЧНО як в таблиці
var SHEET_REG = 'Реєстрація ТТН';     // UA→EU посилки
var SHEET_COURIER = 'Виклик курєра';   // EU→UA посилки
var SHEET_LOGS = 'Логи';                // Логування дій

// ВАЖЛИВО: Якщо назва аркуша містить апостроф (Виклик кур'єра),
// розкоментуй рядок нижче і закоментуй рядок вище:
// var SHEET_COURIER = "Виклик кур\u0027єра";

// Порядок колонок (A-V = 22 колонки, індекс 0-21)
// A:ВО  B:Номер№  C:Номер ТТН  D:Вага  E:Адреса Отримувача  F:Напрямок
// G:Телефон Отримувача  H:Сума Є  I:Статус оплати  J:Оплата
// K:Телефон Реєстратора  L:Примітка  M:Статус посилки  N:ІД  O:ПіБ
// P:дата оформлення  Q:Таймінг  R:Примітка смс  S:Дата отримання
// T:Фото  U:Статус  V:Дата архів
var COL = {
  VO: 0,            // A — ВО (менеджер: Д, Ш, Б)
  NUMBER: 1,        // B — Номер№
  TTN: 2,           // C — Номер ТТН
  WEIGHT: 3,        // D — Вага
  ADDRESS: 4,       // E — Адреса Отримувача
  DIRECTION: 5,     // F — Напрямок
  PHONE: 6,         // G — Телефон Отримувача
  AMOUNT: 7,        // H — Сума Є
  PAY_STATUS: 8,    // I — Статус оплати
  PAYMENT: 9,       // J — Оплата
  PHONE_REG: 10,    // K — Телефон Реєстратора
  NOTE: 11,         // L — Примітка
  PARCEL_STATUS: 12,// M — Статус посилки (для клієнта: Невідомий/Зареєстровано/Оформлено/Кордон/Доставка)
  ID: 13,           // N — ІД
  NAME: 14,         // O — ПіБ
  DATE_REG: 15,     // P — дата оформлення
  TIMING: 16,       // Q — Таймінг
  SMS_NOTE: 17,     // R — Примітка смс
  DATE_RECEIVE: 18, // S — Дата отримання
  PHOTO: 19,        // T — Фото
  STATUS: 20,       // U — Статус (CRM: new/work/route/archived/refused/transferred/deleted)
  DATE_ARCHIVE: 21  // V — Дата архів
};
var TOTAL_COLS = 22;

// Маппінг напрямок → аркуш (ключова прив'язка!)
// UA→EU ліди падають у "Реєстрація ТТН"
// EU→UA ліди падають у "Виклик кур'єра"
function getSheetByDirection(direction) {
  if (direction === 'eu-ua') return SHEET_COURIER;
  return SHEET_REG; // за замовчуванням ua-eu
}

// Маппінг аркуш → напрямок
function getDirectionBySheet(sheetName) {
  if (sheetName === SHEET_COURIER) return 'eu-ua';
  return 'ua-eu';
}

// ============================================
// ГОЛОВНИЙ ОБРОБНИК — doPost
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    switch (action) {
      case 'getAll':
        return respond(getAllPackages());

      case 'getStructure':
        return respond(getStructure());

      case 'updatePackage':
        return respond(updatePackage(data));

      case 'addPackage':
        return respond(addPackage(data));

      case 'deletePackage':
        return respond(deletePackage(data));

      case 'updateStatus':
        return respond(updateStatus(data));

      case 'bulkUpdateStatus':
        return respond(bulkUpdateStatus(data));

      case 'updateField':
        return respond(updateField(data));

      case 'bulkAssignVehicle':
        return respond(bulkAssignVehicle(data));

      default:
        return respond({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// МАППІНГ назв полів CRM → індексів колонок таблиці
// ============================================
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
  dateArchive: COL.DATE_ARCHIVE
};

// ============================================
// getAll — Витягнути ВСІ посилки з обох аркушів
// ============================================
function getAllPackages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allPackages = [];

  // Читаємо обидва аркуші
  var sheetsToRead = [
    { name: SHEET_REG, direction: 'ua-eu' },
    { name: SHEET_COURIER, direction: 'eu-ua' }
  ];

  for (var s = 0; s < sheetsToRead.length; s++) {
    var sheetInfo = sheetsToRead[s];
    var sheet = ss.getSheetByName(sheetInfo.name);

    if (!sheet) {
      // Спробуємо знайти аркуш з апострофом (кур'єра)
      if (sheetInfo.name === SHEET_COURIER) {
        sheet = findSheetFuzzy(ss, 'Виклик кур');
      }
      if (!sheet) continue;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue; // тільки заголовки, даних немає

    // Читаємо ВСІ рядки одразу (один запит до таблиці = швидко)
    var dataRange = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS);
    var values = dataRange.getValues();

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowNum = i + 2; // рядок в таблиці (1-based, +1 за заголовок)

      // Пропускаємо повністю порожні рядки
      if (isEmptyRow(row)) continue;

      // Мінімальна перевірка — має бути хоча б щось ідентифікуюче
      // (ІД, ТТН, телефон або ПіБ)
      var hasIdentity = row[COL.ID] || row[COL.TTN] || row[COL.PHONE] || row[COL.NAME];
      if (!hasIdentity) continue;

      var dateReg = formatDate(row[COL.DATE_REG]);

      // Визначаємо чи лід новий (за останні 24 години)
      var isNew24h = false;
      if (dateReg) {
        try {
          var regDate = new Date(dateReg);
          var now = new Date();
          isNew24h = (now.getTime() - regDate.getTime()) < 86400000; // 24 год
        } catch (e) { /* ігноруємо помилки дат */ }
      }

      // Статус CRM — нормалізуємо
      var crmStatus = String(row[COL.STATUS] || '').toLowerCase().trim();
      // Якщо статус пустий — лід новий (прилетів з бота/SmartSender)
      if (!crmStatus) crmStatus = '';

      allPackages.push({
        // Ідентифікація (прив'язка до аркуша + рядка)
        id: String(row[COL.ID] || ''),
        rowNum: rowNum,
        sheet: sheetInfo.name,

        // Основні поля (22 колонки)
        vo: String(row[COL.VO] || ''),
        number: String(row[COL.NUMBER] || ''),
        ttn: String(row[COL.TTN] || ''),
        weight: String(row[COL.WEIGHT] || ''),
        address: String(row[COL.ADDRESS] || ''),
        directionRaw: String(row[COL.DIRECTION] || ''),
        direction: sheetInfo.direction,  // прив'язано до аркуша!
        phone: String(row[COL.PHONE] || ''),
        amount: String(row[COL.AMOUNT] || ''),
        payStatus: String(row[COL.PAY_STATUS] || ''),
        payment: String(row[COL.PAYMENT] || ''),
        phoneReg: String(row[COL.PHONE_REG] || ''),
        note: String(row[COL.NOTE] || ''),
        parcelStatus: String(row[COL.PARCEL_STATUS] || ''),
        name: String(row[COL.NAME] || ''),
        dateReg: dateReg,
        timing: String(row[COL.TIMING] || ''),
        smsNote: String(row[COL.SMS_NOTE] || ''),
        dateReceive: formatDate(row[COL.DATE_RECEIVE]),
        photo: String(row[COL.PHOTO] || ''),
        status: crmStatus,
        dateArchive: formatDate(row[COL.DATE_ARCHIVE]),

        // Мета
        isNew: isNew24h
      });
    }
  }

  return {
    success: true,
    packages: allPackages,
    count: allPackages.length,
    timestamp: new Date().toISOString()
  };
}

// ============================================
// getStructure — Структура таблиці (для дебагу)
// ============================================
function getStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var result = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var sample = lastRow > 1
      ? sheet.getRange(2, 1, Math.min(2, lastRow - 1), lastCol).getValues()
      : [];

    result.push({
      sheet: name,
      rows: lastRow,
      cols: lastCol,
      headers: headers,
      sample: sample
    });
  }

  return { success: true, sheets: result };
}

// ============================================
// addPackage — Додати нову посилку
// Пише у ПРАВИЛЬНИЙ аркуш на основі напрямку!
// ua-eu → "Реєстрація ТТН"
// eu-ua → "Виклик кур'єра"
// ============================================
function addPackage(data) {
  var fields = data.fields;
  if (!fields) {
    return { success: false, error: 'Missing fields' };
  }

  // Визначаємо аркуш по напрямку
  var direction = fields.direction || 'ua-eu';
  var sheetName = data.sheet || getSheetByDirection(direction);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Пробуємо знайти fuzzy
    if (sheetName === SHEET_COURIER) {
      sheet = findSheetFuzzy(ss, 'Виклик кур');
    }
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + sheetName };
    }
  }

  // Створюємо порожній рядок на 22 колонки
  var newRow = new Array(TOTAL_COLS);
  for (var i = 0; i < TOTAL_COLS; i++) {
    newRow[i] = '';
  }

  // Заповнюємо значення з маппінгу
  for (var field in fields) {
    if (fields.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
      newRow[FIELD_MAP[field]] = fields[field];
    }
  }

  // Якщо дата оформлення не задана — ставимо сьогодні
  if (!newRow[COL.DATE_REG]) {
    newRow[COL.DATE_REG] = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  }

  // Якщо ІД не задано — генеруємо
  if (!newRow[COL.ID]) {
    newRow[COL.ID] = 'crm_' + new Date().getTime();
  }

  // Якщо напрямок не вказано — беремо з аркуша
  if (!newRow[COL.DIRECTION]) {
    newRow[COL.DIRECTION] = direction === 'eu-ua' ? 'EU→UA' : 'UA→EU';
  }

  // Додаємо рядок
  sheet.appendRow(newRow);
  var newRowNum = sheet.getLastRow();

  // Логуємо
  writeLog('addPackage', sheetName, newRowNum, 'new',
    'ПіБ: ' + (fields.name || '') + ' | ТТН: ' + (fields.ttn || '') + ' | Тел: ' + (fields.phone || ''));

  return {
    success: true,
    sheet: sheetName,
    rowNum: newRowNum,
    id: newRow[COL.ID],
    direction: direction
  };
}

// ============================================
// updatePackage — Оновити одну посилку
// Прив'язка: sheet + rowNum (унікальна адреса рядка)
// ============================================
function updatePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var fields = data.fields;

  if (!sheetName || !rowNum || !fields) {
    return { success: false, error: 'Missing sheet, rowNum or fields' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  // Перевіряємо що рядок існує
  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Row ' + rowNum + ' does not exist (lastRow: ' + sheet.getLastRow() + ')' };
  }

  var updated = [];
  for (var field in fields) {
    if (fields.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
      var colIndex = FIELD_MAP[field] + 1; // 1-based для Range
      var value = fields[field];
      sheet.getRange(rowNum, colIndex).setValue(value);
      updated.push(field);
    }
  }

  // Логуємо дію
  writeLog('updatePackage', sheetName, rowNum, updated.join(', '), JSON.stringify(fields));

  return {
    success: true,
    updated: updated,
    sheet: sheetName,
    rowNum: rowNum
  };
}

// ============================================
// deletePackage — Видалити посилку (статус = deleted)
// ============================================
function deletePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;

  if (!sheetName || !rowNum) {
    return { success: false, error: 'Missing sheet or rowNum' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  // Не видаляємо фізично — ставимо статус "deleted"
  sheet.getRange(rowNum, COL.STATUS + 1).setValue('deleted');
  sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
    Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
  );

  writeLog('deletePackage', sheetName, rowNum, 'deleted', '');

  return { success: true, sheet: sheetName, rowNum: rowNum };
}

// ============================================
// updateStatus — Змінити CRM статус однієї посилки
// ============================================
function updateStatus(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var newStatus = data.status;

  if (!sheetName || !rowNum || !newStatus) {
    return { success: false, error: 'Missing sheet, rowNum or status' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  // Якщо архів/відмова/видалення/перенос — ставимо дату
  var archiveStatuses = ['archived', 'refused', 'deleted', 'transferred'];
  if (archiveStatuses.indexOf(newStatus) !== -1) {
    sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
      Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
    );
  }

  writeLog('updateStatus', sheetName, rowNum, newStatus, '');

  return { success: true, sheet: sheetName, rowNum: rowNum, status: newStatus };
}

// ============================================
// bulkUpdateStatus — Масова зміна статусу
// ============================================
function bulkUpdateStatus(data) {
  var items = data.items; // масив { sheet, rowNum }
  var newStatus = data.status;

  if (!items || !items.length || !newStatus) {
    return { success: false, error: 'Missing items or status' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  var archiveStatuses = ['archived', 'refused', 'deleted', 'transferred'];
  var needDate = archiveStatuses.indexOf(newStatus) !== -1;
  var count = 0;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = findSheet(ss, item.sheet);
    if (!sheet) continue;

    sheet.getRange(item.rowNum, COL.STATUS + 1).setValue(newStatus);
    if (needDate) {
      sheet.getRange(item.rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
    }
    count++;
  }

  writeLog('bulkUpdateStatus', 'bulk', 0, newStatus, count + ' items');

  return { success: true, count: count, status: newStatus };
}

// ============================================
// updateField — Оновити одне поле (швидке редагування)
// ============================================
function updateField(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var field = data.field;
  var value = data.value;

  if (!sheetName || !rowNum || !field) {
    return { success: false, error: 'Missing sheet, rowNum or field' };
  }

  if (!FIELD_MAP.hasOwnProperty(field)) {
    return { success: false, error: 'Unknown field: ' + field };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(value);

  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}

// ============================================
// bulkAssignVehicle — Масове призначення авто
// ============================================
function bulkAssignVehicle(data) {
  var items = data.items;
  var vehicle = data.vehicle;

  if (!items || !items.length || !vehicle) {
    return { success: false, error: 'Missing items or vehicle' };
  }

  // Авто зберігається локально в CRM (localStorage)
  // Логуємо для відстеження
  writeLog('bulkAssignVehicle', 'bulk', 0, vehicle, items.length + ' items');

  return { success: true, count: items.length, vehicle: vehicle };
}

// ============================================
// ЛОГУВАННЯ — запис дій в аркуш "Логи"
// ============================================
function writeLog(action, sheetName, rowNum, detail, extra) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(SHEET_LOGS);

    if (!logSheet) {
      // Створюємо аркуш Логи якщо не існує
      logSheet = ss.insertSheet(SHEET_LOGS);
      logSheet.appendRow(['Дата/Час', 'Дія', 'Аркуш', 'Рядок', 'Деталі', 'Дані']);
      // Заморожуємо заголовок
      logSheet.setFrozenRows(1);
    }

    var timestamp = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([
      timestamp,
      action,
      sheetName,
      rowNum,
      detail,
      extra || ''
    ]);
  } catch (e) {
    // Логування не повинно ламати основну логіку
    Logger.log('Log error: ' + e.toString());
  }
}

// ============================================
// ДОПОМІЖНІ ФУНКЦІЇ
// ============================================

// Знайти аркуш по назві (з fallback для апострофів)
function findSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  // Пробуємо fuzzy для "Виклик кур'єра" варіантів
  if (name.indexOf('Виклик кур') === 0 || name === SHEET_COURIER) {
    return findSheetFuzzy(ss, 'Виклик кур');
  }
  return null;
}

// Fuzzy пошук аркуша (для проблем з апострофами)
function findSheetFuzzy(ss, prefix) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName.indexOf(prefix) === 0) {
      return sheets[i];
    }
  }
  return null;
}

// Перевірка чи рядок порожній
function isEmptyRow(row) {
  for (var c = 0; c < row.length && c < TOTAL_COLS; c++) {
    var val = row[c];
    if (val !== '' && val !== null && val !== undefined) {
      return false;
    }
  }
  return true;
}

// Форматування дати до YYYY-MM-DD
function formatDate(value) {
  if (!value) return '';

  // Якщо це об'єкт Date (Google Sheets повертає Date)
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }

  var str = String(value).trim();
  if (!str) return '';

  // Вже YYYY-MM-DD — повертаємо
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }

  // DD.MM.YYYY → YYYY-MM-DD
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(str)) {
    var parts = str.split('.');
    return parts[2] + '-' + parts[1].padStart(2, '0') + '-' + parts[0].padStart(2, '0');
  }

  // DD/MM/YYYY → YYYY-MM-DD
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    var parts2 = str.split('/');
    return parts2[2] + '-' + parts2[1].padStart(2, '0') + '-' + parts2[0].padStart(2, '0');
  }

  // Спроба парсингу
  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd');
    }
  } catch (e) { /* ігноруємо */ }

  return '';
}

// Відповідь у форматі JSON
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// ТЕСТОВІ ФУНКЦІЇ — запускай в редакторі
// ============================================

// Тест 1: Перевірити чи аркуші знаходяться
function testFindSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  Logger.log('=== ВСІ АРКУШІ В ТАБЛИЦІ ===');
  for (var i = 0; i < sheets.length; i++) {
    Logger.log('  [' + i + '] "' + sheets[i].getName() + '" (' + sheets[i].getLastRow() + ' рядків)');
  }

  Logger.log('');
  Logger.log('=== ПОШУК ПОТРІБНИХ АРКУШІВ ===');

  var regSheet = findSheet(ss, SHEET_REG);
  Logger.log('Реєстрація ТТН: ' + (regSheet ? 'ЗНАЙДЕНО ✅ (' + regSheet.getLastRow() + ' рядків)' : 'НЕ ЗНАЙДЕНО ❌'));

  var courierSheet = findSheet(ss, SHEET_COURIER);
  Logger.log('Виклик кур`єра: ' + (courierSheet ? 'ЗНАЙДЕНО ✅ (' + courierSheet.getLastRow() + ' рядків)' : 'НЕ ЗНАЙДЕНО ❌'));

  var logSheet = ss.getSheetByName(SHEET_LOGS);
  Logger.log('Логи: ' + (logSheet ? 'ЗНАЙДЕНО ✅' : 'НЕ ЗНАЙДЕНО (буде створено автоматично)'));
}

// Тест 2: Витягнути всі посилки
function testGetAll() {
  var result = getAllPackages();
  Logger.log('=== РЕЗУЛЬТАТ getAll ===');
  Logger.log('Успіх: ' + result.success);
  Logger.log('Всього посилок: ' + result.count);

  if (result.packages && result.packages.length > 0) {
    Logger.log('');
    Logger.log('=== ПЕРШІ 5 ПОСИЛОК ===');
    for (var i = 0; i < Math.min(5, result.packages.length); i++) {
      var p = result.packages[i];
      Logger.log(
        '  [' + p.sheet + ' #' + p.rowNum + '] ' +
        'ПіБ: ' + (p.name || '—') +
        ' | ТТН: ' + (p.ttn || '—') +
        ' | Тел: ' + (p.phone || '—') +
        ' | Напр: ' + p.direction +
        ' | Статус: ' + (p.status || '(пусто)') +
        ' | Дата: ' + (p.dateReg || '—')
      );
    }

    // Статистика по аркушах
    var regCount = result.packages.filter(function(p) { return p.sheet === SHEET_REG || p.direction === 'ua-eu'; }).length;
    var courCount = result.packages.filter(function(p) { return p.direction === 'eu-ua'; }).length;
    Logger.log('');
    Logger.log('UA→EU (Реєстрація): ' + regCount);
    Logger.log('EU→UA (Виклик кур`єра): ' + courCount);
  }
}

// Тест 3: Подивитись структуру
function testStructure() {
  var result = getStructure();
  Logger.log('=== СТРУКТУРА ТАБЛИЦІ ===');
  for (var i = 0; i < result.sheets.length; i++) {
    var s = result.sheets[i];
    Logger.log('');
    Logger.log('[' + s.sheet + '] ' + s.rows + ' rows, ' + s.cols + ' cols');
    Logger.log('   Колонки: ' + s.headers.join(' | '));
    if (s.sample && s.sample.length > 0) {
      Logger.log('   Приклад: ' + JSON.stringify(s.sample[0]));
    }
  }
}
