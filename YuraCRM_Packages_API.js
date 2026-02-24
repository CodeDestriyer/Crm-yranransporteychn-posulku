// ============================================
// 📦 ЮРА ТРАНСПОРТЕЙШН — CRM ПОСИЛКИ
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

// Назви аркушів — ТОЧНО як в таблиці
var SHEET_REG = 'Реєстрація ТТН';    // UA→EU посилки
var SHEET_COURIER = 'Виклик кур\u0027єра';  // EU→UA посилки
var SHEET_LOGS = 'Логи';               // Логування дій

// Порядок колонок (A-V = 22 колонки, індекс 0-21)
// A:ВО  B:Номер№  C:Номер ТТН  D:Вага  E:Адреса Отримувача  F:Напрямок
// G:Телефон Отримувача  H:Сума Є  I:Статус оплати  J:Оплата
// K:Телефон Реєстратора  L:Примітка  M:Статус посилки  N:ІД  O:ПіБ
// P:дата оформлення  Q:Таймінг  R:Примітка смс  S:Дата отримання
// T:Фото  U:Статус  V:Дата архів

var COL = {
  VO: 0,            // A — ВО (менеджер)
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
  PARCEL_STATUS: 12,// M — Статус посилки
  ID: 13,           // N — ІД
  NAME: 14,         // O — ПіБ
  DATE_REG: 15,     // P — дата оформлення
  TIMING: 16,       // Q — Таймінг
  SMS_NOTE: 17,     // R — Примітка смс
  DATE_RECEIVE: 18, // S — Дата отримання
  PHOTO: 19,        // T — Фото
  STATUS: 20,       // U — Статус (CRM)
  DATE_ARCHIVE: 21  // V — Дата архів
};

var TOTAL_COLS = 22;


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
// getAll — Витягнути всі посилки з обох аркушів
// ============================================
function getAllPackages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allPackages = [];

  // Читаємо обидва аркуші
  var sheets = [
    { name: SHEET_REG, direction: 'ua-eu' },
    { name: SHEET_COURIER, direction: 'eu-ua' }
  ];

  for (var s = 0; s < sheets.length; s++) {
    var sheetInfo = sheets[s];
    var sheet = ss.getSheetByName(sheetInfo.name);
    if (!sheet) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue; // тільки заголовки, даних немає

    var lastCol = Math.max(sheet.getLastColumn(), TOTAL_COLS);
    var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowNum = i + 2; // рядок в таблиці (1-based, мінус заголовок)

      // Пропускаємо повністю порожні рядки
      var hasData = false;
      for (var c = 0; c < TOTAL_COLS; c++) {
        if (row[c] !== '' && row[c] !== null && row[c] !== undefined) {
          hasData = true;
          break;
        }
      }
      if (!hasData) continue;

      var dateReg = formatDate(row[COL.DATE_REG]);
      var isNew24h = false;
      if (dateReg) {
        try {
          var regDate = new Date(dateReg);
          var now = new Date();
          isNew24h = (now - regDate) < 86400000; // 24 години
        } catch (e) {}
      }

      allPackages.push({
        // Ідентифікація
        id: String(row[COL.ID] || ''),
        rowNum: rowNum,
        sheet: sheetInfo.name,

        // Основні дані
        vo: String(row[COL.VO] || ''),
        number: String(row[COL.NUMBER] || ''),
        ttn: String(row[COL.TTN] || ''),
        weight: String(row[COL.WEIGHT] || ''),
        address: String(row[COL.ADDRESS] || ''),
        directionRaw: String(row[COL.DIRECTION] || ''),
        direction: sheetInfo.direction,
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
        status: String(row[COL.STATUS] || '').toLowerCase().trim(),
        dateArchive: formatDate(row[COL.DATE_ARCHIVE]),

        // Мета
        isNew: isNew24h,
        vehicle: '' // авто призначається в CRM, не в таблиці
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
    var sample = lastRow > 1 ? sheet.getRange(2, 1, Math.min(2, lastRow - 1), lastCol).getValues() : [];

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
// updatePackage — Оновити одну посилку
// ============================================
function updatePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var fields = data.fields; // об'єкт з полями для оновлення

  if (!sheetName || !rowNum || !fields) {
    return { success: false, error: 'Missing sheet, rowNum or fields' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  // Маппінг назв полів до індексів колонок
  var fieldMap = {
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

  var updated = [];

  for (var field in fields) {
    if (fields.hasOwnProperty(field) && fieldMap.hasOwnProperty(field)) {
      var colIndex = fieldMap[field] + 1; // 1-based для Range
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
// addPackage — Додати нову посилку
// ============================================
function addPackage(data) {
  var sheetName = data.sheet || SHEET_REG; // за замовчуванням — Реєстрація ТТН
  var fields = data.fields;

  if (!fields) {
    return { success: false, error: 'Missing fields' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  // Створюємо порожній рядок на 22 колонки
  var newRow = [];
  for (var i = 0; i < TOTAL_COLS; i++) {
    newRow.push('');
  }

  // Заповнюємо значення
  if (fields.vo !== undefined) newRow[COL.VO] = fields.vo;
  if (fields.number !== undefined) newRow[COL.NUMBER] = fields.number;
  if (fields.ttn !== undefined) newRow[COL.TTN] = fields.ttn;
  if (fields.weight !== undefined) newRow[COL.WEIGHT] = fields.weight;
  if (fields.address !== undefined) newRow[COL.ADDRESS] = fields.address;
  if (fields.direction !== undefined) newRow[COL.DIRECTION] = fields.direction;
  if (fields.phone !== undefined) newRow[COL.PHONE] = fields.phone;
  if (fields.amount !== undefined) newRow[COL.AMOUNT] = fields.amount;
  if (fields.payStatus !== undefined) newRow[COL.PAY_STATUS] = fields.payStatus;
  if (fields.payment !== undefined) newRow[COL.PAYMENT] = fields.payment;
  if (fields.phoneReg !== undefined) newRow[COL.PHONE_REG] = fields.phoneReg;
  if (fields.note !== undefined) newRow[COL.NOTE] = fields.note;
  if (fields.parcelStatus !== undefined) newRow[COL.PARCEL_STATUS] = fields.parcelStatus;
  if (fields.id !== undefined) newRow[COL.ID] = fields.id;
  if (fields.name !== undefined) newRow[COL.NAME] = fields.name;
  if (fields.dateReg !== undefined) newRow[COL.DATE_REG] = fields.dateReg;
  if (fields.timing !== undefined) newRow[COL.TIMING] = fields.timing;
  if (fields.smsNote !== undefined) newRow[COL.SMS_NOTE] = fields.smsNote;
  if (fields.dateReceive !== undefined) newRow[COL.DATE_RECEIVE] = fields.dateReceive;
  if (fields.photo !== undefined) newRow[COL.PHOTO] = fields.photo;
  if (fields.status !== undefined) newRow[COL.STATUS] = fields.status;
  if (fields.dateArchive !== undefined) newRow[COL.DATE_ARCHIVE] = fields.dateArchive;

  // Якщо дата оформлення не задана — ставимо сьогодні
  if (!newRow[COL.DATE_REG]) {
    newRow[COL.DATE_REG] = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  }

  // Якщо ІД не задано — генеруємо
  if (!newRow[COL.ID]) {
    newRow[COL.ID] = 'crm_' + new Date().getTime();
  }

  // Додаємо рядок
  sheet.appendRow(newRow);
  var newRowNum = sheet.getLastRow();

  // Логуємо
  writeLog('addPackage', sheetName, newRowNum, 'new', JSON.stringify(fields));

  return {
    success: true,
    sheet: sheetName,
    rowNum: newRowNum,
    id: newRow[COL.ID]
  };
}


// ============================================
// deletePackage — Видалити посилку (ставимо статус deleted)
// ============================================
function deletePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;

  if (!sheetName || !rowNum) {
    return { success: false, error: 'Missing sheet or rowNum' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
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
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  // Якщо архів/відмова/видалення — ставимо дату архіву
  if (['archived', 'refused', 'deleted'].indexOf(newStatus) !== -1) {
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
  var count = 0;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = ss.getSheetByName(item.sheet);
    if (!sheet) continue;

    sheet.getRange(item.rowNum, COL.STATUS + 1).setValue(newStatus);

    if (['archived', 'refused', 'deleted'].indexOf(newStatus) !== -1) {
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

  var fieldMap = {
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

  if (!fieldMap.hasOwnProperty(field)) {
    return { success: false, error: 'Unknown field: ' + field };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: 'Sheet not found: ' + sheetName };
  }

  sheet.getRange(rowNum, fieldMap[field] + 1).setValue(value);

  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}


// ============================================
// bulkAssignVehicle — Масове призначення авто
// (Зберігаємо у Примітці СМС тимчасово, або в окремій колонці)
// ============================================
function bulkAssignVehicle(data) {
  var items = data.items; // масив { sheet, rowNum }
  var vehicle = data.vehicle;

  if (!items || !items.length || !vehicle) {
    return { success: false, error: 'Missing items or vehicle' };
  }

  // Поки авто зберігається локально в CRM (localStorage)
  // Але логуємо для відстеження
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

// Форматування дати до YYYY-MM-DD
function formatDate(value) {
  if (!value) return '';

  // Якщо це об'єкт Date
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }

  var str = String(value).trim();
  if (!str) return '';

  // Якщо вже YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }

  // Якщо DD.MM.YYYY
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(str)) {
    var parts = str.split('.');
    return parts[2] + '-' + parts[1] + '-' + parts[0];
  }

  // Спроба парсингу
  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd');
    }
  } catch (e) {}

  return '';
}

// Відповідь у форматі JSON
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================
// ТЕСТОВА ФУНКЦІЯ — запусти в редакторі для перевірки
// ============================================
function testGetAll() {
  var result = getAllPackages();
  Logger.log('Всього посилок: ' + result.count);
  Logger.log('Перші 3:');
  for (var i = 0; i < Math.min(3, result.packages.length); i++) {
    var p = result.packages[i];
    Logger.log(
      '  [' + p.sheet + '] #' + p.rowNum +
      ' | ПіБ: ' + p.name +
      ' | ТТН: ' + p.ttn +
      ' | Напр: ' + p.direction +
      ' | Статус: ' + p.status +
      ' | Дата: ' + p.dateReg
    );
  }
  Logger.log(JSON.stringify(result.packages.slice(0, 2), null, 2));
}

function testStructure() {
  var result = getStructure();
  result.sheets.forEach(function(s) {
    Logger.log('📋 ' + s.sheet + ' — ' + s.rows + ' рядків, ' + s.cols + ' колонок');
    Logger.log('   Колонки: ' + s.headers.join(' | '));
  });
}
