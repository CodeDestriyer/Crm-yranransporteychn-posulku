// ============================================
// ЮРА ТРАНСПОРТЕЙШН — CRM ПОСИЛКИ v1.0
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
// 5. Встав URL в CRM HTML файл
// ============================================

// ============================================
// КОНФІГУРАЦІЯ
// ============================================

// Назви аркушів — ТОЧНО як в таблиці
var SHEET_REG = 'Реєстрація ТТН';     // UA→EU посилки
var SHEET_COURIER = 'Виклик курєра';   // EU→UA посилки
var SHEET_LOGS = 'Логи';               // Логування дій (в архівній таблиці)

// URL архівного скрипта (Crm_Arhiv_1.0)
var ARCHIVE_API_URL = 'https://script.google.com/macros/s/AKfycbwJLGZgYT333VdMW-nM5kPjYs2WIGGjfqkZnDJYjJxUt8nzE8GDGCPm7EzMHhcxNDOn/exec';

// Порядок колонок (A-W = 23 колонки, індекс 0-22)
// A:ВО  B:Номер№  C:Номер ТТН  D:Вага  E:Адреса Отримувача  F:Напрямок
// G:Телефон Отримувача  H:Сума Є  I:Статус оплати  J:Оплата
// K:Телефон Реєстратора  L:Примітка  M:Статус посилки  N:ІД  O:ПіБ
// P:дата оформлення  Q:Таймінг  R:Примітка смс  S:Дата отримання
// T:Фото  U:Статус  V:Дата архів  W:ARCHIVE_ID
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
  PARCEL_STATUS: 12,// M — Статус посилки (Невідомий/Зареєстровано/Оформлено/Кордон/Доставка)
  ID: 13,           // N — ІД
  NAME: 14,         // O — ПіБ
  DATE_REG: 15,     // P — дата оформлення
  TIMING: 16,       // Q — Таймінг
  SMS_NOTE: 17,     // R — Примітка смс
  DATE_RECEIVE: 18, // S — Дата отримання
  PHOTO: 19,        // T — Фото
  STATUS: 20,       // U — Статус (CRM: new/work/route/archived/refused/transferred/deleted)
  DATE_ARCHIVE: 21, // V — Дата архів
  ARCHIVE_ID: 22    // W — ARCHIVE_ID (зв'язок з таблицею Архіви)
};
var TOTAL_COLS = 23;

// Статуси для архівації
var ARCHIVE_STATUSES = ['archived', 'refused', 'deleted', 'transferred'];

// Маппінг напрямок → аркуш
function getSheetByDirection(direction) {
  if (direction === 'eu-ua') return SHEET_COURIER;
  return SHEET_REG;
}

// Маппінг аркуш → напрямок
function getDirectionBySheet(sheetName) {
  if (sheetName === SHEET_COURIER) return 'eu-ua';
  return 'ua-eu';
}

// ============================================
// МАППІНГ полів CRM → індексів колонок
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
  dateArchive: COL.DATE_ARCHIVE,
  archiveId: COL.ARCHIVE_ID
};

// ============================================
// ГОЛОВНИЙ ОБРОБНИК — doPost
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    switch (action) {
      // --- ЧИТАННЯ ---
      case 'getAll':
        return respond(getAllPackages());

      case 'getStructure':
        return respond(getStructure());

      // --- СТВОРЕННЯ ---
      case 'addPackage':
        return respond(addPackage(data));

      // --- ОНОВЛЕННЯ ---
      case 'updatePackage':
        return respond(updatePackage(data));

      case 'updateField':
        return respond(updateField(data));

      case 'updateStatus':
        return respond(updateStatus(data));

      case 'bulkUpdateStatus':
        return respond(bulkUpdateStatus(data));

      case 'bulkAssignVehicle':
        return respond(bulkAssignVehicle(data));

      // --- ВИДАЛЕННЯ ---
      case 'deletePackage':
        return respond(deletePackage(data));

      // --- АРХІВАЦІЯ ---
      case 'archivePackage':
        return respond(archivePackage(data));

      case 'bulkArchive':
        return respond(bulkArchive(data));

      // --- ДУБЛІКАТИ ---
      case 'checkDuplicates':
        return respond(checkDuplicates(data));

      default:
        return respond({ success: false, error: 'Невідома дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// doGet — Перевірка здоров'я
// ============================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'health';

    switch (action) {
      case 'health':
        return respond({
          success: true,
          version: '1.0',
          service: 'CRM Посилки — ЮРА ТРАНСПОРТЕЙШН',
          sheets: [SHEET_REG, SHEET_COURIER],
          totalCols: TOTAL_COLS,
          timestamp: new Date().toISOString()
        });

      default:
        return respond({ success: false, error: 'Невідома GET дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// getAll — Витягнути ВСІ посилки з обох аркушів
// ============================================
function getAllPackages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allPackages = [];

  // Читаємо 2 робочі аркуші (РОБОЧА — не читаємо, лише для бекапу)
  var sheetsToRead = [
    { name: SHEET_REG, direction: 'ua-eu' },
    { name: SHEET_COURIER, direction: 'eu-ua' }
  ];

  for (var s = 0; s < sheetsToRead.length; s++) {
    var sheetInfo = sheetsToRead[s];
    var sheet = findSheet(ss, sheetInfo.name);
    if (!sheet) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    // Читаємо ВСІ рядки одразу (один запит = швидко)
    var dataRange = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS);
    var values = dataRange.getValues();

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowNum = i + 2;

      if (isEmptyRow(row)) continue;

      // Мінімальна ідентифікація
      var hasIdentity = row[COL.ID] || row[COL.TTN] || row[COL.PHONE] || row[COL.NAME];
      if (!hasIdentity) continue;

      var dateReg = formatDate(row[COL.DATE_REG]);

      // Новий лід? (за останні 24 год)
      var isNew24h = false;
      if (dateReg) {
        try {
          var regDate = new Date(dateReg);
          var now = new Date();
          isNew24h = (now.getTime() - regDate.getTime()) < 86400000;
        } catch (e) {}
      }

      var crmStatus = String(row[COL.STATUS] || '').toLowerCase().trim();

      allPackages.push({
        // Ідентифікація
        id: String(row[COL.ID] || ''),
        rowNum: rowNum,
        sheet: sheetInfo.name,

        // Дані (23 колонки)
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
        status: crmStatus,
        dateArchive: formatDate(row[COL.DATE_ARCHIVE]),
        archiveId: String(row[COL.ARCHIVE_ID] || ''),

        // Мета
        isNew: isNew24h,
        isArchived: ARCHIVE_STATUSES.indexOf(crmStatus) !== -1
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
// addPackage — Додати нову посилку
// + перевірка дублікатів перед додаванням
// ============================================
function addPackage(data) {
  var fields = data.fields;
  if (!fields) {
    return { success: false, error: 'Відсутні поля (fields)' };
  }

  var direction = fields.direction || 'ua-eu';
  var sheetName = data.sheet || getSheetByDirection(direction);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  // --- ПЕРЕВІРКА ДУБЛІКАТІВ ---
  var duplicates = [];
  var checkTTN = fields.ttn ? String(fields.ttn).trim() : '';
  var checkPhone = fields.phone ? String(fields.phone).trim() : '';
  var checkId = fields.id ? String(fields.id).trim() : '';

  if (checkTTN || checkPhone || checkId) {
    var sheetsToCheck = [
      findSheet(ss, SHEET_REG),
      findSheet(ss, SHEET_COURIER)
    ];

    for (var s = 0; s < sheetsToCheck.length; s++) {
      var chkSheet = sheetsToCheck[s];
      if (!chkSheet) continue;

      var lastRow = chkSheet.getLastRow();
      if (lastRow < 2) continue;

      var chkValues = chkSheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
      for (var r = 0; r < chkValues.length; r++) {
        var existRow = chkValues[r];
        var existStatus = String(existRow[COL.STATUS] || '').toLowerCase().trim();

        // Пропускаємо архівовані/видалені
        if (ARCHIVE_STATUSES.indexOf(existStatus) !== -1) continue;

        var existTTN = String(existRow[COL.TTN] || '').trim();
        var existPhone = String(existRow[COL.PHONE] || '').trim();
        var existId = String(existRow[COL.ID] || '').trim();

        var isDuplicate = false;
        var reason = '';

        // Дублікат за ІД
        if (checkId && existId && checkId === existId) {
          isDuplicate = true;
          reason = 'ІД: ' + checkId;
        }
        // Дублікат за ТТН
        if (checkTTN && existTTN && checkTTN === existTTN) {
          isDuplicate = true;
          reason = 'ТТН: ' + checkTTN;
        }
        // Дублікат за телефоном + ім'ям (м'яка перевірка)
        if (!isDuplicate && checkPhone && existPhone && checkPhone === existPhone) {
          var checkName = String(fields.name || '').trim().toLowerCase();
          var existName = String(existRow[COL.NAME] || '').trim().toLowerCase();
          if (checkName && existName && checkName === existName) {
            isDuplicate = true;
            reason = 'Телефон+ПіБ: ' + checkPhone;
          }
        }

        if (isDuplicate) {
          duplicates.push({
            sheet: chkSheet.getName(),
            rowNum: r + 2,
            ttn: existTTN,
            phone: existPhone,
            name: String(existRow[COL.NAME] || ''),
            status: existStatus,
            reason: reason
          });
        }
      }
    }
  }

  // Якщо є дублікати і не передано force — повертаємо попередження
  if (duplicates.length > 0 && !data.force) {
    writeLog('addPackage:DUPLICATE', sheetName, 0, 'blocked',
      'Знайдено ' + duplicates.length + ' дублікатів | ' + duplicates[0].reason);

    return {
      success: false,
      error: 'duplicate',
      message: 'Знайдено ' + duplicates.length + ' можливих дублікатів',
      duplicates: duplicates
    };
  }

  // --- СТВОРЕННЯ РЯДКА ---
  var newRow = new Array(TOTAL_COLS);
  for (var i = 0; i < TOTAL_COLS; i++) {
    newRow[i] = '';
  }

  for (var field in fields) {
    if (fields.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
      newRow[FIELD_MAP[field]] = fields[field];
    }
  }

  // Автозаповнення
  if (!newRow[COL.DATE_REG]) {
    newRow[COL.DATE_REG] = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  }
  if (!newRow[COL.ID]) {
    newRow[COL.ID] = 'crm_' + new Date().getTime();
  }
  if (!newRow[COL.DIRECTION]) {
    newRow[COL.DIRECTION] = direction === 'eu-ua' ? 'EU→UA' : 'UA→EU';
  }
  if (!newRow[COL.STATUS]) {
    newRow[COL.STATUS] = 'new';
  }

  sheet.appendRow(newRow);
  var newRowNum = sheet.getLastRow();

  writeLog('addPackage', sheetName, newRowNum, 'new',
    'ПіБ: ' + (fields.name || '') + ' | ТТН: ' + (fields.ttn || '') + ' | Тел: ' + (fields.phone || '') +
    (duplicates.length > 0 ? ' | FORCE (дублікат ігноровано)' : ''));

  return {
    success: true,
    sheet: sheetName,
    rowNum: newRowNum,
    id: newRow[COL.ID],
    direction: direction,
    duplicatesIgnored: duplicates.length
  };
}

// ============================================
// updatePackage — Оновити посилку
// + верифікація по ІД щоб не перезаписати чужий рядок
// ============================================
function updatePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var fields = data.fields;

  if (!sheetName || !rowNum || !fields) {
    return { success: false, error: 'Відсутні sheet, rowNum або fields' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує (lastRow: ' + sheet.getLastRow() + ')' };
  }

  // --- ВЕРИФІКАЦІЯ ПО ІД ---
  // Якщо передано expectedId — перевіряємо що рядок не змінився
  if (data.expectedId) {
    var currentId = String(sheet.getRange(rowNum, COL.ID + 1).getValue() || '').trim();
    if (currentId !== String(data.expectedId).trim()) {
      writeLog('updatePackage:CONFLICT', sheetName, rowNum, 'blocked',
        'Очікувався ІД: ' + data.expectedId + ', фактичний: ' + currentId);

      return {
        success: false,
        error: 'conflict',
        message: 'Рядок змінився (можливо видалено/переміщено). Очікувався ІД: ' + data.expectedId + ', фактичний: ' + currentId
      };
    }
  }

  // --- ПЕРЕВІРКА ЧИ НЕ АРХІВОВАНИЙ ---
  var currentStatus = String(sheet.getRange(rowNum, COL.STATUS + 1).getValue() || '').toLowerCase().trim();
  if (ARCHIVE_STATUSES.indexOf(currentStatus) !== -1 && !data.force) {
    return {
      success: false,
      error: 'archived',
      message: 'Запис вже архівований (статус: ' + currentStatus + '). Використайте force=true для перезапису'
    };
  }

  var updated = [];
  for (var field in fields) {
    if (fields.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
      sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(fields[field]);
      updated.push(field);
    }
  }

  writeLog('updatePackage', sheetName, rowNum, updated.join(', '), JSON.stringify(fields));

  return { success: true, updated: updated, sheet: sheetName, rowNum: rowNum };
}

// ============================================
// updateField — Оновити одне поле
// ============================================
function updateField(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var field = data.field;
  var value = data.value;

  if (!sheetName || !rowNum || !field) {
    return { success: false, error: 'Відсутні sheet, rowNum або field' };
  }

  if (!FIELD_MAP.hasOwnProperty(field)) {
    return { success: false, error: 'Невідоме поле: ' + field };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(value);

  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}

// ============================================
// updateStatus — Змінити CRM статус
// + автоматичне логування дати архіву
// ============================================
function updateStatus(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var newStatus = data.status;

  if (!sheetName || !rowNum || !newStatus) {
    return { success: false, error: 'Відсутні sheet, rowNum або status' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  var oldStatus = String(sheet.getRange(rowNum, COL.STATUS + 1).getValue() || '').toLowerCase().trim();

  // Ставимо новий статус
  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  // Якщо архівний статус — ставимо дату
  if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
    sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
      Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
    );
  }

  writeLog('updateStatus', sheetName, rowNum, oldStatus + ' → ' + newStatus, '');

  return { success: true, sheet: sheetName, rowNum: rowNum, status: newStatus, oldStatus: oldStatus };
}

// ============================================
// bulkUpdateStatus — Масова зміна статусу
// ============================================
function bulkUpdateStatus(data) {
  var items = data.items;
  var newStatus = data.status;

  if (!items || !items.length || !newStatus) {
    return { success: false, error: 'Відсутні items або status' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  var needDate = ARCHIVE_STATUSES.indexOf(newStatus) !== -1;
  var count = 0;
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = findSheet(ss, item.sheet);
    if (!sheet) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Аркуш не знайдено' });
      continue;
    }
    if (item.rowNum > sheet.getLastRow()) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Рядок не існує' });
      continue;
    }

    sheet.getRange(item.rowNum, COL.STATUS + 1).setValue(newStatus);
    if (needDate) {
      sheet.getRange(item.rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
    }
    count++;
  }

  writeLog('bulkUpdateStatus', 'bulk', 0, newStatus, count + '/' + items.length + ' оновлено');

  return {
    success: true,
    count: count,
    total: items.length,
    status: newStatus,
    errors: errors.length > 0 ? errors : undefined
  };
}

// ============================================
// deletePackage — Видалити (статус = deleted + дата)
// ============================================
function deletePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;

  if (!sheetName || !rowNum) {
    return { success: false, error: 'Відсутні sheet або rowNum' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  // Зберігаємо інфо для логу
  var recordId = String(sheet.getRange(rowNum, COL.ID + 1).getValue() || '');
  var recordName = String(sheet.getRange(rowNum, COL.NAME + 1).getValue() || '');

  // Позначаємо як видалений
  sheet.getRange(rowNum, COL.STATUS + 1).setValue('deleted');
  sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
    Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
  );

  writeLog('deletePackage', sheetName, rowNum, 'deleted',
    'ІД: ' + recordId + ' | ПіБ: ' + recordName);

  return { success: true, sheet: sheetName, rowNum: rowNum, id: recordId };
}

// ============================================
// bulkAssignVehicle — Масове призначення авто
// ============================================
function bulkAssignVehicle(data) {
  var items = data.items;
  var vehicle = data.vehicle;

  if (!items || !items.length || !vehicle) {
    return { success: false, error: 'Відсутні items або vehicle' };
  }

  writeLog('bulkAssignVehicle', 'bulk', 0, vehicle, items.length + ' items');

  return { success: true, count: items.length, vehicle: vehicle };
}

// ============================================
// archivePackage — Архівувати одну посилку
// Відправляє в Crm_Arhiv_1.0 через HTTP
// ============================================
function archivePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var user = data.user || 'crm';
  var reason = data.reason || 'manual';

  if (!sheetName || !rowNum) {
    return { success: false, error: 'Відсутні sheet або rowNum' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  // Читаємо поточний рядок
  var rowData = sheet.getRange(rowNum, 1, 1, TOTAL_COLS).getValues()[0];
  var recordId = String(rowData[COL.ID] || '').trim();

  // Перевірка: чи не вже архівований
  var existingArchiveId = String(rowData[COL.ARCHIVE_ID] || '').trim();
  if (existingArchiveId) {
    return {
      success: false,
      error: 'already_archived',
      message: 'Запис вже архівований: ' + existingArchiveId
    };
  }

  // Ставимо статус archived + дату
  sheet.getRange(rowNum, COL.STATUS + 1).setValue('archived');
  sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
    Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
  );

  // Відправляємо в архівний скрипт
  var archiveResult = sendToArchive({
    action: 'archiveByIds',
    source: 'BOT_POSYLKY',
    sheet: sheetName,
    ids: [recordId],
    user: user,
    reason: reason,
    deleteFromSource: false  // не видаляємо — CRM сам керує
  });

  writeLog('archivePackage', sheetName, rowNum, 'archived',
    'ІД: ' + recordId + ' | Архів: ' + (archiveResult.success ? 'OK' : archiveResult.error));

  return {
    success: true,
    sheet: sheetName,
    rowNum: rowNum,
    id: recordId,
    archiveResult: archiveResult
  };
}

// ============================================
// bulkArchive — Масова архівація
// ============================================
function bulkArchive(data) {
  var items = data.items; // масив { sheet, rowNum }
  var user = data.user || 'crm';
  var reason = data.reason || 'bulk';

  if (!items || !items.length) {
    return { success: false, error: 'Відсутні items' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  var archivedIds = [];
  var count = 0;
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = findSheet(ss, item.sheet);
    if (!sheet) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Аркуш не знайдено' });
      continue;
    }
    if (item.rowNum > sheet.getLastRow()) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Рядок не існує' });
      continue;
    }

    var rowData = sheet.getRange(item.rowNum, 1, 1, TOTAL_COLS).getValues()[0];
    var existingArchiveId = String(rowData[COL.ARCHIVE_ID] || '').trim();
    if (existingArchiveId) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Вже архівовано: ' + existingArchiveId });
      continue;
    }

    var recordId = String(rowData[COL.ID] || '').trim();

    // Ставимо статус
    sheet.getRange(item.rowNum, COL.STATUS + 1).setValue('archived');
    sheet.getRange(item.rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);

    if (recordId) {
      archivedIds.push(recordId);
    }
    count++;
  }

  // Відправляємо всі ІД в архів одним запитом
  var archiveResult = { success: false, error: 'Немає ІД для архівації' };
  if (archivedIds.length > 0) {
    archiveResult = sendToArchive({
      action: 'archiveByIds',
      source: 'BOT_POSYLKY',
      ids: archivedIds,
      user: user,
      reason: reason,
      deleteFromSource: false
    });
  }

  writeLog('bulkArchive', 'bulk', 0, 'archived',
    count + '/' + items.length + ' архівовано | Архів: ' + (archiveResult.success ? 'OK' : archiveResult.error));

  return {
    success: true,
    count: count,
    total: items.length,
    errors: errors.length > 0 ? errors : undefined,
    archiveResult: archiveResult
  };
}

// ============================================
// checkDuplicates — Перевірити дублікати
// ============================================
// Параметри:
//   ttn: номер ТТН
//   phone: телефон
//   id: ІД запису
//   name: ПіБ (для м'якої перевірки з телефоном)
//   excludeRow: { sheet, rowNum } — виключити цей рядок
// ============================================
function checkDuplicates(data) {
  var checkTTN = data.ttn ? String(data.ttn).trim() : '';
  var checkPhone = data.phone ? String(data.phone).trim() : '';
  var checkId = data.id ? String(data.id).trim() : '';
  var checkName = data.name ? String(data.name).trim().toLowerCase() : '';
  var excludeSheet = data.excludeRow ? data.excludeRow.sheet : '';
  var excludeRowNum = data.excludeRow ? data.excludeRow.rowNum : 0;

  if (!checkTTN && !checkPhone && !checkId) {
    return { success: true, duplicates: [], count: 0 };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var duplicates = [];

  var sheetsToCheck = [
    { sheet: findSheet(ss, SHEET_REG), name: SHEET_REG },
    { sheet: findSheet(ss, SHEET_COURIER), name: SHEET_COURIER }
  ];

  for (var s = 0; s < sheetsToCheck.length; s++) {
    var chkSheet = sheetsToCheck[s].sheet;
    var chkName = sheetsToCheck[s].name;
    if (!chkSheet) continue;

    var lastRow = chkSheet.getLastRow();
    if (lastRow < 2) continue;

    var values = chkSheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();

    for (var r = 0; r < values.length; r++) {
      var rowNum = r + 2;
      var row = values[r];

      // Виключаємо переданий рядок
      if (chkName === excludeSheet && rowNum === excludeRowNum) continue;

      if (isEmptyRow(row)) continue;

      var existStatus = String(row[COL.STATUS] || '').toLowerCase().trim();
      // Пропускаємо архівовані/видалені
      if (ARCHIVE_STATUSES.indexOf(existStatus) !== -1) continue;

      var existTTN = String(row[COL.TTN] || '').trim();
      var existPhone = String(row[COL.PHONE] || '').trim();
      var existId = String(row[COL.ID] || '').trim();
      var existName = String(row[COL.NAME] || '').trim().toLowerCase();

      var matchReasons = [];

      if (checkId && existId && checkId === existId) {
        matchReasons.push('ІД');
      }
      if (checkTTN && existTTN && checkTTN === existTTN) {
        matchReasons.push('ТТН');
      }
      if (checkPhone && existPhone && checkPhone === existPhone && checkName && existName && checkName === existName) {
        matchReasons.push('Телефон+ПіБ');
      }

      if (matchReasons.length > 0) {
        duplicates.push({
          sheet: chkName,
          rowNum: rowNum,
          id: existId,
          ttn: existTTN,
          phone: existPhone,
          name: String(row[COL.NAME] || ''),
          status: existStatus,
          dateReg: formatDate(row[COL.DATE_REG]),
          matchReasons: matchReasons
        });
      }
    }
  }

  return {
    success: true,
    duplicates: duplicates,
    count: duplicates.length
  };
}

// ============================================
// getStructure — Структура таблиці (дебаг)
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
      try {
        return JSON.parse(body);
      } catch (e) {
        return { success: false, error: 'Невалідна відповідь від архіву' };
      }
    } else {
      return { success: false, error: 'HTTP ' + code + ': ' + body.substring(0, 200) };
    }
  } catch (e) {
    Logger.log('Archive API error: ' + e.toString());
    return { success: false, error: 'Архів недоступний: ' + e.toString() };
  }
}

// ============================================
// ЛОГУВАННЯ — пише в архівну таблицю, аркуш "Логи"
// ============================================
var ARCHIVE_SS_ID_LOG = '1Kmf6NF1sJUi-j3SamrhUqz337pcZSvZCUkGxBzari6U';

function writeLog(action, sheetName, rowNum, detail, extra) {
  try {
    var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID_LOG);
    var logSheet = archiveSS.getSheetByName('Логи');

    if (!logSheet) {
      logSheet = archiveSS.insertSheet('Логи');
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
// ДОПОМІЖНІ ФУНКЦІЇ
// ============================================

// Знайти аркуш (з fallback для апострофів)
function findSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  if (name.indexOf('Виклик кур') === 0 || name === SHEET_COURIER) {
    return findSheetFuzzy(ss, 'Виклик кур');
  }
  return null;
}

// Fuzzy пошук аркуша
function findSheetFuzzy(ss, prefix) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(prefix) === 0) {
      return sheets[i];
    }
  }
  return null;
}

// Перевірка порожнього рядка
function isEmptyRow(row) {
  for (var c = 0; c < row.length && c < TOTAL_COLS; c++) {
    if (row[c] !== '' && row[c] !== null && row[c] !== undefined) {
      return false;
    }
  }
  return true;
}

// Форматування дати → YYYY-MM-DD
function formatDate(value) {
  if (!value) return '';

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }

  var str = String(value).trim();
  if (!str) return '';

  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.substring(0, 10);

  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(str)) {
    var p = str.split('.');
    return p[2] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[0]).slice(-2);
  }

  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    var p2 = str.split('/');
    return p2[2] + '-' + ('0' + p2[1]).slice(-2) + '-' + ('0' + p2[0]).slice(-2);
  }

  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd');
    }
  } catch (e) {}

  return '';
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
  ui.createMenu('CRM Посилки')
    .addItem('Структура таблиці', 'testStructure')
    .addItem('Тест: всі посилки', 'testGetAll')
    .addItem('Тест: знайти аркуші', 'testFindSheets')
    .addItem('Тест: архів зв\'язок', 'testArchiveConnection')
    .addToUi();
}

// ============================================
// ТЕСТИ
// ============================================

function testFindSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  Logger.log('=== ВСІ АРКУШІ ===');
  for (var i = 0; i < sheets.length; i++) {
    Logger.log('  [' + i + '] "' + sheets[i].getName() + '" (' + sheets[i].getLastRow() + ' рядків)');
  }

  Logger.log('');
  var regSheet = findSheet(ss, SHEET_REG);
  Logger.log('Реєстрація ТТН: ' + (regSheet ? 'ЗНАЙДЕНО (' + regSheet.getLastRow() + ' рядків)' : 'НЕ ЗНАЙДЕНО'));

  var courierSheet = findSheet(ss, SHEET_COURIER);
  Logger.log('Виклик курєра: ' + (courierSheet ? 'ЗНАЙДЕНО (' + courierSheet.getLastRow() + ' рядків)' : 'НЕ ЗНАЙДЕНО'));
}

function testGetAll() {
  var result = getAllPackages();
  Logger.log('Всього посилок: ' + result.count);

  if (result.packages && result.packages.length > 0) {
    for (var i = 0; i < Math.min(5, result.packages.length); i++) {
      var p = result.packages[i];
      Logger.log(
        '[' + p.sheet + ' #' + p.rowNum + '] ' +
        'ПіБ: ' + (p.name || '-') +
        ' | ТТН: ' + (p.ttn || '-') +
        ' | Статус: ' + (p.status || '(пусто)') +
        ' | ArchiveID: ' + (p.archiveId || '-')
      );
    }
  }
}

function testStructure() {
  var result = getStructure();
  Logger.log('=== СТРУКТУРА ===');
  for (var i = 0; i < result.sheets.length; i++) {
    var s = result.sheets[i];
    Logger.log('[' + s.sheet + '] ' + s.rows + ' рядків, ' + s.cols + ' колонок');
    Logger.log('  Колонки: ' + s.headers.join(' | '));
  }
}

function testArchiveConnection() {
  Logger.log('=== ТЕСТ ЗВ\'ЯЗКУ З АРХІВОМ ===');
  Logger.log('URL: ' + ARCHIVE_API_URL);

  var result = sendToArchive({ action: 'getStats' });
  if (result.success) {
    Logger.log('Зв\'язок: OK');
    Logger.log('Статистика: ' + JSON.stringify(result.stats, null, 2));
  } else {
    Logger.log('ПОМИЛКА: ' + result.error);
  }
}
