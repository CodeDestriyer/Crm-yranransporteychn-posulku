// ============================================
// ЮРА ТРАНСПОРТЕЙШН — CRM ПАСАЖИРИ v1.0
// Apps Script API для таблиці "Бот ПАСАЖИРИ"
// ID: 1U1deQJvMPZ9fctIEoHCXr8cFQmgWLVe2VRhlzb5IpjI
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Бот ПАСАЖИРИ" → Розширення → Apps Script
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

var SPREADSHEET_ID = '1U1deQJvMPZ9fctIEoHCXr8cFQmgWLVe2VRhlzb5IpjI';

// Таблиця маршрутів пасажирів
var ROUTE_SPREADSHEET_ID = '1iKlD0Bj-5qB3Gc1d5ZBHscbRipcSe5xU7svqBfpB77Y';

// URL архівного скрипта (Crm_Arhiv_1.0)
var ARCHIVE_API_URL = 'https://script.google.com/macros/s/AKfycbwJLGZgYT333VdMW-nM5kPjYs2WIGGjfqkZnDJYjJxUt8nzE8GDGCPm7EzMHhcxNDOn/exec';

// Аркуші
var SHEET_UA_EU = 'Україна-єв';    // UA→EU пасажири
var SHEET_EU_UA = 'Європа-ук';     // EU→UA пасажири
var SHEET_LOGS = 'Логи';

// Маппінг авто → аркуш маршруту
// Нові назви = назва аркуша напряму. Старі назви "Авто N" залишені для зворотної сумісності.
var VEHICLE_TO_ROUTE = {
  'Пас. Маршрут 1': 'Пас. Маршрут 1',
  'Пас. Маршрут 2': 'Пас. Маршрут 2',
  'Пас. Маршрут 3': 'Пас. Маршрут 3',
  'Авто 1': 'Пас. Маршрут 1',
  'Авто 2': 'Пас. Маршрут 2',
  'Авто 3': 'Пас. Маршрут 3'
};

// Google Maps API ключ
var API_KEY = 'AIzaSyCthPzhD6zDM9zR-re0R2ceohyhCRdawNc';

// Точка старту за замовчуванням
var DEFAULT_START = { name: 'Ужгород', lat: 48.6209, lng: 22.2879 };
var MAX_POINTS_PER_MAP = 25;

// Колонки (A-T = 20 колонок, індекс 0-19)
// A:Дата виїзду  B:Адреса Відправки!  C:Адреса прибуття  D:Кількість місць
// E:ПіБ  F:Телефон Пасажира  G:Відмітка  H:Оплата  I:Відсоток  J:Диспечер
// K:ІД  L:Телефон Реєстратора  M:Вага  N:Автомобіль  O:Таймінг
// P:дата оформлення  Q:Примітка  R:Статус  S:Дата архів  T:ARCHIVE_ID
var COL = {
  DATE: 0,          // A — Дата виїзду
  FROM: 1,          // B — Адреса Відправки
  TO: 2,            // C — Адреса прибуття
  SEATS: 3,         // D — Кількість місць
  NAME: 4,          // E — ПіБ
  PHONE: 5,         // F — Телефон Пасажира
  MARK: 6,          // G — Відмітка
  PAYMENT: 7,       // H — Оплата
  PERCENT: 8,       // I — Відсоток
  DISPATCHER: 9,    // J — Диспечер
  ID: 10,           // K — ІД
  PHONE_REG: 11,    // L — Телефон Реєстратора
  WEIGHT: 12,       // M — Вага
  VEHICLE: 13,      // N — Автомобіль
  TIMING: 14,       // O — Таймінг
  DATE_REG: 15,     // P — дата оформлення
  NOTE: 16,         // Q — Примітка
  STATUS: 17,       // R — Статус (new/work/route/archived/refused/transferred/deleted)
  DATE_ARCHIVE: 18, // S — Дата архів
  ARCHIVE_ID: 19    // T — ARCHIVE_ID (зв'язок з таблицею Архіви)
};
var TOTAL_COLS = 20;

// Статуси для архівації
var ARCHIVE_STATUSES = ['archived', 'refused', 'deleted', 'transferred'];

// Маппінг полів CRM → індексів колонок
var FIELD_MAP = {
  date: COL.DATE,
  from: COL.FROM,
  to: COL.TO,
  seats: COL.SEATS,
  name: COL.NAME,
  phone: COL.PHONE,
  mark: COL.MARK,
  payment: COL.PAYMENT,
  percent: COL.PERCENT,
  dispatcher: COL.DISPATCHER,
  id: COL.ID,
  phoneReg: COL.PHONE_REG,
  weight: COL.WEIGHT,
  vehicle: COL.VEHICLE,
  timing: COL.TIMING,
  dateReg: COL.DATE_REG,
  note: COL.NOTE,
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
    var payload = data.payload || data;

    switch (action) {
      // --- ЧИТАННЯ ---
      case 'getAll':
        return respond(getAllPassengers());

      case 'getUaEu':
        return respond(getSheetPassengers(SHEET_UA_EU, 'ua-eu'));

      case 'getEuUa':
        return respond(getSheetPassengers(SHEET_EU_UA, 'eu-ua'));

      case 'getStructure':
        return respond(getStructure());

      // --- СТВОРЕННЯ ---
      case 'addPassenger':
        return respond(addPassenger(payload));

      // --- ОНОВЛЕННЯ ---
      case 'updatePassenger':
        return respond(updatePassenger(payload));

      case 'updateMultiple':
        return respond(updateMultiplePassengers(payload.passengers || payload));

      case 'updateField':
        return respond(updateField(payload));

      case 'updateStatus':
        return respond(updateStatus(payload));

      // --- СТАТУСИ (масові) ---
      case 'archivePassengers':
        return respond(changePassengersStatus(payload, 'archived'));

      case 'restorePassengers':
        return respond(changePassengersStatus(payload, 'work'));

      case 'refusePassengers':
        return respond(changePassengersStatus(payload, 'refused'));

      case 'transferPassengers':
        return respond(changePassengersStatus(payload, 'transferred'));

      case 'deletePassengers':
        return respond(changePassengersStatus(payload, 'deleted'));

      // --- ВИДАЛЕННЯ ---
      case 'deletePassengersPermanently':
        return respond(deletePassengersPermanently(payload));

      // --- АРХІВАЦІЯ (з відправкою в Crm_Arhiv_1.0) ---
      case 'archivePassenger':
        return respond(archivePassenger(payload));

      case 'bulkArchive':
        return respond(bulkArchive(payload));

      // --- ДУБЛІКАТИ ---
      case 'checkDuplicates':
        return respond(checkDuplicates(payload));

      // --- МАРШРУТИ ---
      case 'optimize':
        return respond(optimizeRoute(payload));

      case 'copyToRoute':
        return respond(copyToRouteSheet(payload));

      case 'checkRouteSheets':
        return respond(checkRouteSheets(payload));

      case 'createRouteSheet':
        return respond(createRouteSheet(payload));

      case 'getRoutePassengers':
        return respond(getRoutePassengers(payload));

      case 'getAvailableRoutes':
        return respond(getAvailableRoutes());

      case 'deleteRouteSheet':
        return respond(deleteRouteSheet(payload));

      // --- РОЗСИЛКА ---
      case 'getMailingStatus':
        return respond(getMailingStatus());

      case 'addMailingRecord':
        return respond(addMailingRecord(payload));

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
          service: 'CRM Пасажири — ЮРА ТРАНСПОРТЕЙШН',
          sheets: [SHEET_UA_EU, SHEET_EU_UA],
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
// getAll — Витягнути ВСІ пасажирів з обох аркушів
// ============================================
function getAllPassengers() {
  var uaEu = getSheetPassengers(SHEET_UA_EU, 'ua-eu');
  var euUa = getSheetPassengers(SHEET_EU_UA, 'eu-ua');

  var allPassengers = [];
  if (uaEu.passengers) allPassengers = allPassengers.concat(uaEu.passengers);
  if (euUa.passengers) allPassengers = allPassengers.concat(euUa.passengers);

  return {
    success: true,
    passengers: allPassengers,
    counts: {
      total: allPassengers.length,
      uaEu: uaEu.passengers ? uaEu.passengers.length : 0,
      euUa: euUa.passengers ? euUa.passengers.length : 0
    },
    timestamp: new Date().toISOString()
  };
}

// ============================================
// getSheetPassengers — Читання одного аркуша
// ============================================
function getSheetPassengers(sheetName, direction) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: false, passengers: [], error: 'Аркуш не знайдено: ' + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, passengers: [], count: 0 };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
    var passengers = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.NAME] && !row[COL.PHONE]) continue;

      var dateReg = formatDate(row[COL.DATE_REG]);
      var crmStatus = resolveStatus(row);

      passengers.push({
        id: String(row[COL.ID] || ''),
        rowNum: i + 2,
        sheet: sheetName,
        direction: direction,

        date: formatDate(row[COL.DATE]),
        from: String(row[COL.FROM] || ''),
        to: String(row[COL.TO] || ''),
        seats: parseInt(row[COL.SEATS]) || 1,
        name: String(row[COL.NAME] || ''),
        phone: String(row[COL.PHONE] || ''),
        mark: String(row[COL.MARK] || ''),
        payment: String(row[COL.PAYMENT] || ''),
        percent: String(row[COL.PERCENT] || ''),
        dispatcher: String(row[COL.DISPATCHER] || ''),
        phoneReg: String(row[COL.PHONE_REG] || ''),
        weight: String(row[COL.WEIGHT] || ''),
        vehicle: String(row[COL.VEHICLE] || ''),
        timing: String(row[COL.TIMING] || ''),
        dateReg: dateReg,
        note: String(row[COL.NOTE] || ''),
        status: crmStatus,
        dateArchive: formatDate(row[COL.DATE_ARCHIVE]),
        archiveId: String(row[COL.ARCHIVE_ID] || ''),

        isNew: isRecent(row[COL.DATE_REG] || row[COL.DATE], 24),
        isArchived: ARCHIVE_STATUSES.indexOf(crmStatus) !== -1
      });
    }

    return { success: true, passengers: passengers, count: passengers.length };
  } catch (err) {
    return { success: false, passengers: [], error: err.toString() };
  }
}

// ============================================
// addPassenger — Додати пасажира
// + перевірка дублікатів
// ============================================
function addPassenger(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var direction = payload.direction || 'ua-eu';
    var sheetName = direction === 'eu-ua' ? SHEET_EU_UA : SHEET_UA_EU;
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    // --- ПЕРЕВІРКА ДУБЛІКАТІВ ---
    var checkPhone = payload.phone ? String(payload.phone).trim() : '';
    var checkName = payload.name ? String(payload.name).trim().toLowerCase() : '';
    var checkDate = payload.date ? String(payload.date).trim() : '';
    var duplicates = [];

    if (checkPhone || checkName) {
      var sheetsToCheck = [
        ss.getSheetByName(SHEET_UA_EU),
        ss.getSheetByName(SHEET_EU_UA)
      ];

      for (var s = 0; s < sheetsToCheck.length; s++) {
        var chkSheet = sheetsToCheck[s];
        if (!chkSheet) continue;

        var lastRow = chkSheet.getLastRow();
        if (lastRow < 2) continue;

        var values = chkSheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
        for (var r = 0; r < values.length; r++) {
          var existRow = values[r];
          var existStatus = String(existRow[COL.STATUS] || '').toLowerCase().trim();
          if (ARCHIVE_STATUSES.indexOf(existStatus) !== -1) continue;

          var existPhone = String(existRow[COL.PHONE] || '').trim();
          var existName = String(existRow[COL.NAME] || '').trim().toLowerCase();
          var existDate = formatDate(existRow[COL.DATE]);

          var isDuplicate = false;
          var reason = '';

          // Дублікат: той самий телефон + ім'я + дата
          if (checkPhone && existPhone && checkPhone === existPhone &&
              checkName && existName && checkName === existName &&
              checkDate && existDate && checkDate === existDate) {
            isDuplicate = true;
            reason = 'Телефон+ПіБ+Дата';
          }
          // М'який дублікат: телефон + ім'я (без дати)
          else if (checkPhone && existPhone && checkPhone === existPhone &&
                   checkName && existName && checkName === existName) {
            isDuplicate = true;
            reason = 'Телефон+ПіБ';
          }

          if (isDuplicate) {
            duplicates.push({
              sheet: chkSheet.getName(),
              rowNum: r + 2,
              phone: existPhone,
              name: String(existRow[COL.NAME] || ''),
              date: existDate,
              status: existStatus,
              reason: reason
            });
          }
        }
      }
    }

    // Блокуємо якщо є дублікати і не force
    if (duplicates.length > 0 && !payload.force) {
      writeLog('addPassenger:DUPLICATE', sheetName, 0, 'blocked',
        'Знайдено ' + duplicates.length + ' дублікатів | ' + duplicates[0].reason);

      return {
        success: false,
        error: 'duplicate',
        message: 'Знайдено ' + duplicates.length + ' можливих дублікатів',
        duplicates: duplicates
      };
    }

    // --- СТВОРЕННЯ ---
    var initials = payload.initials || 'CRM';
    var newId = generateId(initials, sheet);
    var today = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');

    var newRow = new Array(TOTAL_COLS);
    for (var i = 0; i < TOTAL_COLS; i++) newRow[i] = '';

    newRow[COL.DATE] = payload.date || '';
    newRow[COL.FROM] = payload.from || '';
    newRow[COL.TO] = payload.to || '';
    newRow[COL.SEATS] = payload.seats || 1;
    newRow[COL.NAME] = payload.name || '';
    newRow[COL.PHONE] = payload.phone || '';
    newRow[COL.MARK] = payload.mark || '';
    newRow[COL.PAYMENT] = payload.payment || '';
    newRow[COL.PERCENT] = payload.percent || '';
    newRow[COL.DISPATCHER] = payload.dispatcher || 'CRM';
    newRow[COL.ID] = newId;
    newRow[COL.PHONE_REG] = payload.phoneReg || '';
    newRow[COL.WEIGHT] = payload.weight || '';
    newRow[COL.VEHICLE] = payload.vehicle || '';
    newRow[COL.TIMING] = payload.timing || '';
    newRow[COL.DATE_REG] = today;
    newRow[COL.NOTE] = payload.note || '';
    newRow[COL.STATUS] = 'new';

    sheet.appendRow(newRow);
    var newRowNum = sheet.getLastRow();

    writeLog('addPassenger', sheetName, newRowNum, 'new',
      'ПіБ: ' + (payload.name || '') + ' | Тел: ' + (payload.phone || '') +
      (duplicates.length > 0 ? ' | FORCE (дублікат ігноровано)' : ''));

    return {
      success: true,
      sheet: sheetName,
      rowNum: newRowNum,
      id: newId,
      direction: direction,
      duplicatesIgnored: duplicates.length
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// updatePassenger — Оновити пасажира
// + верифікація по ІД
// ============================================
function updatePassenger(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(payload.sheet);

    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + payload.sheet };
    }

    var rowNum = parseInt(payload.rowNum);
    if (!rowNum || rowNum < 2) {
      return { success: false, error: 'Невірний rowNum' };
    }
    if (rowNum > sheet.getLastRow()) {
      return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
    }

    // --- ВЕРИФІКАЦІЯ ПО ІД ---
    if (payload.expectedId) {
      var currentId = String(sheet.getRange(rowNum, COL.ID + 1).getValue() || '').trim();
      if (currentId !== String(payload.expectedId).trim()) {
        writeLog('updatePassenger:CONFLICT', payload.sheet, rowNum, 'blocked',
          'Очікувався ІД: ' + payload.expectedId + ', фактичний: ' + currentId);
        return {
          success: false,
          error: 'conflict',
          message: 'Рядок змінився. Очікувався ІД: ' + payload.expectedId + ', фактичний: ' + currentId
        };
      }
    }

    // --- ПЕРЕВІРКА ЧИ НЕ АРХІВОВАНИЙ ---
    var currentStatus = String(sheet.getRange(rowNum, COL.STATUS + 1).getValue() || '').toLowerCase().trim();
    if (ARCHIVE_STATUSES.indexOf(currentStatus) !== -1 && !payload.force) {
      return {
        success: false,
        error: 'archived',
        message: 'Запис архівований (статус: ' + currentStatus + '). Використайте force=true'
      };
    }

    // --- ОНОВЛЕННЯ ПОЛІВ ---
    var updated = [];
    for (var field in payload) {
      if (payload.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
        if (payload[field] !== undefined) {
          sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(payload[field]);
          updated.push(field);
        }
      }
    }

    writeLog('updatePassenger', payload.sheet, rowNum, updated.join(', '), '');

    return { success: true, updated: updated, sheet: payload.sheet, rowNum: rowNum };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// updateMultiple — Масове оновлення
// ============================================
function updateMultiplePassengers(passengers) {
  var updated = 0;
  var errors = [];

  for (var i = 0; i < passengers.length; i++) {
    var result = updatePassenger(passengers[i]);
    if (result.success) {
      updated++;
    } else {
      errors.push((passengers[i].id || 'unknown') + ': ' + result.error);
    }
  }

  return { success: true, updated: updated, errors: errors.length > 0 ? errors : undefined };
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
  if (!sheet) return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Рядок не існує' };

  sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(value);
  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}

// ============================================
// updateStatus — Змінити статус одного пасажира
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

  var oldStatus = String(sheet.getRange(rowNum, COL.STATUS + 1).getValue() || '').toLowerCase().trim();
  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
    sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
      Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
    );
  }

  writeLog('updateStatus', sheetName, rowNum, oldStatus + ' → ' + newStatus, '');

  return { success: true, sheet: sheetName, rowNum: rowNum, status: newStatus, oldStatus: oldStatus };
}

// ============================================
// changePassengersStatus — Масова зміна статусу
// ============================================
function changePassengersStatus(payload, newStatus) {
  try {
    var passengers = payload.passengers || [];
    var note = payload.note || '';

    if (passengers.length === 0) {
      return { success: false, error: 'Немає пасажирів' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var today = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var changed = 0;
    var errors = [];

    for (var i = 0; i < passengers.length; i++) {
      var p = passengers[i];
      var sheet = ss.getSheetByName(p.sheet);
      if (!sheet) { errors.push(p.sheet + ': не знайдено'); continue; }

      var row = parseInt(p.rowNum);
      if (!row || row < 2 || row > sheet.getLastRow()) {
        errors.push('Рядок ' + row + ': не існує');
        continue;
      }

      var oldStatus = String(sheet.getRange(row, COL.STATUS + 1).getValue() || '');
      sheet.getRange(row, COL.STATUS + 1).setValue(newStatus);

      if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
        sheet.getRange(row, COL.DATE_ARCHIVE + 1).setValue(today);
      }

      if (note) {
        var currentNote = String(sheet.getRange(row, COL.NOTE + 1).getValue() || '');
        var newNote = note + (currentNote ? ' | ' + currentNote : '');
        sheet.getRange(row, COL.NOTE + 1).setValue(newNote);
      }

      changed++;
    }

    writeLog('changeStatus:' + newStatus, 'bulk', 0, changed + '/' + passengers.length,
      note || '');

    return {
      success: true,
      changed: changed,
      total: passengers.length,
      status: newStatus,
      errors: errors.length > 0 ? errors : undefined
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
    var passengers = payload.passengers || [];
    if (passengers.length === 0) {
      return { success: false, error: 'Немає пасажирів' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var deleted = 0;

    // Групуємо по аркушам
    var bySheet = {};
    for (var i = 0; i < passengers.length; i++) {
      var p = passengers[i];
      if (!bySheet[p.sheet]) bySheet[p.sheet] = [];
      bySheet[p.sheet].push(parseInt(p.rowNum));
    }

    // Видаляємо знизу вгору
    for (var sheetName in bySheet) {
      if (!bySheet.hasOwnProperty(sheetName)) continue;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;

      var rows = bySheet[sheetName].sort(function(a, b) { return b - a; });
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
// archivePassenger — Архівувати одного пасажира
// Відправляє в Crm_Arhiv_1.0
// ============================================
function archivePassenger(payload) {
  var sheetName = payload.sheet;
  var rowNum = parseInt(payload.rowNum);
  var user = payload.user || 'crm';
  var reason = payload.reason || 'manual';

  if (!sheetName || !rowNum) {
    return { success: false, error: 'Відсутні sheet або rowNum' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Аркуш не знайдено' };
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Рядок не існує' };

  var rowData = sheet.getRange(rowNum, 1, 1, TOTAL_COLS).getValues()[0];

  // Перевірка: вже архівований?
  var existingArchiveId = String(rowData[COL.ARCHIVE_ID] || '').trim();
  if (existingArchiveId) {
    return {
      success: false,
      error: 'already_archived',
      message: 'Вже архівовано: ' + existingArchiveId
    };
  }

  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss');
  var archiveId = generateArchiveId_();

  // === КРОК 1: Пишемо НАПРЯМУ в архівну таблицю ===
  try {
    var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID_LOG);
    var archiveSheet = archiveSS.getSheetByName('Пасажири');
    if (!archiveSheet) {
      return { success: false, error: 'Архівний аркуш "Пасажири" не знайдено' };
    }

    // Будуємо рядок: 23 колонки (A-W)
    // A-R (0-17): дані | S(18): дата | T(19): хто | U(20): причина | V(21): аркуш | W(22): ARCHIVE_ID
    var archiveRow = [];
    for (var i = 0; i < 18; i++) {
      archiveRow.push(rowData[i] !== undefined ? rowData[i] : '');
    }
    archiveRow.push(dateNow);       // S - DATE_ARCHIVE
    archiveRow.push(user);          // T - ARCHIVED_BY
    archiveRow.push(reason);        // U - ARCHIVE_REASON
    archiveRow.push(sheetName);     // V - SOURCE_SHEET
    archiveRow.push(archiveId);     // W - ARCHIVE_ID

    archiveSheet.appendRow(archiveRow);
  } catch (err) {
    return { success: false, error: 'Помилка запису в архів: ' + err.toString() };
  }

  // === КРОК 2: Оновлюємо джерело (тільки після успішного запису) ===
  sheet.getRange(rowNum, COL.STATUS + 1).setValue('archived');
  sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow.substring(0, 10));
  sheet.getRange(rowNum, COL.ARCHIVE_ID + 1).setValue(archiveId);

  var recordId = String(rowData[COL.ID] || '');
  writeLog('archivePassenger', sheetName, rowNum, 'archived',
    'ІД: ' + recordId + ' | ArchiveID: ' + archiveId);

  return {
    success: true,
    sheet: sheetName,
    rowNum: rowNum,
    id: recordId,
    archiveId: archiveId
  };
}

// ============================================
// bulkArchive — Масова архівація
// Пише НАПРЯМУ в архівну таблицю (без HTTP)
// ============================================
function bulkArchive(payload) {
  var items = payload.items || payload.passengers || [];
  var user = payload.user || 'crm';
  var reason = payload.reason || 'bulk';

  if (!items.length) {
    return { success: false, error: 'Немає items' };
  }

  // Відкриваємо архівну таблицю
  var archiveSS;
  var archiveSheet;
  try {
    archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID_LOG);
    archiveSheet = archiveSS.getSheetByName('Пасажири');
    if (!archiveSheet) {
      return { success: false, error: 'Архівний аркуш "Пасажири" не знайдено' };
    }
  } catch (err) {
    return { success: false, error: 'Не вдалося відкрити архів: ' + err.toString() };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss');
  var dateShort = dateNow.substring(0, 10);
  var archiveRows = [];
  var successItems = [];
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = ss.getSheetByName(item.sheet);
    if (!sheet) { errors.push(item.sheet + ': не знайдено'); continue; }

    var rowNum = parseInt(item.rowNum);
    if (rowNum > sheet.getLastRow()) { errors.push('Рядок ' + rowNum + ': не існує'); continue; }

    var rowData = sheet.getRange(rowNum, 1, 1, TOTAL_COLS).getValues()[0];
    if (String(rowData[COL.ARCHIVE_ID] || '').trim()) {
      errors.push('Рядок ' + rowNum + ': вже архівовано');
      continue;
    }

    var archiveId = generateArchiveId_();

    // Будуємо рядок архіву: 23 колонки
    var archiveRow = [];
    for (var j = 0; j < 18; j++) {
      archiveRow.push(rowData[j] !== undefined ? rowData[j] : '');
    }
    archiveRow.push(dateNow);       // S - DATE_ARCHIVE
    archiveRow.push(user);          // T - ARCHIVED_BY
    archiveRow.push(reason);        // U - ARCHIVE_REASON
    archiveRow.push(item.sheet);    // V - SOURCE_SHEET
    archiveRow.push(archiveId);     // W - ARCHIVE_ID

    archiveRows.push(archiveRow);
    successItems.push({ sheet: item.sheet, rowNum: rowNum, archiveId: archiveId, srcSheet: sheet });
  }

  if (archiveRows.length === 0) {
    return { success: true, count: 0, total: items.length, errors: errors.length > 0 ? errors : undefined };
  }

  // === КРОК 1: Batch-запис в архів ===
  try {
    var startRow = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(startRow, 1, archiveRows.length, 23).setValues(archiveRows);
  } catch (err) {
    return { success: false, error: 'Помилка batch-запису в архів: ' + err.toString() };
  }

  // === КРОК 2: Оновлюємо джерело ===
  for (var k = 0; k < successItems.length; k++) {
    var si = successItems[k];
    si.srcSheet.getRange(si.rowNum, COL.STATUS + 1).setValue('archived');
    si.srcSheet.getRange(si.rowNum, COL.DATE_ARCHIVE + 1).setValue(dateShort);
    si.srcSheet.getRange(si.rowNum, COL.ARCHIVE_ID + 1).setValue(si.archiveId);
  }

  writeLog('bulkArchive', 'bulk', 0, 'archived',
    archiveRows.length + '/' + items.length + ' записано в архів');

  return {
    success: true,
    count: archiveRows.length,
    total: items.length,
    errors: errors.length > 0 ? errors : undefined
  };
}

// Генерація ARCHIVE_ID
function generateArchiveId_() {
  var now = new Date();
  var ts = Utilities.formatDate(now, 'Europe/Kiev', 'yyyyMMddHHmmss');
  var rnd = Math.floor(Math.random() * 10000).toString();
  while (rnd.length < 4) rnd = '0' + rnd;
  return 'ARC_' + ts + '_' + rnd;
}

// ============================================
// checkDuplicates — Перевірка дублікатів
// ============================================
function checkDuplicates(payload) {
  var checkPhone = payload.phone ? String(payload.phone).trim() : '';
  var checkName = payload.name ? String(payload.name).trim().toLowerCase() : '';
  var checkDate = payload.date ? String(payload.date).trim() : '';
  var excludeSheet = payload.excludeRow ? payload.excludeRow.sheet : '';
  var excludeRowNum = payload.excludeRow ? payload.excludeRow.rowNum : 0;

  if (!checkPhone && !checkName) {
    return { success: true, duplicates: [], count: 0 };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var duplicates = [];

  var sheetsToCheck = [SHEET_UA_EU, SHEET_EU_UA];
  for (var s = 0; s < sheetsToCheck.length; s++) {
    var sheet = ss.getSheetByName(sheetsToCheck[s]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    var values = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
    for (var r = 0; r < values.length; r++) {
      var rowNum = r + 2;
      if (sheetsToCheck[s] === excludeSheet && rowNum === excludeRowNum) continue;

      var row = values[r];
      var existStatus = String(row[COL.STATUS] || '').toLowerCase().trim();
      if (ARCHIVE_STATUSES.indexOf(existStatus) !== -1) continue;

      var existPhone = String(row[COL.PHONE] || '').trim();
      var existName = String(row[COL.NAME] || '').trim().toLowerCase();
      var matchReasons = [];

      if (checkPhone && existPhone && checkPhone === existPhone &&
          checkName && existName && checkName === existName) {
        matchReasons.push('Телефон+ПіБ');
      }

      if (matchReasons.length > 0) {
        duplicates.push({
          sheet: sheetsToCheck[s],
          rowNum: rowNum,
          id: String(row[COL.ID] || ''),
          phone: existPhone,
          name: String(row[COL.NAME] || ''),
          date: formatDate(row[COL.DATE]),
          status: existStatus,
          matchReasons: matchReasons
        });
      }
    }
  }

  return { success: true, duplicates: duplicates, count: duplicates.length };
}

// ============================================
// МАРШРУТИ — читання пасажирів з маршруту
// ============================================
function getRoutePassengers(payload) {
  try {
    var vehicleName = payload.vehicleName || '';
    var sheetName = payload.sheetName || '';

    if (vehicleName && !sheetName) {
      sheetName = VEHICLE_TO_ROUTE[vehicleName] || vehicleName;
    }
    if (!sheetName) {
      return { success: false, error: 'Не вказано аркуш маршруту' };
    }

    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);
    var sheet = routeSS.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, passengers: [], count: 0, sheetName: sheetName };
    }

    // Маршрут пасажирів: 23 колонки (A-W), читаємо 18 (A-R дані + Статус)
    var readCols = Math.min(sheet.getLastColumn(), 23);
    var dataRange = sheet.getRange(2, 1, lastRow - 1, readCols);
    var data = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();

    var passengers = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.NAME] && !row[COL.PHONE]) continue;

      // Статус водія з Відмітка або кольору рядка
      var driverStatus = resolveDriverStatus(row, backgrounds[i]);

      passengers.push({
        rowNum: i + 2,
        date: formatDate(row[COL.DATE]),
        from: String(row[COL.FROM] || ''),
        to: String(row[COL.TO] || ''),
        seats: parseInt(row[COL.SEATS]) || 1,
        name: String(row[COL.NAME] || ''),
        phone: String(row[COL.PHONE] || ''),
        mark: String(row[COL.MARK] || ''),
        payment: String(row[COL.PAYMENT] || ''),
        percent: String(row[COL.PERCENT] || ''),
        dispatcher: String(row[COL.DISPATCHER] || ''),
        id: String(row[COL.ID] || ''),
        phoneReg: String(row[COL.PHONE_REG] || ''),
        weight: String(row[COL.WEIGHT] || ''),
        vehicle: String(row[COL.VEHICLE] || ''),
        timing: String(row[COL.TIMING] || ''),
        note: String(row[COL.NOTE] || ''),
        status: String(row[COL.STATUS] || ''),
        driverStatus: driverStatus,
        rowColor: backgrounds[i][0],
        sheet: sheetName
      });
    }

    // Статистика
    var stats = { total: passengers.length, pending: 0, inProgress: 0, completed: 0, cancelled: 0 };
    for (var j = 0; j < passengers.length; j++) {
      var ds = passengers[j].driverStatus;
      if (ds === 'pending') stats.pending++;
      else if (ds === 'in-progress') stats.inProgress++;
      else if (ds === 'completed') stats.completed++;
      else if (ds === 'cancelled') stats.cancelled++;
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
// МАРШРУТИ — список доступних
// ============================================
function getAvailableRoutes() {
  try {
    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);
    var sheets = routeSS.getSheets();
    var routes = [];

    var excludePatterns = ['архів', 'template', 'шаблон', 'логи', 'logs', 'провірка', 'перевірка', 'розсилк', 'тест', 'test'];

    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      var nameLower = name.toLowerCase();

      var isExcluded = false;
      for (var e = 0; e < excludePatterns.length; e++) {
        if (nameLower.indexOf(excludePatterns[e]) !== -1) { isExcluded = true; break; }
      }
      if (isExcluded) continue;

      var count = Math.max(0, sheets[i].getLastRow() - 1);
      routes.push({ name: name, count: count, sheetId: sheets[i].getSheetId() });
    }

    return { success: true, routes: routes, count: routes.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// МАРШРУТИ — перевірка існуючих
// ============================================
function checkRouteSheets(payload) {
  try {
    var vehicleNames = payload.vehicleNames || [];
    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);
    var existing = [];

    for (var i = 0; i < vehicleNames.length; i++) {
      var routeSheetName = VEHICLE_TO_ROUTE[vehicleNames[i]];
      if (!routeSheetName) continue;
      var routeSheet = routeSS.getSheetByName(routeSheetName);
      if (!routeSheet) continue;
      var lastRow = routeSheet.getLastRow();
      if (lastRow > 1) {
        existing.push({ vehicle: vehicleNames[i], sheet: routeSheetName, count: lastRow - 1 });
      }
    }

    return { success: true, existing: existing };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// МАРШРУТИ — створення аркуша
// ============================================
function createRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву авто' };

    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);

    var existingSheet = routeSS.getSheetByName(vehicleName);
    if (existingSheet) {
      return { success: true, sheetName: vehicleName, existed: true };
    }

    var templateSheet = routeSS.getSheetByName('Пас. Маршрут 1');
    if (templateSheet) {
      var newSheet = templateSheet.copyTo(routeSS);
      newSheet.setName(vehicleName);
      if (newSheet.getLastRow() > 1) {
        newSheet.deleteRows(2, newSheet.getLastRow() - 1);
      }
    } else {
      var created = routeSS.insertSheet(vehicleName);
      var headers = ['Дата виїзду', 'Адреса Відправки!', 'Адреса прибуття', 'Кількість місць',
                     'ПіБ', 'Телефон Пасажира', 'Відмітка', 'Оплата', 'Відсоток',
                     'Диспечер', 'ІД', 'Телефон Реєстратора', 'Вага', 'Автомобіль',
                     'Таймінг', 'дата оформлення', 'Примітка', 'Статус'];
      created.getRange(1, 1, 1, headers.length).setValues([headers]);
      created.getRange(1, 1, 1, headers.length)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      created.setFrozenRows(1);
    }

    writeLog('createRouteSheet', vehicleName, 0, 'created', '');

    return { success: true, sheetName: vehicleName, vehicleName: vehicleName };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// МАРШРУТИ — копіювання пасажирів
// ============================================
function copyToRouteSheet(payload) {
  try {
    var passengersByVehicle = payload.passengersByVehicle;
    var conflictAction = payload.conflictAction || 'add';

    if (!passengersByVehicle) {
      return { success: false, error: 'Немає пасажирів' };
    }

    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);
    var totalCopied = 0;
    var totalArchived = 0;
    var totalCleared = 0;
    var results = [];

    for (var vehicleName in passengersByVehicle) {
      if (!passengersByVehicle.hasOwnProperty(vehicleName)) continue;
      var passengers = passengersByVehicle[vehicleName];

      var routeSheetName = VEHICLE_TO_ROUTE[vehicleName];
      if (!routeSheetName) {
        // Пробуємо знайти по назві
        var allSheets = routeSS.getSheets();
        for (var s = 0; s < allSheets.length; s++) {
          if (allSheets[s].getName().toLowerCase().indexOf(vehicleName.toLowerCase()) !== -1) {
            routeSheetName = allSheets[s].getName();
            break;
          }
        }
      }
      if (!routeSheetName) {
        results.push({ vehicle: vehicleName, error: 'Невідомий маршрут' });
        continue;
      }

      var routeSheet = routeSS.getSheetByName(routeSheetName);
      if (!routeSheet) {
        results.push({ vehicle: vehicleName, error: 'Аркуш не знайдено: ' + routeSheetName });
        continue;
      }

      var lastRow = routeSheet.getLastRow();
      if (lastRow > 1 && conflictAction !== 'add') {
        if (conflictAction === 'clear') {
          totalCleared += lastRow - 1;
          routeSheet.deleteRows(2, lastRow - 1);
        } else if (conflictAction === 'archive') {
          totalArchived += lastRow - 1;
          // Позначаємо як архівовані в маршруті
          var oldData = routeSheet.getRange(2, 1, lastRow - 1, routeSheet.getLastColumn()).getValues();
          for (var a = 0; a < oldData.length; a++) {
            oldData[a][COL.MARK] = 'Архів';
          }
          routeSheet.getRange(2, 1, lastRow - 1, routeSheet.getLastColumn()).setValues(oldData);
        }
      }

      for (var p = 0; p < passengers.length; p++) {
        var pass = passengers[p];
        var row = [
          pass.date || '', pass.from || '', pass.to || '',
          pass.seats || 1, pass.name || '', pass.phone || '',
          pass.mark || '', pass.payment || '', pass.percent || '',
          pass.dispatcher || '', pass.id || '', pass.phoneReg || '',
          pass.weight || '', vehicleName, pass.timing || '',
          pass.dateReg || '', pass.note || ''
        ];
        routeSheet.appendRow(row);
        totalCopied++;
      }

      results.push({ vehicle: vehicleName, sheet: routeSheetName, copied: passengers.length });
    }

    writeLog('copyToRoute', 'route', 0, 'copied: ' + totalCopied,
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
// МАРШРУТИ — видалення аркуша
// ============================================
function deleteRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву авто' };

    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);
    var sheetName = VEHICLE_TO_ROUTE[vehicleName] || vehicleName;
    var sheet = routeSS.getSheetByName(sheetName);

    if (!sheet) {
      return { success: true, message: 'Аркуш не існує', deleted: false };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow > 1 && !payload.force) {
      return {
        success: false,
        error: 'Аркуш містить ' + (lastRow - 1) + ' записів. Використайте force=true',
        recordsCount: lastRow - 1
      };
    }

    routeSS.deleteSheet(sheet);
    writeLog('deleteRouteSheet', sheetName, 0, 'deleted', '');

    return { success: true, sheetName: sheetName, deleted: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// РОЗСИЛКА — статус
// ============================================
function getMailingStatus() {
  try {
    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);
    var sheet = routeSS.getSheetByName('Провірка розсилки');

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, mailingIds: [], count: 0 };
    }

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    var mailingIds = [];

    for (var i = 0; i < data.length; i++) {
      var id = data[i][1];
      if (id && String(id).indexOf('dd.mm') === -1) {
        mailingIds.push(String(id).trim());
      }
    }

    return { success: true, mailingIds: mailingIds, count: mailingIds.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// РОЗСИЛКА — додати запис
// ============================================
function addMailingRecord(payload) {
  try {
    var routeSS = SpreadsheetApp.openById(ROUTE_SPREADSHEET_ID);
    var sheet = routeSS.getSheetByName('Провірка розсилки');

    if (!sheet) {
      sheet = routeSS.insertSheet('Провірка розсилки');
      sheet.getRange(1, 1, 1, 2).setValues([['Дата виїзду', 'ІД']]);
    }

    var records = payload.records || [];
    var userName = payload.userName || 'CRM';

    if (records.length === 0) {
      return { success: false, error: 'Немає записів' };
    }

    var today = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');
    var rows = [];
    for (var i = 0; i < records.length; i++) {
      rows.push([records[i].date || today, userName]);
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 2).setValues(rows);

    return { success: true, added: rows.length };
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
        if (coords) {
          passengers[g].coords = coords;
          geocodedCount++;
        } else {
          notGeocodedList.push({ id: passengers[g].id, name: passengers[g].name, address: passengers[g].rawAddress, uid: passengers[g].uid });
        }
      } catch (e) {
        notGeocodedList.push({ id: passengers[g].id, name: passengers[g].name, address: passengers[g].rawAddress, uid: passengers[g].uid });
      }
      Utilities.sleep(150);
    }

    // Старт
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

    // Оптимізація
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
      var idx = optimizedOrder[o];
      orderedPassengers.push(validPassengers[idx].originalData);
      orderedForMap.push(validPassengers[idx]);
    }
    for (var inv = 0; inv < invalidPassengers.length; inv++) {
      var d = invalidPassengers[inv].originalData;
      d._notGeocoded = true;
      orderedPassengers.push(d);
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
// getStructure — Структура таблиці (дебаг)
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

// Визначення CRM статусу з рядка
function resolveStatus(row) {
  var statusValue = String(row[COL.STATUS] || '').toLowerCase().trim();

  if (statusValue === 'archived' || statusValue === 'архів') return 'archived';
  if (statusValue === 'refused' || statusValue === 'відмова') return 'refused';
  if (statusValue === 'transferred' || statusValue === 'пересадка') return 'transferred';
  if (statusValue === 'deleted' || statusValue === 'видалено') return 'deleted';
  if (statusValue === 'route' || statusValue === 'маршрут') return 'route';
  if (statusValue === 'optimize' || statusValue === 'оптимізація') return 'optimize';
  if (statusValue === 'work' || statusValue === 'в роботі') return 'work';
  if (statusValue === 'new' || statusValue === 'новий') return 'new';

  // Fallback — визначаємо по Відмітка + Автомобіль
  var mark = String(row[COL.MARK] || '').toLowerCase();
  var vehicle = row[COL.VEHICLE];

  if (mark.indexOf('архів') !== -1 || mark === 'archived') return 'archived';
  if (mark.indexOf('маршрут') !== -1 || mark === 'route') return 'route';
  if (mark.indexOf('оптиміз') !== -1) return 'optimize';
  if (vehicle) return 'work';

  return 'new';
}

// Визначення статусу водія (для маршрутних аркушів)
function resolveDriverStatus(row, bgColors) {
  var markValue = String(row[COL.MARK] || '').toLowerCase().trim();

  if (markValue === 'completed' || markValue === 'готово') return 'completed';
  if (markValue === 'in-progress' || markValue === 'в процесі') return 'in-progress';
  if (markValue === 'cancelled' || markValue === 'відмова' || markValue === 'скасовано') return 'cancelled';
  if (markValue === 'archived' || markValue === 'архів') return 'archived';

  // По кольору рядка
  if (bgColors && bgColors[0]) {
    var color = bgColors[0].toLowerCase();
    if (color === '#00ff00' || color === '#b6d7a8' || color === '#93c47d') return 'completed';
    if (color === '#6fa8dc' || color === '#a4c2f4' || color === '#3d85c6') return 'in-progress';
    if (color === '#e06666' || color === '#ea9999' || color === '#cc0000') return 'cancelled';
  }

  return 'pending';
}

// Генерація ІД
function generateId(initials, sheet) {
  var prefix = initials.toUpperCase().substring(0, 3);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + '-001';

  var ids = sheet.getRange(2, COL.ID + 1, lastRow - 1, 1).getValues();
  var maxNum = 0;

  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0] || '');
    if (id.indexOf(prefix + '-') === 0) {
      var num = parseInt(id.split('-')[1]) || 0;
      if (num > maxNum) maxNum = num;
    }
  }

  var next = String(maxNum + 1);
  while (next.length < 3) next = '0' + next;
  return prefix + '-' + next;
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
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd');
  } catch (e) {}
  return '';
}

// Перевірка чи запис свіжий (за останні N годин)
function isRecent(dateValue, hours) {
  if (!dateValue) return true;
  try {
    var d = new Date(dateValue);
    var now = new Date();
    return (now.getTime() - d.getTime()) < (hours * 3600000);
  } catch (e) {
    return false;
  }
}

// Геокодування через Google Maps
function geocodeAddress(address) {
  try {
    var url = 'https://maps.googleapis.com/maps/api/geocode/json'
      + '?address=' + encodeURIComponent(address)
      + '&key=' + API_KEY + '&language=uk';
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());
    if (json.status === 'OK' && json.results && json.results.length > 0) {
      var loc = json.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
    return null;
  } catch (e) {
    return null;
  }
}

// Directions API оптимізація
function optimizeWithDirectionsAPI(passengers, startCoords, endCoords) {
  try {
    var allCoords = [];
    for (var i = 0; i < passengers.length; i++) {
      allCoords.push(passengers[i].coords.lat + ',' + passengers[i].coords.lng);
    }
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
      + '?origin=' + encodeURIComponent(origin)
      + '&destination=' + encodeURIComponent(destination)
      + '&key=' + API_KEY + '&language=uk';

    if (waypoints.length > 0) {
      url += '&waypoints=optimize:true|' + waypoints.join('|');
    }

    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());
    if (json.status !== 'OK') return null;

    var waypointOrder = json.routes[0].waypoint_order;
    if (endCoords) return waypointOrder;

    var result = waypointOrder.slice();
    result.push(passengers.length - 1);
    return result;
  } catch (e) {
    return null;
  }
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

// Генерація посилань на Google Maps
function generateMapLinks(orderedPassengers, startCoords, endCoords) {
  var links = [];
  if (orderedPassengers.length === 0) return links;

  var chunkStart = 0;
  while (chunkStart < orderedPassengers.length) {
    var chunkEnd = Math.min(chunkStart + MAX_POINTS_PER_MAP - 1, orderedPassengers.length);
    var chunkItems = orderedPassengers.slice(chunkStart, chunkEnd);

    var origin = chunkStart === 0
      ? startCoords.name
      : orderedPassengers[chunkStart - 1].cleanAddress;

    var destination, waypointItems;
    if (chunkEnd >= orderedPassengers.length && endCoords) {
      destination = endCoords.name;
      waypointItems = chunkItems;
    } else {
      destination = chunkItems[chunkItems.length - 1].cleanAddress;
      waypointItems = chunkItems.slice(0, chunkItems.length - 1);
    }

    var url = 'https://www.google.com/maps/dir/' + encodeURIComponent(origin);
    for (var w = 0; w < waypointItems.length; w++) {
      url += '/' + encodeURIComponent(waypointItems[w].cleanAddress);
    }
    url += '/' + encodeURIComponent(destination);

    links.push({ url: url, from: chunkStart + 1, to: chunkEnd, total: orderedPassengers.length });
    chunkStart = chunkEnd;
  }
  return links;
}

// Haversine формула
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
// МЕНЮ В GOOGLE SHEETS
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CRM Пасажири')
    .addItem('Всі пасажири', 'testGetAll')
    .addItem('Структура таблиці', 'testStructure')
    .addItem('Тест архів зв\'язок', 'testArchiveConnection')
    .addToUi();
}

// ============================================
// ТЕСТИ
// ============================================
function testGetAll() {
  var result = getAllPassengers();
  Logger.log('Всього: ' + result.counts.total);
  Logger.log('UA→EU: ' + result.counts.uaEu);
  Logger.log('EU→UA: ' + result.counts.euUa);

  for (var i = 0; i < Math.min(5, result.passengers.length); i++) {
    var p = result.passengers[i];
    Logger.log('[' + p.sheet + ' #' + p.rowNum + '] ' + p.name + ' | ' + p.phone + ' | ' + p.status + ' | ArcID: ' + p.archiveId);
  }
}

function testStructure() {
  var result = getStructure();
  for (var i = 0; i < result.sheets.length; i++) {
    var s = result.sheets[i];
    Logger.log('[' + s.sheet + '] ' + s.rows + ' рядків, ' + s.cols + ' колонок');
    Logger.log('  Колонки: ' + s.headers.join(' | '));
  }
}

function testArchiveConnection() {
  Logger.log('URL: ' + ARCHIVE_API_URL);
  var result = sendToArchive({ action: 'getStats' });
  Logger.log(result.success ? 'Архів: OK' : 'Архів: ПОМИЛКА — ' + result.error);
}
