// ============================================
// 🚗 ЮРА ТРАНСПОРТЕЙШН — МАРШРУТ ПОСИЛКИ
// Єдиний скрипт: Yura Drive (водії) + Package CRM
// Таблиця: "Маршрут Посилки"
// ID: 1Pd3nv3fbwZ_0YSzdG4cda-q52BQT57E0hDe7eQej6z8
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Маршрут Посилки" → Розширення → Apps Script
// 2. Видали весь старий код (BOTI DRIVER) і встав цей файл
// 3. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Скопіюй URL деплоя
// 5. Встав URL в HTML файл як ROUTE_API_URL
//
// Цей скрипт обслуговує ОБА додатки:
// ✅ Yura Drive (водії) — doGet: отримання посилок, doPost: оновлення статусу
// ✅ Package CRM — doPost: управління маршрутами, копіювання, статистика
// ============================================

// ============================================
// КОНФІГУРАЦІЯ
// ============================================
var CONFIG = {
  SPREADSHEET_ID: '1Pd3nv3fbwZ_0YSzdG4cda-q52BQT57E0hDe7eQej6z8',
  LOG_SHEET: 'Маршрути водіїв',
  MAILING_SHEET: 'Провірка розсилки',
  ROUTES: ['Братислава марш.', 'Нітра марш.', 'Словаччина марш.', 'Кошице+прешов марш.'],

  // Кольори статусів (для водіїв)
  COLORS: {
    'pending': '#ffc107',
    'in-progress': '#2196F3',
    'completed': '#4CAF50',
    'cancelled': '#dc3545'
  },
  BACKGROUNDS: {
    'pending': '#fffbf0',
    'in-progress': '#e3f2fd',
    'completed': '#e8f5e9',
    'cancelled': '#ffebee'
  }
};

// Службові аркуші — НЕ показуємо як маршрути
var EXCLUDE_SHEETS = ['Маршрути водіїв', 'Провірка розсилки', 'Логи'];

// ============================================
// КОЛОНИ — порядок стовпців в аркушах маршрутів
// A:ВО B:Номер№ C:ТТН D:Вага E:Адреса F:Напрямок
// G:Телефон H:Сума I:Статус оплати J:Оплата
// K:Тел.реєстратора L:Примітка M:Статус посилки N:ІД O:ПіБ
// P:Дата оформлення Q:Таймінг R:Примітка смс S:Дата отримання
// T:Фото
// ============================================
var COL = {
  VO: 0,
  NUMBER: 1,
  TTN: 2,
  WEIGHT: 3,
  ADDRESS: 4,
  DIRECTION: 5,
  PHONE: 6,
  AMOUNT: 7,
  PAYMENT_STATUS: 8,
  PAYMENT: 9,
  REGISTRAR_PHONE: 10,
  NOTE: 11,
  STATUS: 12,
  ID: 13,
  NAME: 14,
  CREATED_AT: 15,
  TIMING: 16,
  SMS_NOTE: 17,
  RECEIVE_DATE: 18,
  PHOTO: 19
};
var TOTAL_COLS = 20;

// Маппінг полів CRM → індексів колонок
var FIELD_MAP = {
  vo: COL.VO,
  number: COL.NUMBER,
  ttn: COL.TTN,
  weight: COL.WEIGHT,
  address: COL.ADDRESS,
  direction: COL.DIRECTION,
  phone: COL.PHONE,
  amount: COL.AMOUNT,
  payStatus: COL.PAYMENT_STATUS,
  payment: COL.PAYMENT,
  phoneReg: COL.REGISTRAR_PHONE,
  note: COL.NOTE,
  parcelStatus: COL.STATUS,
  id: COL.ID,
  name: COL.NAME,
  dateReg: COL.CREATED_AT,
  timing: COL.TIMING,
  smsNote: COL.SMS_NOTE,
  dateReceive: COL.RECEIVE_DATE,
  photo: COL.PHOTO
};

// ============================================
// doGet — YURA DRIVE (водії отримують посилки)
// ============================================
function doGet(e) {
  try {
    if (!e || !e.parameter) {
      return sendJSON({ error: 'Немає параметрів' });
    }
    var action = e.parameter.action || 'getDeliveries';
    var sheet = e.parameter.sheet || 'Братислава марш.';

    debugLog('GET: action=' + action + ', sheet=' + sheet);

    if (action === 'getDeliveries') {
      return getDeliveries(sheet);
    } else {
      return sendJSON({ error: 'Невідома дія: ' + action });
    }
  } catch (error) {
    debugLog('doGet Error: ' + error.message);
    return sendJSON({ error: error.message });
  }
}

// ============================================
// doPost — РОЗУМНИЙ РОУТИНГ
// Якщо є data.action → CRM запит
// Якщо немає action → Yura Drive (водій оновлює статус)
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // CRM запити (мають поле action)
    if (data.action) {
      return handleCrmAction(data);
    }

    // Yura Drive запити (без action — legacy формат водія)
    logStatusChange(data);
    return sendJSON({ success: true, message: 'Статус записано' });

  } catch (error) {
    debugLog('doPost Error: ' + error.message);
    return sendJSON({ error: error.message });
  }
}

// ============================================
// CRM РОУТЕР — обробка дій від Package CRM
// ============================================
function handleCrmAction(data) {
  var action = data.action;
  var payload = data.payload || {};
  var response;

  switch (action) {
    // --- Маршрути ---
    case 'getAvailableRoutes':
      response = getAvailableRoutes();
      break;

    case 'getRoutePassengers':
      response = getRoutePackages(payload);
      break;

    case 'checkRouteSheets':
      response = checkRouteSheets(payload);
      break;

    case 'copyToRoute':
      response = copyToRoute(payload);
      break;

    case 'createRouteSheet':
      response = createRouteSheet(payload);
      break;

    case 'deleteRouteSheet':
      response = deleteRouteSheet(payload);
      break;

    // --- Розсилка ---
    case 'getMailingStatus':
      response = getMailingStatus();
      break;

    case 'addMailingRecord':
      response = addMailingRecord(payload);
      break;

    // --- Невідома дія ---
    default:
      response = { success: false, error: 'Невідома CRM дія: ' + action };
  }

  return sendJSON(response);
}

// ============================================
// === YURA DRIVE ФУНКЦІЇ (для водіїв) ===
// ============================================

// Отримати посилки для маршруту (doGet)
function getDeliveries(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return sendJSON({ error: 'Лист не знайдено: ' + sheetName });
    }

    var data = sheet.getDataRange().getValues();
    var deliveries = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Пропускаємо рядки без ВО або номера
      if (!row[COL.VO]) continue;
      var internalNumber = row[COL.NUMBER] ? row[COL.NUMBER].toString().trim() : '';
      if (!internalNumber) continue;

      deliveries.push({
        internalNumber: internalNumber,
        address: (row[COL.ADDRESS] || '').toString().trim(),
        phone: (row[COL.PHONE] || '').toString().trim(),
        name: row[COL.NAME] || '',
        ttn: row[COL.TTN] || '',
        weight: row[COL.WEIGHT] || '',
        direction: row[COL.DIRECTION] || '',
        price: (row[COL.AMOUNT] || '').toString().trim(),
        paymentStatus: row[COL.PAYMENT_STATUS] || '',
        payment: row[COL.PAYMENT] || '',
        registrarPhone: row[COL.REGISTRAR_PHONE] || '',
        note: row[COL.NOTE] || '',
        status: row[COL.STATUS] || 'pending',
        id: row[COL.ID] || '',
        createdAt: row[COL.CREATED_AT] || '',
        timing: row[COL.TIMING] || '',
        smsNote: row[COL.SMS_NOTE] || '',
        receiveDate: row[COL.RECEIVE_DATE] || '',
        photo: row[COL.PHOTO] || '',
        coords: { lat: 48.1486, lng: 17.1077 }
      });
    }

    return sendJSON({
      success: true,
      count: deliveries.length,
      deliveries: deliveries
    });
  } catch (error) {
    return sendJSON({ error: 'Помилка: ' + error.message });
  }
}

// Записати зміну статусу від водія (doPost legacy)
function logStatusChange(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // --- ЛОГУЄМО ---
    var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
    if (!logSheet) {
      throw new Error('Лист логування не знайдено');
    }

    logSheet.appendRow([
      data.date,
      data.time,
      data.driverId,
      data.routeName,
      data.deliveryNumber,
      data.address,
      data.status,
      data.cancelReason || '',
      data.phone,
      data.price
    ]);

    debugLog('Логовано: ' + data.deliveryNumber + ' -> ' + data.status);

    // --- ОНОВЛЮЄМО СТАТУС В АРКУШІ ---
    var deliverySheet = ss.getSheetByName(data.routeName);
    if (!deliverySheet) {
      throw new Error('Маршрут не знайдено: ' + data.routeName);
    }

    var allData = deliverySheet.getDataRange().getValues();
    var rowsToUpdate = [];

    for (var i = 1; i < allData.length; i++) {
      var deliveryNum = allData[i][COL.NUMBER] ? allData[i][COL.NUMBER].toString().trim() : '';
      if (deliveryNum === data.deliveryNumber) {
        rowsToUpdate.push(i + 1);
      }
    }

    if (rowsToUpdate.length === 0) {
      throw new Error('Посилка не знайдена: ' + data.deliveryNumber);
    }

    // Оновлюємо кожен знайдений рядок
    for (var r = 0; r < rowsToUpdate.length; r++) {
      var rowNum = rowsToUpdate[r];
      var statusCell = deliverySheet.getRange(rowNum, COL.STATUS + 1);
      statusCell.setValue(data.status);

      // Фарбуємо рядок
      var rowColor = CONFIG.BACKGROUNDS[data.status] || '#ffffff';
      var borderColor = CONFIG.COLORS[data.status] || '#000000';
      var rangeToColor = deliverySheet.getRange(rowNum, 1, 1, TOTAL_COLS);
      rangeToColor.setBackground(rowColor);
      rangeToColor.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      statusCell.setFontColor(borderColor);
      statusCell.setFontWeight('bold');
    }

    return true;
  } catch (error) {
    debugLog('logStatusChange Error: ' + error.message);
    throw error;
  }
}

// ============================================
// === CRM ФУНКЦІЇ (для Package CRM) ===
// ============================================

// --- Список доступних маршрутів ---
function getAvailableRoutes() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var routes = [];

    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (EXCLUDE_SHEETS.indexOf(name) !== -1) continue;

      var lastRow = sheets[i].getLastRow();
      var count = lastRow > 1 ? lastRow - 1 : 0;

      routes.push({
        name: name,
        type: 'package',
        count: count,
        sheetId: sheets[i].getSheetId()
      });
    }

    debugLog('getAvailableRoutes: ' + routes.length + ' маршрутів');

    return {
      success: true,
      routes: routes,
      count: routes.length
    };
  } catch (error) {
    debugLog('getAvailableRoutes Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// --- Отримати посилки маршруту (для CRM) ---
function getRoutePackages(payload) {
  try {
    var vehicleName = payload.vehicleName;
    var sheetName = payload.sheetName || vehicleName;

    if (!sheetName) {
      return { success: false, error: 'Не вказано аркуш маршруту' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        success: true,
        packages: [],
        count: 0,
        sheetName: sheetName,
        vehicleName: vehicleName || '',
        stats: { total: 0, pending: 0, inProgress: 0, completed: 0, cancelled: 0, archived: 0 }
      };
    }

    var colsToRead = Math.min(TOTAL_COLS, sheet.getLastColumn());
    var dataRange = sheet.getRange(2, 1, lastRow - 1, colsToRead);
    var data = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();

    var packages = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // Пропускаємо повністю порожні рядки
      if (!row[COL.NAME] && !row[COL.PHONE] && !row[COL.NUMBER] && !row[COL.VO]) continue;

      // Визначаємо статус водія
      var driverStatus = 'pending';
      var statusVal = String(row[COL.STATUS] || '').toLowerCase().trim();

      if (statusVal === 'completed' || statusVal === 'готово' || statusVal === 'доставлено') {
        driverStatus = 'completed';
      } else if (statusVal === 'in-progress' || statusVal === 'в процесі' || statusVal === 'доставка') {
        driverStatus = 'in-progress';
      } else if (statusVal === 'cancelled' || statusVal === 'відмова' || statusVal === 'скасовано') {
        driverStatus = 'cancelled';
      } else if (statusVal === 'archived' || statusVal === 'архів') {
        driverStatus = 'archived';
      }

      // Fallback по кольору рядка
      if (driverStatus === 'pending' && backgrounds[i]) {
        var rowColor = backgrounds[i][0];
        if (rowColor === '#00ff00' || rowColor === '#b6d7a8' || rowColor === '#93c47d') {
          driverStatus = 'completed';
        } else if (rowColor === '#6fa8dc' || rowColor === '#a4c2f4' || rowColor === '#3d85c6') {
          driverStatus = 'in-progress';
        } else if (rowColor === '#e06666' || rowColor === '#ea9999' || rowColor === '#cc0000') {
          driverStatus = 'cancelled';
        }
      }

      packages.push({
        rowNum: i + 2,
        vo: String(row[COL.VO] || ''),
        number: String(row[COL.NUMBER] || ''),
        ttn: String(row[COL.TTN] || ''),
        weight: String(row[COL.WEIGHT] || ''),
        address: String(row[COL.ADDRESS] || ''),
        direction: String(row[COL.DIRECTION] || ''),
        phone: String(row[COL.PHONE] || ''),
        amount: String(row[COL.AMOUNT] || ''),
        payStatus: String(row[COL.PAYMENT_STATUS] || ''),
        payment: String(row[COL.PAYMENT] || ''),
        phoneReg: String(row[COL.REGISTRAR_PHONE] || ''),
        note: String(row[COL.NOTE] || ''),
        parcelStatus: String(row[COL.STATUS] || ''),
        id: String(row[COL.ID] || ''),
        name: String(row[COL.NAME] || ''),
        dateReg: formatDate(row[COL.CREATED_AT]),
        timing: String(row[COL.TIMING] || ''),
        smsNote: String(row[COL.SMS_NOTE] || ''),
        dateReceive: formatDate(row[COL.RECEIVE_DATE]),
        photo: String(row[COL.PHOTO] || ''),
        driverStatus: driverStatus,
        rowColor: backgrounds[i] ? backgrounds[i][0] : '#ffffff'
      });
    }

    var stats = {
      total: packages.length,
      pending: packages.filter(function(p) { return p.driverStatus === 'pending'; }).length,
      inProgress: packages.filter(function(p) { return p.driverStatus === 'in-progress'; }).length,
      completed: packages.filter(function(p) { return p.driverStatus === 'completed'; }).length,
      cancelled: packages.filter(function(p) { return p.driverStatus === 'cancelled'; }).length,
      archived: packages.filter(function(p) { return p.driverStatus === 'archived'; }).length
    };

    debugLog('getRoutePackages: ' + sheetName + ' → ' + packages.length + ' посилок');

    return {
      success: true,
      packages: packages,
      count: packages.length,
      sheetName: sheetName,
      vehicleName: vehicleName || '',
      stats: stats
    };
  } catch (error) {
    debugLog('getRoutePackages Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// --- Перевірка існуючих записів ---
function checkRouteSheets(payload) {
  try {
    var vehicleNames = payload.vehicleNames || [];
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var existing = [];

    for (var i = 0; i < vehicleNames.length; i++) {
      var vName = vehicleNames[i];
      var sheet = findRouteSheet(ss, vName);
      if (!sheet) continue;

      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        existing.push({
          vehicle: vName,
          sheet: sheet.getName(),
          count: lastRow - 1
        });
      }
    }

    return { success: true, existing: existing };
  } catch (error) {
    debugLog('checkRouteSheets Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// --- Копіювання посилок в маршрут ---
function copyToRoute(payload) {
  try {
    var packagesByVehicle = payload.packagesByVehicle;
    var conflictAction = payload.conflictAction || 'add';

    if (!packagesByVehicle || Object.keys(packagesByVehicle).length === 0) {
      return { success: false, error: 'Немає посилок для копіювання' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var totalCopied = 0;
    var totalArchived = 0;
    var totalCleared = 0;
    var results = [];

    for (var vehicleName in packagesByVehicle) {
      if (!packagesByVehicle.hasOwnProperty(vehicleName)) continue;

      var pkgs = packagesByVehicle[vehicleName];
      var sheet = findRouteSheet(ss, vehicleName);

      // Якщо аркуш не існує — створюємо
      if (!sheet) {
        sheet = createNewRouteSheet(ss, vehicleName);
        if (!sheet) {
          results.push({ vehicle: vehicleName, error: 'Не вдалось створити аркуш' });
          continue;
        }
      }

      // Обробка конфлікту (якщо є існуючі дані)
      var lastRow = sheet.getLastRow();
      if (lastRow > 1 && conflictAction !== 'add') {
        if (conflictAction === 'clear') {
          totalCleared += lastRow - 1;
          sheet.deleteRows(2, lastRow - 1);
        } else if (conflictAction === 'archive') {
          // Помічаємо старі записи як "Архів"
          var oldData = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
          for (var a = 0; a < oldData.length; a++) {
            oldData[a][COL.STATUS] = 'archived';
          }
          sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).setValues(oldData);
          totalArchived += lastRow - 1;
        }
      }

      // Копіюємо нові посилки
      for (var p = 0; p < pkgs.length; p++) {
        var pkg = pkgs[p];
        var newRow = new Array(TOTAL_COLS);
        for (var c = 0; c < TOTAL_COLS; c++) newRow[c] = '';

        // Маппінг полів з CRM формату
        newRow[COL.VO] = pkg.vo || '';
        newRow[COL.NUMBER] = pkg.number || '';
        newRow[COL.TTN] = pkg.ttn || '';
        newRow[COL.WEIGHT] = pkg.weight || '';
        newRow[COL.ADDRESS] = pkg.address || '';
        newRow[COL.DIRECTION] = pkg.direction || pkg.directionRaw || '';
        newRow[COL.PHONE] = pkg.phone || '';
        newRow[COL.AMOUNT] = pkg.amount || '';
        newRow[COL.PAYMENT_STATUS] = pkg.payStatus || '';
        newRow[COL.PAYMENT] = pkg.payment || '';
        newRow[COL.REGISTRAR_PHONE] = pkg.phoneReg || '';
        newRow[COL.NOTE] = pkg.note || '';
        newRow[COL.STATUS] = pkg.parcelStatus || 'pending';
        newRow[COL.ID] = pkg.id || '';
        newRow[COL.NAME] = pkg.name || '';
        newRow[COL.CREATED_AT] = pkg.dateReg || '';
        newRow[COL.TIMING] = pkg.timing || '';
        newRow[COL.SMS_NOTE] = pkg.smsNote || '';
        newRow[COL.RECEIVE_DATE] = pkg.dateReceive || '';
        newRow[COL.PHOTO] = pkg.photo || '';

        sheet.appendRow(newRow);
        totalCopied++;
      }

      results.push({ vehicle: vehicleName, sheet: sheet.getName(), copied: pkgs.length });
      debugLog(vehicleName + ': ' + pkgs.length + ' посилок скопійовано');
    }

    return {
      success: true,
      copied: totalCopied,
      archived: totalArchived,
      cleared: totalCleared,
      details: results
    };
  } catch (error) {
    debugLog('copyToRoute Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// --- Створити аркуш маршруту ---
function createRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) {
      return { success: false, error: 'Не вказано назву авто' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var existingSheet = ss.getSheetByName(vehicleName);

    if (existingSheet) {
      return {
        success: true,
        sheetName: vehicleName,
        vehicleName: vehicleName,
        existed: true
      };
    }

    var sheet = createNewRouteSheet(ss, vehicleName);
    if (!sheet) {
      return { success: false, error: 'Не вдалось створити аркуш' };
    }

    debugLog('Створено аркуш: ' + vehicleName);

    return {
      success: true,
      sheetName: sheet.getName(),
      vehicleName: vehicleName
    };
  } catch (error) {
    debugLog('createRouteSheet Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// --- Видалити аркуш маршруту ---
function deleteRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) {
      return { success: false, error: 'Не вказано назву авто' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = findRouteSheet(ss, vehicleName);

    if (!sheet) {
      return {
        success: true,
        message: 'Аркуш не існує',
        sheetName: vehicleName,
        deleted: false
      };
    }

    var lastRow = sheet.getLastRow();
    var hasData = lastRow > 1;

    if (hasData && !payload.force) {
      return {
        success: false,
        error: 'Аркуш містить ' + (lastRow - 1) + ' записів. Використайте force: true.',
        sheetName: sheet.getName(),
        recordsCount: lastRow - 1
      };
    }

    var sheetName = sheet.getName();
    ss.deleteSheet(sheet);

    debugLog('Видалено аркуш: ' + sheetName);

    return {
      success: true,
      message: 'Аркуш видалено',
      sheetName: sheetName,
      deleted: true
    };
  } catch (error) {
    debugLog('deleteRouteSheet Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// --- Статус розсилки ---
function getMailingStatus() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.MAILING_SHEET);

    if (!sheet) {
      return { success: true, mailingIds: [], mailingData: [], count: 0 };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, mailingIds: [], mailingData: [], count: 0 };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var mailingData = [];

    for (var i = 0; i < data.length; i++) {
      var date = data[i][0];
      var id = data[i][1];

      if (!id || String(id).indexOf('dd.mm.yyyy') !== -1) continue;

      mailingData.push({
        date: date ? formatDate(date) : '',
        id: String(id).trim()
      });
    }

    var mailingIds = mailingData.map(function(m) { return m.id; });

    debugLog('getMailingStatus: ' + mailingIds.length + ' записів');

    return {
      success: true,
      mailingData: mailingData,
      mailingIds: mailingIds,
      count: mailingIds.length
    };
  } catch (error) {
    debugLog('getMailingStatus Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// --- Додати запис розсилки ---
function addMailingRecord(payload) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.MAILING_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.MAILING_SHEET);
      sheet.getRange(1, 1, 1, 2).setValues([['Дата виїзду', 'ІД']]);
    }

    var records = payload.records || [];
    var userName = payload.userName || 'Невідомий';

    if (records.length === 0) {
      return { success: false, error: 'Немає записів для додавання' };
    }

    var today = formatMailingDate(new Date());
    var rowsToAdd = records.map(function(record) {
      return [record.date || today, userName];
    });

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 2).setValues(rowsToAdd);

    debugLog('addMailingRecord: ' + rowsToAdd.length + ' записів від ' + userName);

    return {
      success: true,
      added: rowsToAdd.length,
      userName: userName
    };
  } catch (error) {
    debugLog('addMailingRecord Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================
// ДОПОМІЖНІ ФУНКЦІЇ
// ============================================

// Знайти аркуш маршруту по назві авто
function findRouteSheet(ss, vehicleName) {
  // Спочатку точне співпадіння
  var sheet = ss.getSheetByName(vehicleName);
  if (sheet) return sheet;

  // Пошук по включенню
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.toLowerCase().indexOf(vehicleName.toLowerCase()) !== -1) {
      return sheets[i];
    }
  }

  return null;
}

// Створити новий аркуш маршруту
function createNewRouteSheet(ss, name) {
  try {
    // Пробуємо скопіювати з першого маршрутного аркуша як шаблон
    var templateSheet = null;
    for (var i = 0; i < CONFIG.ROUTES.length; i++) {
      templateSheet = ss.getSheetByName(CONFIG.ROUTES[i]);
      if (templateSheet) break;
    }

    var newSheet;
    if (templateSheet) {
      newSheet = templateSheet.copyTo(ss);
      newSheet.setName(name);
      if (newSheet.getLastRow() > 1) {
        newSheet.deleteRows(2, newSheet.getLastRow() - 1);
      }
    } else {
      newSheet = ss.insertSheet(name);
      var headers = ['ВО', 'Номер№', 'Номер ТТН', 'Вага', 'Адреса Отримувача', 'Напрямок',
                     'Телефон Отримувача', 'Сума Є', 'Статус оплати', 'Оплата',
                     'Телефон Реєстратора', 'Примітка', 'Статус посилки', 'ІД', 'ПіБ',
                     'Дата оформлення', 'Таймінг', 'Примітка смс', 'Дата отримання', 'Фото'];
      newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      newSheet.setFrozenRows(1);
    }

    return newSheet;
  } catch (e) {
    debugLog('createNewRouteSheet Error: ' + e.toString());
    return null;
  }
}

// Форматування дати
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
    var parts = str.split('.');
    return parts[2] + '-' + parts[1].padStart(2, '0') + '-' + parts[0].padStart(2, '0');
  }
  return str;
}

function formatMailingDate(date) {
  if (!date) return '';
  if (date instanceof Date) {
    var d = date.getDate().toString().padStart(2, '0');
    var m = (date.getMonth() + 1).toString().padStart(2, '0');
    var y = date.getFullYear();
    return d + '.' + m + '.' + y;
  }
  return String(date);
}

// JSON відповідь
function sendJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Логування
function debugLog(msg) {
  Logger.log('[Route Packages API] ' + msg);
}

// ============================================
// ТЕСТИ
// ============================================
function testGetDeliveries() {
  var result = getDeliveries('Братислава марш.');
  Logger.log('Deliveries: ' + JSON.stringify(result).substring(0, 500));
}

function testGetAvailableRoutes() {
  var result = getAvailableRoutes();
  Logger.log('Routes: ' + JSON.stringify(result));
}

function testGetRoutePackages() {
  var result = getRoutePackages({ sheetName: 'Братислава марш.' });
  Logger.log('Packages: ' + result.count + ' | Stats: ' + JSON.stringify(result.stats));
}

function testLogStatus() {
  var testData = {
    date: new Date().toLocaleDateString('uk-UA'),
    time: new Date().toLocaleTimeString('uk-UA'),
    driverId: 'Водій',
    routeName: 'Братислава марш.',
    deliveryNumber: '188',
    address: 'Test',
    status: 'in-progress',
    cancelReason: '',
    phone: '+421951497677',
    price: '100'
  };

  try {
    logStatusChange(testData);
    Logger.log('Test OK');
  } catch (error) {
    Logger.log('Test ERROR: ' + error.message);
  }
}

// ============================================
// МЕНЮ
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Route Packages API')
    .addItem('📋 Тест: Посилки Братислава', 'testGetDeliveries')
    .addItem('🚐 Тест: Список маршрутів', 'testGetAvailableRoutes')
    .addItem('📊 Тест: Пакети маршруту', 'testGetRoutePackages')
    .addItem('✅ Тест: Оновлення статусу', 'testLogStatus')
    .addToUi();
}
