// ============================================
// 📦 ЮРА ТРАНСПОРТЕЙШН - ROUTE API (ПОСИЛКИ)
// Google Apps Script для маршрутів посилок
// Працює і з CRM, і з BOTI Driver
// ============================================

const CONFIG = {
  COMPANY: 'ЮРА ТРАНСПОРТЕЙШН — ПОСИЛКИ',

  SPREADSHEET_ID: '1Pd3nv3fbwZ_0YSzdG4cda-q52BQT57E0hDe7eQej6z8',

  LOG_SHEET: 'Маршрути водіїв',
  MAILING_SHEET: 'Розсилка',

  // Кольори статусів
  COLORS: {
    'pending':     '#ffc107',
    'in-progress': '#2196F3',
    'completed':   '#4CAF50',
    'cancelled':   '#dc3545'
  },
  BACKGROUNDS: {
    'pending':     '#fffbf0',
    'in-progress': '#e3f2fd',
    'completed':   '#e8f5e9',
    'cancelled':   '#ffebee'
  }
};

// ============================================
// КОЛОНИ МАРШРУТНОГО АРКУША (такі ж як в CRM)
// ============================================
const COLUMNS = {
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
  PHOTO: 19,
  PARCEL_STATUS: 20
};

const HEADERS = [
  'ВО', 'Номер', 'ТТН', 'Вага', 'Адреса', 'Напрямок',
  'Телефон', 'Сума', 'Статус оплати', 'Оплата',
  'Тел. реєстратора', 'Примітка', 'Статус', 'ID',
  'Ім\'я', 'Дата оформлення', 'Таймінг', 'СМС примітка',
  'Дата отримання', 'Фото', 'Статус посилки'
];

// ============================================
// doGet — ВОДІЇ (BOTI Driver) ОТРИМУЮТЬ ПОСИЛКИ
// ============================================
function doGet(e) {
  try {
    if (!e || !e.parameter) {
      return sendJSON({ error: 'Немає параметрів' });
    }

    const action = e.parameter.action || 'getDeliveries';
    const sheet = e.parameter.sheet || '';

    Logger.log('✅ GET: action=' + action + ', sheet=' + sheet);

    if (action === 'getDeliveries') {
      if (!sheet) return sendJSON({ error: 'Не вказано маршрут (sheet)' });
      return getDeliveries(sheet);
    }

    return sendJSON({ error: 'Невідома дія: ' + action });
  } catch (error) {
    Logger.log('❌ doGet Error: ' + error.message);
    return sendJSON({ error: error.message });
  }
}

// ============================================
// doPost — CRM + ВОДІЇ
// ============================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'updateStatus';

    Logger.log('📥 POST action: ' + action);

    switch (action) {
      // ---- CRM ACTIONS ----
      case 'copyToRoute':
        return handleCopyToRoute(data.payload);

      case 'checkRouteSheets':
        return handleCheckRouteSheets(data.payload);

      case 'getRoutePassengers':
        return handleGetRoutePassengers(data.payload);

      case 'getAvailableRoutes':
        return handleGetAvailableRoutes();

      case 'deleteRouteSheet':
        return handleDeleteRouteSheet(data.payload);

      case 'addMailingRecord':
        return handleAddMailingRecord(data.payload);

      case 'getMailingStatus':
        return handleGetMailingStatus();

      // ---- DRIVER ACTIONS ----
      case 'updateStatus':
        return handleDriverStatusUpdate(data);

      default:
        return sendJSON({ error: 'Невідома дія: ' + action });
    }
  } catch (error) {
    Logger.log('❌ doPost Error: ' + error.message);
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 📋 GET DELIVERIES (для водіїв)
// ============================================
function getDeliveries(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return sendJSON({ error: 'Лист не знайдено: ' + sheetName });
    }

    const data = sheet.getDataRange().getValues();
    const deliveries = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (!row[COLUMNS.VO] && !row[COLUMNS.NUMBER]) continue;

      const internalNumber = row[COLUMNS.NUMBER] ? row[COLUMNS.NUMBER].toString().trim() : '';
      if (!internalNumber) continue;

      deliveries.push({
        internalNumber: internalNumber,
        vo: (row[COLUMNS.VO] || '').toString().trim(),
        ttn: (row[COLUMNS.TTN] || '').toString().trim(),
        weight: (row[COLUMNS.WEIGHT] || '').toString().trim(),
        address: (row[COLUMNS.ADDRESS] || '').toString().trim(),
        direction: (row[COLUMNS.DIRECTION] || '').toString().trim(),
        phone: (row[COLUMNS.PHONE] || '').toString().trim(),
        price: (row[COLUMNS.AMOUNT] || '').toString().trim(),
        paymentStatus: (row[COLUMNS.PAYMENT_STATUS] || '').toString().trim(),
        payment: (row[COLUMNS.PAYMENT] || '').toString().trim(),
        registrarPhone: (row[COLUMNS.REGISTRAR_PHONE] || '').toString().trim(),
        note: (row[COLUMNS.NOTE] || '').toString().trim(),
        status: (row[COLUMNS.STATUS] || 'pending').toString().trim(),
        id: (row[COLUMNS.ID] || '').toString().trim(),
        name: (row[COLUMNS.NAME] || '').toString().trim(),
        createdAt: (row[COLUMNS.CREATED_AT] || '').toString().trim(),
        timing: (row[COLUMNS.TIMING] || '').toString().trim(),
        smsNote: (row[COLUMNS.SMS_NOTE] || '').toString().trim(),
        receiveDate: (row[COLUMNS.RECEIVE_DATE] || '').toString().trim(),
        photo: (row[COLUMNS.PHOTO] || '').toString().trim(),
        parcelStatus: (row[COLUMNS.PARCEL_STATUS] || '').toString().trim(),
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

// ============================================
// 📤 COPY TO ROUTE (CRM → маршрутний аркуш)
// ============================================
function handleCopyToRoute(payload) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const packagesByVehicle = payload.packagesByVehicle;
    const conflictAction = payload.conflictAction || 'add';

    let totalCopied = 0;
    let totalArchived = 0;
    let totalCleared = 0;

    for (const vehicleName in packagesByVehicle) {
      const packages = packagesByVehicle[vehicleName];
      if (!packages || !packages.length) continue;

      // Знаходимо або створюємо аркуш маршруту
      const sheetName = vehicleName + ' марш.';
      let sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        // Записуємо заголовки
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
        sheet.getRange(1, 1, 1, HEADERS.length)
          .setBackground('#1a1a2e')
          .setFontColor('#ffffff')
          .setFontWeight('bold');
        sheet.setFrozenRows(1);
      }

      // Обробка конфліктів
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        if (conflictAction === 'clear') {
          totalCleared += lastRow - 1;
          sheet.deleteRows(2, lastRow - 1);
        } else if (conflictAction === 'archive') {
          totalArchived += lastRow - 1;
          // Переносимо старі записи в архівний аркуш
          archiveRouteData(ss, sheetName, sheet);
        }
        // 'add' — просто дописуємо нижче
      }

      // Записуємо посилки
      const rows = packages.map(pkg => [
        pkg.vo || '',
        pkg.number || '',
        pkg.ttn || '',
        pkg.weight || '',
        pkg.address || '',
        pkg.direction || '',
        pkg.phone || '',
        pkg.amount || '',
        pkg.payStatus || '',
        pkg.payment || '',
        pkg.phoneReg || '',
        pkg.note || '',
        'pending',           // STATUS — початковий статус
        pkg.id || '',
        pkg.name || '',
        pkg.dateReg || '',
        pkg.timing || '',
        pkg.smsNote || '',
        pkg.dateReceive || '',
        pkg.photo || '',
        pkg.parcelStatus || ''
      ]);

      if (rows.length > 0) {
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);

        // Фарбуємо pending рядки
        const range = sheet.getRange(startRow, 1, rows.length, HEADERS.length);
        range.setBackground(CONFIG.BACKGROUNDS['pending']);

        totalCopied += rows.length;
      }

      // Авторозмір колонок
      try {
        sheet.autoResizeColumns(1, HEADERS.length);
      } catch(e) {}
    }

    Logger.log('✅ Copied: ' + totalCopied + ', Archived: ' + totalArchived + ', Cleared: ' + totalCleared);

    return sendJSON({
      success: true,
      copied: totalCopied,
      archived: totalArchived,
      cleared: totalCleared
    });
  } catch (error) {
    Logger.log('❌ copyToRoute Error: ' + error.message);
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 🔍 CHECK ROUTE SHEETS (чи є дані)
// ============================================
function handleCheckRouteSheets(payload) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const vehicleNames = payload.vehicleNames || [];
    const existing = [];

    for (const vName of vehicleNames) {
      const sheetName = vName + ' марш.';
      const sheet = ss.getSheetByName(sheetName);

      if (sheet && sheet.getLastRow() > 1) {
        existing.push({
          vehicle: vName,
          sheet: sheetName,
          count: sheet.getLastRow() - 1
        });
      }
    }

    return sendJSON({
      success: true,
      existing: existing
    });
  } catch (error) {
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 📖 GET ROUTE PASSENGERS (посилки з маршруту)
// ============================================
function handleGetRoutePassengers(payload) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const vehicleName = payload.vehicleName || '';
    const sheetName = payload.sheetName || (vehicleName ? vehicleName + ' марш.' : '');

    if (!sheetName) {
      return sendJSON({ error: 'Не вказано маршрут' });
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return sendJSON({ success: true, passengers: [], count: 0 });
    }

    const data = sheet.getDataRange().getValues();
    const passengers = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[COLUMNS.NUMBER] && !row[COLUMNS.VO]) continue;

      passengers.push({
        rowNum: i + 1,
        vo: row[COLUMNS.VO] || '',
        number: (row[COLUMNS.NUMBER] || '').toString().trim(),
        ttn: row[COLUMNS.TTN] || '',
        weight: row[COLUMNS.WEIGHT] || '',
        address: row[COLUMNS.ADDRESS] || '',
        direction: row[COLUMNS.DIRECTION] || '',
        phone: (row[COLUMNS.PHONE] || '').toString().trim(),
        amount: row[COLUMNS.AMOUNT] || '',
        payStatus: row[COLUMNS.PAYMENT_STATUS] || '',
        payment: row[COLUMNS.PAYMENT] || '',
        phoneReg: row[COLUMNS.REGISTRAR_PHONE] || '',
        note: row[COLUMNS.NOTE] || '',
        status: row[COLUMNS.STATUS] || 'pending',
        id: row[COLUMNS.ID] || '',
        name: row[COLUMNS.NAME] || '',
        dateReg: row[COLUMNS.CREATED_AT] || '',
        timing: row[COLUMNS.TIMING] || '',
        smsNote: row[COLUMNS.SMS_NOTE] || '',
        dateReceive: row[COLUMNS.RECEIVE_DATE] || '',
        photo: row[COLUMNS.PHOTO] || '',
        parcelStatus: row[COLUMNS.PARCEL_STATUS] || '',
        sheet: sheetName
      });
    }

    return sendJSON({
      success: true,
      passengers: passengers,
      count: passengers.length,
      sheetName: sheetName
    });
  } catch (error) {
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 📋 GET AVAILABLE ROUTES (список маршрутів)
// ============================================
function handleGetAvailableRoutes() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const routes = [];

    for (const sheet of sheets) {
      const name = sheet.getName();
      // Маршрутні аркуші мають суфікс " марш."
      if (name.endsWith(' марш.')) {
        const count = Math.max(0, sheet.getLastRow() - 1);
        routes.push({
          name: name,
          vehicle: name.replace(' марш.', ''),
          count: count
        });
      }
    }

    return sendJSON({
      success: true,
      routes: routes
    });
  } catch (error) {
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 🗑️ DELETE ROUTE SHEET
// ============================================
function handleDeleteRouteSheet(payload) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const vehicleName = payload.vehicleName || '';
    const force = payload.force || false;
    const sheetName = vehicleName + ' марш.';

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return sendJSON({ success: true, message: 'Аркуш не існує' });
    }

    const rowCount = sheet.getLastRow() - 1;
    if (rowCount > 0 && !force) {
      return sendJSON({
        success: false,
        error: 'Аркуш містить ' + rowCount + ' записів. Використайте force=true',
        count: rowCount
      });
    }

    ss.deleteSheet(sheet);
    Logger.log('🗑️ Видалено аркуш: ' + sheetName);

    return sendJSON({
      success: true,
      message: 'Аркуш видалено: ' + sheetName
    });
  } catch (error) {
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 📨 MAILING — ЗАПИСАТИ РОЗСИЛКУ
// ============================================
function handleAddMailingRecord(payload) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.MAILING_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.MAILING_SHEET);
      sheet.getRange(1, 1, 1, 6).setValues([['Дата', 'Час', 'Користувач', 'Телефон', 'Повідомлення', 'ID посилки']]);
      sheet.getRange(1, 1, 1, 6)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    const records = payload.records || [];
    const userName = payload.userName || 'CRM';
    const now = new Date();
    const date = now.toLocaleDateString('uk-UA');
    const time = now.toLocaleTimeString('uk-UA');

    const rows = records.map(r => [
      date,
      time,
      userName,
      r.phone || '',
      r.message || '',
      r.packageId || ''
    ]);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    }

    return sendJSON({
      success: true,
      count: rows.length,
      message: 'Записано ' + rows.length + ' записів розсилки'
    });
  } catch (error) {
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 📨 MAILING — ОТРИМАТИ СТАТУС РОЗСИЛКИ
// ============================================
function handleGetMailingStatus() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.MAILING_SHEET);

    if (!sheet || sheet.getLastRow() <= 1) {
      return sendJSON({ success: true, mailingIds: [] });
    }

    const data = sheet.getDataRange().getValues();
    const mailingIds = [];

    for (let i = 1; i < data.length; i++) {
      const packageId = data[i][5]; // колонка F — ID посилки
      if (packageId) {
        mailingIds.push(packageId.toString().trim());
      }
    }

    // Унікальні ID
    const uniqueIds = [...new Set(mailingIds)];

    return sendJSON({
      success: true,
      mailingIds: uniqueIds
    });
  } catch (error) {
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 🚗 DRIVER STATUS UPDATE (від водіїв)
// ============================================
function handleDriverStatusUpdate(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 1. Логуємо
    let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.LOG_SHEET);
      logSheet.getRange(1, 1, 1, 10).setValues([[
        'Дата', 'Час', 'Водій', 'Маршрут', 'Номер посилки',
        'Адреса', 'Статус', 'Причина скасування', 'Телефон', 'Сума'
      ]]);
      logSheet.getRange(1, 1, 1, 10)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    const logRow = [
      data.date || new Date().toLocaleDateString('uk-UA'),
      data.time || new Date().toLocaleTimeString('uk-UA'),
      data.driverId || '',
      data.routeName || '',
      data.deliveryNumber || '',
      data.address || '',
      data.status || '',
      data.cancelReason || '',
      data.phone || '',
      data.price || ''
    ];
    logSheet.appendRow(logRow);
    Logger.log('✅ Логовано: ' + data.deliveryNumber + ' -> ' + data.status);

    // 2. Оновлюємо статус у маршрутному аркуші
    const routeSheet = ss.getSheetByName(data.routeName);
    if (!routeSheet) {
      return sendJSON({ success: true, message: 'Логовано (маршрут не знайдено)' });
    }

    const allData = routeSheet.getDataRange().getValues();
    const rowsToUpdate = [];

    for (let i = 1; i < allData.length; i++) {
      const num = allData[i][COLUMNS.NUMBER] ? allData[i][COLUMNS.NUMBER].toString().trim() : '';
      if (num === data.deliveryNumber) {
        rowsToUpdate.push(i + 1);
      }
    }

    if (rowsToUpdate.length === 0) {
      return sendJSON({ success: true, message: 'Логовано (посилку не знайдено в маршруті)' });
    }

    // Оновлюємо кожен знайдений рядок
    for (const rowNum of rowsToUpdate) {
      const statusCell = routeSheet.getRange(rowNum, COLUMNS.STATUS + 1);
      statusCell.setValue(data.status);

      // Кольори
      const bgColor = CONFIG.BACKGROUNDS[data.status] || '#ffffff';
      const borderColor = CONFIG.COLORS[data.status] || '#000000';

      const rangeToColor = routeSheet.getRange(rowNum, 1, 1, HEADERS.length);
      rangeToColor.setBackground(bgColor);
      rangeToColor.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);

      statusCell.setFontColor(borderColor);
      statusCell.setFontWeight('bold');
    }

    Logger.log('✅ Оновлено ' + rowsToUpdate.length + ' рядків');

    return sendJSON({
      success: true,
      message: 'Статус записано',
      updatedRows: rowsToUpdate.length
    });
  } catch (error) {
    Logger.log('❌ Status update error: ' + error.message);
    return sendJSON({ error: error.message });
  }
}

// ============================================
// 📦 ARCHIVE ROUTE DATA (допоміжна)
// ============================================
function archiveRouteData(ss, sheetName, sheet) {
  try {
    const archiveName = 'Архів ' + sheetName;
    let archiveSheet = ss.getSheetByName(archiveName);

    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveName);
      archiveSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      archiveSheet.getRange(1, 1, 1, HEADERS.length)
        .setBackground('#4a4a4a')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
      const values = dataRange.getValues();

      // Додаємо в архів
      const archiveStart = archiveSheet.getLastRow() + 1;
      archiveSheet.getRange(archiveStart, 1, values.length, HEADERS.length).setValues(values);

      // Очищаємо оригінал
      sheet.deleteRows(2, lastRow - 1);
    }
  } catch (error) {
    Logger.log('❌ Archive error: ' + error.message);
  }
}

// ============================================
// ДОПОМІЖНІ ФУНКЦІЇ
// ============================================
function sendJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// ТЕСТИ
// ============================================
function testGetDeliveries() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const routeSheets = sheets.filter(s => s.getName().endsWith(' марш.'));

  if (routeSheets.length > 0) {
    const result = getDeliveries(routeSheets[0].getName());
    Logger.log(result.getContent());
  } else {
    Logger.log('Немає маршрутних аркушів');
  }
}

function testGetAvailableRoutes() {
  const result = handleGetAvailableRoutes();
  Logger.log(result.getContent());
}

function testCheckSheets() {
  const result = handleCheckRouteSheets({
    vehicleNames: ['А-Братислава', 'А-Нітра']
  });
  Logger.log(result.getContent());
}

// ============================================
// МЕНЮ В GOOGLE SHEETS
// ============================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Посилки - Маршрути')
    .addItem('📋 Маршрути', 'testGetAvailableRoutes')
    .addItem('🔍 Перевірити аркуші', 'testCheckSheets')
    .addItem('📦 Тест доставок', 'testGetDeliveries')
    .addToUi();
}
