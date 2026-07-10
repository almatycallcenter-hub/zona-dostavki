// ╔══════════════════════════════════════════════════════════╗
// ║   ТӘП-ТӘТТІ · Колл-центр · Google Apps Script           ║
// ║   Вставьте этот код в Apps Script вашей таблицы          ║
// ╚══════════════════════════════════════════════════════════╝
//
// КАК ПОДКЛЮЧИТЬ:
// 1. Создайте Google Таблицу
// 2. Расширения → Apps Script → вставьте этот код
// 3. Замените SHEET_ID ниже на ID вашей таблицы
//    (ID — часть URL: docs.google.com/spreadsheets/d/ВОТ_ЭТО/edit)
// 4. Нажмите «Развернуть» → Новое развёртывание
//    Тип: Веб-приложение | Выполнять от: Я | Доступ: Все
// 5. Скопируйте URL и вставьте в ⚙️ Настройки сайта

const SHEET_ID = 'ВСТАВЬТЕ_ID_ТАБЛИЦЫ_СЮДА';

// ── Названия листов ──────────────────────────────────────────
const SHEET_SHIFTS       = 'Смены';
const SHEET_TRANSACTIONS = 'Транзакции';

// ════════════════════════════════════════════════════════════
//  GET — получение данных (история смен)
// ════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'getShifts') {
    return jsonResponse(getShifts());
  }
  if (action === 'getTransactions') {
    const date = e.parameter.date || '';
    return jsonResponse(getTransactions(date));
  }

  if (action === 'getCarryoverPrepays') {
    const date = e.parameter.date || '';
    return jsonResponse(getCarryoverPrepays(date));
  }

  if (action === 'getProducts') {
    return jsonResponse(getProducts());
  }

  return jsonResponse({ status: 'ok', message: 'ТӘП-ТӘТТІ API работает' });
}

// ════════════════════════════════════════════════════════════
//  POST — сохранение данных
// ════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.type === 'shift') {
      saveShift(data);
      return jsonResponse({ status: 'ok', message: 'Смена сохранена' });
    }

    if (data.type === 'transactions') {
      saveTransactions(data);
      return jsonResponse({ status: 'ok', message: 'Транзакции сохранены' });
    }

    return jsonResponse({ status: 'error', message: 'Неизвестный тип данных' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ════════════════════════════════════════════════════════════
//  Сохранение смены (upsert: обновляет если дата уже есть)
// ════════════════════════════════════════════════════════════
function saveShift(data) {
  const headers = [
    'Дата', 'Оператор', 'Итого Kaspi (₸)', 'Дневные (₸)',
    'Предоплаты сегодня (₸)', 'Предоплаты будущие (₸)',
    'Ожидаемый iiko (₸)', 'Фактический iiko (₸)', 'Расхождение iiko (₸)',
    'Кол-во транзакций', 'Статус сверки', 'Статус iiko', 'Комментарий',
    'Время записи'
  ];
  const sheet = getOrCreateSheet(SHEET_SHIFTS, headers);

  const kaspi      = data.totalKaspi   || 0;
  const daily      = data.totalDaily   || 0;
  const prepToday  = data.prepayToday  || 0;
  const prepFuture = data.prepayFuture || 0;
  const expIiko    = data.expectedIiko || 0;
  const actIiko    = data.actualIiko   || 0;
  const diffIiko   = actIiko > 0 ? actIiko - expIiko : '';

  const sverkaStatus = (kaspi === daily + (data.totalPrepay || 0))
    ? '✅ Сходится' : '❌ Расхождение';
  const iikoStatus = actIiko > 0
    ? (diffIiko === 0 ? '✅ Сходится' : '❌ ' + formatNum(Math.abs(diffIiko)) + ' ₸')
    : '—';

  const newRow = [
    asText(data.date),
    data.operator    || '',
    kaspi, daily, prepToday, prepFuture, expIiko,
    actIiko || '', diffIiko,
    data.txnCount    || 0,
    sverkaStatus, iikoStatus,
    data.comment     || '',
    new Date().toLocaleString('ru-RU')
  ];

  // Ищем существующую строку с такой же датой → заменяем, иначе добавляем
  const tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
  const allRows = sheet.getDataRange().getValues();
  let dataRows = allRows.slice(1);

  const idx = dataRows.findIndex(r => String(normalizeCell('Дата', r[0])) === String(data.date));
  if (idx !== -1) {
    dataRows[idx] = newRow;
  } else {
    dataRows.push(newRow);
  }

  // Нормализуем дату (колонка 0) каждой строки → чистая текстовая строка yyyy-MM-dd
  dataRows = dataRows.map(r => {
    const raw = r[0];
    const dateStr = raw instanceof Date
      ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd')
      : String(raw).replace(/^'/, '').trim();
    return [dateStr, ...r.slice(1)];
  });

  // Сортируем в JS по дате (ISO-строки сравниваются корректно)
  dataRows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  // Перезаписываем лист: очищаем старые данные, пишем отсортированные
  const lastRowOld = sheet.getLastRow();
  const numCols = sheet.getLastColumn();
  if (lastRowOld > 1) {
    sheet.getRange(2, 1, lastRowOld - 1, numCols).clearContent();
  }
  SpreadsheetApp.flush();

  if (dataRows.length > 0) {
    const withPrefix = dataRows.map(r => ["'" + r[0], ...r.slice(1)]);
    sheet.getRange(2, 1, withPrefix.length, withPrefix[0].length).setNumberFormat('@').setValues(withPrefix);
  }
}

// ════════════════════════════════════════════════════════════
//  Сохранение транзакций (upsert: удаляет старые за эту дату)
// ════════════════════════════════════════════════════════════
function saveTransactions(data) {
  const sheet = getOrCreateSheet(SHEET_TRANSACTIONS, [
    'Дата', 'Оператор', 'Время', 'Сумма (₸)',
    'Тип', 'Дата заказа', 'Клиент', 'Заметка'
  ]);

  // Столбцы с датами хранятся как текст, чтобы таблица не превращала их в объект Date
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1).setNumberFormat('@');

  // Удаляем все строки за эту дату (снизу вверх, чтобы не сбить индексы)
  const allRows = sheet.getDataRange().getValues();
  for (let i = allRows.length - 1; i >= 1; i--) {
    if (String(normalizeCell('Дата', allRows[i][0])) === String(data.date)) {
      sheet.deleteRow(i + 1);
    }
  }

  // Добавляем новые строки
  const newRows = (data.transactions || []).map(t => [
    asText(data.date),
    data.operator || '',
    t.time        || '',
    t.amount      || 0,
    t.type === 'prepay' ? 'Предоплата' : t.type === 'daily' ? 'Дневной' : '—',
    asText(t.orderDate),
    t.client      || '',
    t.note        || ''
  ]);
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }

  // Сортируем по дате, затем по времени
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .sort([{ column: 1, ascending: true }, { column: 3, ascending: true }]);
  }
}

// ════════════════════════════════════════════════════════════
//  Получение истории смен
// ════════════════════════════════════════════════════════════
function getShifts() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_SHIFTS);
    if (!sheet) return { status: 'ok', data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { status: 'ok', data: [] };

    const headers = rows[0];
    const data = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = normalizeCell(h, row[i]));
      return obj;
    });

    // Дата должна быть уникальной. Если в листе остались дубликаты
    // (например, из-за старого бага со сравнением дат), оставляем
    // только запись с самым последним «Время записи» для каждой даты.
    const byDate = {};
    data.forEach(row => {
      const d = row['Дата'];
      const existing = byDate[d];
      if (!existing || String(row['Время записи']) >= String(existing['Время записи'])) {
        byDate[d] = row;
      }
    });
    const deduped = Object.values(byDate).sort((a, b) => String(a['Дата']).localeCompare(String(b['Дата'])));

    return { status: 'ok', data: deduped };
  } catch (err) {
    return { status: 'error', message: err.message, data: [] };
  }
}

// ════════════════════════════════════════════════════════════
//  Получение транзакций за дату
// ════════════════════════════════════════════════════════════
function getTransactions(date) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TRANSACTIONS);
    if (!sheet) return { status: 'ok', data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { status: 'ok', data: [] };

    const headers = rows[0];
    const data = rows.slice(1)
      .filter(row => !date || String(normalizeCell(headers[0], row[0])) === date)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = normalizeCell(h, row[i]));
        return obj;
      });

    return { status: 'ok', data };
  } catch (err) {
    return { status: 'error', message: err.message, data: [] };
  }
}

// ════════════════════════════════════════════════════════════
//  Приводит значения дат к 'yyyy-MM-dd' (старые строки, сохранённые
//  до фикса с apostrophe-префиксом, могли превратиться в объект Date —
//  таблица интерпретирует текстовую дату по СВОЕМУ часовому поясу,
//  поэтому для восстановления исходной даты берём именно его, а не
//  Asia/Almaty)
// ════════════════════════════════════════════════════════════
function normalizeCell(header, value) {
  if (value instanceof Date && (header === 'Дата' || header === 'Дата заказа')) {
    const tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  return value;
}

// ════════════════════════════════════════════════════════════
//  Принудительно сохраняет значение как текст (апостроф — стандартный
//  для Google Sheets маркер «не превращать в дату/число»; при чтении
//  апостроф не виден и в значении не присутствует)
// ════════════════════════════════════════════════════════════
function asText(value) {
  return value ? ("'" + value) : (value || '');
}

// ════════════════════════════════════════════════════════════
//  Перенесённые предоплаты из прошлых смен
//  Возвращает сумму транзакций где:
//    Дата заказа = запрошенная дата
//    Дата смены  < запрошенная дата
// ════════════════════════════════════════════════════════════
function getCarryoverPrepays(date) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TRANSACTIONS);
    if (!sheet) return { status: 'ok', total: 0, count: 0 };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { status: 'ok', total: 0, count: 0 };

    const headers = rows[0];
    const idxDate      = headers.indexOf('Дата');
    const idxOrderDate = headers.indexOf('Дата заказа');
    const idxAmount    = headers.indexOf('Сумма (₸)');
    const idxType      = headers.indexOf('Тип');

    const matches = rows.slice(1).filter(row =>
      String(normalizeCell('Дата заказа', row[idxOrderDate])) === date &&   // заказ на нужную дату
      String(normalizeCell('Дата', row[idxDate]))             <  date &&    // записано в прошлой смене
      row[idxType] === 'Предоплата'
    );

    const total = matches.reduce((s, row) => s + (Number(row[idxAmount]) || 0), 0);
    return { status: 'ok', total, count: matches.length };
  } catch (err) {
    return { status: 'error', message: err.message, total: 0, count: 0 };
  }
}

// ════════════════════════════════════════════════════════════
//  Разовая очистка дублей в листе «Смены» (запускать вручную
//  из редактора Apps Script: выбрать функцию cleanupDuplicateShifts
//  в выпадающем списке вверху и нажать ▶ Выполнить).
//  Оставляет для каждой даты строку с самым последним «Время записи».
// ════════════════════════════════════════════════════════════
function cleanupDuplicateShifts() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_SHIFTS);
  if (!sheet) return;

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 2) return;

  const headers = rows[0];
  const idxDate    = headers.indexOf('Дата');
  const idxLogTime = headers.indexOf('Время записи');

  // Для каждой даты находим индекс строки (1-based, с учётом заголовка) с самым поздним «Время записи»
  const bestRowForDate = {};
  for (let i = 1; i < rows.length; i++) {
    const d = String(normalizeCell('Дата', rows[i][idxDate]));
    const logTime = String(rows[i][idxLogTime]);
    if (!bestRowForDate[d] || logTime >= bestRowForDate[d].logTime) {
      bestRowForDate[d] = { rowIndex: i, logTime };
    }
  }
  const keepRowIndexes = new Set(Object.values(bestRowForDate).map(v => v.rowIndex));

  // Удаляем снизу вверх все строки, которые не вошли в keepRowIndexes
  for (let i = rows.length - 1; i >= 1; i--) {
    if (!keepRowIndexes.has(i)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ════════════════════════════════════════════════════════════
//  Каталог продуктов — чтение и запись листа «Каталог»
// ════════════════════════════════════════════════════════════
function getProducts() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Каталог');
    if (!sheet) return { status: 'ok', data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { status: 'ok', data: [] };

    const data = rows.slice(1)
      .filter(r => r[7] !== false && String(r[7]).toLowerCase() !== 'false' && String(r[7]) !== '0')
      .map(r => {
        let dims = {};
        try { dims = r[3] ? JSON.parse(r[3]) : {}; } catch(e) {}
        return {
          name:   String(r[0] || ''),
          cat:    String(r[1] || ''),
          weight: String(r[2] || ''),
          dims,
          ingr:   String(r[4] || ''),
          desc:   String(r[5] || ''),
          shelf:  String(r[6] || '')
        };
      });

    return { status: 'ok', data };
  } catch(err) {
    return { status: 'error', message: err.message, data: [] };
  }
}

function initCatalogSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Каталог');
  if (sheet) { SpreadsheetApp.getUi().alert('Лист «Каталог» уже существует.'); return; }

  sheet = ss.insertSheet('Каталог');
  const headers = ['Атауы (название)', 'Категория', 'Салмақ (вес)', 'Өлшемдер JSON', 'Құрамы (состав)', 'Сипаттамасы (описание)', 'Сақтау (срок хранения)', 'Белсенді (TRUE/FALSE)'];
  sheet.appendRow(headers);
  const hr = sheet.getRange(1, 1, 1, headers.length);
  hr.setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 200);
  SpreadsheetApp.getUi().alert('Лист «Каталог» создан! Теперь запустите функцию fillCatalogFromCode чтобы заполнить его текущими продуктами.');
}

// ════════════════════════════════════════════════════════════
//  Вспомогательные функции
// ════════════════════════════════════════════════════════════
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    // Стиль заголовка
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatNum(n) {
  return Number(n).toLocaleString('ru-RU');
}

// ════════════════════════════════════════════════════════════
//  Разовый сдвиг всех дат на +1 день в листах «Смены» и «Транзакции»
//  Запустить вручную: выбрать shiftAllDatesForward и нажать ▶ Выполнить
// ════════════════════════════════════════════════════════════
function shiftAllDatesForward() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();

  function shiftSheet(sheetName, dateColumns) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;
    const headers = rows[0];
    const colIndexes = dateColumns.map(h => headers.indexOf(h));

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      colIndexes.forEach(ci => {
        if (ci === -1) return;
        const raw = row[ci];
        let dateStr = raw instanceof Date
          ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd')
          : String(raw).trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
          const d = new Date(dateStr + 'T00:00:00Z');
          d.setUTCDate(d.getUTCDate() + 1);
          const shifted = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
          sheet.getRange(i + 1, ci + 1).setNumberFormat('@').setValue("'" + shifted);
        }
      });
    }
  }

  shiftSheet(SHEET_SHIFTS,       ['Дата']);
  shiftSheet(SHEET_TRANSACTIONS, ['Дата', 'Дата заказа']);

  SpreadsheetApp.getUi().alert('Готово! Все даты сдвинуты на +1 день.');
}

// ════════════════════════════════════════════════════════════
//  Сдвиг всех дат на -1 день (отмена предыдущего сдвига)
//  Запустить вручную: выбрать shiftAllDatesBackward и нажать ▶ Выполнить
// ════════════════════════════════════════════════════════════
function shiftAllDatesBackward() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();

  function shiftSheet(sheetName, dateColumns) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;
    const headers = rows[0];
    const colIndexes = dateColumns.map(h => headers.indexOf(h));

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      colIndexes.forEach(ci => {
        if (ci === -1) return;
        const raw = row[ci];
        let dateStr = raw instanceof Date
          ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd')
          : String(raw).trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
          const d = new Date(dateStr + 'T00:00:00Z');
          d.setUTCDate(d.getUTCDate() - 1);
          const shifted = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
          sheet.getRange(i + 1, ci + 1).setNumberFormat('@').setValue("'" + shifted);
        }
      });
    }
  }

  shiftSheet(SHEET_SHIFTS,       ['Дата']);
  shiftSheet(SHEET_TRANSACTIONS, ['Дата', 'Дата заказа']);

  SpreadsheetApp.getUi().alert('Готово! Все даты сдвинуты на -1 день.');
}
