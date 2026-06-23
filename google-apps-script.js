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
    data.date        || '',
    data.operator    || '',
    kaspi, daily, prepToday, prepFuture, expIiko,
    actIiko || '', diffIiko,
    data.txnCount    || 0,
    sverkaStatus, iikoStatus,
    data.comment     || '',
    new Date().toLocaleString('ru-RU')
  ];

  // Ищем существующую строку с такой же датой → заменяем
  const allRows = sheet.getDataRange().getValues();
  let replaced = false;
  for (let i = 1; i < allRows.length; i++) {
    if (String(allRows[i][0]) === String(data.date)) {
      sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
      replaced = true;
      break;
    }
  }
  if (!replaced) sheet.appendRow(newRow);

  // Сортируем по дате по возрастанию (столбец A, пропускаем заголовок)
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .sort({ column: 1, ascending: true });
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

  // Удаляем все строки за эту дату (снизу вверх, чтобы не сбить индексы)
  const allRows = sheet.getDataRange().getValues();
  for (let i = allRows.length - 1; i >= 1; i--) {
    if (String(allRows[i][0]) === String(data.date)) {
      sheet.deleteRow(i + 1);
    }
  }

  // Добавляем новые строки
  const newRows = (data.transactions || []).map(t => [
    data.date     || '',
    data.operator || '',
    t.time        || '',
    t.amount      || 0,
    t.type === 'prepay' ? 'Предоплата' : t.type === 'daily' ? 'Дневной' : '—',
    t.orderDate   || '',
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
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });

    return { status: 'ok', data };
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
      .filter(row => !date || String(row[0]) === date)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });

    return { status: 'ok', data };
  } catch (err) {
    return { status: 'error', message: err.message, data: [] };
  }
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
      String(row[idxOrderDate]) === date &&   // заказ на нужную дату
      String(row[idxDate])      <  date &&    // записано в прошлой смене
      row[idxType] === 'Предоплата'
    );

    const total = matches.reduce((s, row) => s + (Number(row[idxAmount]) || 0), 0);
    return { status: 'ok', total, count: matches.length };
  } catch (err) {
    return { status: 'error', message: err.message, total: 0, count: 0 };
  }
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
