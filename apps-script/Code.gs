/**
 * Приём заявок с лендинга «Свои запуски» в Google-таблицу.
 *
 * Как подключить — подробно в README.md, коротко:
 *   1. Создать Google-таблицу, открыть Расширения → Apps Script.
 *   2. Вставить этот код, сохранить.
 *   3. Развернуть → Новое развёртывание → Веб-приложение.
 *      Запуск от имени: «Я». Доступ: «Все».
 *   4. Скопировать URL вида https://script.google.com/macros/s/…/exec
 *      и вписать его в CONFIG.sheetEndpoint в index.html.
 *
 * Права доступа: скрипт привязан к таблице и пишет только в неё.
 */

// Лист, в который складываются заявки. Создаётся автоматически.
var SHEET_NAME = 'Заявки';

var HEADERS = ['Дата и время', 'Тип', 'Контакт', 'Страница', 'Источник', 'User-Agent'];

/**
 * Возвращает лист заявок, создавая его вместе с шапкой при первом запуске.
 */
function getSheet_() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = book.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // дата и время
    sheet.setColumnWidth(2, 90);  // тип
    sheet.setColumnWidth(3, 200); // контакт
  }

  return sheet;
}

/**
 * Принимает заявку с лендинга.
 *
 * Лендинг шлёт JSON с Content-Type: text/plain, чтобы браузер счёл запрос простым
 * и не отправлял preflight OPTIONS — Apps Script на него не отвечает.
 */
function doPost(e) {
  try {
    var payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    var contact = String(payload.contact || '').trim();
    if (!contact) {
      return json_({ ok: false, error: 'empty contact' });
    }

    var type = payload.type === 'phone' ? 'Телефон' : 'Телеграм';

    // Блокировка на случай, если несколько заявок придут одновременно:
    // без неё две записи могут лечь в одну строку.
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      getSheet_().appendRow([
        new Date(),
        type,
        contact,
        String(payload.page || ''),
        String(payload.referrer || ''),
        String(payload.userAgent || '')
      ]);
    } finally {
      lock.releaseLock();
    }

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * Открытие /exec в браузере — быстрая проверка, что развёртывание живо.
 */
function doGet() {
  return json_({ ok: true, service: 'svoi-zapuski registrations' });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
