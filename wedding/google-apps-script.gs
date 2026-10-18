/**
 * Анкета гостей → Google Таблица
 * Свадебное приглашение: Михаил & Виктория, 17.07.2027
 *
 * КАК ПОДКЛЮЧИТЬ (5 минут):
 * 1. Создайте новую Google Таблицу (sheets.new) и назовите её, например, «Ответы гостей».
 * 2. В таблице: Расширения → Apps Script. Удалите весь код и вставьте этот файл целиком. Сохраните (Ctrl+S).
 * 3. Нажмите «Развернуть» → «Новое развёртывание» → тип «Веб-приложение»:
 *      • Выполнять от имени: «Я»
 *      • У кого есть доступ: «Все»
 *    Нажмите «Развернуть», разрешите доступ (Google попросит подтвердить аккаунт;
 *    если появится «Приложение не проверено» — «Дополнительно» → «Перейти … (небезопасно)»: это ваш собственный скрипт).
 * 4. Скопируйте «URL веб-приложения» (оканчивается на /exec).
 * 5. В файле index.html найдите строку
 *      var GOOGLE_SCRIPT_URL = 'REPLACE_WITH_YOUR_GOOGLE_SCRIPT_URL';
 *    и вставьте туда скопированный адрес.
 * 6. Проверка: откройте адрес /exec в браузере — должно появиться «Анкета работает».
 *    Заполните анкету на сайте — в таблице появится лист «Ответы» с новой строкой.
 *
 * ВАЖНО: если вы меняете этот код, после правок нужно снова «Развернуть» → «Управление развёртываниями» →
 * ✏️ → «Версия: Новая версия» → «Развернуть». Адрес /exec при этом не меняется.
 */

var SHEET_NAME   = 'Ответы';                 // имя листа в таблице
var TIMEZONE     = 'Asia/Yekaterinburg';     // Орск: UTC+5
var NOTIFY_EMAIL = '';                       // например 'mirror1825@gmail.com' — тогда на почту придёт письмо о каждом ответе. Пусто = без писем.

var HEADERS = ['Дата ответа', 'ФИО', 'Придёт', 'Состав', 'Гость 1', 'Гость 2', 'Гость 3', 'Напитки', 'Вино', 'Аллергия'];

/** Принимает данные анкеты с сайта (POST). */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);                      // чтобы одновременные ответы не перемешались
  try {
    var p = (e && e.parameter) || {};
    var row = [
      Utilities.formatDate(new Date(), TIMEZONE, 'dd.MM.yyyy HH:mm'),
      p.fio, p.attending, p.plus,
      p.guest1, p.guest2, p.guest3,
      p.drink, p.wine, p.allergy
    ].map(safe_);

    getSheet_().appendRow(row);
    if (NOTIFY_EMAIL) notify_(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Проверка из браузера. */
function doGet() {
  return ContentService.createTextOutput('Анкета работает');
}

/** Лист с ответами; при первом запуске создаёт заголовки. */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#F6E3C9');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, HEADERS.length, 150);
    sh.setColumnWidth(10, 260);
  }
  return sh;
}

/** Защита от формул: значения, начинающиеся с = + - @, записываются как текст. */
function safe_(v) {
  v = (v === undefined || v === null) ? '' : String(v).trim();
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function notify_(row) {
  var lines = HEADERS.map(function (h, i) { return h + ': ' + (row[i] || '—'); });
  MailApp.sendEmail(NOTIFY_EMAIL, 'Новый ответ на приглашение: ' + row[1], lines.join('\n'));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
