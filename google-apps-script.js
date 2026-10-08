/**
 * Приём ответов гостей с сайта-приглашения в Google Таблицу.
 *
 * Как подключить:
 * 1. Создайте новую Google Таблицу (sheets.new).
 * 2. Меню «Расширения» → «Apps Script». Удалите всё в редакторе и вставьте этот файл целиком.
 * 3. «Начать развёртывание» → «Новое развёртывание» → тип «Веб-приложение».
 *      Запуск от имени: «От моего имени»
 *      У кого есть доступ: «Все»
 * 4. Нажмите «Развернуть», разрешите доступ к таблице и скопируйте URL веб-приложения
 *    (он заканчивается на /exec).
 * 5. Вставьте этот URL в index.html в строку:  const SHEETS_URL = '';
 *
 * Если позже измените этот код — создайте новую версию развёртывания
 * («Управление развёртываниями» → карандаш → «Новая версия»), иначе изменения не применятся.
 */

const HEADERS = ['Дата ответа', 'Имя', 'Ответ', 'Кол-во гостей', 'Пожелания'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      clean(data.name, 80),
      clean(data.attending, 20),
      Math.max(0, Math.min(20, Number(data.guests) || 0)),
      clean(data.message, 500)
    ]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, info: 'RSVP endpoint is running' });
}

// Обрезаем длину и не даём тексту гостя превратиться в формулу таблицы
function clean(value, max) {
  const s = String(value == null ? '' : value).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
