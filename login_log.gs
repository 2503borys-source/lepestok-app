/**
 * Сервер Lepestok mini app: проверка входа + журнал входов.
 * Вставить в Google Apps Script, привязанный к Google Таблице (Расширения → Apps Script).
 * Каждая попытка входа = строка в таблице + сообщение тебе в Telegram (если заданы BOT_TOKEN и CHAT_ID).
 *
 * Настройка: Project Settings → Script properties:
 *   APP_LOGIN  = логин для входа в приложение
 *   APP_PASS   = пароль для входа в приложение
 *   ADMIN_PASS = пароль для просмотра журнала в приложении
 *   SECRET     = любая длинная случайная строка (подпись для «Запомнить меня»)
 *   BOT_TOKEN  = токен бота от @BotFather (необязательно, для уведомлений)
 *   CHAT_ID    = твой Telegram id, узнать у @userinfobot (необязательно)
 * Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → скопировать URL в API_URL в index.html
 * Логины и пароли живут только здесь — в публичном коде приложения их нет.
 */
const COLS = ['Время', 'Успех', 'Логин', 'Комментарий', 'TG id', 'TG username', 'Имя', 'Широта', 'Долгота', 'Точность, м', 'Карта', 'Платформа', 'Устройство'];

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const p = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // вход по логину и паролю
  if (d.action === 'login') {
    const ok = !!p.getProperty('APP_LOGIN') && d.login === p.getProperty('APP_LOGIN') && d.pass === p.getProperty('APP_PASS');
    const geo = !!d.lat;
    log(sheet, p, d, ok && geo, d.login, !ok ? 'неверный пароль' : !geo ? 'нет геолокации' : '');
    if (!ok) return json({error: 'Неверный логин или пароль'});
    if (!geo) return json({error: 'Разреши доступ к местоположению — без него вход невозможен'});
    return json({ok: true, token: sign(p, d.tgId)});
  }

  // «Запомнить меня»: проверка сохранённого токена
  if (d.action === 'check') {
    const ok = !!d.token && d.token === sign(p, d.tgId);
    log(sheet, p, d, ok, 'запомнен', ok ? 'автовход' + (d.lat ? '' : ', без геолокации') : 'недействительный токен');
    return json(ok ? {ok: true} : {error: 'expired'});
  }

  // админ: чтение журнала
  if (d.action === 'list') {
    const admin = p.getProperty('ADMIN_PASS');
    if (!admin || d.pass !== admin) return json({error: 'Неверный пароль'});
    const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getDisplayValues() : [];
    return json({rows: rows.reverse().slice(0, 300)});
  }

  return json({error: 'unknown action'});
}

// токен = HMAC(SECRET + логин + пароль, tgId): смена пароля или SECRET сбрасывает все «запомненные» входы
function sign(p, tgId) {
  const key = (p.getProperty('SECRET') || '') + '|' + p.getProperty('APP_LOGIN') + '|' + p.getProperty('APP_PASS');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(tgId || 'web'), key));
}

function log(sheet, p, d, success, login, note) {
  if (sheet.getLastRow() === 0) sheet.appendRow(COLS);
  const map = d.lat ? 'https://maps.google.com/?q=' + d.lat + ',' + d.lon : '';
  const time = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([time, success ? 'да' : 'нет', login, note, d.tgId, d.tgUser, d.tgName, d.lat, d.lon, d.acc, map, d.platform, d.ua]);

  const token = p.getProperty('BOT_TOKEN'), chat = p.getProperty('CHAT_ID');
  if (token && chat) {
    const text = (success ? '✅ Вход' : '⚠️ Попытка входа') + '\n' + time +
      '\n👤 ' + (d.tgName || '?') + (d.tgUser ? ' @' + d.tgUser : '') + ' (id ' + (d.tgId || '?') + ')' +
      '\nЛогин: ' + login + (note ? '\n' + note : '') +
      (map ? '\n📍 ' + map + ' (±' + d.acc + ' м)' : '\n📍 без геолокации');
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({chat_id: chat, text: text, disable_web_page_preview: true})
    });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
