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
const COLS = ['Время', 'Успех', 'TG ник', 'Имя', 'TG id', 'Комментарий'];

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const p = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  const who = String(d.tgId || 'web');

  // вход по логину и паролю
  if (d.action === 'login') {
    if (isLocked(p, who)) {
      log(sheet, p, d, false, d.login, 'заблокирован');
      return json({error: LOCK_MSG, locked: true});
    }
    const ok = !!p.getProperty('APP_LOGIN') && d.login === p.getProperty('APP_LOGIN') && d.pass === p.getProperty('APP_PASS');
    if (!ok) {
      const left = fail(p, who);
      log(sheet, p, d, false, d.login, 'неверный пароль' + (left > 0 ? ', осталось попыток: ' + left : ' — ЗАБЛОКИРОВАН'));
      return left > 0 ? json({error: 'Неверный логин или пароль. Осталось попыток: ' + left})
                      : json({error: LOCK_MSG, locked: true});
    }
    p.deleteProperty('fails_' + who);
    log(sheet, p, d, true, d.login, '');
    return json({ok: true, token: sign(p, d.tgId)});
  }

  // «Запомнить меня»: проверка сохранённого токена
  if (d.action === 'check') {
    if (isLocked(p, who)) return json({error: 'expired'});
    const ok = !!d.token && d.token === sign(p, d.tgId);
    log(sheet, p, d, ok, 'запомнен', ok ? 'автовход' : 'недействительный токен');
    return json(ok ? {ok: true} : {error: 'expired'});
  }

  // админ: журнал, список блокировок, снятие блокировок
  if (d.action === 'list' || d.action === 'unlock') {
    const admin = p.getProperty('ADMIN_PASS');
    if (!admin || d.pass !== admin) return json({error: 'Неверный пароль'});
    if (d.action === 'unlock') unlock(p, d.target);
    const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, COLS.length).getDisplayValues() : [];
    return json({rows: rows.reverse().slice(0, 300), locks: locks(p)});
  }

  return json({error: 'unknown action'});
}

// ---------- защита от подбора ----------
// 3 неверных пароля с одного Telegram-аккаунта → он заблокирован.
// 10 неверных паролей всего (с любых аккаунтов) → заблокирован вход для всех.
// Снять блокировку можно только в журнале входов по ADMIN_PASS.
const MAX_FAILS = 3, MAX_TOTAL = 10;
const LOCK_MSG = 'Слишком много неверных попыток. Вход заблокирован — обратись к администратору.';

function isLocked(p, who) {
  return !!p.getProperty('locked_' + who) || !!p.getProperty('locked_ALL');
}
function fail(p, who) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const n = Number(p.getProperty('fails_' + who) || 0) + 1;
    const total = Number(p.getProperty('fails_total') || 0) + 1;
    p.setProperty('fails_' + who, String(n));
    p.setProperty('fails_total', String(total));
    const now = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd HH:mm');
    if (n >= MAX_FAILS) p.setProperty('locked_' + who, now);
    if (total >= MAX_TOTAL) p.setProperty('locked_ALL', now);
    return MAX_FAILS - n;
  } finally { lock.releaseLock(); }
}
function locks(p) {
  const all = p.getProperties();
  return Object.keys(all).filter(k => k.indexOf('locked_') === 0)
    .map(k => ({id: k.slice(7), since: all[k]}));
}
function unlock(p, target) {
  const all = p.getProperties();
  Object.keys(all).forEach(k => {
    const id = k.replace(/^(locked_|fails_)/, '');
    if ((k.indexOf('locked_') === 0 || k.indexOf('fails_') === 0) && (target === '*' || id === target || (target === 'ALL' && k === 'fails_total')))
      p.deleteProperty(k);
  });
  if (target === 'ALL' || target === '*') p.deleteProperty('fails_total');
}

// токен = HMAC(SECRET + логин + пароль, tgId): смена пароля или SECRET сбрасывает все «запомненные» входы
function sign(p, tgId) {
  const key = (p.getProperty('SECRET') || '') + '|' + p.getProperty('APP_LOGIN') + '|' + p.getProperty('APP_PASS');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(tgId || 'web'), key));
}

function log(sheet, p, d, success, login, note) {
  if (sheet.getLastRow() === 0) sheet.appendRow(COLS);
  const time = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd HH:mm:ss');
  if (!success && login !== 'запомнен') note += ' (логин: ' + login + ')';
  sheet.appendRow([time, success ? 'да' : 'нет', d.tgUser || '', d.tgName || '', d.tgId || '', note]);

  const token = p.getProperty('BOT_TOKEN'), chat = p.getProperty('CHAT_ID');
  if (token && chat) {
    const text = (success ? '✅ Вход' : '⚠️ Попытка входа') + ' — ' + time +
      '\n👤 ' + (d.tgUser ? '@' + d.tgUser : 'без ника') + (d.tgName ? ' (' + d.tgName + ')' : '') +
      (note ? '\n' + note : '');
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({chat_id: chat, text: text, disable_web_page_preview: true})
    });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
