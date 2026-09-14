/*************************************************************
 *  ЗАГРУЗЧИК — вставляется в Apps Script ОДИН РАЗ
 *
 *  Дальше сервер сам берёт свежий Code.gs из вашего репозитория,
 *  и вставлять код заново больше не нужно. Никогда.
 *
 *  ЧТО ДЕЛАТЬ:
 *   1. Apps Script → выделить весь код (Ctrl+A) → вставить этот файл.
 *   2. Сохранить.
 *   3. Выбрать функцию updateNow и нажать «Выполнить».
 *      Google попросит разрешения (доступ в интернет) — согласиться.
 *      В окне покажется версия загруженного кода.
 *   4. Развернуть → Управление развёртываниями → карандаш на строке
 *      «Веб-приложение» → Версия: Новая → Развернуть.
 *
 *  Всё. Больше сюда заходить не придётся.
 *
 *  ЧЕСТНО О РИСКЕ: сервер будет выполнять код из вашего репозитория
 *  на GitHub автоматически. Репозиторий ваш. Если кто-то получит к нему
 *  доступ, он получит и выполнение кода в вашем Google-аккаунте.
 *  Выключить самообновление в любой момент: поставить SAFE_MODE = true
 *  и сохранить — сервер продолжит работать на последней рабочей копии.
 *************************************************************/

var CODE_URL  = 'https://raw.githubusercontent.com/yessenali2001-rgb/aqtoberobotics/main/Code.gs';
var CODE_TTL  = 300;        // секунд держим скачанное в памяти (5 минут)
var SAFE_MODE = false;      // true — не ходить в интернет, работать на сохранённой копии

/* ---------- запасная копия в свойствах проекта ---------- */
var CHUNK = 8000;

function saveBackup_(src) {
  var pr = PropertiesService.getScriptProperties();
  var n = Math.ceil(src.length / CHUNK), old = Number(pr.getProperty('srcN') || 0), data = {};
  for (var i = 0; i < n; i++) data['src' + i] = src.substr(i * CHUNK, CHUNK);
  data.srcN = String(n);
  pr.setProperties(data);
  for (var j = n; j < old; j++) pr.deleteProperty('src' + j);   // хвост от прошлой версии
}

function loadBackup_() {
  var pr = PropertiesService.getScriptProperties();
  var n = Number(pr.getProperty('srcN') || 0);
  if (!n) return '';
  var out = '';
  for (var i = 0; i < n; i++) out += pr.getProperty('src' + i) || '';
  return out;
}

/* ---------- сборка и проверка кода ---------- */
function make_(src, fnName) {
  return new Function('__args',
    src + '\nreturn typeof ' + fnName + ' === "function" ? ' + fnName + '.apply(null, __args) : ' +
    '(function(){ throw new Error("В коде нет функции ' + fnName + '"); })();');
}

/** Код годится, только если он похож на наш и вообще собирается. */
function looksValid_(src) {
  if (!src || src.length < 5000) return false;
  if (src.indexOf('function doGet') < 0 || src.indexOf('var VERSION') < 0) return false;
  try { make_(src, 'doGet'); return true; } catch (e) { return false; }
}

function fetchCode_() {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache && !SAFE_MODE) {
    var hit = cache.get('src');
    if (hit) return hit;
  }
  if (!SAFE_MODE) {
    try {
      var res = UrlFetchApp.fetch(CODE_URL + '?t=' + Date.now(),
        { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() === 200) {
        var src = res.getContentText();
        if (looksValid_(src)) {
          if (cache) { try { cache.put('src', src, CODE_TTL); } catch (e) {} }
          saveBackup_(src);
          return src;
        }
      }
    } catch (e) {}       // нет связи или GitHub недоступен — идём на запасную копию
  }
  var backup = loadBackup_();
  if (!backup) throw new Error('Код ещё не загружен. Запустите функцию updateNow.');
  return backup;
}

/** Выполнить функцию из загруженного кода. Если он вдруг сломан — берём запасную копию. */
function run_(fnName, args) {
  var src = fetchCode_();
  try {
    return make_(src, fnName)(args || []);
  } catch (err) {
    var backup = loadBackup_();
    if (backup && backup !== src && looksValid_(backup)) {
      try { CacheService.getScriptCache().remove('src'); } catch (e) {}
      return make_(backup, fnName)(args || []);
    }
    throw err;
  }
}

/* ---------- то, что вызывает сам сервер ---------- */
function doGet(e)  { return run_('doGet',  [e]); }
function doPost(e) { return run_('doPost', [e]); }

/* ---------- то, что запускают руками ---------- */
function updateNow() {
  try { CacheService.getScriptCache().remove('src'); } catch (e) {}
  var src = fetchCode_();
  var v = (src.match(/var VERSION = '([^']+)'/) || [])[1] || 'без метки';
  var msg = 'Код обновлён. Версия: ' + v + '\nРазмер: ' + Math.round(src.length / 1024) + ' КБ';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

function showVersion()        { return updateNow(); }
function setup()              { return run_('setup'); }
function setupAll()           { return run_('setup'); }
function resetAdminPin()      { return run_('resetAdminPin'); }
function fixPeopleSheet()     { return run_('fixPeopleSheet'); }
function importBirthdays()    { return run_('importBirthdays'); }
function importSchoolAwards() { return run_('importSchoolAwards'); }
function installDailyMail()   { return run_('installDailyMail'); }
function removeDailyMail()    { return run_('removeDailyMail'); }
function sendDailyMail()      { return run_('sendDailyMail'); }
