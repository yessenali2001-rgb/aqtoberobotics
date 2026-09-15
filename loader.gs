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
 *  Если что-то не вышло — запустите checkInternet: она честно скажет,
 *  пускает ли Google этот скрипт в интернет и что отвечают адреса.
 *
 *  ЧЕСТНО О РИСКЕ: сервер будет выполнять код из вашего репозитория
 *  на GitHub автоматически. Репозиторий ваш. Если кто-то получит к нему
 *  доступ, он получит и выполнение кода в вашем Google-аккаунте.
 *  Выключить самообновление в любой момент: поставить SAFE_MODE = true
 *  и сохранить — сервер продолжит работать на последней рабочей копии.
 *************************************************************/

/* Два адреса одного и того же файла: если первый недоступен, берём второй. */
var CODE_URLS = [
  'https://raw.githubusercontent.com/yessenali2001-rgb/aqtoberobotics/main/Code.gs',
  'https://yessenali2001-rgb.github.io/aqtoberobotics/Code.gs'
];
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

/** Пусто — код годится. Иначе — причина отказа человеческими словами. */
function whyBad_(src) {
  if (!src) return 'пустой ответ';
  if (src.length < 5000) return 'слишком короткий: ' + src.length + ' знаков, ждём больше 5000';
  if (src.indexOf('function doGet') < 0) return 'нет функции doGet — скачалась не та страница';
  if (src.indexOf('var VERSION') < 0) return 'нет метки версии';
  try { make_(src, 'doGet'); } catch (e) { return 'не собирается: ' + errText_(e); }
  return '';
}

function looksValid_(src) { return !whyBad_(src); }

function errText_(e) { return (e && e.message) ? e.message : String(e); }

/**
 * Достаёт код: память → интернет → запасная копия.
 * В log (массив) складывает, что именно произошло, — это читает updateNow.
 */
function fetchCode_(log) {
  log = log || [];
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}

  if (cache && !SAFE_MODE) {
    var hit = cache.get('src');
    if (hit) { log.push('Взят из памяти (держим ' + CODE_TTL + ' секунд).'); return hit; }
  }

  if (SAFE_MODE) {
    log.push('SAFE_MODE = true — в интернет не ходим.');
  } else {
    for (var i = 0; i < CODE_URLS.length; i++) {
      var url = CODE_URLS[i];
      try {
        var res  = UrlFetchApp.fetch(url + '?t=' + Date.now(),
                     { muteHttpExceptions: true, followRedirects: true });
        var code = res.getResponseCode();
        if (code !== 200) { log.push('нет: ' + url + ' — ответ ' + code); continue; }

        var src = res.getContentText();
        var bad = whyBad_(src);
        if (bad) { log.push('нет: ' + url + ' — скачалось, но ' + bad); continue; }

        log.push('да: ' + url + ' — ' + Math.round(src.length / 1024) + ' КБ');
        if (cache) { try { cache.put('src', src, CODE_TTL); } catch (e) {} }
        try { saveBackup_(src); }
        catch (e) { log.push('   запасную копию сохранить не вышло: ' + errText_(e)); }
        return src;
      } catch (e) {
        log.push('нет: ' + url + ' — ' + errText_(e));
      }
    }
  }

  var backup = loadBackup_();
  if (backup) { log.push('Взята запасная копия из свойств проекта.'); return backup; }
  throw new Error('Код не загружен, и запасной копии нет.\n' + log.join('\n'));
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
function tell_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}   // из редактора окна нет — это нормально
  return msg;
}

function updateNow() {
  try { CacheService.getScriptCache().remove('src'); } catch (e) {}
  var log = [];
  try {
    var src = fetchCode_(log);
    var v = (src.match(/var VERSION = '([^']+)'/) || [])[1] || 'без метки';
    return tell_('Код на месте. Версия: ' + v +
                 '\nРазмер: ' + Math.round(src.length / 1024) + ' КБ\n\n' + log.join('\n'));
  } catch (e) {
    return tell_(
      'НЕ ПОЛУЧИЛОСЬ.\n\n' + errText_(e) +
      '\n\nЧастые причины:\n' +
      '1. Не выданы разрешения. Запустите updateNow ещё раз и пройдите все окна\n' +
      '   до кнопки «Разрешить» («Дополнительные настройки» → «Перейти к проекту»).\n' +
      '2. Аккаунт школьный или рабочий, и администратор запретил скриптам выходить\n' +
      '   в интернет. Признак — в строках выше слова про доступ к сервису.\n' +
      '   Тогда переносите проект в личный аккаунт Google.\n' +
      '3. Адреса недоступны. Откройте их в браузере: они должны показать текст кода.');
  }
}

/** Отдельная проверка: пускают ли этот скрипт в интернет и что отвечают адреса. */
function checkInternet() {
  var out = [];
  for (var i = 0; i < CODE_URLS.length; i++) {
    try {
      var r = UrlFetchApp.fetch(CODE_URLS[i] + '?t=' + Date.now(), { muteHttpExceptions: true });
      var t = r.getContentText();
      out.push(CODE_URLS[i] + '\n   ответ ' + r.getResponseCode() + ', ' + t.length + ' знаков' +
               (whyBad_(t) ? ', но ' + whyBad_(t) : ', код годится'));
    } catch (e) {
      out.push(CODE_URLS[i] + '\n   ' + errText_(e));
    }
  }
  var pr = PropertiesService.getScriptProperties();
  out.push('Запасная копия: ' + (Number(pr.getProperty('srcN') || 0) ? 'есть' : 'нет'));
  return tell_(out.join('\n\n'));
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
