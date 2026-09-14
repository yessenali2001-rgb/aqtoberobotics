/*************************************************************
 *  ПОСЕЩАЕМОСТЬ КРУЖКА — ВСЁ В ОДНОМ ФАЙЛЕ
 *
 *  Этот файл собран из Code.gs и import-club-data.gs,
 *  чтобы его можно было вставить в Apps Script одной вставкой.
 *
 *  ЧТО ДЕЛАТЬ:
 *   1. Создайте Google-таблицу → Расширения → Apps Script.
 *   2. Удалите всё, что там есть, вставьте этот файл целиком, сохраните.
 *   3. Вверху выберите функцию setupAll и нажмите «Выполнить».
 *      Google спросит разрешения — согласитесь.
 *   4. Запишите код администратора из появившегося окна.
 *   5. Развернуть → Новое развёртывание → Веб-приложение,
 *      «Запуск от имени: Я», «Доступ: У всех» → скопируйте ссылку /exec.
 *
 *  Не редактируйте этот файл руками: он пересобирается
 *  командой sh build-setup-all.sh
 *************************************************************/

/*************************************************************
 *  ПОСЕЩАЕМОСТЬ КРУЖКА — серверная часть (Google Apps Script)
 *
 *  Все данные хранятся в этой же Google-таблице, на листах:
 *    people   — ученики и сотрудники (админ, помощник)
 *    marks    — отметки посещаемости
 *    settings — расписание, учебный год, выходные дни
 *    sessions — активные входы (токены)
 *
 *  Установка описана в attendance/README.md
 *************************************************************/

var VERSION = '2026-09-18';         // метка версии кода: видна в ответе сервера
var SHEET_ID = '';                    // пусто = скрипт привязан к таблице
var FIRST_ADMIN_NAME = 'Администратор';
var SESSION_DAYS = 90;                // сколько дней держится вход
var MAX_FAILS = 7;                    // неверных PIN подряд до блокировки
var LOCK_MINUTES = 15;                // на сколько блокируется вход

var SHEETS = {
  people:   ['id','role','name','pinHash','salt','active','since','createdAt','failCount','lockUntil','note',
             'cls','team','category','teamRole','birth','children','email'],
  marks:    ['id','date','studentId','status','hours','note','updatedAt','updatedBy'],
  settings: ['key','value'],
  sessions: ['token','personId','createdAt','expiresAt'],
  awards:   ['id','studentId','date','title','subject','level','result','note','updatedAt','updatedBy'],
  topics:   ['id','studentId','subject','title','status','due','note','updatedAt','updatedBy'],
  school:   ['id','year','category','event','team','award','result','note','updatedAt','updatedBy'],
  photos:   ['studentId','data','buf','updatedAt','updatedBy']
};

var STATUSES = ['p','l','e','a'];     // был / опоздал / уважительная / пропуск

/* ================= доступ к таблице ================= */

/* Обращения к таблице — самая дорогая часть запроса, поэтому книга,
   листы и уже прочитанные строки кешируются на время одного вызова. */
var _book = null, _tz = null, _sheet = {}, _rows = {};

function book_() {
  if (!_book) _book = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
  return _book;
}
function tz_() {
  if (!_tz) _tz = book_().getSpreadsheetTimeZone() || 'Etc/GMT';
  return _tz;
}
function forget_(name) { delete _rows[name]; }

var HDR_CHECKED = {};

function sheet_(name) {
  if (_sheet[name]) return _sheet[name];
  var bk = book_();
  var sh = bk.getSheetByName(name);
  if (!sh) {
    sh = bk.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
    if (name === 'marks') sh.getRange('B:B').setNumberFormat('@');
    if (name === 'people') sh.getRange('G:H').setNumberFormat('@');
    if (name === 'people') sh.getRange('P:P').setNumberFormat('@');
    HDR_CHECKED[name] = true;
    _sheet[name] = sh;
    return sh;
  }
  // лист создан более старой версией скрипта — дописываем недостающие колонки
  if (!HDR_CHECKED[name]) {
    HDR_CHECKED[name] = true;
    if (sh.getLastColumn() < SHEETS[name].length) {
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
    }
  }
  _sheet[name] = sh;
  return sh;
}

function rows_(name) {
  if (_rows[name]) return _rows[name];
  var vals = sheet_(name).getDataRange().getValues();
  var head = vals.shift() || SHEETS[name];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === '') continue;
    var o = { _row: i + 2 };
    for (var j = 0; j < head.length; j++) o[head[j]] = vals[i][j];
    out.push(o);
  }
  _rows[name] = out;
  return out;
}

function append_(name, obj) {
  sheet_(name).appendRow(SHEETS[name].map(function (h) {
    return obj[h] === undefined ? '' : obj[h];
  }));
  forget_(name);
}

function update_(name, row, patch) {
  var sh = sheet_(name), cols = SHEETS[name];
  var cur = sh.getRange(row, 1, 1, cols.length).getValues()[0];
  var out = cols.map(function (h, i) { return patch[h] === undefined ? cur[i] : patch[h]; });
  sh.getRange(row, 1, 1, out.length).setValues([out]);
  forget_(name);
}

function drop_(name, row) { sheet_(name).deleteRow(row); forget_(name); }

/* ================= мелкие утилиты ================= */

function dstr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}
/* Дату в таблице могут вписать как угодно: 24.03.2012, 24/03/2012, 2012-03-24
   или обычной датой Google. Приводим всё к виду ГГГГ-ММ-ДД. */
function looseDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  var t = String(v == null ? '' : v).trim();
  if (!t) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  var m = t.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{4})$/);   // 24.03.2012
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = t.match(/^(\d{4})[.\/](\d{1,2})[.\/](\d{1,2})$/);            // 2012/03/24
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return '';
}

function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function nowIso_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function newId_(prefix) {
  return prefix + Date.now().toString(36) + Math.floor(Math.random() * 1679616).toString(36);
}
function rndStr_(n) {
  var a = 'abcdefghijkmnpqrstuvwxyz23456789', s = '';
  for (var i = 0; i < n; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}
function genPin_(len) {
  var s = '';
  for (var i = 0; i < len; i++) s += String(Math.floor(Math.random() * 10));
  if (s.charAt(0) === '0') s = '1' + s.substring(1);
  return s;
}
function hashPin_(pin, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    salt + '|' + String(pin), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
function num_(v, def) {
  var n = Number(v);
  return isNaN(n) ? (def || 0) : n;
}

/* ================= настройки ================= */

function defaultSettings_() {
  var y = new Date().getFullYear(), m = new Date().getMonth();
  var start = (m >= 7 ? y : y - 1) + '-09-01';
  var end = (m >= 7 ? y + 1 : y) + '-05-31';
  return {
    groupName: 'Кружок',
    // часы занятий по дням недели: 0 — воскресенье … 6 — суббота
    schedule: [0, 1, 1, 1, 1, 1, 4],
    yearStart: start,
    yearEnd: end,
    holidays: [],                      // [{date:'2026-01-01', note:'Каникулы'}]
    siteUrl: 'https://yessenali2001-rgb.github.io/aqtoberobotics/',
    mailEnabled: '0',                  // рассылка родителям выключена, пока не включат
    mailHour: '19',                    // во сколько уходит письмо
    competitions: [],                  // [{title:'WRO', when:'April - May', term:'4 term', status:''}]
    plan: []                           // [{month:'SEPTEMBER', week:'I', n:'1', role:'Builder', task:'…'}]
  };
}

function readSettings_() {
  var s = defaultSettings_();
  rows_('settings').forEach(function (r) {
    var k = String(r.key);
    if (!(k in s)) return;
    var v = r.value;
    if (k === 'schedule' || k === 'holidays' || k === 'competitions' || k === 'plan') {
      try { v = JSON.parse(v); } catch (e) { return; }
    } else { v = dstr_(v); }
    s[k] = v;
  });
  if (!Array.isArray(s.schedule) || s.schedule.length !== 7) s.schedule = defaultSettings_().schedule;
  s.schedule = s.schedule.map(function (h) { return Math.max(0, num_(h, 0)); });
  if (!Array.isArray(s.holidays)) s.holidays = [];
  if (!Array.isArray(s.competitions)) s.competitions = [];
  s.competitions = s.competitions.map(function (c, i) {
    if (!c.id) c.id = 'c' + i + '-' + String(c.title || '').replace(/[^0-9A-Za-zА-Яа-я]/g, '').slice(0, 12);
    return c;
  });
  if (!Array.isArray(s.plan)) s.plan = [];
  return s;
}

function writeSetting_(key, value) {
  var v = (typeof value === 'object') ? JSON.stringify(value) : String(value);
  var found = null;
  rows_('settings').forEach(function (r) { if (String(r.key) === key) found = r; });
  if (found) update_('settings', found._row, { key: key, value: v });
  else append_('settings', { key: key, value: v });
}

/* ================= люди и вход ================= */

function personPub_(p) {
  return {
    id: String(p.id), role: String(p.role), name: String(p.name),
    active: String(p.active) !== '0' && p.active !== false,
    since: looseDate_(p.since), createdAt: dstr_(p.createdAt), note: String(p.note || ''),
    cls: String(p.cls || ''), team: String(p.team || ''),
    category: String(p.category || ''), teamRole: String(p.teamRole || ''),
    birth: looseDate_(p.birth),
    children: String(p.children || '').split(',').map(function (x) { return x.trim(); })
      .filter(function (x) { return x; }),
    email: String(p.email || '').trim()
  };
}

function findPerson_(id) {
  var all = rows_('people'), hit = null;
  all.forEach(function (p) { if (String(p.id) === String(id)) hit = p; });
  return hit;
}

function createPerson_(role, name, pin, since, extra) {
  name = String(name || '').trim();
  if (!name) throw new Error('Не указано имя.');
  extra = extra || {};
  var salt = rndStr_(12);
  var row = {
    id: newId_(role === 'student' ? 's' : 'u'),
    role: role, name: name,
    pinHash: hashPin_(pin, salt), salt: salt,
    active: extra.active === 0 ? 0 : 1, since: since || today_(), createdAt: today_(),
    failCount: 0, lockUntil: '', note: String(extra.note || ''),
    cls: String(extra.cls || ''), team: String(extra.team || ''),
    category: String(extra.category || ''), teamRole: String(extra.teamRole || ''),
    birth: dstr_(extra.birth),
    children: (extra.children || []).join(','),
    email: String(extra.email || '').trim()
  };
  append_('people', row);
  dropBootCache_();
  return row;
}

function makeSession_(personId) {
  var t = rndStr_(10) + rndStr_(10) + rndStr_(10);
  var exp = new Date(Date.now() + SESSION_DAYS * 864e5);
  append_('sessions', {
    token: t, personId: personId, createdAt: nowIso_(),
    expiresAt: Utilities.formatDate(exp, tz_(), 'yyyy-MM-dd HH:mm:ss')
  });
  return t;
}

function auth_(token) {
  token = String(token || '');
  if (!token) throw new Error('AUTH');
  var hit = null;
  rows_('sessions').forEach(function (r) { if (String(r.token) === token) hit = r; });
  if (!hit) throw new Error('AUTH');
  var exp = hit.expiresAt instanceof Date ? hit.expiresAt : new Date(String(hit.expiresAt).replace(' ', 'T'));
  if (exp && exp.getTime() < Date.now()) { drop_('sessions', hit._row); throw new Error('AUTH'); }
  var p = findPerson_(hit.personId);
  if (!p || String(p.active) === '0') throw new Error('AUTH');
  return p;
}

function requireStaff_(p) {
  if (p.role !== 'admin' && p.role !== 'assistant') throw new Error('Недостаточно прав.');
}
function childrenOf_(p) {
  return String(p.children || '').split(',').map(function (x) { return x.trim(); })
    .filter(function (x) { return x; });
}
function requireAdmin_(p) {
  if (p.role !== 'admin') throw new Error('Это может сделать только администратор.');
}

/* ================= выборка данных ================= */

function marksPub_(all, onlyStudentId) {
  var out = [];
  all.forEach(function (m) {
    var sid = String(m.studentId);
    if (onlyStudentId && sid !== String(onlyStudentId)) return;
    out.push({ d: dstr_(m.date), u: sid, s: String(m.status), h: num_(m.hours, 0), n: String(m.note || '') });
  });
  return out;
}

function awardsPub_(all, onlyStudentId) {
  var out = [];
  all.forEach(function (a) {
    var sid = String(a.studentId);
    if (onlyStudentId && sid !== String(onlyStudentId)) return;
    out.push({
      id: String(a.id), u: sid, date: dstr_(a.date), title: String(a.title || ''),
      subject: String(a.subject || ''), level: String(a.level || ''),
      result: String(a.result || ''), note: String(a.note || '')
    });
  });
  return out;
}

function topicsPub_(all, onlyStudentId) {
  var out = [];
  all.forEach(function (t) {
    var sid = String(t.studentId);
    if (onlyStudentId && sid !== String(onlyStudentId)) return;
    out.push({
      id: String(t.id), u: sid, subject: String(t.subject || ''), title: String(t.title || ''),
      status: String(t.status || 'todo'), due: dstr_(t.due), note: String(t.note || '')
    });
  });
  return out;
}

/** Награды всего кружка — без идентификаторов, только для общей сводки. */
function clubAwards_(people) {
  var byId = {};
  people.forEach(function (p) { byId[p.id] = p; });
  return awardsPub_(rows_('awards')).map(function (a) {
    var p = byId[a.u] || {};
    return { name: p.name || '', team: p.team || '', cls: p.cls || '',
             title: a.title, result: a.result, level: a.level };
  }).filter(function (a) { return a.name; });
}

function stateFor_(person) {
  var settings = readSettings_();
  var people = rows_('people').map(personPub_);
  var me = personPub_(person);

  if (person.role === 'parent') {
    var kids = childrenOf_(person);
    var mine = people.filter(function (x) { return kids.indexOf(x.id) >= 0; });
    var marks = marksPub_(rows_('marks')).filter(function (m) { return kids.indexOf(m.u) >= 0; });
    var photos = photoMap_(), myPhotos = {};
    kids.forEach(function (id) { if (photos[id]) myPhotos[id] = photos[id]; });
    return {
      me: me, settings: settings, people: mine, marks: marks,
      awards: awardsPub_(rows_('awards')).filter(function (a) { return kids.indexOf(a.u) >= 0; }),
      topics: topicsPub_(rows_('topics')).filter(function (t) { return kids.indexOf(t.u) >= 0; }),
      school: rows_('school').map(schoolPub_),
      clubAwards: clubAwards_(people),
      photoMap: myPhotos,
      version: VERSION
    };
  }
  if (person.role === 'student') {
    return {
      me: me, settings: settings,
      people: people.filter(function (p) { return p.id === me.id; }),
      marks: marksPub_(rows_('marks'), me.id),
      awards: awardsPub_(rows_('awards'), me.id),
      topics: topicsPub_(rows_('topics'), me.id),
      birthdays: people.filter(function (x) { return x.role === 'student' && x.active && x.birth; })
        .map(function (x) { return { name: x.name, birth: x.birth, team: x.team }; }),
      clubAwards: clubAwards_(people),
      school: rows_('school').map(schoolPub_),
      photoMap: photoMap_(),
      version: VERSION
    };
  }
  var withBirth = people.filter(function (x) { return x.role === 'student' && x.birth; }).length;
  return {
    me: me, settings: settings, people: people,
    marks: marksPub_(rows_('marks')),
    awards: awardsPub_(rows_('awards')),
    topics: topicsPub_(rows_('topics')),
    school: rows_('school').map(schoolPub_),
    photoMap: photoMap_(),
    version: VERSION,
    stats: { students: people.filter(function (x) { return x.role === 'student' && x.active; }).length,
             withBirth: withBirth, awards: rows_('awards').length,
             competitions: settings.competitions.length, plan: settings.plan.length,
             school: rows_('school').length,
             photos: Object.keys(photoMap_()).length,
             parents: people.filter(function (x) { return x.role === 'parent' && x.active; }).length }
  };
}

/* ================= API ================= */

var API = {};

/** Сброс кеша списка входа — после любых изменений людей или настроек. */
function dropBootCache_() {
  try { CacheService.getScriptCache().remove('boot'); } catch (e) {}
}

/** Список имён для экрана входа. Меняется редко, поэтому кешируется. */
API.boot = function () {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache) {
    var hit = cache.get('boot');
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  }
  var s = readSettings_();
  var roster = rows_('people')
    .filter(function (p) { return String(p.active) !== '0'; })
    .map(function (p) { return { id: String(p.id), name: String(p.name), role: String(p.role) }; })
    .sort(function (a, b) { return a.name.localeCompare(b.name, 'ru'); });
  var out = { groupName: s.groupName, roster: roster, needSetup: roster.length === 0, version: VERSION };
  if (cache) { try { cache.put('boot', JSON.stringify(out), 300); } catch (e) {} }
  return out;
};

API.login = function (personId, pin) {
  var p = findPerson_(personId);
  if (!p || String(p.active) === '0') throw new Error('Такого участника нет в списке.');
  var lockUntil = p.lockUntil ? new Date(String(p.lockUntil).replace(' ', 'T')) : null;
  if (lockUntil && lockUntil.getTime() > Date.now()) {
    var min = Math.ceil((lockUntil.getTime() - Date.now()) / 60000);
    throw new Error('Слишком много неверных попыток. Подождите ' + min + ' мин.');
  }
  if (hashPin_(String(pin || '').trim(), p.salt) !== String(p.pinHash)) {
    var fails = num_(p.failCount, 0) + 1;
    var patch = { failCount: fails };
    if (fails >= MAX_FAILS) {
      patch.lockUntil = Utilities.formatDate(new Date(Date.now() + LOCK_MINUTES * 60000), tz_(), 'yyyy-MM-dd HH:mm:ss');
      patch.failCount = 0;
    }
    update_('people', p._row, patch);
    throw new Error('Неверный код.');
  }
  update_('people', p._row, { failCount: 0, lockUntil: '' });
  var st = stateFor_(p);
  st.token = makeSession_(p.id);
  return st;
};

API.state = function (token) { return stateFor_(auth_(token)); };

API.logout = function (token) {
  rows_('sessions').forEach(function (r) { if (String(r.token) === String(token)) drop_('sessions', r._row); });
  return true;
};

API.changePin = function (token, oldPin, newPin) {
  var p = auth_(token);
  if (hashPin_(String(oldPin || '').trim(), p.salt) !== String(p.pinHash)) throw new Error('Старый код неверный.');
  newPin = String(newPin || '').trim();
  if (!/^\d{4,8}$/.test(newPin)) throw new Error('Новый код — от 4 до 8 цифр.');
  var salt = rndStr_(12);
  update_('people', p._row, { salt: salt, pinHash: hashPin_(newPin, salt) });
  return true;
};

/* ---- отметки ---- */

function setOneMark_(staff, date, studentId, status, hours, note) {
  date = dstr_(date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Неверная дата.');
  var st = String(studentId);
  var target = findPerson_(st);
  if (!target || target.role !== 'student') throw new Error('Ученик не найден.');

  var existing = null;
  rows_('marks').forEach(function (m) {
    if (dstr_(m.date) === date && String(m.studentId) === st) existing = m;
  });

  if (status === 'clear' || status === '' || status == null) {
    if (existing) drop_('marks', existing._row);
    return null;
  }
  if (STATUSES.indexOf(status) < 0) throw new Error('Неизвестный статус.');
  var h = Math.max(0, num_(hours, 0));
  if (status === 'a' || status === 'e') h = 0;

  var patch = {
    date: date, studentId: st, status: status, hours: h,
    note: String(note || ''), updatedAt: nowIso_(), updatedBy: String(staff.name)
  };
  if (existing) update_('marks', existing._row, patch);
  else { patch.id = newId_('m'); append_('marks', patch); }
  return { d: date, u: st, s: status, h: h, n: patch.note };
}

API.setMark = function (token, date, studentId, status, hours, note) {
  var p = auth_(token); requireStaff_(p);
  return setOneMark_(p, date, studentId, status, hours, note);
};

API.setMarks = function (token, date, items) {
  var p = auth_(token); requireStaff_(p);
  var out = [];
  (items || []).forEach(function (it) {
    out.push(setOneMark_(p, date, it.u, it.s, it.h, it.n));
  });
  return out;
};

/* ---- ученики и сотрудники ---- */

API.addStudent = function (token, name, since, extra) {
  var p = auth_(token); requireStaff_(p);
  var pin = genPin_(4);
  var row = createPerson_('student', name, pin, dstr_(since) || today_(), extra);
  return { person: personPub_(row), pin: pin };
};

API.addStudents = function (token, names, since, extra) {
  var p = auth_(token); requireStaff_(p);
  var out = [];
  (names || []).forEach(function (n) {
    n = String(n || '').trim();
    if (!n) return;
    var pin = genPin_(4);
    out.push({ person: personPub_(createPerson_('student', n, pin, dstr_(since) || today_(), extra)), pin: pin });
  });
  return out;
};

/** Родитель: свой код, видит только своих детей. */
API.addParent = function (token, name, childIds, email) {
  var p = auth_(token); requireStaff_(p);
  var kids = (childIds || []).map(String).filter(function (id) {
    var c = findPerson_(id);
    return c && c.role === 'student';
  });
  if (!kids.length) throw new Error('Выберите хотя бы одного ребёнка.');
  var pin = genPin_(4);
  var row = createPerson_('parent', name, pin, today_(), { children: kids, email: email });
  return { person: personPub_(row), pin: pin };
};

API.addStaff = function (token, name, role) {
  var p = auth_(token); requireAdmin_(p);
  if (role !== 'admin' && role !== 'assistant') throw new Error('Роль должна быть admin или assistant.');
  var pin = genPin_(6);
  return { person: personPub_(createPerson_(role, name, pin)), pin: pin };
};

API.savePerson = function (token, id, patch) {
  var p = auth_(token); requireStaff_(p);
  var t = findPerson_(id);
  if (!t) throw new Error('Участник не найден.');
  patch = patch || {};
  var out = {};
  if (patch.name !== undefined) {
    var nm = String(patch.name).trim();
    if (!nm) throw new Error('Имя не может быть пустым.');
    out.name = nm;
  }
  if (patch.since !== undefined) out.since = dstr_(patch.since);
  if (patch.note !== undefined) out.note = String(patch.note);
  if (patch.birth !== undefined) out.birth = dstr_(patch.birth);
  if (patch.email !== undefined) out.email = String(patch.email).trim();
  if (patch.children !== undefined) {
    out.children = (patch.children || []).map(String).filter(function (id) {
      var c = findPerson_(id);
      return c && c.role === 'student';
    }).join(',');
  }
  ['cls', 'team', 'category', 'teamRole'].forEach(function (k) {
    if (patch[k] !== undefined) out[k] = String(patch[k]).trim();
  });
  if (patch.active !== undefined || patch.role !== undefined) {
    requireAdmin_(p);
    if (patch.active !== undefined) out.active = patch.active ? 1 : 0;
    if (patch.role !== undefined) {
      if (['admin', 'assistant', 'student', 'parent'].indexOf(patch.role) < 0) throw new Error('Неизвестная роль.');
      out.role = patch.role;
    }
  }
  if (t.role !== 'student' && p.role !== 'admin') throw new Error('Менять сотрудников может только администратор.');
  if (String(t.id) === String(p.id) && out.active === 0) throw new Error('Нельзя отключить самого себя.');
  update_('people', t._row, out);
  dropBootCache_();
  return personPub_(findPerson_(id));
};

API.resetPin = function (token, id) {
  var p = auth_(token); requireStaff_(p);
  var t = findPerson_(id);
  if (!t) throw new Error('Участник не найден.');
  if (t.role !== 'student' && t.role !== 'parent') requireAdmin_(p);
  var pin = genPin_(t.role === 'student' || t.role === 'parent' ? 4 : 6);
  var salt = rndStr_(12);
  update_('people', t._row, { salt: salt, pinHash: hashPin_(pin, salt), failCount: 0, lockUntil: '' });
  return { id: String(t.id), pin: pin };
};

API.deletePerson = function (token, id) {
  var p = auth_(token); requireAdmin_(p);
  var t = findPerson_(id);
  if (!t) throw new Error('Участник не найден.');
  if (String(t.id) === String(p.id)) throw new Error('Нельзя удалить самого себя.');
  rows_('marks').slice().reverse().forEach(function (m) {
    if (String(m.studentId) === String(id)) drop_('marks', m._row);
  });
  rows_('photos').slice().reverse().forEach(function (f) {
    if (String(f.studentId) === String(id)) drop_('photos', f._row);
  });
  rows_('awards').slice().reverse().forEach(function (a) {
    if (String(a.studentId) === String(id)) drop_('awards', a._row);
  });
  rows_('topics').slice().reverse().forEach(function (t) {
    if (String(t.studentId) === String(id)) drop_('topics', t._row);
  });
  rows_('sessions').slice().reverse().forEach(function (s) {
    if (String(s.personId) === String(id)) drop_('sessions', s._row);
  });
  drop_('people', t._row);
  dropBootCache_();
  return true;
};

/* ---- фотографии учеников ----
   Картинка уменьшается прямо в браузере и приходит частями: за один запрос
   Apps Script принимает около 8 КБ. Храним в таблице, наружу не отдаём —
   только тем, кто вошёл в систему. */

var PHOTO_MAX = 60000;   // предел на одну ячейку таблицы

function findPhoto_(studentId) {
  var hit = null;
  rows_('photos').forEach(function (r) { if (String(r.studentId) === String(studentId)) hit = r; });
  return hit;
}

/** Кусок фотографии. part — с нуля, total — сколько всего кусков. */
API.photoChunk = function (token, studentId, part, total, chunk) {
  var p = auth_(token); requireStaff_(p);
  checkStudent_(studentId);
  part = num_(part, 0); total = Math.max(1, num_(total, 1));
  chunk = String(chunk || '');

  var hit = findPhoto_(studentId);
  var buf = (part === 0 || !hit) ? chunk : String(hit.buf || '') + chunk;
  if (buf.length > PHOTO_MAX) throw new Error('Фотография слишком большая.');

  var last = (part + 1 >= total);
  var patch = last
    ? { data: buf, buf: '', updatedAt: nowIso_(), updatedBy: String(p.name) }
    : { buf: buf, updatedAt: nowIso_(), updatedBy: String(p.name) };

  if (hit) update_('photos', hit._row, patch);
  else {
    patch.studentId = String(studentId);
    if (!patch.data) patch.data = '';
    append_('photos', patch);
  }
  return { got: part + 1, of: total, done: last, size: buf.length };
};

/** Сами картинки — отдельным запросом, чтобы не утяжелять загрузку. */
API.photos = function (token, ids) {
  auth_(token);
  var want = {};
  (ids || []).forEach(function (x) { want[String(x)] = true; });
  var out = {};
  rows_('photos').forEach(function (r) {
    var id = String(r.studentId);
    if (!want[id]) return;
    var d = String(r.data || '');
    if (d) out[id] = d;
  });
  return out;
};

API.deletePhoto = function (token, studentId) {
  var p = auth_(token); requireStaff_(p);
  var hit = findPhoto_(studentId);
  if (hit) drop_('photos', hit._row);
  return true;
};

/** Кто с фотографией и когда она обновлялась — чтобы браузер знал, что перечитать. */
function photoMap_() {
  var out = {};
  rows_('photos').forEach(function (r) {
    if (String(r.data || '')) out[String(r.studentId)] = dstr_(r.updatedAt) || '1';
  });
  return out;
}

/* ---- достижения школы (командные, по годам) ---- */

function schoolPub_(r) {
  return {
    id: String(r.id), year: String(r.year || ''), category: String(r.category || ''),
    event: String(r.event || ''), team: String(r.team || ''),
    award: String(r.award || ''), result: String(r.result || ''), note: String(r.note || '')
  };
}

API.saveSchoolAward = function (token, rec) {
  var p = auth_(token); requireStaff_(p);
  rec = rec || {};
  var patch = {
    year: String(rec.year || '').trim(), category: String(rec.category || '').trim(),
    event: String(rec.event || '').trim(), team: String(rec.team || '').trim(),
    award: String(rec.award || '').trim(),
    result: RESULTS.indexOf(rec.result) >= 0 ? rec.result : '',
    note: String(rec.note || ''), updatedAt: nowIso_(), updatedBy: String(p.name)
  };
  if (!patch.award && !patch.event) throw new Error('Укажите соревнование или награду.');
  var hit = null;
  if (rec.id) rows_('school').forEach(function (r) { if (String(r.id) === String(rec.id)) hit = r; });
  if (hit) { update_('school', hit._row, patch); patch.id = String(hit.id); }
  else { patch.id = newId_('sa'); append_('school', patch); }
  return schoolPub_(patch);
};

API.deleteSchoolAward = function (token, id) {
  var p = auth_(token); requireStaff_(p);
  var hit = null;
  rows_('school').forEach(function (r) { if (String(r.id) === String(id)) hit = r; });
  if (hit) drop_('school', hit._row);
  return true;
};

/* ---- календарь соревнований ---- */

var COMP_STATUS = ['plan', 'prep', 'done', 'skip'];

function compPub_(c) {
  c = c || {};
  return {
    id: String(c.id || ''),
    title: String(c.title || '').trim(),
    when: String(c.when || '').trim(),
    term: String(c.term || '').trim(),
    dateFrom: dstr_(c.dateFrom),
    dateTo: dstr_(c.dateTo),
    category: String(c.category || '').trim(),
    place: String(c.place || '').trim(),
    status: COMP_STATUS.indexOf(c.status) >= 0 ? c.status : '',
    note: String(c.note || ''),
    members: (c.members || []).map(String)
  };
}

/** Добавить или изменить соревнование. */
API.saveCompetition = function (token, comp) {
  var p = auth_(token); requireStaff_(p);
  var one = compPub_(comp);
  if (!one.title) throw new Error('Укажите название соревнования.');
  var list = readSettings_().competitions.map(compPub_);
  var found = false;
  if (one.id) {
    list = list.map(function (c) {
      if (c.id && c.id === one.id) { found = true; return one; }
      return c;
    });
  }
  if (!found) {
    if (!one.id) one.id = newId_('c');
    list.push(one);
  }
  writeSetting_('competitions', list);
  return list;
};

API.deleteCompetition = function (token, id) {
  var p = auth_(token); requireStaff_(p);
  var list = readSettings_().competitions.map(compPub_).filter(function (c) { return c.id !== String(id); });
  writeSetting_('competitions', list);
  return list;
};

/** Результаты целого соревнования: по записи на каждого участника. */
API.saveAwards = function (token, awards) {
  var p = auth_(token); requireStaff_(p);
  var out = [];
  (awards || []).forEach(function (a) { out.push(API.saveAward(token, a)); });
  return out;
};

/* ---- олимпиады и достижения ---- */

var LEVELS = ['school','city','region','republic','international'];
var RESULTS = ['gold','silver','bronze','honor','part'];

function checkStudent_(id) {
  var t = findPerson_(id);
  if (!t || t.role !== 'student') throw new Error('Ученик не найден.');
  return t;
}

API.saveAward = function (token, award) {
  var p = auth_(token); requireStaff_(p);
  award = award || {};
  checkStudent_(award.u);
  var patch = {
    studentId: String(award.u),
    date: dstr_(award.date),
    title: String(award.title || '').trim(),
    subject: String(award.subject || '').trim(),
    level: LEVELS.indexOf(award.level) >= 0 ? award.level : '',
    result: RESULTS.indexOf(award.result) >= 0 ? award.result : '',
    note: String(award.note || ''),
    updatedAt: nowIso_(), updatedBy: String(p.name)
  };
  if (!patch.title) throw new Error('Укажите название олимпиады.');
  var hit = null;
  if (award.id) rows_('awards').forEach(function (a) { if (String(a.id) === String(award.id)) hit = a; });
  if (hit) { update_('awards', hit._row, patch); patch.id = String(hit.id); }
  else { patch.id = newId_('a'); append_('awards', patch); }
  return { id: patch.id, u: patch.studentId, date: patch.date, title: patch.title,
           subject: patch.subject, level: patch.level, result: patch.result, note: patch.note };
};

API.deleteAward = function (token, id) {
  var p = auth_(token); requireStaff_(p);
  var hit = null;
  rows_('awards').forEach(function (a) { if (String(a.id) === String(id)) hit = a; });
  if (hit) drop_('awards', hit._row);
  return true;
};

/* ---- темы для изучения ---- */

var TOPIC_STATUS = ['todo', 'doing', 'done'];

API.saveTopic = function (token, topic) {
  var p = auth_(token); requireStaff_(p);
  topic = topic || {};
  checkStudent_(topic.u);
  var patch = {
    studentId: String(topic.u),
    subject: String(topic.subject || '').trim(),
    title: String(topic.title || '').trim(),
    status: TOPIC_STATUS.indexOf(topic.status) >= 0 ? topic.status : 'todo',
    due: dstr_(topic.due),
    note: String(topic.note || ''),
    updatedAt: nowIso_(), updatedBy: String(p.name)
  };
  if (!patch.title) throw new Error('Укажите тему.');
  var hit = null;
  if (topic.id) rows_('topics').forEach(function (t) { if (String(t.id) === String(topic.id)) hit = t; });
  if (hit) { update_('topics', hit._row, patch); patch.id = String(hit.id); }
  else { patch.id = newId_('t'); append_('topics', patch); }
  return { id: patch.id, u: patch.studentId, subject: patch.subject, title: patch.title,
           status: patch.status, due: patch.due, note: patch.note };
};

/** Одна и та же тема сразу нескольким ученикам (команде, роли, всем). */
API.saveTopics = function (token, topic, studentIds) {
  var p = auth_(token); requireStaff_(p);
  var out = [];
  (studentIds || []).forEach(function (u) {
    var one = {};
    for (var k in topic) if (topic.hasOwnProperty(k)) one[k] = topic[k];
    one.id = '';
    one.u = u;
    out.push(API.saveTopic(token, one));
  });
  return out;
};

/** Ученик может сам отметить тему изученной — остальное меняют только педагоги. */
API.setTopicStatus = function (token, id, status) {
  var p = auth_(token);
  if (TOPIC_STATUS.indexOf(status) < 0) throw new Error('Неизвестный статус темы.');
  var hit = null;
  rows_('topics').forEach(function (t) { if (String(t.id) === String(id)) hit = t; });
  if (!hit) throw new Error('Тема не найдена.');
  if (p.role === 'student' && String(hit.studentId) !== String(p.id)) throw new Error('Недостаточно прав.');
  update_('topics', hit._row, { status: status, updatedAt: nowIso_(), updatedBy: String(p.name) });
  return true;
};

API.deleteTopic = function (token, id) {
  var p = auth_(token); requireStaff_(p);
  var hit = null;
  rows_('topics').forEach(function (t) { if (String(t.id) === String(id)) hit = t; });
  if (hit) drop_('topics', hit._row);
  return true;
};

/* ---- настройки ---- */

API.saveSettings = function (token, patch) {
  var p = auth_(token); requireAdmin_(p);
  patch = patch || {};
  if (patch.groupName !== undefined) {
    writeSetting_('groupName', String(patch.groupName).trim() || 'Кружок');
    dropBootCache_();
  }
  if (patch.yearStart !== undefined) writeSetting_('yearStart', dstr_(patch.yearStart));
  if (patch.yearEnd !== undefined) writeSetting_('yearEnd', dstr_(patch.yearEnd));
  if (patch.schedule !== undefined) {
    var sc = patch.schedule;
    if (!Array.isArray(sc) || sc.length !== 7) throw new Error('Расписание — это 7 чисел (Вс…Сб).');
    writeSetting_('schedule', sc.map(function (h) { return Math.max(0, num_(h, 0)); }));
  }
  if (patch.competitions !== undefined) {
    writeSetting_('competitions', (patch.competitions || []).map(compPub_).filter(function (c) { return c.title; }));
  }
  if (patch.plan !== undefined) {
    writeSetting_('plan', (patch.plan || []).map(function (x) {
      return { month: String(x.month || ''), week: String(x.week || ''),
               n: String(x.n || ''), role: String(x.role || ''), task: String(x.task || '') };
    }).filter(function (x) { return x.task; }));
  }
  if (patch.siteUrl !== undefined) writeSetting_('siteUrl', String(patch.siteUrl).trim());
  if (patch.mailEnabled !== undefined) writeSetting_('mailEnabled', patch.mailEnabled ? '1' : '0');
  if (patch.mailHour !== undefined) {
    var h = Math.min(23, Math.max(0, num_(patch.mailHour, 19)));
    writeSetting_('mailHour', String(h));
  }
  if (patch.holidays !== undefined) {
    var hs = (patch.holidays || []).filter(function (h) { return h && dstr_(h.date); })
      .map(function (h) { return { date: dstr_(h.date), note: String(h.note || '') }; });
    writeSetting_('holidays', hs);
  }
  return readSettings_();
};

/*************************************************************
 *  ПИСЬМА РОДИТЕЛЯМ
 *
 *  Раз в день — одно письмо на родителя со сводкой за сегодня:
 *  пропустил, опоздал или отсутствовал по уважительной причине.
 *  Чтобы письма уходили сами, запустите один раз installDailyMail().
 *************************************************************/

var MAIL_WORDS = { a: 'пропустил занятие', l: 'опоздал', e: 'отсутствовал по уважительной причине' };

function dailyMailText_(date, items, settings) {
  var d = parseDate_(date);
  var lines = items.map(function (it) {
    return '• ' + it.name + ' — ' + (MAIL_WORDS[it.status] || it.status) +
      (it.hours ? ' (' + it.hours + ' ч)' : '');
  });
  return 'Здравствуйте!\n\n' +
    'Сводка за ' + date + ' по кружку «' + settings.groupName + '»:\n\n' +
    lines.join('\n') + '\n\n' +
    'Подробности — в личном кабинете родителя:\n' + (settings.siteUrl || '') + '\n\n' +
    'Если отсутствие по уважительной причине, сообщите руководителю кружка.\n\n' +
    'Это письмо отправлено автоматически, отвечать на него не нужно.';
}

function parseDate_(ds) { var p = String(ds).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }

/** Собрать и разослать сводку за день. Возвращает, что получилось. */
function sendDailyMail(dateStr) {
  var settings = readSettings_();
  if (String(settings.mailEnabled) !== '1') return 'Рассылка выключена в настройках сайта.';

  var date = dstr_(dateStr) || today_();
  var people = rows_('people').map(personPub_);
  var byId = {};
  people.forEach(function (p) { byId[p.id] = p; });

  /* что случилось сегодня */
  var todayMarks = {};
  rows_('marks').forEach(function (m) {
    if (dstr_(m.date) !== date) return;
    var st = String(m.status);
    if (st !== 'a' && st !== 'l' && st !== 'e') return;
    var plan = Number(settings.schedule[parseDate_(date).getDay()]) || 0;
    todayMarks[String(m.studentId)] = { status: st, hours: st === 'l' ? 0 : plan };
  });
  if (!Object.keys(todayMarks).length) return 'За ' + date + ' пропусков нет — писать не о чем.';

  var sent = 0, skipped = [];
  people.forEach(function (par) {
    if (par.role !== 'parent' || !par.active) return;
    if (!par.email) { skipped.push(par.name + ' (нет адреса)'); return; }
    var items = [];
    par.children.forEach(function (kid) {
      var m = todayMarks[kid];
      if (!m) return;
      items.push({ name: (byId[kid] || {}).name || kid, status: m.status, hours: m.hours });
    });
    if (!items.length) return;
    try {
      MailApp.sendEmail(par.email,
        'Кружок «' + settings.groupName + '»: сводка за ' + date,
        dailyMailText_(date, items, settings));
      sent++;
    } catch (e) {
      skipped.push(par.name + ' (' + e.message + ')');
    }
  });

  return 'Писем отправлено: ' + sent +
    (skipped.length ? '\nНе отправлено: ' + skipped.join(', ') : '');
}

/** Кнопка «Отправить сейчас» на сайте. */
API.sendMailNow = function (token, dateStr) {
  var p = auth_(token); requireAdmin_(p);
  return sendDailyMail(dateStr);
};

/** Запустить один раз: включает ежедневную отправку. */
function installDailyMail() {
  removeDailyMail();
  var hour = Math.min(23, Math.max(0, num_(readSettings_().mailHour, 19)));
  ScriptApp.newTrigger('sendDailyMail').timeBased().atHour(hour).everyDays(1).create();
  return tell_('Ежедневная рассылка включена: письма уходят около ' + hour + ':00.\n' +
    'Не забудьте включить её и в настройках сайта.');
}

function removeDailyMail() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyMail') ScriptApp.deleteTrigger(t);
  });
  return 'Ежедневная рассылка выключена.';
}

/* ================= точка входа ================= */

/** Сообщение пользователю: в журнал и, если можно, окном. */
var SILENT = false;
function tell_(msg) {
  Logger.log(msg);
  if (!SILENT) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) {} }
  return msg;
}

/* Эти команды только читают данные, поэтому выполняются без блокировки. */
var READ_ONLY = { boot: true, state: true };

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var out;
  try {
    var raw = (e && e.parameter && e.parameter.p) || '';
    if (!raw) return json_({ ok: true, version: VERSION,
      hello: 'Сервер посещаемости работает. Вставьте этот адрес в index.html.' });
    var p = JSON.parse(raw);
    var fn = p.fn, args = p.args || [];
    if (!API.hasOwnProperty(fn)) throw new Error('Неизвестная команда: ' + fn);
    if (READ_ONLY[fn]) {
      // чтение ничего не меняет — очередь не нужна, ответ приходит быстрее
      out = { ok: true, result: API[fn].apply(null, args) };
    } else {
      var lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try { out = { ok: true, result: API[fn].apply(null, args) }; }
      finally { lock.releaseLock(); }
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return json_(out);
}

function doPost(e) { return doGet(e); }

/* ================= первый запуск =================
 * Запустите функцию setup() один раз из редактора Apps Script.
 * Она создаст листы и выдаст PIN администратора.
 */
function setup() {
  Object.keys(SHEETS).forEach(function (n) { sheet_(n); });
  var s = readSettings_();
  writeSetting_('groupName', s.groupName);
  writeSetting_('schedule', s.schedule);
  writeSetting_('yearStart', s.yearStart);
  writeSetting_('yearEnd', s.yearEnd);
  writeSetting_('holidays', s.holidays);
  if (!rows_('settings').some(function (r) { return String(r.key) === 'competitions'; })) writeSetting_('competitions', s.competitions);
  if (!rows_('settings').some(function (r) { return String(r.key) === 'plan'; })) writeSetting_('plan', s.plan);

  var admins = rows_('people').filter(function (p) { return String(p.role) === 'admin'; });
  var msg;
  if (admins.length) {
    msg = 'Администратор уже есть: ' + admins[0].name +
          '.\nЗабыли код? Запустите функцию resetAdminPin().';
  } else {
    var pin = genPin_(6);
    createPerson_('admin', FIRST_ADMIN_NAME, pin);
    msg = 'Готово!\n\nАдминистратор: ' + FIRST_ADMIN_NAME + '\nКод для входа: ' + pin +
          '\n\nЗапишите код — он больше нигде не показывается.';
  }
  return tell_(msg);
}

/*************************************************************
 *  Разовый импорт дат рождения из таблицы кружка.
 *  Запустите функцию importBirthdays() один раз.
 *  Имена сверяются без учёта ә/ғ/қ/ң/ө/ұ/ү/і, поэтому
 *  «Амангос» и «Аманғос» считаются одним человеком.
 *************************************************************/
var BIRTHDAYS = [
  ['Аманғос Омарәлім', '2012-03-24'],
  ['Базарбай Өміржан', '2011-12-15'],
  ['Жандосұлы Жансерік', '2012-02-20'],
  ['Хажиолиев Төлеген', '2012-02-17'],
  ['Ғабит Абубакр', '2012-01-21'],
  ['Али Ахмет', '2011-08-02'],
  ['Еркін Диас', '2011-04-25'],
  ['Мұрзақан Мадияр', '2011-07-13'],
  ['Нығмет Әбдірасул', '2011-06-20'],
  ['Нурмухамедов Асмир', '2010-09-27'],
  ['Бисенгалиев Ислам', '2011-07-17'],
  ['Таяу Сүлеймен', '2011-04-18'],
  ['Тулегенов Расул', '2010-11-12'],
  ['Қазанбаев Сүлеймен', '2009-12-16'],
  ['Теңелғали Бекарыстан', '2009-08-27'],
  ['Давлетов Есентай', '2010-02-15'],
  ['Қыдырбай Нұрислам', '2009-11-01'],
  ['Нұрмұхан Нұрали', '2009-08-21'],
  ['Барлық Исламбек', '2009-10-04'],
  ['Суйесинов Алмаз', '2010-04-03'],
];

function plainName_(s) {
  s = String(s || '').toLowerCase();
  var from = 'әғқңөұүһіё', to = 'агкноуухие';
  var out = '';
  for (var i = 0; i < s.length; i++) {
    var j = from.indexOf(s.charAt(i));
    out += j >= 0 ? to.charAt(j) : s.charAt(i);
  }
  return out.replace(/\s+/g, ' ').trim();
}

function importBirthdays() {
  var byName = {};
  rows_('people').forEach(function (p) {
    if (String(p.role) === 'student') byName[plainName_(p.name)] = p;
  });
  var done = 0, missed = [];
  BIRTHDAYS.forEach(function (row) {
    var p = byName[plainName_(row[0])];
    if (!p) { missed.push(row[0]); return; }
    update_('people', p._row, { birth: row[1] });
    done++;
  });
  return tell_('Даты рождения проставлены: ' + done +
    (missed.length ? '\nНе нашлись: ' + missed.join(', ') : ''));
}

/*************************************************************
 *  Разовый импорт достижений школы (командные, 2024-2026).
 *  Запустите функцию importSchoolAwards() один раз.
 *************************************************************/
var SCHOOL_AWARDS = [
  ['2024', 'First Lego League', 'Aktobe First regional', 'White Hill', 'Robot design award finalist', 'honor'],
  ['2024', 'First Lego League', 'Aktobe First regional', 'Gambit', 'Robot design award winner', 'honor'],
  ['2025', 'First Lego League', 'Batys First Championship', 'Gambit', 'Robot design award winner', 'honor'],
  ['2025', 'First Lego League', 'Almaty Tech Cup (CAFC)', 'Gambit', 'Mechanical innovation award', 'honor'],
  ['2025', 'Fibonacci', 'FootBot League', 'Robotics team', '2nd place', 'silver'],
  ['2025', 'Fibonacci', 'FootBot League', 'Robotics team', '2nd place', 'silver'],
  ['2025', 'Infomatrix ASIA', 'Lego race', 'Robotics team', '1st place', 'gold'],
  ['2025', 'Infomatrix ASIA', 'Lego race', 'Robotics team', '2nd place', 'silver'],
  ['2025', 'WRO', 'Regional competition', 'Robotics team', '2nd place', 'silver'],
  ['2025', 'WRO', 'Regional competition', 'Robotics team', '3rd place', 'bronze'],
  ['2025', 'First Lego League', 'Aktobe First regional', 'Aqsonix', 'Core values award finalist', 'honor'],
  ['2025', 'First Lego League', 'Aktobe First regional', 'AqBIL_Innovators', 'Robot performance award winner', 'honor'],
  ['2025', 'First Lego League', 'Aktobe First regional', 'Gambit jr', 'Champion\'s award winner', 'honor'],
  ['2025', 'First Lego League', 'Bishkek First regional', 'AqBIL_Innovators', 'Robot performance award winner', 'honor'],
  ['2026', 'First Lego League', 'Batys first championship', 'Gambit jr', 'Engineering excellence award', 'honor'],
  ['2026', 'First Lego League', 'Batys first championship', 'Aqsonix', 'Engineering excellence award', 'honor'],
  ['2026', 'First Tech Challenge', 'Batys first championship', 'Gambit', 'Sustain award winner', 'honor'],
  ['2026', 'Fibonacci', 'FootBot lego 9-12', 'AqBIL_Innovators', 'Judge awards', 'honor'],
  ['2026', 'Fibonacci', 'Maze solving 9-12', 'AqBIL_Innovators', 'Judge awards', 'honor'],
  ['2026', 'Fibonacci', 'Line following 9-12', 'AqBIL_Innovators', 'Judge awards', 'honor'],
  ['2026', 'Fibonacci', 'Maze solving 5-8', 'Aqsonix', 'Judge awards', 'honor'],
  ['2026', 'Fibonacci', 'Autonomous car 5-8', 'Aqsonix', '3rd place', 'bronze'],
  ['2026', 'First Lego League', 'Astana CAFC', 'Aqsonix', 'Technical excellence award', 'honor'],
  ['2026', 'First Lego League', 'Astana CAFC', 'BILinnovators', 'Young innovators award', 'honor'],
  ['2026', 'First Lego League', 'Astana CAFC', 'Gambit jr', 'Engineering excellence award (quota to Bulgaria)', 'honor'],
  ['2026', 'First Tech Challenge', 'Astana CAFC', 'Gambit', 'First premier event (quota to America)', 'honor'],
  ['2026', 'Infomatrix ASIA', 'AI Hackathon', 'Gambit', '1st place', 'gold'],
  ['2026', 'Infomatrix ASIA', 'Arduino Hackaton', 'Aqsonix', '3rd place', 'bronze'],
  ['2026', 'Fibonacci Eurasian Championship', 'FootBot lego 5-8', 'Aqsonix', '3rd place', 'bronze'],
  ['2026', 'Fibonacci Eurasian Championship', 'Maze solving 5-8', 'Aqsonix', 'Judge award', 'honor'],
  ['2026', 'Fibonacci Eurasian Championship', 'Lego Line following 5-8', 'Aqsonix', '3rd place', 'bronze'],
  ['2026', 'Fibonacci Eurasian Championship', 'Autonomous car 5-8', 'Aqsonix', 'Judge award', 'honor'],
  ['2026', 'Fibonacci Eurasian Championship', 'Lego Line following 9-12', 'Aqsonix', '2nd place', 'silver'],
  ['2026', 'First Lego League', 'Bulgaria FLL', 'Gambit jr', '3rd place', 'bronze'],
  ['2026', 'WRO', 'Robo sport', 'Aqsonix', '2nd place', 'silver'],
  ['2026', 'WRO', 'Robomission senior', 'AqBIL_Innovators', '2nd place', 'silver'],
  ['2026', 'WRO', 'Robomission senior', 'AqBIL_Innovators', '1st place', 'gold'],
  ['2026', 'First Tech Challenge', 'Western Edge', 'Gambit', 'Think award 2nd place', 'silver'],
  ['2026', 'WRO', 'Robomission senior (respa)', 'AqBIL_Innovators', 'participants', 'part'],
  ['2026', 'FTC', 'Almaty Tech Cup', 'Gambit', 'Judges choice award', 'honor'],
];

function importSchoolAwards() {
  var have = {};
  rows_('school').forEach(function (r) {
    have[[r.year, r.event, r.team, r.award].join('|')] = true;
  });
  var added = 0;
  SCHOOL_AWARDS.forEach(function (row) {
    var key = [row[0], row[2], row[3], row[4]].join('|');
    if (have[key]) return;
    have[key] = true;
    append_('school', {
      id: newId_('sa'), year: row[0], category: row[1], event: row[2],
      team: row[3], award: row[4], result: row[5], note: '',
      updatedAt: nowIso_(), updatedBy: 'Импорт'
    });
    added++;
  });
  return tell_('Достижений школы добавлено: ' + added + '\nВсего в таблице: ' + rows_('school').length);
}

/*************************************************************
 *  Если строки в лист people добавляли руками, у них нет id и кода —
 *  сайт такие строки не показывает. Эта функция их чинит:
 *  проставляет id, роль, дату зачисления и выдаёт коды для входа.
 *  Запускать по мере надобности, лишнего не трогает.
 *************************************************************/
function fixPeopleSheet() {
  var sh = sheet_('people');
  var vals = sh.getDataRange().getValues();
  var head = vals[0], col = {};
  head.forEach(function (h, i) { col[h] = i; });
  var fixed = [], renamedPins = [];

  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    var name = String(row[col.name] || '').trim();
    if (!name) continue;                       // пустая строка — пропускаем
    var changed = false;

    if (!String(row[col.id] || '').trim()) {
      row[col.id] = newId_(String(row[col.role]) === 'student' || !row[col.role] ? 's' : 'u');
      changed = true;
    }
    if (!String(row[col.role] || '').trim()) { row[col.role] = 'student'; changed = true; }
    if (String(row[col.active]) === '') { row[col.active] = 1; changed = true; }
    if (!String(row[col.since] || '').trim()) { row[col.since] = today_(); changed = true; }
    if (!String(row[col.createdAt] || '').trim()) { row[col.createdAt] = today_(); changed = true; }
    if (!String(row[col.pinHash] || '').trim() || !String(row[col.salt] || '').trim()) {
      var pin = genPin_(String(row[col.role]) === 'student' ? 4 : 6);
      var salt = rndStr_(12);
      row[col.salt] = salt;
      row[col.pinHash] = hashPin_(pin, salt);
      renamedPins.push({ name: name, pin: pin });
      changed = true;
    }
    if (changed) { sh.getRange(r + 1, 1, 1, head.length).setValues([row]); fixed.push(name); }
  }

  forget_('people');      // строки читались до правок — перечитаем заново
  dropBootCache_();
  var msg = fixed.length
    ? 'Починено строк: ' + fixed.length + '\n' + fixed.join(', ') +
      (renamedPins.length ? '\n\nВыданы коды:\n' +
        renamedPins.map(function (x) { return x.name + ' — ' + x.pin; }).join('\n') : '')
    : 'Всё в порядке, чинить нечего.';
  return tell_(msg);
}

/** Сброс кода администратора, если он потерян. */
function resetAdminPin() {
  var admins = rows_('people').filter(function (p) { return String(p.role) === 'admin'; });
  if (!admins.length) return setup();
  var pin = genPin_(6), salt = rndStr_(12);
  update_('people', admins[0]._row, { salt: salt, pinHash: hashPin_(pin, salt), failCount: 0, lockUntil: '', active: 1 });
  return tell_('Новый код для «' + admins[0].name + '»: ' + pin);
}

/*************************************************************
 *  ИМПОРТ ДАННЫХ КРУЖКА РОБОТОТЕХНИКИ
 *
 *  Данные взяты из файла robotics_club.xlsx (листы «spisok 26-27»,
 *  «Жарыстан орын алған оқушылар», «Robotics жарыстар 26-27», «Annual plan»).
 *  ИИН, даты рождения и номера документов НЕ переносились.
 *
 *  Как пользоваться: добавьте этот файл в тот же проект Apps Script,
 *  что и Code.gs, и один раз запустите функцию importClubData().
 *  Повторный запуск ничего не испортит: существующие записи не дублируются.
 *************************************************************/

var IMPORT_SINCE = '2026-09-01';        // с какой даты считать посещаемость текущего состава
var IMPORT_WRITE_PIN_SHEET = true;   // выписать коды учеников на отдельный лист

/* имя, класс, команда, категория, в составе сейчас (1) или в архиве (0) */
var IMPORT_STUDENTS = [
  ['Базарбай Өміржан', '9А', 'Aqsonix', 'FLL', 1],
  ['Ғабит Абубакр', '9А', 'Aqsonix', 'FLL', 1],
  ['Танкиев Ескендір', '9А', 'Aqsonix', 'FLL', 1],
  ['Хажиолиев Төлеген', '9А', 'Aqsonix', 'FLL', 1],
  ['Турдыбеков Алихан', '9А', 'Aqsonix', 'FLL', 1],
  ['Ахмет Султан', '9Б', '', 'FLL', 1],
  ['Елеусіз Бекарыс', '9Б', '', 'FLL', 1],
  ['Азмуханов Алим', '10А', '', '', 1],
  ['Али Ахмет', '10А', '', 'FTC', 1],
  ['Бисенгалиев Ислам', '10А', '', 'FTC', 1],
  ['Мұрзақан Мадияр', '10А', '', 'FTC', 1],
  ['Нығмет Әбдірасул', '10А', '', 'FTC', 1],
  ['Әменов Омар', '10А', '', 'FTC', 1],
  ['Тлеумбетов Ерамир', '10А', '', 'FTC', 1],
  ['Тулегенов Расул', '10Б', '', 'FTC', 1],
  ['Таяу Сүлеймен', '10Б', '', 'FTC', 1],
  ['Майор Әсет', '10Б', '', 'FTC', 1],
  ['Нагашбаев Амир', '10Б', '', 'FTC', 1],
  ['Ақылбек Сардар', '9А', 'GAMBIT', 'FTC', 1],
  ['Аманғос Омарәлім', '9А', 'GAMBIT', 'FTC', 1],
  ['Еркін Диас', '10А', 'GAMBIT', 'FTC', 1],
  ['Саматов Ерасыл', '10Б', 'GAMBIT', 'FTC', 1],
  ['Есберген Қайсар', '10Б', 'GAMBIT', 'FTC', 1],
  ['Қазанбаев Сүлеймен', '11А', 'GAMBIT', 'FTC', 1],
  ['Қыдырбай Нұрислам', '11А', 'GAMBIT', 'FTC', 1],
  ['Мухтубаев Али', '11А', 'GAMBIT', 'FTC', 1],
  ['Санжар Жұмабек', '11А', 'GAMBIT', 'FTC', 1],
  ['Рат Мадияр', '11А', '', 'FTC', 1],
  ['Барлық Исламбек', '', '', '', 1],
  ['Алпысбай Аслан', '11А', 'GAMBIT', '', 0],
  ['Аманбай Исабек', '11А', 'GAMBIT', '', 0],
  ['Давлетов Есентай', '10А', 'GAMBIT jr', '', 0],
  ['Жалғасбай Әділ', '10Б', 'GAMBIT', '', 0],
  ['Жандосұлы Жансерік', '8A', 'Aqsonix', '', 0],
  ['Наурызғали Бекнар', '11А', 'GAMBIT', '', 0],
  ['Нурмухамедов Асмир', '9A', 'BIL innovators', '', 0],
  ['Нұрмұхан Нұрали', '10А', 'GAMBIT jr', '', 0],
  ['Оспанов Даурен', '11А', 'GAMBIT', '', 0],
  ['Сағынғали Абдуссаттар', '10Б', 'GAMBIT', '', 0],
  ['Серіков Темірлан', '8A', 'Aqsonix', '', 0],
  ['Серікқали Ғани', '11А', 'GAMBIT', '', 0],
  ['Суйесинов Алмаз', '10А', 'GAMBIT jr', '', 0],
  ['Теңелғали Бекарыстан', '10А', 'GAMBIT jr', '', 0],
  ['Төребай Қасымхан', '11А', 'GAMBIT', '', 0],
  ['Умаров Қайсар', '11А', 'GAMBIT', '', 0],
  ['Шүйншбай Ақтілек', '10А', 'GAMBIT', '', 0],
];

/* имя, соревнование, уровень, результат, как записано в таблице */
var IMPORT_AWARDS = [
  ['Ақылбек Сардар', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Ақылбек Сардар', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Ақылбек Сардар', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Ақылбек Сардар', 'INFOMATRIX', 'international', 'bronze', '3rd place'],
  ['Ақылбек Сардар', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Ақылбек Сардар', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Аманғос Омарәлім', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Аманғос Омарәлім', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Аманғос Омарәлім', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Аманғос Омарәлім', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Аманғос Омарәлім', 'INFOMATRIX', 'international', 'bronze', '3rd place'],
  ['Аманғос Омарәлім', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Аманғос Омарәлім', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Базарбай Өміржан', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Базарбай Өміржан', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Базарбай Өміржан', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Базарбай Өміржан', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Базарбай Өміржан', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Базарбай Өміржан', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Ғабит Абубакр', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Ғабит Абубакр', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Ғабит Абубакр', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Ғабит Абубакр', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Ғабит Абубакр', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Жандосұлы Жансерік', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Жандосұлы Жансерік', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Жандосұлы Жансерік', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Жандосұлы Жансерік', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Жандосұлы Жансерік', 'INFOMATRIX', 'international', 'part', 'participant'],
  ['Серіков Темірлан', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Серіков Темірлан', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Танкиев Ескендір', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Танкиев Ескендір', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Танкиев Ескендір', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Танкиев Ескендір', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Хажиолиев Төлеген', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Хажиолиев Төлеген', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Хажиолиев Төлеген', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Хажиолиев Төлеген', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Хажиолиев Төлеген', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Турдыбеков Алихан', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Турдыбеков Алихан', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Турдыбеков Алихан', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Ахмет Султан', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Елеусіз Бекарыс', 'Fibonacci Turkey', 'international', 'silver', '2nd place'],
  ['Али Ахмет', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Али Ахмет', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Али Ахмет', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Али Ахмет', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Али Ахмет', 'WRO', 'international', 'gold', '1st place'],
  ['Еркін Диас', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Еркін Диас', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Еркін Диас', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Еркін Диас', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Еркін Диас', 'WRO', 'international', 'gold', '1st place'],
  ['Бисенгалиев Ислам', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Бисенгалиев Ислам', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Бисенгалиев Ислам', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Бисенгалиев Ислам', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Бисенгалиев Ислам', 'WRO', 'international', 'honor', '4th place'],
  ['Мұрзақан Мадияр', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Мұрзақан Мадияр', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Мұрзақан Мадияр', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Мұрзақан Мадияр', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Мұрзақан Мадияр', 'WRO', 'international', 'honor', '4th place'],
  ['Нығмет Әбдірасул', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Нығмет Әбдірасул', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Нығмет Әбдірасул', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Нығмет Әбдірасул', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Нығмет Әбдірасул', 'WRO', 'international', 'honor', '4th place'],
  ['Нурмухамедов Асмир', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Нурмухамедов Асмир', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Нурмухамедов Асмир', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Нурмухамедов Асмир', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Әменов Омар', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Әменов Омар', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Әменов Омар', 'WRO', 'international', 'gold', '1st place'],
  ['Таяу Сүлеймен', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Таяу Сүлеймен', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Таяу Сүлеймен', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Таяу Сүлеймен', 'WRO', 'international', 'silver', '2nd place'],
  ['Тулегенов Расул', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Тулегенов Расул', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Тлеумбетов Ерамир', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Тлеумбетов Ерамир', 'WRO', 'international', 'honor', '4th place'],
  ['Саматов Ерасыл', 'WRO', 'international', 'honor', '4th place'],
  ['Есберген Қайсар', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Есберген Қайсар', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Есберген Қайсар', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Есберген Қайсар', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Есберген Қайсар', 'WRO', 'international', 'silver', '2nd place'],
  ['Барлық Исламбек', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Барлық Исламбек', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Барлық Исламбек', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Барлық Исламбек', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Давлетов Есентай', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Давлетов Есентай', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Давлетов Есентай', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Давлетов Есентай', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Қазанбаев Сүлеймен', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Қазанбаев Сүлеймен', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Қазанбаев Сүлеймен', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Қыдырбай Нұрислам', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Қыдырбай Нұрислам', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Қыдырбай Нұрислам', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Қыдырбай Нұрислам', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Нұрмұхан Нұрали', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Нұрмұхан Нұрали', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Нұрмұхан Нұрали', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Суйесинов Алмаз', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Суйесинов Алмаз', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Суйесинов Алмаз', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Теңелғали Бекарыстан', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Теңелғали Бекарыстан', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Теңелғали Бекарыстан', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Рат Мадияр', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Рат Мадияр', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Рат Мадияр', 'INFOMATRIX', 'international', 'gold', '1st place'],
  ['Шүйншбай Ақтілек', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Мухтубаев Али', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Мухтубаев Али', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Мухтубаев Али', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Сағынғали Абдуссаттар', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Сағынғали Абдуссаттар', 'INFOMATRIX', 'international', 'gold', '1st place'],
  ['Жалғасбай Әділ', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Жалғасбай Әділ', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Жалғасбай Әділ', 'INFOMATRIX', 'international', 'gold', '1st place'],
  ['Наурызғали Бекнар', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Наурызғали Бекнар', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Аманбай Исабек', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Аманбай Исабек', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Оспанов Даурен', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Оспанов Даурен', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Серікқали Ғани', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Серікқали Ғани', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Умаров Қайсар', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Умаров Қайсар', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Төребай Қасымхан', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Төребай Қасымхан', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Алпысбай Аслан', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
];

/* соревнование, когда, четверть, статус */
var IMPORT_COMPETITIONS = [
  ['Облыстық жоба тақырып ұсыну', '20-25 август', '1 term', 'Done'],
  ['Облыстық жоба жарысы', '10-11 ноябрь', '2 term', ''],
  ['Республикалық жоба жарысы', '25-27 февраль', '3 term', ''],
  ['Fibonacci Italy', 'November 27-5', '2 term', 'Loading'],
  ['First Aktobe Regional (FLL, FTC)', 'December 12-13', '2 term', ''],
  ['Batys First Aktobe (FLL, FTC)', 'January 16-17', '3 term', ''],
  ['Fibonacci Astana', '', '3 term', ''],
  ['Astana CAFC (FLL, FTC)', '', '3 term', ''],
  ['INFOMATRIX', 'March', '3 term', ''],
  ['World robotics olympiad (WRO)', 'April - May', '4 term', ''],
  ['VEX', '', '', ''],
  ['SeaPerch', '', '', ''],
  ['Drone симулятор', '', '', ''],
  ['Startup projects', '', '', ''],
  ['STEAM Baku', '', '', ''],
  ['Летный академия', '', '', ''],
];

/* месяц, неделя, номер недели, роль, задача */
var IMPORT_PLAN = [
  ['SEPTEMBER', 'I', '1', 'Команда жұмысы', 'Introduction, тренермен танысу'],
  ['SEPTEMBER', 'I', '1', 'Captain', 'Дайындыққа оқушылардың келгенін қадағалау'],
  ['SEPTEMBER', 'I', '1', 'Manager', 'Инста аккаунт ашылғанын күту'],
  ['SEPTEMBER', 'I', '1', 'Builder', 'pinterest/youtube idea izdeu'],
  ['SEPTEMBER', 'I', '1', 'Programmer', 'prime lessons'],
  ['SEPTEMBER', 'I', '1', 'SMM', 'Инстаграм аккаунт ашу'],
  ['SEPTEMBER', 'I', '1', 'Slider/Layer', 'Canva pro aккаунт ашу'],
  ['SEPTEMBER', 'II', '2', 'Команда жұмысы', 'Картамен танысу'],
  ['SEPTEMBER', 'II', '2', 'Captain', 'Қадағалау'],
  ['SEPTEMBER', 'II', '2', 'Manager', 'мектепте, сыныпта аккаут актуалность жинау'],
  ['SEPTEMBER', 'II', '2', 'Builder', 'Миссия қанша очко правило оқу'],
  ['SEPTEMBER', 'II', '2', 'Programmer', 'prime lessons'],
  ['SEPTEMBER', 'II', '2', 'SMM', 'Лого дайындау'],
  ['SEPTEMBER', 'II', '2', 'Slider/Layer', 'Canva сабақтарын қарау'],
  ['SEPTEMBER', 'III', '3', 'Команда жұмысы', 'Инновационный проект тақырыбымен танысу'],
  ['SEPTEMBER', 'III', '3', 'Captain', 'Қадағалау'],
  ['SEPTEMBER', 'III', '3', 'Manager', 'ағылшынша speaking дамыту'],
  ['SEPTEMBER', 'III', '3', 'Builder', 'robot building idea tauip құрастырып бастау'],
  ['SEPTEMBER', 'III', '3', 'Programmer', 'youtube'],
  ['SEPTEMBER', 'III', '3', 'SMM', 'Сторис салу'],
  ['SEPTEMBER', 'III', '3', 'Slider/Layer', 'Canva сабақтарын қарау'],
  ['SEPTEMBER', 'IV', '4', 'Команда жұмысы', 'Инновационный проект тақырыбын зерттеу'],
  ['SEPTEMBER', 'IV', '4', 'Captain', 'Қадағалау'],
  ['SEPTEMBER', 'IV', '4', 'Manager', 'FLL Қазақстан бойынша қанша команда бар білу'],
  ['SEPTEMBER', 'IV', '4', 'Builder', 'Насадка ойлап табу'],
  ['SEPTEMBER', 'IV', '4', 'Programmer', 'youtube'],
  ['SEPTEMBER', 'IV', '4', 'SMM', '2 пост салу'],
  ['SEPTEMBER', 'IV', '4', 'Slider/Layer', 'Canva сабақтарын қарау'],
  ['OCTOBER', 'I', '5', 'Команда жұмысы', 'Команданы First тіркеу'],
  ['OCTOBER', 'I', '5', 'Captain', 'Команданы First тіркеу'],
  ['OCTOBER', 'I', '5', 'Manager', '1 митинг жасау'],
  ['OCTOBER', 'I', '5', 'Builder', 'Mission 1-2-3'],
  ['OCTOBER', 'I', '5', 'Programmer', 'Mission 1-2-3'],
  ['OCTOBER', 'I', '5', 'SMM', '2 пост салу'],
  ['OCTOBER', 'I', '5', 'Slider/Layer', 'Power point сабақтарын қарау'],
  ['OCTOBER', 'II', '6', 'Команда жұмысы', 'Активити жасау'],
  ['OCTOBER', 'II', '6', 'Captain', 'Дұрыс орындалмай жатқан жұмысты өзі жасау'],
  ['OCTOBER', 'II', '6', 'Manager', '2 митинг Қазақстандағы мықты командалармен'],
  ['OCTOBER', 'II', '6', 'Builder', 'points көбейту'],
  ['OCTOBER', 'II', '6', 'Programmer', 'points көбейту'],
  ['OCTOBER', 'II', '6', 'SMM', 'Команда таныстыру видео салу'],
  ['OCTOBER', 'II', '6', 'Slider/Layer', 'Power point сабақтарын қарау'],
  ['OCTOBER', 'III', '7', 'Команда жұмысы', 'Robot game уақытқа жасау'],
  ['OCTOBER', 'III', '7', 'Captain', 'Активити жасау'],
  ['OCTOBER', 'III', '7', 'Manager', 'Археолог адаммен сөйлесу'],
  ['OCTOBER', 'III', '7', 'Builder', '100 points жеткізу'],
  ['OCTOBER', 'III', '7', 'Programmer', '100 points жеткізу'],
  ['OCTOBER', 'III', '7', 'SMM', '150 подписчик жинау'],
  ['OCTOBER', 'III', '7', 'Slider/Layer', 'Power point сабақтарын қарау'],
  ['OCTOBER', 'IV', '8', 'Команда жұмысы', 'Жарысқа тіркелу'],
  ['OCTOBER', 'IV', '8', 'Captain', 'Активити жасау'],
  ['OCTOBER', 'IV', '8', 'Manager', 'Музейге бару соны ұйымдастыру'],
  ['OCTOBER', 'IV', '8', 'Builder', '120 points жеткізу'],
  ['OCTOBER', 'IV', '8', 'Programmer', '120 points жеткізу'],
  ['OCTOBER', 'IV', '8', 'SMM', '2 пост 150 лайк жинау'],
  ['OCTOBER', 'IV', '8', 'Slider/Layer', 'Инновационный проект слайд бітіру'],
  ['NOVEMBER', 'I', '9', 'Команда жұмысы', 'Балалардың спич тексеру'],
  ['NOVEMBER', 'I', '9', 'Captain', 'Активити жасау'],
  ['NOVEMBER', 'I', '9', 'Manager', 'Мектепке облысына команданы реклама жасау'],
  ['NOVEMBER', 'I', '9', 'Builder', '200 points жеткізу'],
  ['NOVEMBER', 'I', '9', 'Programmer', '200 points жеткізу'],
  ['NOVEMBER', 'I', '9', 'SMM', 'Әр пост 100 лайктан аз болмау'],
  ['NOVEMBER', 'I', '9', 'Slider/Layer', 'Робот дизайн слайд жасау'],
  ['NOVEMBER', 'II', '10', 'Команда жұмысы', 'Жарыс алдындағы дайындық'],
  ['NOVEMBER', 'II', '10', 'Captain', 'Общий команда Robot game дайындық'],
  ['NOVEMBER', 'II', '10', 'Manager', 'Мектепке Ақтөбе облысына команданы реклама жасау'],
  ['NOVEMBER', 'II', '10', 'Builder', '250 points жеткізу'],
  ['NOVEMBER', 'II', '10', 'Programmer', '250 points жеткізу'],
  ['NOVEMBER', 'II', '10', 'SMM', 'Сторис салу'],
  ['NOVEMBER', 'II', '10', 'Slider/Layer', 'Журнал толтыру'],
  ['NOVEMBER', 'III', '11', 'Команда жұмысы', '2 video call, 2 рет музейге бару, Robot game 265 points, Инновационный проект макет дайын болу керек'],
  ['NOVEMBER', 'III', '11', 'Manager', 'Мектепке Ақтөбе облысына команданы реклама жасау'],
  ['NOVEMBER', 'III', '11', 'Builder', '265 points жеткізу'],
  ['NOVEMBER', 'III', '11', 'Programmer', '265 points жеткізу'],
  ['NOVEMBER', 'III', '11', 'SMM', 'Сторис салу'],
  ['NOVEMBER', 'III', '11', 'Slider/Layer', 'Бос адамдармен макет жасау'],
  ['NOVEMBER', 'IV', '12', 'Команда жұмысы', '3 video call, 2 рет музейге бару, Robot game 265 points, Инновационный проект макет, robot design slide дайын болу керек'],
  ['NOVEMBER', 'IV', '12', 'Manager', 'FLL November region jarys'],
  ['NOVEMBER', 'IV', '12', 'Builder', 'Проверьте робота на поле 10+ раз подряд на стабильность.'],
  ['NOVEMBER', 'IV', '12', 'SMM', 'Подготовьте единый стиль команды (футболки, логотип, плакат).'],
  ['DECEMBER', 'I', '13', 'Команда жұмысы', 'Спикерлердің сөйлеу мәнерін (intonation, gestures, smile) жаттықтыру.)'],
  ['DECEMBER', 'I', '13', 'Manager', 'Инновацияны жергілікті мысалмен байланыстыру (мысалы: Kazakhstan minerals, Smart irrigation system).'],
  ['DECEMBER', 'I', '13', 'Programmer', 'Роботтың қозғалысын тұрақты ету (PID calibration, gyro turning).'],
  ['DECEMBER', 'I', '13', 'Slider/Layer', 'PowerPoint-та дизайнды STEM стилінде әдемі қыл.'],
  ['DECEMBER', 'II', '14', 'Команда жұмысы', 'Команда ұраны ойлап тап: “Dig Deep. Think Smart. Build the Future!”'],
  ['DECEMBER', 'II', '14', 'Captain', 'Футболка / худи дизайнын жаса (логотип, ұран, түстер).'],
  ['DECEMBER', 'II', '14', 'Manager', '3D модель немесе слайд-анимация жасап, презентацияға қос.'],
  ['DECEMBER', 'II', '14', 'Builder', 'Миссияларды топтастыру (бір сапарда 2–3 миссия орындау).'],
  ['DECEMBER', 'II', '14', 'Programmer', 'FLL SPIKE Prime Gyro Alignment'],
  ['DECEMBER', 'II', '14', 'SMM', '📸 Инстаграм парақшасын белсенді жүргіз: Story, reels, robot progress, team spirit.'],
  ['DECEMBER', 'III', '15', 'Manager', 'Нағыз экспертпен (археолог, инженер, эколог) сұхбаттас — жюриге ұнайды.'],
  ['DECEMBER', 'III', '15', 'Builder', 'Күн сайын 15–20 минут “code + test + fix”'],
  ['DECEMBER', 'III', '15', 'SMM', 'Қысқа видео-интервью түсір (әр қатысушы өз рөлін таныстырады).'],
];

function normName_(s) {
  return String(s || '').toLowerCase().replace(/\u0451/g, '\u0435').replace(/\s+/g, ' ').trim();
}

function importClubData() {
  var existing = {}, byId = {};
  rows_('people').forEach(function (p) {
    if (String(p.role) === 'student') existing[normName_(p.name)] = p;
    byId[String(p.id)] = p;
  });

  var created = [], updated = 0;
  IMPORT_STUDENTS.forEach(function (row) {
    var name = row[0], extra = { cls: row[1], team: row[2], category: row[3], active: row[4] ? 1 : 0 };
    if (!row[4]) extra.note = 'Из прошлогоднего состава';
    var hit = existing[normName_(name)];
    if (hit) {
      update_('people', hit._row, {
        cls: extra.cls, team: extra.team, category: extra.category,
        active: extra.active, note: hit.note || extra.note || ''
      });
      updated++;
      return;
    }
    var pin = genPin_(4);
    var p = createPerson_('student', name, pin, IMPORT_SINCE, extra);
    existing[normName_(name)] = p;
    created.push({ name: name, pin: pin, cls: extra.cls, team: extra.team });
  });

  /* --- олимпиады и награды --- */
  var have = {};
  rows_('awards').forEach(function (a) { have[String(a.studentId) + '|' + String(a.title)] = true; });
  var addedAwards = 0, skipped = [];
  IMPORT_AWARDS.forEach(function (row) {
    var p = existing[normName_(row[0])];
    if (!p) { skipped.push(row[0]); return; }
    var key = String(p.id) + '|' + row[1];
    if (have[key]) return;
    have[key] = true;
    append_('awards', {
      id: newId_('a'), studentId: String(p.id), date: '', title: row[1],
      subject: '\u0420\u043e\u0431\u043e\u0442\u043e\u0442\u0435\u0445\u043d\u0438\u043a\u0430',
      level: row[2], result: row[3], note: row[4],
      updatedAt: nowIso_(), updatedBy: '\u0418\u043c\u043f\u043e\u0440\u0442'
    });
    addedAwards++;
  });

  /* --- календарь соревнований и годовой план --- */
  writeSetting_('competitions', IMPORT_COMPETITIONS.map(function (c) {
    return { title: c[0], when: c[1], term: c[2], status: c[3] };
  }));
  writeSetting_('plan', IMPORT_PLAN.map(function (x) {
    return { month: x[0], week: x[1], n: x[2], role: x[3], task: x[4] };
  }));

  /* --- лист с кодами --- */
  if (IMPORT_WRITE_PIN_SHEET && created.length) {
    var bk = book_();
    var sh = bk.getSheetByName('\u041a\u043e\u0434\u044b') || bk.insertSheet('\u041a\u043e\u0434\u044b');
    sh.clear();
    sh.appendRow(['\u0423\u0447\u0435\u043d\u0438\u043a', '\u041a\u043b\u0430\u0441\u0441', '\u041a\u043e\u043c\u0430\u043d\u0434\u0430', '\u041a\u043e\u0434']);
    created.forEach(function (c) { sh.appendRow([c.name, c.cls, c.team, c.pin]); });
    sh.getRange('D:D').setNumberFormat('@');
  }

  var msg = '\u0418\u043c\u043f\u043e\u0440\u0442 \u0437\u0430\u0432\u0435\u0440\u0448\u0451\u043d.\n\n' +
    '\u0414\u043e\u0431\u0430\u0432\u043b\u0435\u043d\u043e \u0443\u0447\u0435\u043d\u0438\u043a\u043e\u0432: ' + created.length + '\n' +
    '\u041e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u043e: ' + updated + '\n' +
    '\u0417\u0430\u043f\u0438\u0441\u0435\u0439 \u043e \u043d\u0430\u0433\u0440\u0430\u0434\u0430\u0445: ' + addedAwards + '\n' +
    '\u0421\u043e\u0440\u0435\u0432\u043d\u043e\u0432\u0430\u043d\u0438\u0439 \u0432 \u043a\u0430\u043b\u0435\u043d\u0434\u0430\u0440\u0435: ' + IMPORT_COMPETITIONS.length + '\n' +
    '\u041f\u0443\u043d\u043a\u0442\u043e\u0432 \u0433\u043e\u0434\u043e\u0432\u043e\u0433\u043e \u043f\u043b\u0430\u043d\u0430: ' + IMPORT_PLAN.length +
    (created.length ? '\n\n\u041a\u043e\u0434\u044b \u0443\u0447\u0435\u043d\u0438\u043a\u043e\u0432 \u2014 \u043d\u0430 \u043b\u0438\u0441\u0442\u0435 \u00ab\u041a\u043e\u0434\u044b\u00bb. \u0420\u0430\u0437\u0434\u0430\u0439\u0442\u0435 \u0438\u0445 \u0438 \u0443\u0434\u0430\u043b\u0438\u0442\u0435 \u043b\u0438\u0441\u0442.' : '');
  return tell_(msg);
}

/*************************************************************
 *  Запустите эту функцию один раз — она сделает всё сразу:
 *  создаст листы, заведёт администратора и перенесёт данные кружка.
 *************************************************************/
function setupAll() {
  SILENT = true;
  var a, b;
  try {
    a = setup();
    b = importClubData();
  } finally {
    SILENT = false;
  }
  return tell_(a + '\n\n— — —\n\n' + b);
}
