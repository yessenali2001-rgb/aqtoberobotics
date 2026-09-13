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

var VERSION = '2026-09-13-2';         // метка версии кода: видна в ответе сервера
var SHEET_ID = '';                    // пусто = скрипт привязан к таблице
var FIRST_ADMIN_NAME = 'Администратор';
var SESSION_DAYS = 90;                // сколько дней держится вход
var MAX_FAILS = 7;                    // неверных PIN подряд до блокировки
var LOCK_MINUTES = 15;                // на сколько блокируется вход

var SHEETS = {
  people:   ['id','role','name','pinHash','salt','active','since','createdAt','failCount','lockUntil','note',
             'cls','team','category','teamRole','birth'],
  marks:    ['id','date','studentId','status','hours','note','updatedAt','updatedBy'],
  settings: ['key','value'],
  sessions: ['token','personId','createdAt','expiresAt'],
  awards:   ['id','studentId','date','title','subject','level','result','note','updatedAt','updatedBy'],
  topics:   ['id','studentId','subject','title','status','due','note','updatedAt','updatedBy']
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
    since: dstr_(p.since), createdAt: dstr_(p.createdAt), note: String(p.note || ''),
    cls: String(p.cls || ''), team: String(p.team || ''),
    category: String(p.category || ''), teamRole: String(p.teamRole || ''),
    birth: dstr_(p.birth)
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
    birth: dstr_(extra.birth)
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

function stateFor_(person) {
  var settings = readSettings_();
  var people = rows_('people').map(personPub_);
  var me = personPub_(person);
  if (person.role === 'student') {
    return {
      me: me, settings: settings,
      people: people.filter(function (p) { return p.id === me.id; }),
      marks: marksPub_(rows_('marks'), me.id),
      awards: awardsPub_(rows_('awards'), me.id),
      topics: topicsPub_(rows_('topics'), me.id),
      birthdays: people.filter(function (x) { return x.role === 'student' && x.active && x.birth; })
        .map(function (x) { return { name: x.name, birth: x.birth, team: x.team }; })
    };
  }
  return {
    me: me, settings: settings, people: people,
    marks: marksPub_(rows_('marks')),
    awards: awardsPub_(rows_('awards')),
    topics: topicsPub_(rows_('topics'))
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
  ['cls', 'team', 'category', 'teamRole'].forEach(function (k) {
    if (patch[k] !== undefined) out[k] = String(patch[k]).trim();
  });
  if (patch.active !== undefined || patch.role !== undefined) {
    requireAdmin_(p);
    if (patch.active !== undefined) out.active = patch.active ? 1 : 0;
    if (patch.role !== undefined) {
      if (['admin', 'assistant', 'student'].indexOf(patch.role) < 0) throw new Error('Неизвестная роль.');
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
  if (t.role !== 'student') requireAdmin_(p);
  var pin = genPin_(t.role === 'student' ? 4 : 6);
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
  if (patch.holidays !== undefined) {
    var hs = (patch.holidays || []).filter(function (h) { return h && dstr_(h.date); })
      .map(function (h) { return { date: dstr_(h.date), note: String(h.note || '') }; });
    writeSetting_('holidays', hs);
  }
  return readSettings_();
};

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

/** Сброс кода администратора, если он потерян. */
function resetAdminPin() {
  var admins = rows_('people').filter(function (p) { return String(p.role) === 'admin'; });
  if (!admins.length) return setup();
  var pin = genPin_(6), salt = rndStr_(12);
  update_('people', admins[0]._row, { salt: salt, pinHash: hashPin_(pin, salt), failCount: 0, lockUntil: '', active: 1 });
  return tell_('Новый код для «' + admins[0].name + '»: ' + pin);
}
