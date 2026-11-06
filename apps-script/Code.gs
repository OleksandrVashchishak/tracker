/**
 * Бекенд для застосунку «Калорії», версія 2.
 *
 * Що робить:
 *  - віддає всі дані (pull) і приймає пакети змін (push);
 *  - веде номер версії таблиці (rev): пристрій, який не бачив останньої версії, не може писати
 *    (він спершу підтягує нові дані);
 *  - не дозволяє заміну цілої вкладки без явного force і масові видалення (понад 20% рядків);
 *  - робить резервні копії таблиці у Google Drive (папка «Калорії: резервні копії»);
 *  - веде журнал змін у вкладці «Журнал».
 *
 * Не змінюйте порядок колонок у вкладках: скрипт читає їх за позицією.
 */

// У репозиторії тут заглушки. Справжні значення стоять лише в розгорнутому скрипті
// (Розширення → Apps Script у таблиці). Після вставки цього коду впишіть їх назад.
const TOKEN = 'ВСТАВТЕ_ТОКЕН';
const SPREADSHEET_ID = 'ВСТАВТЕ_ID_ТАБЛИЦІ';

const BACKUP_FOLDER = 'Калорії: резервні копії';
const JOURNAL = 'Журнал';
const TZ = 'Europe/Kyiv';
const KEEP_BACKUPS = 40;       // скільки копій тримати
const KEEP_JOURNAL = 1000;     // скільки рядків журналу тримати
const BACKUP_GAP_MS = 3 * 3600 * 1000;       // звичайна копія: не частіше, ніж раз на 3 години
const BACKUP_GAP_URGENT_MS = 15 * 60 * 1000; // перед видаленнями: не частіше, ніж раз на 15 хвилин

// Назва таблиці в застосунку -> вкладка, колонки за порядком, ключ
const TABLES = {
  entries:   { sheet: 'Записи',       cols: ['id', 'd', 'm', 'n', 'g', 'q', 'u', 'k'],                key: 'id' },
  products:  { sheet: 'Продукти',     cols: ['id', 'n', 'k', 'p', 'u', 'ug', 'c', 'r', 'w'],          key: 'id' },
  cats:      { sheet: 'Категорії',    cols: ['id', 'n', 'o'],                                         key: 'id' },
  templates: { sheet: 'Шаблони',      cols: ['id', 'n', 'm', 'items'],                                key: 'id' },
  weights:   { sheet: 'Вага',         cols: ['d', 'kg'],                                              key: 'd' },
  settings:  { sheet: 'Налаштування', cols: ['key', 'value', 'note'],                                 key: 'key' },
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function props_() { return PropertiesService.getScriptProperties(); }
function getRev_() { return Number(props_().getProperty('rev') || 0); }

function sheet_(name) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('Немає вкладки: ' + name);
  return sh;
}

// Рядки вкладки як масиви (без заголовка). Порожні клітинки -> null.
function readRows_(t) {
  const sh = sheet_(t.sheet), n = t.cols.length;
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, n).getValues()
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => r.map(v => (v === '' ? null : v)));
}

function toObj_(t, row) {
  const o = {};
  t.cols.forEach((c, i) => { o[c] = row[i]; });
  return o;
}

function toRow_(t, obj) {
  return t.cols.map(c => (obj[c] === undefined || obj[c] === null ? '' : obj[c]));
}

function writeRows_(t, rows) {
  const sh = sheet_(t.sheet), n = t.cols.length;
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, n).clearContent();
  const need = rows.length + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 200);
  if (rows.length) sh.getRange(2, 1, rows.length, n).setValues(rows);
}

// Резервна копія всієї таблиці у Drive. Повертає true, якщо копію створено.
function maybeBackup_(urgent) {
  const now = Date.now();
  const last = Number(props_().getProperty('lastBackup') || 0);
  if (now - last < (urgent ? BACKUP_GAP_URGENT_MS : BACKUP_GAP_MS)) return false;
  return backup_(now);
}

function backup_(now) {
  try {
    const file = DriveApp.getFileById(SPREADSHEET_ID);
    const it = DriveApp.getFoldersByName(BACKUP_FOLDER);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
    file.makeCopy('Калорії ' + Utilities.formatDate(new Date(now), TZ, 'yyyy-MM-dd HH:mm'), folder);
    const files = [];
    const fi = folder.getFiles();
    while (fi.hasNext()) files.push(fi.next());
    files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
    files.slice(KEEP_BACKUPS).forEach(f => f.setTrashed(true));
    props_().setProperty('lastBackup', String(now));
    props_().deleteProperty('backupError');
    return true;
  } catch (err) {
    props_().setProperty('backupError', String(err));
    return false;
  }
}

function log_(rev, body, plan, backed) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sh = ss.getSheetByName(JOURNAL);
    if (!sh) {
      sh = ss.insertSheet(JOURNAL);
      sh.appendRow(['Час', 'Версія', 'Пристрій', 'Застосунок', 'Зміни', 'Резервна копія']);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    const parts = Object.keys(plan)
      .filter(n => plan[n].added || plan[n].changed || plan[n].removed)
      .map(n => n + ': +' + plan[n].added + ' ~' + plan[n].changed + ' −' + plan[n].removed);
    sh.appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), rev, body.dev || '', body.v || '',
      parts.join('; ') + (body.force ? ' [force]' : ''), backed ? 'так' : '',
    ]);
    const last = sh.getLastRow();
    if (last > KEEP_JOURNAL + 1) sh.deleteRows(2, last - KEEP_JOURNAL - 1);
  } catch (err) { /* журнал не повинен ламати запис */ }
}

// Читання: ?action=pull&token=...   Перевірка: ?action=ping&token=...
function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    const p = (e && e.parameter) || {};
    if (p.token !== TOKEN) return json_({ ok: false, error: 'auth' });
    if (p.action === 'ping') return json_({ ok: true, ver: 2 });
    lock.waitLock(15000);   // читаємо, коли ніхто не пише, щоб не побачити половину запису
    const out = {};
    Object.keys(TABLES).forEach(name => {
      const t = TABLES[name];
      out[name] = readRows_(t).map(r => toObj_(t, r));
    });
    return json_({
      ok: true, ver: 2, tables: out, rev: getRev_(),
      lastBackup: Number(props_().getProperty('lastBackup') || 0),
      backupError: props_().getProperty('backupError') || '',
      time: new Date().toISOString(),
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Запис: тіло (text/plain) = JSON { token, ops: [...], baseRev, force, dev, v }
 *   { t: 'entries', op: 'set', row: {...} }        додати або замінити рядок за ключем
 *   { t: 'entries', op: 'del', key: 123 }          видалити рядок за ключем
 *   { t: 'cats',    op: 'replace', rows: [...] }   замінити всю вкладку (лише з force)
 * baseRev: версія таблиці, яку бачив пристрій. Якщо вона не збігається з поточною, запис
 * відхиляється (error: 'stale'), і пристрій спершу підтягує нові дані.
 * Операції повторюються безпечно: повторна відправка дасть той самий результат.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return json_({ ok: false, error: 'auth' });

    const rev = getRev_();
    if (body.baseRev !== undefined && body.baseRev !== null && Number(body.baseRev) !== rev) {
      return json_({ ok: false, error: 'stale', rev: rev });
    }
    const force = body.force === true;

    const byTable = {};
    (body.ops || []).forEach(op => {
      if (!TABLES[op.t]) return;
      (byTable[op.t] = byTable[op.t] || []).push(op);
    });

    // 1. Рахуємо результат у пам'яті, нічого не пишучи.
    const plan = {};
    let hasDel = false;
    for (const name of Object.keys(byTable)) {
      const t = TABLES[name], kIdx = t.cols.indexOf(t.key);
      const before = readRows_(t);
      let rows = before.map(r => r.slice());
      const pos = {};
      rows.forEach((r, i) => { pos[String(r[kIdx])] = i; });

      for (const op of byTable[name]) {
        if (op.op === 'replace') {
          if (!force) return json_({ ok: false, error: 'forbidden_replace', table: name });
          rows = (op.rows || []).map(o => toRow_(t, o));
          Object.keys(pos).forEach(k => delete pos[k]);
          rows.forEach((r, i) => { pos[String(r[kIdx])] = i; });
        } else if (op.op === 'set') {
          const row = toRow_(t, op.row), k = String(row[kIdx]);
          if (k in pos) rows[pos[k]] = row; else { pos[k] = rows.length; rows.push(row); }
        } else if (op.op === 'del') {
          const k = String(op.key);
          if (k in pos) { rows[pos[k]] = null; delete pos[k]; }
        }
      }
      rows = rows.filter(r => r !== null).map(r => r.map(v => (v === null ? '' : v)));

      const norm = r => JSON.stringify(r.map(v => (v === null ? '' : v)));
      const bm = {};
      before.forEach(r => { bm[String(r[kIdx])] = norm(r); });
      let added = 0, changed = 0, removed = 0;
      const am = {};
      rows.forEach(r => {
        const k = String(r[kIdx]);
        am[k] = 1;
        if (!(k in bm)) added++; else if (bm[k] !== norm(r)) changed++;
      });
      Object.keys(bm).forEach(k => { if (!(k in am)) removed++; });
      if (removed > 0) hasDel = true;

      // Захист від масових видалень
      if (!force && before.length >= 10 && removed > Math.max(5, before.length * 0.2)) {
        return json_({ ok: false, error: 'too_many_deletes', table: name, removed: removed, total: before.length });
      }
      plan[name] = { t: t, rows: rows, added: added, changed: changed, removed: removed };
    }

    // 2. Копія перед записом, потім сам запис.
    const names = Object.keys(plan).filter(n => plan[n].added || plan[n].changed || plan[n].removed);
    const backed = names.length ? maybeBackup_(hasDel || force) : false;
    names.forEach(n => writeRows_(plan[n].t, plan[n].rows));

    let newRev = rev;
    if (names.length) {
      newRev = rev + 1;
      props_().setProperty('rev', String(newRev));
      log_(newRev, body, plan, backed);
    }
    return json_({ ok: true, rev: newRev, changed: names.length, backup: backed });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Запустіть один раз вручну в редакторі, щоб надати скрипту доступ до Drive (для резервних копій). */
function authorize() {
  DriveApp.getRootFolder();
  SpreadsheetApp.openById(SPREADSHEET_ID);
  PropertiesService.getScriptProperties();
  LockService.getScriptLock();
}

/** Створити резервну копію просто зараз (для перевірки). */
function backupNow() {
  const ok = backup_(Date.now());
  Logger.log(ok ? 'Копію створено' : 'Помилка: ' + props_().getProperty('backupError'));
}
