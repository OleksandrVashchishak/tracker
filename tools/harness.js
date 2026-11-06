// Макет Google для перевірок: запускає справжній apps-script/Code.gs з підробленими
// SpreadsheetApp / DriveApp / Lock / Properties і віддає застосунок, що дивиться на цей макет,
// а не на справжню таблицю. Використання: node tools/harness.js [порт]   (за замовчуванням 8767)
//
// Адреси для тестів:
//   /                 застосунок (index.html із підміненою адресою бекенду та тестовим токеном)
//   /api              те саме, що /exec скрипта (GET pull/ping, POST зміни)
//   /admin/reset      очистити «таблицю»
//   /admin/dump       кількість рядків по вкладках, версія, копії, журнал
//   /admin/sheet?n=…  вміст вкладки
//   /admin/age        «постарити» останню резервну копію на добу
const http = require('http'), fs = require('fs'), vm = require('vm'), path = require('path');
const ROOT = path.resolve(__dirname, '..');
const PORT = +process.argv[2] || 8767;
const TEST_TOKEN = 'test-token';

const CODE = fs.readFileSync(path.join(ROOT, 'apps-script', 'Code.gs'), 'utf8')
  .replace(/const TOKEN = '[^']*';/, `const TOKEN = '${TEST_TOKEN}';`);
const APP = () => fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8')
  .replace(/url:'https:\/\/script\.google\.com[^']*'/, `url:'http://localhost:${PORT}/api'`)
  .replace(/token:'[^']*'/, `token:'${TEST_TOKEN}'`);

const HEAD = {
  'Записи': ['ID', 'Дата', 'Прийом', 'Назва', 'Грами', 'Кількість', 'Одиниця', 'Ккал'],
  'Продукти': ['ID', 'Назва', 'a', 'b', 'c', 'd', 'e', 'f', 'g'],
  'Категорії': ['ID', 'Назва', 'Порядок'],
  'Шаблони': ['ID', 'Назва', 'Прийом', 'Склад'],
  'Вага': ['Дата', 'Вага'],
  'Налаштування': ['Ключ', 'Значення', 'Опис'],
};

let world;
function reset() {
  const sheets = {};
  for (const [n, h] of Object.entries(HEAD)) sheets[n] = mkSheet(n, [h]);
  world = { sheets, props: {}, folders: [], copies: [], now: Date.now() };
  buildContext();
}
function mkSheet(name, rows) {
  return {
    name, data: rows.map(r => r.slice()), maxRows: 1000,   // як у справжній таблиці: 1000 рядків за замовчуванням
    getLastRow() { let l = 0; this.data.forEach((r, i) => { if (r.some(v => v !== '' && v !== null && v !== undefined)) l = i + 1; }); return l; },
    getMaxRows() { return this.maxRows; },
    insertRowsAfter(pos, n) { this.maxRows += n; },
    getRange(r, c, nr = 1, nc = 1) {
      const self = this;
      const chk = () => { if (r + nr - 1 > self.maxRows) throw new Error('The coordinates of the range are outside the dimensions of the sheet.'); };
      return {
        getValues() { const o = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (self.data[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined || v === null ? '' : v); } o.push(row); } return o; },
        setValues(v) { chk(); for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { while (self.data.length < r + i) self.data.push([]); self.data[r - 1 + i][c - 1 + j] = v[i][j]; } },
        clearContent() { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (self.data[r - 1 + i]) self.data[r - 1 + i][c - 1 + j] = ''; },
        setFontWeight() { return this; },
      };
    },
    appendRow(a) { this.data.push(a.slice()); },
    deleteRows(s, n) { this.data.splice(s - 1, n); },
    setFrozenRows() {},
  };
}
function buildContext() {
  const props = world.props;
  const ctx = {
    console, JSON, Date, Math, Object, Array, String, Number, Error, Promise,
    Logger: { log: m => console.log('[Logger]', m) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; }, getContent() { return s; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }) },
    Utilities: { formatDate: d => new Date(d).toISOString().replace('T', ' ').slice(0, 19) },
    SpreadsheetApp: { openById: () => ({ getSheetByName: n => world.sheets[n] || null, insertSheet: n => (world.sheets[n] = mkSheet(n, [])) }) },
    DriveApp: {
      getRootFolder() {},
      getFileById: () => ({ makeCopy(name, folder) { const f = { name, created: world.now++, trashed: false, getDateCreated() { return this.created; }, setTrashed(v) { this.trashed = v; } }; folder.files.push(f); world.copies.push(name); } }),
      getFoldersByName: n => { const l = world.folders.filter(f => f.name === n); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; },
      createFolder: n => { const f = { name: n, files: [], getFiles() { const l = this.files.filter(x => !x.trashed); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; } }; world.folders.push(f); return f; },
    },
  };
  vm.createContext(ctx);
  vm.runInContext(CODE, ctx);
  world.ctx = ctx;
}
reset();

const dump = () => Object.fromEntries(Object.entries(world.sheets).map(([n, s]) => [n, s.data.filter(r => r.some(v => v !== '')).length - 1]));
const send = (res, code, body, type = 'application/json') => { res.writeHead(code, { 'Content-Type': type + '; charset=utf-8', 'Access-Control-Allow-Origin': '*' }); res.end(body); };

http.createServer((req, res) => {
  const u = new URL(req.url, 'http://x');
  let body = '';
  req.on('data', d => body += d);
  req.on('end', () => {
    try {
      if (u.pathname === '/' || u.pathname === '/index.html') return send(res, 200, APP(), 'text/html');
      if (u.pathname === '/api') {
        if (req.method === 'POST') return send(res, 200, world.ctx.doPost({ postData: { contents: body } }).getContent());
        return send(res, 200, world.ctx.doGet({ parameter: Object.fromEntries(u.searchParams) }).getContent());
      }
      if (u.pathname === '/admin/reset') { reset(); return send(res, 200, '{"ok":true}'); }
      if (u.pathname === '/admin/dump') return send(res, 200, JSON.stringify({ counts: dump(), rev: world.props.rev || 0, copies: world.copies, lastBackup: world.props.lastBackup || null, journal: world.sheets['Журнал'] ? world.sheets['Журнал'].data.slice(1) : [] }));
      if (u.pathname === '/admin/sheet') return send(res, 200, JSON.stringify(world.sheets[u.searchParams.get('n')].data));
      if (u.pathname === '/admin/age') { world.props.lastBackup = String(Date.now() - 24 * 3600 * 1000); return send(res, 200, '{"ok":true}'); }
      send(res, 404, '{}');
    } catch (e) { send(res, 500, JSON.stringify({ error: String(e), stack: e.stack })); }
  });
}).listen(PORT, () => console.log('макет на порту', PORT, '→ http://localhost:' + PORT + '/'));
