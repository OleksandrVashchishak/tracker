// Перевірки протоколу бекенду на макеті. Спершу запустіть макет: node tools/harness.js
// Потім: node tools/test-server.js [порт]
// Перевіряє: токен, версію таблиці (stale), заборону заміни й масових видалень без force,
// резервні копії та їхній інтервал, журнал, ліміт у 1000 рядків, сумісність зі старим клієнтом.
const B = 'http://localhost:' + (process.argv[2] || 8767), T = 'test-token';
const post = async b => (await fetch(B + '/api', { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ token: T, ...b }) })).json();
const get = async q => (await fetch(B + '/api?token=' + T + '&' + q)).json();
const dump = async () => (await fetch(B + '/admin/dump')).json();
let fails = 0;
const ok = (n, c, x) => { if (!c) fails++; console.log((c ? 'PASS ' : 'FAIL ') + n + (x ? '  ' + x : '')); };
const row = (id, extra = {}) => ({ id, d: '2026-10-01', m: 'Обід', n: 'x' + id, g: null, q: null, u: null, k: 100, ...extra });

(async () => {
  await fetch(B + '/admin/reset');
  ok('поганий токен', (await (await fetch(B + '/api?token=x&action=pull')).json()).error === 'auth');
  const rows = Array.from({ length: 30 }, (_, i) => ({ t: 'entries', op: 'set', row: row(i + 1) }));
  let r = await post({ ops: rows, baseRev: 0, dev: 'A' });
  ok('запис 30 рядків, версія 1', r.ok && r.rev === 1, JSON.stringify(r));
  let p = await get('action=pull');
  ok('читання бачить 30 і версію', p.tables.entries.length === 30 && p.rev === 1);
  r = await post({ ops: [{ t: 'entries', op: 'set', row: row(99) }], baseRev: 0 });
  ok('застарілий запис відхилено', r.error === 'stale' && r.rev === 1, JSON.stringify(r));
  p = await get('action=pull'); ok('відхилений запис нічого не змінив', p.tables.entries.length === 30);
  r = await post({ ops: [{ t: 'entries', op: 'replace', rows: [] }], baseRev: 1 });
  ok('заміна без force відхилена', r.error === 'forbidden_replace');
  const dels = Array.from({ length: 10 }, (_, i) => ({ t: 'entries', op: 'del', key: String(i + 1) }));
  r = await post({ ops: dels, baseRev: 1 });
  ok('масове видалення відхилене', r.error === 'too_many_deletes', JSON.stringify(r));
  p = await get('action=pull'); ok('масове видалення нічого не змінило', p.tables.entries.length === 30);
  r = await post({ ops: [{ t: 'entries', op: 'del', key: '1' }, { t: 'entries', op: 'del', key: '2' }], baseRev: 1 });
  ok('мале видалення (2 з 30) проходить', r.ok && r.rev === 2);
  r = await post({ ops: dels, baseRev: 2, force: true });
  ok('force дозволяє масове видалення', r.ok && r.rev === 3, JSON.stringify(r));
  let d = await dump(); ok('резервна копія існує', d.copies.length >= 1, 'копій: ' + d.copies.length);
  r = await post({ ops: [{ t: 'cats', op: 'replace', rows: [{ id: 'a', n: 'A', o: 0 }] }], baseRev: 3, force: true });
  ok('заміна з force проходить', r.ok);
  d = await dump(); ok('журнал веде записи', d.journal.length >= 4, d.journal.length + ' рядків');
  r = await post({ ops: [{ t: 'entries', op: 'set', row: row(500) }] });
  ok('старий клієнт без baseRev працює (лише set)', r.ok);
  const many = Array.from({ length: 1500 }, (_, i) => ({ t: 'entries', op: 'set', row: row(10000 + i) }));
  r = await post({ ops: many, baseRev: (await get('action=pull')).rev });
  ok('1500 рядків понад ліміт 1000', r.ok, JSON.stringify(r).slice(0, 100));
  p = await get('action=pull'); ok('усі рядки на місці', p.tables.entries.length >= 1500, String(p.tables.entries.length));
  d = await dump(); const c1 = d.copies.length;
  await post({ ops: [{ t: 'entries', op: 'del', key: '10000' }], baseRev: (await get('action=pull')).rev });
  d = await dump(); ok('копії не частіше за інтервал', d.copies.length === c1, c1 + ' → ' + d.copies.length);
  await fetch(B + '/admin/age');
  await post({ ops: [{ t: 'entries', op: 'del', key: '10001' }], baseRev: (await get('action=pull')).rev });
  d = await dump(); ok('нова копія після інтервалу', d.copies.length === c1 + 1, c1 + ' → ' + d.copies.length);
  console.log(fails ? `\nПровалів: ${fails}` : '\nУсе пройшло');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.log('ПОМИЛКА', e.message, '\nМакет запущено? node tools/harness.js'); process.exit(2); });
