// Швидка перевірка синтаксису скрипта всередині index.html: node tools/check.js
const fs = require('fs'), os = require('os'), path = require('path'), { spawnSync } = require('child_process');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const m = html.match(/<script>([\s\S]*)<\/script>/);
if (!m) { console.log('Не знайдено <script> в index.html'); process.exit(2); }
const tmp = path.join(os.tmpdir(), 'kcal-inline-check.js');
fs.writeFileSync(tmp, m[1]);
const r = spawnSync(process.execPath, ['--check', tmp], { encoding: 'utf8' });
if (r.status === 0) console.log('Синтаксис в порядку');
else { console.log(r.stderr); process.exit(1); }
