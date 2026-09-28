// 配置図グリッド(JSON) + 機種マスタ(CSV) → ページ用データ(JSON)
// 解析処理は gas/Code.gs の buildData_ をそのまま使う（スプレッドシート版と同じ結果になる）
// 使い方: node tools/build.js grid.json data/master.csv 出力.json [取り込み元のシート名]
const fs = require('fs');
const path = require('path');
const code = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Code.gs'), 'utf8');
eval(code.replace(/^const /gm, 'var '));

function readCsv(file) {
  const text = fs.readFileSync(file, 'utf8').replace(/^﻿/, '');
  const rows = []; let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
      else if (ch === '"') q = false;
      else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n') { row.push(cell.replace(/\r$/, '')); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

const [gridFile, masterFile, outFile, sourceName] = process.argv.slice(2);
const grid = JSON.parse(fs.readFileSync(gridFile, 'utf8')).map(r => r.map(v => (v && v.__date ? new Date(v.__date) : v)));
const master = readCsv(masterFile).slice(1).filter(r => r[0])
  .map(r => [r[0], r[1] === 'pachi' ? 'パチンコ' : r[1] === 'slot' ? 'スロット' : r[1], r[2], r[3], r[4]]);
const data = buildData_(grid, master, sourceName || '');
const unknown = data.unknownModels;
delete data.unknownModels;
fs.writeFileSync(outFile, JSON.stringify(data));
console.log(`機種 ${data.machines.length} / 台 ${data.machines.reduce((a, m) => a + m.nos.length, 0)} / ${data.sheet} (${data.updated})`);
if (unknown.length) {
  console.log('機種マスタ未登録の型式（data/master.csv に追記が必要）:');
  unknown.forEach(u => console.log(`  ${u.name}\t${u.cat}\t${u.type || '確認中'}`));
  process.exitCode = 2;
}
