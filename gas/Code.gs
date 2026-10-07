/**
 * 上峰店 設置機種ガイド データ更新スクリプト（Google Apps Script）
 *
 * このスクリプトは「設置機種ガイド データ」スプレッドシートに貼り付けて使います。
 *  - 「配置図」シートに配置図の内容を貼り付ける → 自動でガイドに反映
 *  - 配置図のExcelを「設置機種ガイド_配置図アップロード」フォルダに入れる → 10分以内に自動で取り込み
 *  - ガイドページは、このスクリプトをWebアプリとして公開したURLからデータを読み込みます
 *
 * 初回だけ、メニュー「設置機種ガイド」→「初期設定」を実行してください。
 */

const SHEET_LAYOUT = '配置図';
const SHEET_MASTER = '機種マスタ';
const FOLDER_UPLOAD = '設置機種ガイド_配置図アップロード';
const FOLDER_DONE = '取り込み済み';
const CACHE_KEY = 'guide_data_v1';
const CACHE_SEC = 21600; // 6時間（貼り付け・取り込み時は即時に作り直します）

// 配置図の島ラベル → ページのレート区分
const RATE_MAP = { '4P': '4P', '1P': '1P', 'プラス1': '1P', '20S': '20S', 'プラス5': '5S' };

// シリーズ名（機種マスタの「シリーズ」欄が空のときに使う）
const SERIES = [
  ['ジャグラー', 'ジャグラー'], ['海物語', '海物語'], ['エヴァンゲリオン', 'エヴァンゲリオン'], ['番長', '番長'],
  ['北斗', '北斗の拳'], ['禁書目録', 'とある'], ['超電磁砲', 'とある'], ['とある', 'とある'], ['牙狼', '牙狼'],
  ['慶次', '花の慶次'], ['カバネリ', 'カバネリ'], ['バイオハザード', 'バイオハザード'], ['鬼武者', '鬼武者'],
  ['シンフォギア', 'シンフォギア'], ['からくりサーカス', 'からくりサーカス'], ['沖ドキ', '沖ドキ'],
  ['モンキーターン', 'モンキーターン'], ['ダンまち', 'ダンまち'], ['アズールレーン', 'アズールレーン'],
  ['炎炎ノ消防隊', '炎炎ノ消防隊'], ['甲賀', 'バジリスク'], ['バジリスク', 'バジリスク'], ['ガンダム', 'ガンダム'],
  ['パルサー', 'パルサー'], ['秘宝伝', '秘宝伝'], ['ゴッドイーター', 'ゴッドイーター'], ['ブラックジャック', 'ブラックジャック'],
  ['ヤマト', '宇宙戦艦ヤマト'], ['東京リベンジャーズ', '東京リベンジャーズ'], ['ゼロから始める', 'Re:ゼロ'],
  ['この素晴らしい', 'このすば'], ['エウレカ', 'エウレカセブン'], ['一騎当千', '一騎当千'], ['ギンギラ', 'ギンギラパラダイス'],
  ['わんわん', 'わんわん'], ['クィーン', 'クィーン'], ['モンスターハンター', 'モンスターハンター'], ['ハナビ', 'ハナビ'],
  ['クランキー', 'クランキー'], ['ひぐらし', 'ひぐらし'], ['頭文字D', '頭文字D'], ['七つの大罪', '七つの大罪'],
  ['銭形', '銭形'], ['麻雀格闘倶楽部', '麻雀格闘倶楽部'], ['ファミスタ', 'ファミスタ'], ['タイガーマスク', 'タイガーマスク'],
  ['花月', '花月'], ['ゴジラ', 'ゴジラ']
];

const SPEC_VALUES = ['甘デジ', 'ライトミドル', 'ミドル', '設定付き', 'ノーマル(Aタイプ)', 'AT・ART', 'スマスロ'];

// 配置図に書かれていてもガイドに載せない台番号（店の指示：64番台・60番台は実在しない扱い）
const EXCLUDE_NOS = [60, 64];

/* ===================== メニュー・初期設定 ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('設置機種ガイド')
    .addItem('ガイドを今すぐ更新', 'menuRefresh')
    .addItem('アップロードフォルダから取り込む', 'importFromFolder')
    .addSeparator()
    .addItem('初期設定', 'setup')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SHEET_LAYOUT)) ss.insertSheet(SHEET_LAYOUT, 0);
  if (!ss.getSheetByName(SHEET_MASTER)) {
    const m = ss.insertSheet(SHEET_MASTER);
    m.getRange(1, 1, 1, 5).setValues([['配置図の機種名（正式名称）', '区分', 'スペック／タイプ', 'P-WORLD機種ID', 'シリーズ（空欄なら自動・「-」でなし）']]);
  }
  // 「スペック／タイプ」欄は選択式にする
  const master = ss.getSheetByName(SHEET_MASTER);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(SPEC_VALUES.concat(['確認中']), true).setAllowInvalid(false).build();
  master.getRange(2, 3, Math.max(master.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  master.setFrozenRows(1);

  // アップロード用フォルダ（スプレッドシートと同じ場所に作る）
  const folder = getUploadFolder_(true);

  // トリガー（重複しないよう作り直す）
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChangeTrigger').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('importFromFolder').timeBased().everyMinutes(10).create();

  refresh_();
  SpreadsheetApp.getUi().alert(
    '初期設定が終わりました。\n\n' +
    '・コピペ：「' + SHEET_LAYOUT + '」シートに配置図を貼り付けると自動で反映されます。\n' +
    '・アップロード：Googleドライブの「' + folder.getName() + '」フォルダに配置図のExcelを入れると、10分以内に取り込まれます。\n\n' +
    '次に「デプロイ」→「新しいデプロイ」→「ウェブアプリ」で公開し、表示されたURLを連絡してください。');
}

function menuRefresh() {
  const r = refresh_();
  SpreadsheetApp.getActive().toast(r.machines.length + '機種を反映しました。' + (r.unknown ? '機種マスタに未登録の型式が' + r.unknown + '件あります（黄色の行）。' : ''), '設置機種ガイド', 8);
}

/* ===================== 自動反映 ===================== */

function onEditTrigger(e) {
  const name = e && e.range ? e.range.getSheet().getName() : '';
  if (name === SHEET_LAYOUT) PropertiesService.getScriptProperties().deleteProperty('sourceName'); // 貼り付け時は取り込み元の名前をリセット
  if (name === SHEET_LAYOUT || name === SHEET_MASTER) refresh_();
}

function onChangeTrigger(e) {
  // 行の挿入・削除などでも作り直す（編集は onEditTrigger で処理）
  if (e && (e.changeType === 'INSERT_ROW' || e.changeType === 'REMOVE_ROW' || e.changeType === 'OTHER')) refresh_();
}

/** データを作り直してキャッシュし、未登録の型式を機種マスタに追記する */
function refresh_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const data = buildFromSheets_(ss);
    const unknown = appendUnknownToMaster_(ss, data.unknownModels);
    delete data.unknownModels;
    CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(data), CACHE_SEC);
    return { machines: data.machines, unknown: unknown };
  } finally {
    lock.releaseLock();
  }
}

/** Webアプリ：ガイドページがここからデータを読み込む */
function doGet() {
  let json = CacheService.getScriptCache().get(CACHE_KEY);
  if (!json) {
    const data = buildFromSheets_(SpreadsheetApp.getActive());
    delete data.unknownModels;
    json = JSON.stringify(data);
    CacheService.getScriptCache().put(CACHE_KEY, json, CACHE_SEC);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ===================== アップロード取り込み ===================== */

function getUploadFolder_(create) {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = ssFile.getParents();
  const base = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const it = base.getFoldersByName(FOLDER_UPLOAD);
  if (it.hasNext()) return it.next();
  if (!create) return null;
  const f = base.createFolder(FOLDER_UPLOAD);
  f.createFolder(FOLDER_DONE);
  return f;
}

/** アップロードフォルダの最新の配置図を「配置図」シートに取り込む（10分ごとに自動実行） */
function importFromFolder() {
  const folder = getUploadFolder_(false);
  if (!folder) return;
  const doneIt = folder.getFoldersByName(FOLDER_DONE);
  const done = doneIt.hasNext() ? doneIt.next() : folder.createFolder(FOLDER_DONE);

  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const mt = f.getMimeType();
    if (mt === MimeType.MICROSOFT_EXCEL || mt === MimeType.MICROSOFT_EXCEL_LEGACY || mt === MimeType.GOOGLE_SHEETS) files.push(f);
  }
  if (!files.length) return;
  files.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  const latest = files[0];

  let tmpId = null, ss;
  if (latest.getMimeType() === MimeType.GOOGLE_SHEETS) {
    ss = SpreadsheetApp.openById(latest.getId());
  } else {
    // Excel を一時的にGoogleスプレッドシートへ変換（「サービス」で Drive API を追加しておく必要があります）
    const copy = Drive.Files.copy({ mimeType: MimeType.GOOGLE_SHEETS }, latest.getId());
    tmpId = copy.id;
    ss = SpreadsheetApp.openById(tmpId);
  }
  const src = pickLayoutSheet_(ss);
  const values = src.getDataRange().getValues();

  const target = SpreadsheetApp.getActive().getSheetByName(SHEET_LAYOUT) || SpreadsheetApp.getActive().insertSheet(SHEET_LAYOUT, 0);
  target.clearContents();
  if (target.getMaxRows() < values.length) target.insertRowsAfter(target.getMaxRows(), values.length - target.getMaxRows());
  if (target.getMaxColumns() < values[0].length) target.insertColumnsAfter(target.getMaxColumns(), values[0].length - target.getMaxColumns());
  target.getRange(1, 1, values.length, values[0].length).setValues(values);
  PropertiesService.getScriptProperties().setProperty('sourceName', src.getName());

  if (tmpId) DriveApp.getFileById(tmpId).setTrashed(true);
  files.forEach(f => f.moveTo(done));
  refresh_();
}

/** B1が「入替日」のシートのうち、D1の日付が最も新しいものを選ぶ */
function pickLayoutSheet_(ss) {
  let best = null, bestTime = -Infinity;
  ss.getSheets().forEach(sh => {
    if (String(sh.getRange('B1').getValue()).trim() !== '入替日') return;
    const t = toDate_(sh.getRange('D1').getValue());
    const time = t ? t.getTime() : 0;
    if (time > bestTime) { best = sh; bestTime = time; }
  });
  return best || ss.getSheets()[0];
}

/* ===================== 解析（ここから下はシートに依存しない処理） ===================== */

function buildFromSheets_(ss) {
  const layout = ss.getSheetByName(SHEET_LAYOUT);
  const grid = layout ? layout.getDataRange().getValues() : [[]];
  const masterSheet = ss.getSheetByName(SHEET_MASTER);
  const masterRows = masterSheet && masterSheet.getLastRow() > 1
    ? masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 5).getValues() : [];
  const sourceName = PropertiesService.getScriptProperties().getProperty('sourceName') || '';
  return buildData_(grid, masterRows, sourceName, EXCLUDE_NOS);
}

function nf_(s) { return String(s == null ? '' : s).normalize('NFKC'); }
function key_(s) { return nf_(s).replace(/\s/g, ''); }
function isNum_(v) {
  if (typeof v === 'number') return isFinite(v);
  return typeof v === 'string' && /^\s*\d{1,4}\s*$/.test(v);
}
function isName_(v) {
  if (typeof v !== 'string') return false;
  const s = nf_(v).trim();
  return s.length >= 3 && /^(PA|P|e|E|S|L)/.test(s) && !/JC|台|\n/.test(v);
}
function toDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && v > 30000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)); // Excelの日付シリアル
  if (typeof v === 'string') {
    const s = nf_(v).trim();
    if (/^\d{5}$/.test(s)) return toDate_(+s);
    let m = s.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{1,2})[\/月](\d{1,2})日?/); // 年なし（例：9月15日、9/15）は今年とみなす
    if (m) return new Date(new Date().getFullYear(), +m[1] - 1, +m[2]);
  }
  return null;
}
function parseLabel_(label) {
  const s = nf_(label).replace(/\n/g, '');
  const out = [];
  const re = /(20S|プラス5|プラス1|1P|4P)\s*(\d+)?台?/g;
  let m;
  while ((m = re.exec(s))) out.push([m[1], m[2] ? +m[2] : 999]);
  return out;
}

/**
 * 配置図のグリッドから機種データを作る
 * grid: 2次元配列（シートの値） / masterRows: [型式名, 区分, スペック・タイプ, P-WORLD ID, シリーズ]
 */
function buildData_(grid, masterRows, sourceName, excludeNos) {
  const exclude = excludeNos || [];
  const cell = (r, c) => (grid[r] && grid[r][c] !== undefined ? grid[r][c] : '');
  const nRows = grid.length;
  const nCols = grid.reduce((a, row) => Math.max(a, row.length), 0);

  // 「導入機種」の表より上だけを読む
  let stop = nRows;
  for (let r = 0; r < nRows; r++) if (String(cell(r, 1)).trim() === '導入機種') { stop = r; break; }

  const units = []; // 1台ごと
  for (let r = 0; r < stop; r++) {
    const names = [];
    for (let c = 0; c < nCols; c++) if (isName_(cell(r, c))) names.push(c);
    if (names.length < 3) continue;
    // 台番号の行（上下のうち数字が多い方）
    let best = null;
    [r - 1, r + 1].forEach(rr => {
      if (rr < 0 || rr >= nRows) return;
      const cnt = names.filter(c => isNum_(cell(rr, c))).length;
      if (!best || cnt > best.cnt) best = { rr: rr, cnt: cnt };
    });
    // 島ラベル（B〜G列の「台」を含むセル）
    let label = '';
    for (let c = 1; c <= 6; c++) { const v = cell(r, c); if (typeof v === 'string' && v.indexOf('台') >= 0) { label = v; break; } }
    const segs = parseLabel_(label);
    let idx = 0;
    const rateOf = [];
    segs.forEach(([rate, cnt]) => { for (let k = 0; k < cnt && idx < names.length; k++, idx++) rateOf.push(RATE_MAP[rate]); });
    const last = segs.length ? RATE_MAP[segs[segs.length - 1][0]] : '';
    names.forEach((c, i) => {
      const noRaw = best ? cell(best.rr, c) : '';
      const no = isNum_(noRaw) ? +noRaw : null;
      if (no != null && exclude.indexOf(no) >= 0) return; // 載せない台番号
      units.push({ name: nf_(cell(r, c)).trim(), no: no, rate: rateOf[i] || last });
    });
  }

  // 機種マスタ
  const master = {};
  masterRows.forEach(row => {
    const k = key_(row[0]);
    if (k) master[k] = { cat: String(row[1] || ''), type: String(row[2] || '').trim(), pw: String(row[3] || '').trim(), series: String(row[4] || '').trim() };
  });

  // 型式×レートでまとめる
  const groups = {}, order = [], unknownModels = {};
  units.forEach(u => {
    if (!u.rate) return;
    const k = key_(u.name) + '|' + u.rate;
    if (!groups[k]) { groups[k] = { name: u.name, rate: u.rate, nos: [] }; order.push(k); }
    if (u.no != null) groups[k].nos.push(u.no);
  });

  const machines = order.map(k => {
    const g = groups[k];
    const s = g.name;
    const cat = /^[Pe]/.test(s) ? 'pachi' : 'slot';
    const m = master[key_(s)];
    let type = m && SPEC_VALUES.indexOf(m.type) >= 0 ? m.type : '';
    if (!type && cat === 'slot' && /^L/.test(s)) type = 'スマスロ';
    if (!m) unknownModels[key_(s)] = { name: s, cat: cat, type: type };
    let series = m && m.series ? m.series : null;
    if (/^[-－ー―‐]$/.test(series)) series = null; // 「-」はシリーズなし（キーワードの誤一致を防ぐ）
    else if (!series) for (let i = 0; i < SERIES.length; i++) if (s.indexOf(SERIES[i][0]) >= 0) { series = SERIES[i][1]; break; }
    const d = { name: s, full: s, cat: cat, type: type || '確認中', series: series, rate: g.rate, nos: g.nos.sort((a, b) => a - b) };
    if (/^e/.test(s)) d.smart = true;
    if (m && /^\d+$/.test(m.pw)) d.pw = m.pw;
    return d;
  });
  const rateOrder = { '4P': 0, '1P': 1, '20S': 2, '5S': 3 };
  machines.sort((a, b) => (a.cat === b.cat ? 0 : a.cat === 'pachi' ? -1 : 1) || rateOrder[a.rate] - rateOrder[b.rate] || b.nos.length - a.nos.length || (key_(a.name) < key_(b.name) ? -1 : 1));

  // 更新日・表示名（入替日 D1、開店日は入替日の2日後）
  const d1 = toDate_(cell(0, 3));
  let updated = '', sheet = sourceName;
  if (d1) {
    const open = new Date(d1.getTime() + 2 * 86400000);
    updated = open.getFullYear() + '-' + ('0' + (open.getMonth() + 1)).slice(-2) + '-' + ('0' + open.getDate()).slice(-2);
    if (!sheet) sheet = (open.getMonth() + 1) + '月' + open.getDate() + '日開店';
  }
  return { updated: updated, sheet: sheet || '配置図', machines: machines, unknownModels: Object.keys(unknownModels).map(k => unknownModels[k]) };
}

/** 機種マスタにない型式を末尾に追記（黄色）。スペック・IDを入れると次回から反映 */
function appendUnknownToMaster_(ss, list) {
  if (!list || !list.length) return 0;
  const sh = ss.getSheetByName(SHEET_MASTER);
  const start = sh.getLastRow() + 1;
  const rows = list.map(u => [u.name, u.cat === 'pachi' ? 'パチンコ' : 'スロット', u.type || '確認中', '', '']);
  sh.getRange(start, 1, rows.length, 5).setValues(rows).setBackground('#FFF3C4');
  return rows.length;
}
