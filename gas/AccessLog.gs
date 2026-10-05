/**
 * 設置機種ガイド（お試し p5/）のアクセス記録
 *
 * 使い方（記録用のスプレッドシートに設置）
 * 1. 新しいスプレッドシートを作る → 拡張機能 → Apps Script → このコードを貼り付けて保存
 * 2. 関数「setup」を一度実行（「ログ」「集計」シートができる。初回は権限の許可が必要）
 * 3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    次のユーザーとして実行: 自分 ／ アクセスできるユーザー: 全員 → デプロイ
 * 4. 表示されたウェブアプリのURL（…/exec）を Claude に伝える
 *
 * 記録する内容: 日時・種類（open=ページを開いた / qr=台のQRから開いた / view=機種の詳細を開いた / vote=投票ボタン）
 *             ・台番号・機種名・レート・セッション（ランダムな英数字。個人は特定しない）
 */
const LOG_SHEET = 'ログ';
const SUM_SHEET = '集計';
const EVENTS = ['open', 'qr', 'view', 'vote'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const ev = String(p.ev || '');
  if (EVENTS.indexOf(ev) < 0) return done_();
  const no = /^\d{1,4}$/.test(String(p.no || '')) ? Number(p.no) : '';
  const rate = /^(4P|1P|20S|5S)$/.test(String(p.rate || '')) ? String(p.rate) : '';
  // 先頭が = + - @ だとスプレッドシートが数式として扱うため、文字として記録する
  let name = String(p.name || '').replace(/[\r\n\t]/g, ' ').slice(0, 80);
  if (/^[=+\-@]/.test(name)) name = "'" + name;
  const sid = String(p.sid || '').replace(/[^0-9a-z]/gi, '').slice(0, 16);
  const page = String(p.page || '').replace(/[^0-9a-z]/gi, '').slice(0, 10);
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET) || setup();
      sh.appendRow([new Date(), ev, no, name, rate, sid, page]);
    } finally {
      lock.releaseLock();
    }
  }
  return done_();
}

function done_() {
  return ContentService.createTextOutput('ok');
}

/** 最初に一度だけ実行（シートと集計表を作る） */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.appendRow(['日時', '種類', '台番号', '機種名', 'レート', 'セッション', 'ページ']);
    log.setFrozenRows(1);
    log.getRange('D:D').setNumberFormat('@');
    log.getRange('F:G').setNumberFormat('@');
  }
  let sum = ss.getSheetByName(SUM_SHEET);
  if (!sum) sum = ss.insertSheet(SUM_SHEET);
  sum.clear();
  const R = "'" + LOG_SHEET + "'!A:G";
  const blocks = [
    ['機種別（QR＋詳細を開いた回数）', `=QUERY(${R},"select E, D, count(B) where (B='qr' or B='view') and D<>'' group by E, D order by count(B) desc label E 'レート', D '機種名', count(B) '回数'",1)`],
    ['台番号別（台のQRを読んだ回数）', `=QUERY(${R},"select C, D, count(B) where B='qr' and C is not null group by C, D order by count(B) desc label C '台番号', D '機種名', count(B) '回数'",1)`],
    ['日別', `=QUERY(${R},"select toDate(A), count(B) where B='qr' or B='open' group by toDate(A) order by toDate(A) desc label toDate(A) '日付', count(B) '開いた回数'",1)`],
    ['投票ボタン（機種別）', `=QUERY(${R},"select D, count(B) where B='vote' group by D order by count(B) desc label D '開いていた機種（空欄=一覧から）', count(B) '回数'",1)`],
  ];
  blocks.forEach(([title, f], i) => {
    const col = 1 + i * 4;
    sum.getRange(1, col).setValue(title).setFontWeight('bold');
    sum.getRange(2, col).setFormula(f);
  });
  return log;
}
