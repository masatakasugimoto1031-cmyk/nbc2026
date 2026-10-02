/**
 * ニュービジネス・クリエーション 2026　提出の受け口（Google Apps Script ウェブアプリ）
 *
 * ・POST：ワークベンチ（index.html）の「提出する」から届いた提出を、「提出」シートに1行追加する
 * ・GET ：集計ボード（board.html）の「スプレッドシートから読み込む」に、全行を返す（BOARD_PASSWORD が合うときだけ）
 *
 * このファイルは GitHub には上げない（BOARD_PASSWORD が書いてあるため）。
 */

// 学生用の合言葉。config.js の classCode と同じ値にする（公開ページに載るので秘密ではない）
const CLASS_CODE = '788ilo1r';
// 集計ボードの「読み込み用の合言葉」。提出の中身を守る鍵。人に見せない
const BOARD_PASSWORD = 'kl6n0nrhnhp5cu5f';

const SHEET_NAME = '提出';
const SCOPES = ['am', 'pm', 'iv', 'd2', 'pr', 'post'];
const CHUNK = 45000;          // 1セルに入れる文字数（セルの上限は50,000字）
const MAX_CODE = 45000 * 8;   // 提出データの上限（8セル分）
const HEAD = ['受付日時', 'チーム', '名前', '回', '回の名前', '文字数', 'データ1', 'データ2', 'データ3', 'データ4', 'データ5', 'データ6', 'データ7', 'データ8'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// セルの先頭が = + - @ だと式として扱われるため、文字として入れる
function safe_(v, n) {
  let s = String(v == null ? '' : v).slice(0, n || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function stamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) return json_({ ok: false, error: 'データが空です' });
    let d;
    try { d = JSON.parse(raw); } catch (err) { return json_({ ok: false, error: '形が読めません' }); }
    if (String(d.cls || '') !== CLASS_CODE) return json_({ ok: false, error: '合言葉が違います' });
    const code = String(d.code || '');
    if (!/^[zr][A-Za-z0-9_-]+$/.test(code)) return json_({ ok: false, error: '提出データの形が違います' });
    if (code.length > MAX_CODE) return json_({ ok: false, error: '提出データが長すぎます（' + code.length + '字）' });
    const scope = String(d.scope || '');
    if (SCOPES.indexOf(scope) < 0) return json_({ ok: false, error: '提出の回が違います' });
    const nick = String(d.nick || '').trim();
    if (!nick) return json_({ ok: false, error: '名前がありません' });

    const parts = [];
    for (let i = 0; i < code.length; i += CHUNK) parts.push(code.slice(i, i + CHUNK));
    const at = stamp_();
    const row = [at, safe_(d.team), safe_(nick), scope, safe_(d.label), code.length].concat(parts);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = sheet_();
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row.map(String)]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, at: at });
  } catch (err) {
    return json_({ ok: false, error: '受け口でエラーが起きました' });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.ping) return json_({ ok: true, at: stamp_() });
  if (String(p.pw || '') !== BOARD_PASSWORD) return json_({ ok: false, error: '合言葉が違います' });
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: true, rows: [] });
  const vals = sh.getRange(2, 1, n, HEAD.length).getDisplayValues();
  const rows = vals.map(function (v) {
    return { at: v[0], team: v[1], nick: v[2], scope: v[3], label: v[4], code: v.slice(6).join('') };
  }).filter(function (x) { return x.code; });
  return json_({ ok: true, rows: rows });
}

// 公開前に一度だけ実行すると、「提出」シートを作り、許可を求める画面が出る
function setup() {
  sheet_();
}
