// パッと計算 ランキング用 Google Apps Script
// スプレッドシートの「拡張機能 → Apps Script」に貼り付けて、ウェブアプリとして公開する
const SHEET = 'records';
const BOARDS = ['ta30','ta60','ta90','en10','en20','en30'];
const LIMIT = 100; // 1ボードで返す上位件数

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['time', 'board', 'name', 'score', 'acc']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 記録の登録（全件そのまま追記。自己ベスト判定や重複排除はしない）
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    const board = String(d.board);
    const score = Number(d.score), acc = Number(d.acc);
    if (BOARDS.indexOf(board) < 0 || !isFinite(score) || score <= 0) return json_({ ok: false });
    let name = String(d.name || '').trim().slice(0, 12).replace(/^[=+\-@]+/, '') || '名無し';
    lock.waitLock(10000);
    sheet_().appendRow([new Date(), board, name, score, isFinite(acc) ? acc : 0]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// ランキング取得：?board=ta30 など。スコア降順→正答率降順→古い順
function doGet(e) {
  const board = String((e.parameter || {}).board || '');
  if (BOARDS.indexOf(board) < 0) return json_([]);
  const sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return json_([]);
  const rows = sh.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => r[1] === board)
    .map(r => ({ t: new Date(r[0]).getTime(), name: String(r[2]), score: Number(r[3]), acc: Number(r[4]) }));
  rows.sort((a, b) => b.score - a.score || b.acc - a.acc || a.t - b.t);
  return json_(rows.slice(0, LIMIT).map(r => ({ name: r.name, score: r.score, acc: r.acc })));
}
