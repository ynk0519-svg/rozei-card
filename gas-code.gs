/**
 * 租税教室「政策カード ワークシート」 結果受け取りスクリプト
 * ───────────────────────────────────────────────
 * これを使うと、各学校のタブレットから送られた結果が
 * 運営のGoogleスプレッドシートに自動で集まります。
 *
 * 【設定手順（5分）】
 *  1. Googleドライブで新しいスプレッドシートを作る（名前はなんでもOK。例：租税教室 集計）
 *  2. メニューの「拡張機能」→「Apps Script」を開く
 *  3. 出てきたコードを全部消して、このファイルの中身を貼り付けて保存
 *  4. 右上の「デプロイ」→「新しいデプロイ」→ 種類の歯車で「ウェブアプリ」を選ぶ
 *  5. 「次のユーザーとして実行」＝自分、「アクセスできるユーザー」＝全員　にする
 *  6. 「デプロイ」を押し、表示された  https://script.google.com/macros/s/..../exec  をコピー
 *  7. アプリの「運営 → 設定 → 送信先URL」に貼り付ける（これで全端末の送信先になります）
 *
 * ※ コードを直したときは「デプロイ」→「デプロイを管理」→ 鉛筆マーク →
 *    バージョン「新バージョン」→ デプロイ　で更新します（URLは変わりません）。
 */

var SHEET_NAME = 'results';

var HEADERS = ['gid','学校コード','学校名','クラス','班','人数','メンバー','個人の選択',
               'グループの政策','星合計','予算合計','理由','発表者','状態','更新日時','data'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function progress_(g) {
  if (g.done) return '発表ずみ';
  if (g.presenter) return '発表じゅんび';
  if (g.groupReason) return '理由きまり';
  if (g.picks && g.picks.length) return '話し合い中';
  if (g.revealed) return '見せ合い';
  var p = (g.members || []).filter(function (m) { return m.cardId; }).length;
  return 'えらび中 ' + p + '/' + (g.members || []).length;
}

function row_(g) {
  var members = (g.members || []);
  return [
    g.gid || '',
    g.school || '',
    g.schoolName || '',
    g.className || '',
    g.groupName || '',
    members.length,
    members.map(function (m, i) { return m.name || (i + 1) + '番'; }).join(' / '),
    members.map(function (m) { return m.cardId || '-'; }).join(' / '),
    (g.picks || []).join(' / '),
    g.starsTotal || '',
    g.yenTotal || '',
    (g.groupReason || '').replace(/\n/g, ' '),
    g.presenter || '',
    progress_(g),
    new Date(g.updated || Date.now()),
    JSON.stringify(g)
  ];
}

/** 結果の受け取り（アプリの「運営へ送信」から呼ばれます） */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    var groups = body.groups || [];
    var sh = getSheet_();
    var last = sh.getLastRow();
    var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    var added = 0, updated = 0;

    groups.forEach(function (g) {
      var r = row_(g);
      var at = ids.indexOf(g.gid);
      if (at >= 0) {
        sh.getRange(at + 2, 1, 1, HEADERS.length).setValues([r]);
        updated++;
      } else {
        sh.appendRow(r);
        ids.push(g.gid);
        added++;
      }
    });

    return json_({ ok: true, added: added, updated: updated });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 取り出し：  ....../exec?all=1  で全件を返します（運営画面の「クラウドから取りこむ」） */
function doGet(e) {
  try {
    var sh = getSheet_();
    var last = sh.getLastRow();
    if (last < 2) return json_({ ok: true, groups: [] });
    var col = HEADERS.indexOf('data') + 1;
    var vals = sh.getRange(2, col, last - 1, 1).getValues();
    var groups = [];
    vals.forEach(function (v) {
      if (!v[0]) return;
      try { groups.push(JSON.parse(v[0])); } catch (x) {}
    });
    // ?school=コード で学校をしぼりこめます
    var sc = e && e.parameter && e.parameter.school;
    if (sc) groups = groups.filter(function (g) { return g.school === sc; });
    return json_({ ok: true, groups: groups });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 動作確認用：メニューから1回実行すると、テスト行が1行入ります */
function テスト書き込み() {
  doPost({ postData: { contents: JSON.stringify({ type: 'results', groups: [{
    gid: 'test-1', school: 'sample', schoolName: 'テスト小学校', className: '6年1組',
    groupName: '1班', members: [{ name: 'あ', cardId: 1, reason: 'テスト' }],
    picks: [1], groupReason: 'テストです', presenter: 'あ', done: true, updated: Date.now()
  }] }) } });
}
