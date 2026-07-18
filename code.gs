/* ============================================================
   子どもの進路、親の心配タイプ診断｜GAS（設定配信＋ログ記録）
   使い方：
   1. 新規スプレッドシートを作成し、拡張機能→Apps Script にこのコードを貼る
   2. 関数 setup を1回実行（シート「リンク設定」「カード画像」「ログ」を自動作成）
   3. デプロイ→新しいデプロイ→ウェブアプリ
      「自分として実行」「全員がアクセス可能」で公開
   4. 発行されたURLを index.html / demo.html の GAS_URL に貼る
   運用ルール：このスプレッドシートに私的な情報は書かない
============================================================ */

const TYPE_KEYS = ['sakimawari','hikaku','kyorikan','kasane','namae','moya','mimamori'];

const DEFAULT_CTA = [
  ['sakimawari','先回り型',   'この子の進路をみてもらう',       'https://coconala.com/services/4188585', '', 'ON'],
  ['hikaku',    '比較型',     'うちの子の光をみてもらう',       'https://coconala.com/services/4188585', '', 'ON'],
  ['kyorikan',  '距離感型',   'ちょうどいい距離をみてもらう',   'https://coconala.com/services/4188585', '', 'ON'],
  ['kasane',    '自分重ね型', 'この子自身の道をみてもらう',     'https://coconala.com/services/4188585', '', 'ON'],
  ['namae',     '名前まだ型', 'この心配の正体をみてもらう',     'https://coconala.com/services/4188585', '', 'ON'],
  ['moya',      '心配のもや', '心配をまるごとみてもらう',       'https://coconala.com/services/4188585', '', 'ON'],
  ['mimamori',  '見守り上手', '', '', '', 'OFF']
];

const CARD_NAMES = ['愚者','魔術師','女教皇','女帝','皇帝','教皇','恋人','戦車','力','隠者','運命の輪','正義','吊るされた男','死神','節制','悪魔','塔','星','月','太陽','審判','世界'];

/* ---------- 初期セットアップ（1回だけ実行） ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let link = ss.getSheetByName('リンク設定');
  if (!link) {
    link = ss.insertSheet('リンク設定');
    link.getRange(1, 1, 1, 6).setValues([['key','タイプ名','ボタン文言','リンク先URL','サムネ画像URL','表示(ON/OFF)']]);
    link.getRange(2, 1, DEFAULT_CTA.length, 6).setValues(DEFAULT_CTA);
    link.setFrozenRows(1);
  }

  let cards = ss.getSheetByName('カード画像');
  if (!cards) {
    cards = ss.insertSheet('カード画像');
    cards.getRange(1, 1, 1, 2).setValues([['カード名','画像URL（空ならリポジトリ同梱の画像を使用）']]);
    cards.getRange(2, 1, CARD_NAMES.length, 1).setValues(CARD_NAMES.map(n => [n]));
    cards.setFrozenRows(1);
  }

  let log = ss.getSheetByName('ログ');
  if (!log) {
    log = ss.insertSheet('ログ');
    log.getRange(1, 1, 1, 5).setValues([['日時','event','main','sub','card']]);
    log.setFrozenRows(1);
  }

  const s1 = ss.getSheetByName('シート1');
  if (s1 && ss.getSheets().length > 3) ss.deleteSheet(s1);
}

/* ---------- 設定配信（GET ?action=config） ---------- */
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';
  if (action !== 'config') {
    return ContentService.createTextOutput('ok');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = { cta: {}, cards: {} };

  const link = ss.getSheetByName('リンク設定');
  if (link && link.getLastRow() > 1) {
    const rows = link.getRange(2, 1, link.getLastRow() - 1, 6).getValues();
    rows.forEach(r => {
      const key = String(r[0]).trim();
      if (!key || TYPE_KEYS.indexOf(key) === -1) return;
      out.cta[key] = {
        label: String(r[2]),
        url: String(r[3]),
        thumb: String(r[4]),
        show: String(r[5]).toUpperCase() === 'ON'
      };
    });
  }

  const cards = ss.getSheetByName('カード画像');
  if (cards && cards.getLastRow() > 1) {
    const rows = cards.getRange(2, 1, cards.getLastRow() - 1, 2).getValues();
    rows.forEach(r => {
      const name = String(r[0]).trim();
      const url = String(r[1]).trim();
      if (name && url) out.cards[name] = url;
    });
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- ログ記録（POST） ---------- */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = ss.getSheetByName('ログ');
    if (!log) return ContentService.createTextOutput('no sheet');

    const allowed = ['complete','cta_click','tell_click','cross_click'];
    const ev = allowed.indexOf(String(data.event)) !== -1 ? String(data.event) : 'unknown';

    log.appendRow([
      new Date(),
      sanitize(ev),
      sanitize(data.main),
      sanitize(data.sub),
      sanitize(data.card)
    ]);
  } catch (err) {}
  return ContentService.createTextOutput('ok');
}

/* 数式インジェクション対策＋50字制限 */
function sanitize(v) {
  let s = String(v == null ? '' : v).slice(0, 50);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}
