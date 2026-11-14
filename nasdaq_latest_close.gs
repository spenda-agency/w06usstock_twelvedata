/**
 * 「Nasdaq最新」シートの A列（銘柄コード）を読み取り、
 * Twelve Data の日足（interval=1day, outputsize=1）＝直近1日分の終値を B列に出力する。
 *
 * 使い方:
 *   1. スプレッドシート → 拡張機能 → Apps Script にこのファイルを貼り付ける
 *   2. プロジェクトの設定 → スクリプト プロパティに TWELVEDATA_API_KEY を登録
 *      （https://twelvedata.com/account の API Keys で確認できるキー）
 *   3. 必要ならプランに合わせて CHUNK_SIZE を変更
 *   4. startFetch() を実行（初回は権限承認が必要）
 *      - 6分の実行時間制限に達すると、進捗を保存して1分後に自動で続きを実行する
 *      - 完了すると自動トリガーは削除される
 */

const SHEET_NAME = 'Nasdaq最新';
const START_ROW = 2;          // 1行目は見出し（CODE）
// 1分あたりのAPIクレジット上限（1銘柄=1クレジット）。Basic(無料)=8, Grow=55, Pro=610 など
const CHUNK_SIZE = 8;
const WAIT_MS = 61 * 1000;    // 1分ごとのレート制限リセットを待つ
const MAX_RUN_MS = 5 * 60 * 1000; // 6分制限の手前で止めて続きはトリガーに任せる
const PROP_NEXT = 'TD_NEXT_ROW';

function startFetch() {
  PropertiesService.getScriptProperties().setProperty(PROP_NEXT, String(START_ROW));
  getSheet_().getRange('B1').setValue('終値（取得中）');
  fetchLatestClose();
}

function fetchLatestClose() {
  const started = Date.now();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('TWELVEDATA_API_KEY');
  if (!apiKey) throw new Error('スクリプトプロパティ TWELVEDATA_API_KEY が未設定です');

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  let row = Number(props.getProperty(PROP_NEXT) || START_ROW);

  while (row <= lastRow) {
    if (Date.now() - started > MAX_RUN_MS) {
      props.setProperty(PROP_NEXT, String(row));
      scheduleNext_();
      return;
    }

    const n = Math.min(CHUNK_SIZE, lastRow - row + 1);
    const codes = sheet.getRange(row, 1, n, 1).getValues().map(r => String(r[0]).trim());
    const symbols = codes.filter(c => c);
    const results = symbols.length ? requestDaily_(symbols, apiKey) : {};

    const out = codes.map(c => [c ? (results[c] !== undefined ? results[c] : 'N/A') : '']);
    sheet.getRange(row, 2, n, 1).setValues(out);
    SpreadsheetApp.flush();

    row += n;
    props.setProperty(PROP_NEXT, String(row));
    if (row <= lastRow) Utilities.sleep(WAIT_MS);
  }

  // 完了
  props.deleteProperty(PROP_NEXT);
  deleteTriggers_();
  sheet.getRange('B1').setValue('終値 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 取得');
}

/** 複数銘柄をまとめて取得し { SYMBOL: close } を返す。取得できない銘柄はエラーメッセージ。 */
function requestDaily_(symbols, apiKey) {
  const url = 'https://api.twelvedata.com/time_series'
    + '?symbol=' + encodeURIComponent(symbols.join(','))
    + '&interval=1day&outputsize=1&apikey=' + apiKey;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());

  if (json.status === 'error' && json.code) {
    // リクエスト全体のエラー（APIキー不正・クレジット切れ等）
    throw new Error('Twelve Data エラー: ' + json.message);
  }

  // 1銘柄のときはオブジェクトがそのまま返る
  const bySymbol = symbols.length === 1 ? { [symbols[0]]: json } : json;
  const out = {};
  symbols.forEach(s => {
    const d = bySymbol[s];
    if (d && d.status === 'ok' && d.values && d.values.length) {
      out[s] = Number(d.values[0].close);
    } else {
      out[s] = 'ERR: ' + (d && d.message ? d.message : 'no data');
    }
  });
  return out;
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('シート「' + SHEET_NAME + '」が見つかりません');
  return sheet;
}

function scheduleNext_() {
  deleteTriggers_();
  ScriptApp.newTrigger('fetchLatestClose').timeBased().after(60 * 1000).create();
}

function deleteTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'fetchLatestClose')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
