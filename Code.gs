/**
 * GAS × 生成AI サンプル
 * スプレッドシートの「お問い合わせ」シート（A列:日時 / B列:本文）に対して、
 * AIで「分類」と「要約」を生成し、C列・D列に書き出すデモ用スクリプト。
 *
 * 使い方:
 * 1. スプレッドシートに紐づくApps Scriptプロジェクトにこのコードを貼る
 * 2. プロジェクトの設定 > スクリプトプロパティ に OPENAI_API_KEY を登録する
 * 3. メニュー「AI処理」>「未処理の行を分類・要約する」を実行する
 *
 * ※ デモ用のサンプルコードです。実案件の納品物ではありません。
 */

var MODEL = 'gpt-4o-mini';
var SHEET_NAME = 'お問い合わせ';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI処理')
    .addItem('未処理の行を分類・要約する', 'classifyInquiries')
    .addToUi();
}

function classifyInquiries() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var updated = 0;

  for (var i = 0; i < values.length; i++) {
    var body = values[i][1];
    var category = values[i][2];
    if (!body || category) continue; // 本文なし、または処理済みはスキップ

    var result = askAi(body);
    if (result) {
      sheet.getRange(i + 2, 3).setValue(result.category);
      sheet.getRange(i + 2, 4).setValue(result.summary);
      updated++;
    }
  }
  SpreadsheetApp.getUi().alert(updated + ' 件を処理しました。');
}

/**
 * OpenAI APIに分類と要約を依頼し、JSONで受け取る。
 * 失敗時は null を返す（行は未処理のまま残る）。
 */
function askAi(text) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('スクリプトプロパティ OPENAI_API_KEY が未設定です');

  var prompt = [
    '次のお問い合わせ文を分類し、要約してください。',
    '分類は「料金」「使い方」「不具合」「その他」のいずれか1つ。',
    '出力はJSONのみ: {"category": "...", "summary": "40文字以内の要約"}',
    '',
    text
  ].join('\n');

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: MODEL,
      messages: [{ role: 'user', content: prompt }],
      response_format: { type: 'json_object' }
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) return null;
  var content = JSON.parse(response.getContentText()).choices[0].message.content;
  try {
    return JSON.parse(content);
  } catch (e) {
    return null;
  }
}
