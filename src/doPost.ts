import { register } from './register';
import { start } from './start';
import { stop } from './stop';

export function doGet() {
  return ContentService.createTextOutput('Hello, World!, from doGet'); // レスポンスとしてテキストを返す
}

export function doPost(e: GoogleAppsScript.Events.DoPost) {
  // TODO 月ごとにシートを分ける処理を入れる
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 現在開いているスプレッドシートを取得
  const sheet = ss.getSheetByName('シート1'); // 書き込むシートを指定（シート名を変更してください）
  const slackID = e.parameter.user_id;
  const command = e.parameter.command;
  if (!sheet) {
    return ContentService.createTextOutput(
      '無効なコマンドです。管理者に問い合わせてください'
    );
  }
  if (command === '/start') {
    return start(sheet, slackID);
  } else if (command === '/stop') {
    return stop(sheet, slackID);
  } else if (command === '/register') {
    return register(sheet, slackID, e.parameter.text);
  } else {
    // 無効なコマンド。ここに来ることはないはず。
    return ContentService.createTextOutput(
      '無効なコマンドです。管理者に問い合わせてください'
    );
  }
}
