export function stop(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  slackID: string
) {
  // slackID検索
  let target_row = 1;
  while (sheet.getRange(target_row, 1).getValue()) {
    if (slackID === sheet.getRange(target_row, 1).getValue()) {
      break;
    }
    target_row += 1;
  }
  // slackIDが見つからない場合
  if (!sheet.getRange(target_row, 1).getValue()) {
    return ContentService.createTextOutput(
      'アカウントが見つかりませんでした。/registerコマンドで登録してください'
    );
  }

  // slackIDの行の末尾に開始時刻を追加
  let target_index = 3; // 出退勤時刻は3列目から
  while (sheet.getRange(target_row, target_index).getValue()) {
    target_index += 1;
  }

  // 末尾が偶数列(退勤状態)のとき
  // target_indexは末尾の次の列を取る
  if (target_index % 2 === 1) {
    return ContentService.createTextOutput(
      '出勤されていません。出勤状態を確認してください。'
    );
  }

  // slackIDが登録されていて、退勤状態のときのみ出勤登録できる
  const currentDate = new Date();
  const formattedDate = Utilities.formatDate(
    currentDate,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
  sheet.getRange(target_row, target_index).setValue(formattedDate); // A1セルにデータを書き込む（セル範囲を必要に応じて変更してください）

  return ContentService.createTextOutput(formattedDate + 'で退勤登録しました');
}
