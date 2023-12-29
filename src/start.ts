export function start(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  slackID: string
) {
  let target_row = 1;
  while (sheet.getRange('A' + target_row).getValue()) {
    if (sheet.getRange('A' + target_row).getValue() === slackID) {
      break;
    }
    if (target_row === 100) {
      break;
    }
    target_row += 1;
  }
  if (!sheet.getRange(target_row, 1).getValue() || target_row === 100) {
    sheet.getRange(1, 10).setValue(target_row);
    sheet.getRange(1, 10).setValue(slackID);
    return ContentService.createTextOutput(
      'アカウントが見つかりませんでした。/registerコマンドで登録してください'
    );
  }

  let target_index = 3;
  while (sheet.getRange(target_row, target_index).getValue()) {
    target_index += 1;
  }
  if (target_index % 2 === 0) {
    return ContentService.createTextOutput(
      '退勤されていません。出勤状態を確認してください。'
    );
  }
  const currentDate = new Date();
  const formattedDate = Utilities.formatDate(
    currentDate,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
  sheet.getRange(target_row, target_index).setValue(formattedDate); // A1セルにデータを書き込む（セル範囲を必要に応じて変更してください）

  return ContentService.createTextOutput(formattedDate + 'で出勤登録しました');
}
