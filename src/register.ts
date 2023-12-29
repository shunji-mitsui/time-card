export function register(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  slackID: string,
  name: string
) {
  if (!name) {
    return ContentService.createTextOutput('名前が適切でありません。');
  }
  let cellRow = 1;
  while (sheet.getRange('A' + cellRow).getValue()) {
    if (sheet.getRange('A' + cellRow).getValue() === slackID) {
      const prevName = sheet.getRange('B' + cellRow).getValue();
      sheet.getRange('B' + cellRow).setValue(name);
      return ContentService.createTextOutput(
        '表示名を' + prevName + 'から' + name + 'に変更しました'
      );
    }
    cellRow += 1;
  }
  sheet.getRange('A' + cellRow).setValue(slackID);
  sheet.getRange('B' + cellRow).setValue(name);
  return ContentService.createTextOutput(name + 'として新規登録しました。');
}
