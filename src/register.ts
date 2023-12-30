/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
export function register(slackID: string, name: string) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 現在開いているスプレッドシートを取得
  const sheet = ss.getSheetByName('登録アカウント一覧'); // 書き込むシートを指定（シート名を変更してください）

  if (!sheet) {
    return ContentService.createTextOutput(
      'シートがありません。管理者に問い合わせてください。'
    );
  }

  if (!name) {
    return ContentService.createTextOutput('名前が適切でありません。');
  }

  let currentRow = 1;
  while (sheet.getRange('A' + currentRow).getValue()) {
    if (sheet.getRange('A' + currentRow).getValue() === slackID) {
      const prevName = sheet.getRange('B' + currentRow).getValue();
      sheet.getRange('B' + currentRow).setValue(name);
      return ContentService.createTextOutput(
        '表示名を' + prevName + 'から' + name + 'に変更しました'
      );
    }
    currentRow += 1;
  }
  sheet.getRange('A' + currentRow).setValue(slackID);
  sheet.getRange('B' + currentRow).setValue(name);
  return ContentService.createTextOutput(name + 'として新規登録しました。');
}
