import {
  sendErrorMessageToSlack,
  sendSuccessMessageToSlack,
} from './sendMessageToSlack';

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
    sendErrorMessageToSlack(
      slackID,
      'シートがありません。管理者に問い合わせてください。'
    );
    return ContentService.createTextOutput();
  }

  if (!name) {
    sendErrorMessageToSlack(
      slackID,
      '名前の入力が不正です。正しく入力してください。'
    );
    return ContentService.createTextOutput();
  }

  let currentRow = 1;
  while (sheet.getRange('A' + currentRow).getValue()) {
    if (sheet.getRange('A' + currentRow).getValue() === slackID) {
      const prevName = sheet.getRange('B' + currentRow).getValue();
      sheet.getRange('B' + currentRow).setValue(name);
      sendSuccessMessageToSlack(
        slackID,
        `表示名を${prevName}から${name}に変更しました`
      );
      return ContentService.createTextOutput();
    }
    currentRow += 1;
  }
  sheet.getRange('A' + currentRow).setValue(slackID);
  sheet.getRange('B' + currentRow).setValue(name);
  sendSuccessMessageToSlack(slackID, `${name}として新規登録しました。`);
  return ContentService.createTextOutput();
}
