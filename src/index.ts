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

import { register } from './register';
import { start } from './start';
import { stop } from './stop';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(e: GoogleAppsScript.Events.DoPost) {
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
