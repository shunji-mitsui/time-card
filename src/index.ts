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

import { createNewMonthSheet } from './utils/createNewMonthSheet';
import { register } from './register';
import { sendErrorMessageToSlack } from './utils/sendMessageToSlack';
import { start } from './start';
import { stop } from './stop';
import { formatDate } from './utils/formatDate';

// TODO:リンクから飛ぶと稼働表が見れるようにする？
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doGet() {
  //
}

// MEMO:月末にスケジュールして、翌月のシートを作成
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createNextMonthSheet() {
  const date = new Date();
  date.setMonth(date.getMonth() + 1);
  const sheetName = formatDate(date, 'yyyyMM');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createNewMonthSheet(ss, sheetName);
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(e: GoogleAppsScript.Events.DoPost) {
  const startTime = new Date();
  const slackID = e.parameter.user_id;
  const command = e.parameter.command;
  const text = e.parameter.text;
  try {
    switch (command) {
      case '/start':
        start(slackID, text);
        break;
      case '/stop':
        stop(slackID, text);
        break;
      case '/register':
        register(slackID, text);
        break;
      default:
        sendErrorMessageToSlack(
          slackID,
          '無効なコマンドです。管理者に問い合わせてください'
        );
        break;
    }
  } catch (error) {
    sendErrorMessageToSlack(slackID, error as string);
  }
  const endTime = new Date();
  return ContentService.createTextOutput(
    `実行時間:${endTime.getTime() - startTime.getTime()}ms`
  );
}
