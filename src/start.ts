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

import { createNewMonthSheet } from './createNewMonthSheet';
import {
  sendErrorMessageToSlack,
  sendSuccessMessageToSlack,
} from './sendMessageToSlack';

export function start(slackID: string, time: string | undefined) {
  const date = time ? new Date(time) : new Date();
  const sheetName = Utilities.formatDate(
    date,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'yyyyMM'
  );
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName(sheetName) ?? createNewMonthSheet(ss, sheetName);

  // startコマンドで不正な時刻が指定されたとき
  if (time && isNaN(date.getTime())) {
    sendErrorMessageToSlack(
      slackID,
      '入力された日付が不正でした。正常な日付を入力してください。'
    );
    return ContentService.createTextOutput();
  }

  // slackIDの列を探索
  let currentRow = 1;
  while (sheet.getRange(currentRow, 1).getValue()) {
    if (sheet.getRange(currentRow, 1).getValue() === slackID) {
      break;
    }
    currentRow += 1;
  }
  // なかった場合。登録アカウント一覧から探索https://script.google.com/macros/s/AKfycbwb40iaWBmJR4EWlr90mmj0aXreqL5VmejMnGUSpeqUvwsyroACO_fobhqk8GS_mEMBzg/exec
  if (!sheet.getRange(currentRow, 1).getValue()) {
    const accountListSheet = ss.getSheetByName('登録アカウント一覧');
    if (!accountListSheet) {
      sendErrorMessageToSlack(
        slackID,
        'シートがありません。管理者に問い合わせてください。'
      );
      return ContentService.createTextOutput();
    }
    let accountListCurrentRow = 1;
    while (accountListSheet.getRange(accountListCurrentRow, 1).getValue()) {
      if (
        accountListSheet.getRange(accountListCurrentRow, 1).getValue() ===
        slackID
      ) {
        // アカウント一覧にあったらもとのシートの末尾に追加
        sheet
          .getRange(currentRow, 1)
          .setValue(
            accountListSheet.getRange(accountListCurrentRow, 1).getValue()
          );
        sheet
          .getRange(currentRow, 2)
          .setValue(
            accountListSheet.getRange(accountListCurrentRow, 2).getValue()
          );
        break;
      }
      accountListCurrentRow += 1;
    }
    // アカウント一覧になかった場合
    if (!accountListSheet.getRange(currentRow, 1).getValue()) {
      sendErrorMessageToSlack(
        slackID,
        'アカウントが見つかりませんでした。/registerコマンドで登録してください。\nアカウントを登録している場合は管理者に問い合わせてください。'
      );
      return ContentService.createTextOutput();
    }
  }

  // slackIDの行を探索。出退勤時刻は3列目から
  let currentIndex = 3;
  while (sheet.getRange(currentRow, currentIndex).getValue()) {
    currentIndex += 1;
  }
  // 最後の列(currentIndex-1)が奇数だったら出勤中
  if (currentIndex % 2 === 0) {
    sendErrorMessageToSlack(
      slackID,
      '退勤されていません。出勤状態を確認してください。'
    );
    return ContentService.createTextOutput();
  }

  const formattedDate = Utilities.formatDate(
    date,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
  sheet.getRange(currentRow, currentIndex).setValue(formattedDate);

  sendSuccessMessageToSlack(slackID, '出勤しました');
  return ContentService.createTextOutput();
}
