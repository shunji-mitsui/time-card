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

export function stop(slackID: string, time: string | undefined) {
  // 登録する時刻を取得
  const date = time ? new Date(time) : new Date();
  const sheetName = Utilities.formatDate(
    date,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'yyyyMM'
  );
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName(sheetName) ?? createNewMonthSheet(ss, sheetName);

  // stopコマンドで不正な時刻が指定されたとき
  // TODO : 時間が一個前の出勤時刻よりも過去の場合のハンドリングをする
  if (time && isNaN(date.getTime())) {
    sendErrorMessageToSlack(
      slackID,
      '入力された日付が不正でした。正常な日付を入力してください。'
    );
    return ContentService.createTextOutput();
  }

  // slackID検索
  let currentRow = 1;
  while (sheet.getRange(currentRow, 1).getValue()) {
    if (slackID === sheet.getRange(currentRow, 1).getValue()) {
      break;
    }
    currentRow += 1;
  }

  // slackIDが見つからない場合
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
        'アカウントが見つかりませんでした。/registerコマンドで登録してください。\n アカウントを登録してあるばあいは管理者に問い合わせて下さい。'
      );
      return ContentService.createTextOutput();
    }
  }

  // slackIDの行を探索。出退勤時刻は3列目から
  let currentIndex = 3;
  while (sheet.getRange(currentRow, currentIndex).getValue()) {
    currentIndex += 1;
  }

  // 最後の列(currentIndex-1)が偶数だったら退勤中
  if (currentIndex % 2 === 1) {
    sendErrorMessageToSlack(
      slackID,
      '出勤されていません。出勤状態を確認してください。'
    );
    return ContentService.createTextOutput();
  }

  const formattedDate = Utilities.formatDate(
    date,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
  // slackIDが登録されていて、退勤状態のときのみ出勤登録できる
  sheet.getRange(currentRow, currentIndex).setValue(formattedDate);

  sendSuccessMessageToSlack(slackID, '退勤しました');
  return ContentService.createTextOutput();
}
