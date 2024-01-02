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
import { formatDate } from './utils/formatDate';

export function stop(slackID: string, time: string | undefined) {
  // 登録する時刻を取得
  const date = time ? new Date(time) : new Date();

  const sheetName = formatDate(date, 'yyyyMM');
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
    if (currentIndex === 3) {
      const previousMonth = new Date(date);
      previousMonth.setMonth(previousMonth.getMonth() - 1);
      const previousMonthSheetName = formatDate(previousMonth, 'yyyyMM');
      const previousMonthSheet = ss.getSheetByName(previousMonthSheetName);
      const lastColumn = previousMonthSheet?.getLastColumn() ?? 1;
      const lastRow = previousMonthSheet?.getLastRow() ?? 1;
      const range = previousMonthSheet
        ?.getRange(1, 1, lastRow, lastColumn)
        .getValues() ?? [[]];
      const row = range.find(row => row[0] === slackID);
      const rowNumber = range.findIndex(row => row[0] === slackID);
      const lastCellIndex = row?.filter(cell => !!cell)?.length ?? -43;
      if (!lastCellIndex) {
        return ContentService.createTextOutput(row?.join(',') ?? '');
      }
      if (lastCellIndex % 2 === 1) {
        const endOfPreviousDate = new Date(date);
        endOfPreviousDate.setDate(date.getDate() - 1);
        endOfPreviousDate.setHours(23, 59, 59, 999);
        previousMonthSheet
          ?.getRange(rowNumber + 1, lastCellIndex + 1)
          .setValue(formatDate(endOfPreviousDate, 'yyyy-MM-dd HH:mm:ss'));
      }
      const startOfDate = new Date(date);
      startOfDate.setHours(0, 0, 0, 0);
      sheet
        .getRange(currentRow, currentIndex)
        .setValue(formatDate(startOfDate, 'yyyy-MM-dd HH:mm:ss'));
      const formattedDate = formatDate(date, 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(currentRow, currentIndex + 1).setValue(formattedDate);

      return ContentService.createTextOutput(rowNumber as unknown as string);
    }
    sendErrorMessageToSlack(
      slackID,
      '出勤されていません。出勤状態を確認してください。'
    );
    return ContentService.createTextOutput();
  }

  const formattedDate = formatDate(date, 'yyyy-MM-dd HH:mm:ss');
  const prevCell = new Date(
    sheet.getRange(currentRow, currentIndex - 1).getValue()
  );

  if (prevCell.getDate() !== date.getDate()) {
    const endOfPreviousDate = new Date(date);
    endOfPreviousDate.setDate(date.getDate() - 1);
    endOfPreviousDate.setHours(23, 59, 59, 999);

    const startOfCurrentDate = new Date(date);
    startOfCurrentDate.setHours(0, 0, 0, 0);

    sheet
      .getRange(currentRow, currentIndex)
      .setValue(formatDate(endOfPreviousDate, 'yyyy-MM-dd HH:mm:ss'));
    sheet
      .getRange(currentRow, currentIndex + 1)
      .setValue(formatDate(startOfCurrentDate, 'yyyy-MM-dd HH:mm:ss'));
    sheet.getRange(currentRow, currentIndex + 2).setValue(formattedDate);
    return ContentService.createTextOutput('date に入っているよ');
  }
  // slackIDが登録されていて、退勤状態のときのみ出勤登録できる
  sheet.getRange(currentRow, currentIndex).setValue(formattedDate);
  sendSuccessMessageToSlack(slackID, '退勤しました');
  return ContentService.createTextOutput();
}
