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
import { sendSuccessMessageToSlack } from './sendMessageToSlack';
import { formatDate } from './utils/formatDate';
import { getSheetRange } from './utils/getSheetRange';
import { setValueToCell } from './utils/setValueToCell';
import { validateInputDate } from './utils/validateInputDate';

export function stop(slackID: string, time: string | undefined) {
  // 登録する時刻を取得
  const date = time ? new Date(time) : new Date();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = formatDate(date, 'yyyyMM');

  let testSheetRange = getSheetRange(ss, sheetName);

  // stopコマンドで不正な時刻が指定されたとき
  // TODO : 時間が一個前の出勤時刻よりも過去の場合のハンドリングをする
  validateInputDate(time, date);

  // slackID検索
  const currentRow =
    testSheetRange.findIndex(row => row[0] === slackID) !== -1
      ? testSheetRange.findIndex(row => row[0] === slackID)
      : testSheetRange.length;

  // slackIDが見つからない場合
  if (currentRow === testSheetRange.length) {
    const accountListSheet = getSheetRange(ss, '登録アカウント一覧');
    const accountListCurrentRow = accountListSheet.findIndex(
      row => row[0] === slackID
    );
    if (!accountListCurrentRow) {
      throw new Error();
    }
    // 対象のシートに新しくアカウント一覧からコピーしてくる
    setValueToCell(
      ss,
      sheetName,
      currentRow,
      0,
      accountListSheet[accountListCurrentRow][0]
    );
    setValueToCell(
      ss,
      sheetName,
      currentRow,
      1,
      accountListSheet[accountListCurrentRow][1]
    );
    testSheetRange = getSheetRange(ss, sheetName);
  }

  // slackIDの行を探索。
  const currentIndex =
    testSheetRange[currentRow].findIndex(cell => !cell) !== -1
      ? testSheetRange[currentRow].findIndex(cell => !cell)
      : testSheetRange[currentRow].length;

  // 最後の列(currentIndex)が偶数だったら退勤中
  if (currentIndex % 2 === 0) {
    if (currentIndex !== 2) {
      throw new Error('出勤されていません。出勤状態を確認してください。');
    }
    const previousMonth = new Date(date);
    previousMonth.setMonth(previousMonth.getMonth() - 1);
    const previousMonthSheetName = formatDate(previousMonth, 'yyyyMM');
    const range = getSheetRange(ss, previousMonthSheetName);
    const row = range.find(row => row[0] === slackID);
    const rowIndex = range.findIndex(row => row[0] === slackID);
    const lastCellIndex = row?.filter(cell => !!cell)?.length;
    // TODO
    if (!(lastCellIndex && lastCellIndex % 2 === 1)) {
      throw new Error('出勤されていません。出勤状態を確認してください。');
    }
    const endOfPreviousDate = new Date(date);
    endOfPreviousDate.setDate(date.getDate() - 1);
    endOfPreviousDate.setHours(23, 59, 59, 999);
    setValueToCell(
      ss,
      previousMonthSheetName,
      rowIndex,
      lastCellIndex,
      formatDate(endOfPreviousDate, 'yyyy-MM-dd HH:mm:ss')
    );
    const startOfDate = new Date(date);
    startOfDate.setHours(0, 0, 0, 0);
    setValueToCell(
      ss,
      sheetName,
      currentRow,
      currentIndex,
      formatDate(startOfDate, 'yyyy-MM-dd HH:mm:ss')
    );
    setValueToCell(
      ss,
      sheetName,
      currentRow,
      currentIndex + 1,
      formatDate(date, 'yyyy-MM-dd HH:mm:ss')
    );
    return;
  }

  const prevCell = new Date(testSheetRange[currentRow][currentIndex - 1]);

  if (prevCell.getDate() !== date.getDate()) {
    const endOfPreviousDate = new Date(date);
    endOfPreviousDate.setDate(date.getDate() - 1);
    endOfPreviousDate.setHours(23, 59, 59, 999);

    const startOfCurrentDate = new Date(date);
    startOfCurrentDate.setHours(0, 0, 0, 0);

    setValueToCell(
      ss,
      sheetName,
      currentRow,
      currentIndex,
      formatDate(endOfPreviousDate, 'yyyy-MM-dd HH:mm:ss')
    );
    setValueToCell(
      ss,
      sheetName,
      currentRow,
      currentIndex + 1,
      formatDate(startOfCurrentDate, 'yyyy-MM-dd HH:mm:ss')
    );
    setValueToCell(
      ss,
      sheetName,
      currentRow,
      currentIndex + 2,
      formatDate(date, 'yyyy-MM-dd HH:mm:ss')
    );
    return;
  }
  setValueToCell(
    ss,
    sheetName,
    currentRow,
    currentIndex,
    formatDate(date, 'yyyy-MM-dd HH:mm:ss')
  );
  sendSuccessMessageToSlack(slackID, '退勤しました');
  return;
}
