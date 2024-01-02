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

import { copyRowFromAccountList } from './utils/copyRowFromAccountList';
import { getLastColumnOfAccount } from './utils/getLastColumnOfAccount';
import { getSlackIdRow } from './utils/getSlackIdIndex';
import { sendSuccessMessageToSlack } from './utils/sendMessageToSlack';
import { formatDate } from './utils/formatDate';
import { getSheetRange } from './utils/getSheetRange';
import { setValueToCell } from './utils/setValueToCell';
import { validateInputDate } from './utils/validateInputDate';

export function start(slackID: string, time: string | undefined) {
  const date = time ? new Date(time) : new Date();
  const sheetName = formatDate(date, 'yyyyMM');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let testSheetRange = getSheetRange(ss, sheetName);

  validateInputDate(time, date);

  // slackIDの列を探索
  const currentRow = getSlackIdRow(testSheetRange, slackID);

  // なかった場合。登録アカウント一覧から探索
  if (currentRow === testSheetRange.length) {
    testSheetRange = copyRowFromAccountList(ss, sheetName, currentRow, slackID);
  }

  // slackIDの行を探索。出退勤時刻は3列目から
  const currentIndex = getLastColumnOfAccount(testSheetRange, currentRow);
  // 最後の列(currentIndex)が奇数だったら出勤中
  if (currentIndex % 2 === 1) {
    throw new Error('退勤されていません。出勤状態を確認してください。');
  }

  setValueToCell(
    ss,
    sheetName,
    currentRow,
    currentIndex,
    formatDate(date, 'yyyy-MM-dd HH:mm:ss')
  );

  sendSuccessMessageToSlack(slackID, '出勤しました');
}
