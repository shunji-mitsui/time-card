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
import { getSheetRange } from './getSheetRange';
import { setValueToCell } from './setValueToCell';

export function copyRowFromAccountList(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: string,
  targetRow: number,
  slackID: string
) {
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
    targetRow,
    0,
    accountListSheet[accountListCurrentRow][0]
  );
  setValueToCell(
    ss,
    sheetName,
    targetRow,
    1,
    accountListSheet[accountListCurrentRow][1]
  );
  return getSheetRange(ss, sheetName);
}
