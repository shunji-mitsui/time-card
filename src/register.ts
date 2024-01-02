import { getSlackIdRow } from './utils/getSlackIdIndex';
import { sendSuccessMessageToSlack } from './utils/sendMessageToSlack';
import { getSheetRange } from './utils/getSheetRange';
import { setValueToCell } from './utils/setValueToCell';

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
  if (!name) {
    throw new Error('名前の入力が不正です。正しく入力してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRange = getSheetRange(ss, '登録アカウント一覧');

  const index = getSlackIdRow(sheetRange, slackID);

  if (index === sheetRange.length) {
    setValueToCell(ss, '登録アカウント一覧', index, 0, slackID);
    setValueToCell(ss, '登録アカウント一覧', index, 1, name);
    sendSuccessMessageToSlack(slackID, `${name}として新規登録しました。`);
    return;
  }
  sheetRange[index][0];

  const prevName = sheetRange[index][1];
  setValueToCell(ss, '登録アカウント一覧', index, 1, name);
  sendSuccessMessageToSlack(
    slackID,
    `表示名を${prevName}から${name}に変更しました`
  );
}
