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
import { sendErrorMessageToSlack } from './sendMessageToSlack';
import { start } from './start';
import { stop } from './stop';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function doPost(e: GoogleAppsScript.Events.DoPost) {
  const slackID = e.parameter.user_id;
  const command = e.parameter.command;
  const text = e.parameter.text;
  switch (command) {
    case '/start':
      return start(slackID, text);
    case '/stop':
      return stop(slackID, text);
    case '/register':
      return register(slackID, text);
    default:
      sendErrorMessageToSlack(
        slackID,
        '無効なコマンドです。管理者に問い合わせてください'
      );
      return ContentService.createTextOutput();
  }
}
