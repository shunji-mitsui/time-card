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
function sendMessageToSlack(userId: string, message: string, color: string) {
  const slackWebhookURL =
    'https://hooks.slack.com/services/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';
  const payload = {
    attachments: [
      {
        color,
        text: `<@${userId}> \n${message}`,
      },
    ],
  };

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(slackWebhookURL, options);
}

export function sendSuccessMessageToSlack(userId: string, message: string) {
  sendMessageToSlack(userId, message, '2EB886');
}

export function sendErrorMessageToSlack(userId: string, message: string) {
  sendMessageToSlack(userId, message, 'E01E5A');
}
