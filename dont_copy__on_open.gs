/**
 * Copyright 2024 Google LLC
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      https://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific lan  guage governing permissions and
 *  limitations under the License.
 */

const MCC_ID = SpreadsheetApp.getActiveSpreadsheet()
  .getRangeByName(NamedRanges.MCC_ID)
  .getValue();


const SUPPORT_EMAIL = 'eladb+cad_feedback@google.com';

/**
 * Creates a custom menu with an option to send feedback directly in the tool
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('CAD')
    .addItem('Send Feedback', 'sendFeedback')
    .addItem('Consent screen', 'showConsentScreen')
    .addToUi();
  logFunctionRun("Opened CAD");
  if (isFirstOpen()) {
    showConsentScreen();
  }
}

/**
 * Checks if this is first time the spreadsheet has been opened
 */
function showConsentScreen() {
  const message = `Would you like to grant view access to CAD Google developers?

  • Help debugging when needed
  • Investigate errors in app-script (not ads-script)

  You will be able to remove this access at any time from the native user sharing menu.`;

  const userResponse = SpreadsheetApp.getUi().alert(message, SpreadsheetApp.getUi().ButtonSet.YES_NO);

  if (userResponse == SpreadsheetApp.getUi().Button.YES) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.addViewer(SUPPORT_EMAIL);
      SpreadsheetApp.getUi().alert('Sharing was a success');
    } catch (err) {
      SpreadsheetApp.getUi().alert(`Error: ${err.message}. Please manually add ${SUPPORT_EMAIL} as a viewer.`);
    }
  }
}


/**
 * Checks if this is first time the spreadsheet has been opened
 */
function isFirstOpen() {
  const ps = PropertiesService.getScriptProperties();
  let loginCheck = ps.getProperty('First Login');
  if (!loginCheck) {
    ps.setProperty('First Login', 'YES');
    return true;
  } else {
    return false;
  }
}


/**
 * Adds ability for users to send a feedback email directly from the UI
 */
function sendFeedback() {
  try {
    let ui = SpreadsheetApp.getUi();
    response = ui.alert(
      'You can provide feedback on this solution by sending an email to '+SUPPORT_EMAIL+'.\n Alternatively, you can provide feedback directly via this dialogue box.  Would like to proceed with that option?',
      ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      let subject = ui.prompt('Please give a short title for your feedback');
      if (subject.getSelectedButton() != ui.Button.OK) {
        return;
      }
      let body = ui.prompt('Please provide your full feedback message below');
      if (body.getSelectedButton() != ui.Button.OK) {
        return;
      }
      MailApp.sendEmail({
        to: SUPPORT_EMAIL,
        subject: subject.getResponseText(),
        body: body.getResponseText()
      });
    }
    logFunctionRun('Send Feedback');
  } catch (err) {
    Browser.msgBox('ERROR', err.message, Browser.Buttons.OK);
    logFunctionRun('Send Feedback', err);
  }
}