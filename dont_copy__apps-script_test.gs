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
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */

function setCustomDates_test() {
  setCustomDates(4, 21, "33");
}

function setCustomDates_changeCurrent_test() {
  setCustomDates(4, 20, "11");
}

function setCustomDates_changeRangeByName_test() {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const avg_weekday_past = spreadSheet.getRangeByName(appScriptConsts.AVG_WEEKDAY_PAST);
  setCustomDates(avg_weekday_past.getColumn(), avg_weekday_past.getRow(), "5");
}



function togglePeriodsVisibility_test() {
  togglePeriodsVisibility(3, 14, "Daily Avg. Full days.");
}

function timeUtils_test(){
  const minusFive = TimeUtils.getInstance().getLastQueryableHourMinusHours(5);
  var a = "1";
}