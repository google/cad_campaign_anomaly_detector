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

/** Configuration to be used for the Campaign Anomaly Detector. */
const CONFIG = {
  /**
   * URL of the copy from the default spreadsheet template.
   * https://docs.google.com/spreadsheets/d/1ki-fYL3CjKsU-ems174M42NJ4TMfaWo3SPw-oZYuOvs/copy
   @const {string}
   */
  spreadsheet_url: "-----",

  /**
   Should print debug log
   @public @const {boolean}
   */
  is_debug_log: true,

  /**
   More reporting options can be found at
   https://developers.google.com/google-ads/scripts/docs/reference/adsapp/adsapp#report_2
   @public @const {{apiVersion: string}}
   */
  reporting_options: {
    // Comment out the following line to default to the latest reporting
    // version.
    apiVersion: "v20",
  },
};
