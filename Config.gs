/** Configuration to be used for the Campaign Anomaly Detector. */
const CONFIG = {
  /**
   * URL of the copy from the default spreadsheet template.
   * https://docs.google.com/spreadsheets/d/1ki-fYL3CjKsU-ems174M42NJ4TMfaWo3SPw-oZYuOvs/copy
   @const {string}
   */
  spreadsheet_url: '-----',


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
      'apiVersion': 'v14'
  }
};
