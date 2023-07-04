const appScriptConsts = {
  "AVG_TYPE_DROPDOWN": "AVG_TYPE",

  "AVG_WEEKDAY_PAST": "AVG_WEEKDAY_PAST",

  "AVG_DAILY_CURRENT": "AVG_DAILY_CURRENT",
  "AVG_DAILY_PAST": "AVG_DAILY_PAST",

  "AVG_WEEKLY_CURRENT": "AVG_WEEKLY_CURRENT",
  "AVG_WEEKLY_PAST": "AVG_WEEKLY_PAST",

  "AVG_TODAY_VS_YESTERDAY": "AVG_TODAY_VS_YESTERDAY",

  "AVG_LAST_HOUR_VS_BEFORE": "AVG_LAST_HOUR_VS_BEFORE",

  "CURRENT_PERIOD_LENGTH": "current_period_length",
  "CURRENT_PERIOD_UNIT": "current_period_unit",

  "CURRENT_ENDED_LENGTH_AGO": "current_ended_length_ago",
  "CURRENT_END_UNIT": "current_end_unit",

  "PAST_PERIOD_LENGTH": "past_period_length",
  "PAST_PERIOD_UNIT": "past_period_unit",

  "PAST_ENDED_LENGTH_AGO": "past_ended_length_ago",
  "PAST_END_UNIT": "past_end_unit",

  "DATA_AGGREGATION": "data_aggregation",
  "ALL_CAMPAIGN_FOR_ACCOUNTS": "all_campaigns_for_accounts"
}

const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const sheet = spreadSheet.getActiveSheet();


let appScriptNamedRange = {
      isInit: true,
      current_period_length: spreadSheet.getRangeByName(appScriptConsts.CURRENT_PERIOD_LENGTH),
      current_period_unit: spreadSheet.getRangeByName(appScriptConsts.CURRENT_PERIOD_UNIT),
      current_ended_length_ago: spreadSheet.getRangeByName(appScriptConsts.CURRENT_ENDED_LENGTH_AGO),
      current_end_unit: spreadSheet.getRangeByName(appScriptConsts.CURRENT_END_UNIT),
      past_period_length: spreadSheet.getRangeByName(appScriptConsts.PAST_PERIOD_LENGTH),
      past_period_unit: spreadSheet.getRangeByName(appScriptConsts.PAST_PERIOD_UNIT),
      past_ended_length_ago: spreadSheet.getRangeByName(appScriptConsts.PAST_ENDED_LENGTH_AGO),
      past_end_unit: spreadSheet.getRangeByName(appScriptConsts.PAST_END_UNIT),
      avg_type_dropdown: spreadSheet.getRangeByName(appScriptConsts.AVG_TYPE_DROPDOWN),
      avg_weekday_past: spreadSheet.getRangeByName(appScriptConsts.AVG_WEEKDAY_PAST),
      avg_weekly_current: spreadSheet.getRangeByName(appScriptConsts.AVG_WEEKLY_CURRENT),
      avg_daily_current: spreadSheet.getRangeByName(appScriptConsts.AVG_DAILY_CURRENT),
      avg_daily_past: spreadSheet.getRangeByName(appScriptConsts.AVG_DAILY_PAST),
      avg_weekly_past: spreadSheet.getRangeByName(appScriptConsts.AVG_WEEKLY_PAST),
      today_vs_yes: spreadSheet.getRangeByName(appScriptConsts.AVG_TODAY_VS_YESTERDAY),
      last_hour_vs_before: spreadSheet.getRangeByName(appScriptConsts.AVG_LAST_HOUR_VS_BEFORE),
      data_aggregation: spreadSheet.getRangeByName(appScriptConsts.DATA_AGGREGATION),
      all_campaign_for_accounts : spreadSheet.getRangeByName(appScriptConsts.ALL_CAMPAIGN_FOR_ACCOUNTS)
    }

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEdit(e) {
  const range = e.range;
  var row = range.getRow();
  var column = range.getColumn();
  var value = range.getValue();

  if (!appScriptNamedRange.isInit) {
    appScriptNamedRange = initNamedRanges();
  }

  toggleAggregationVisibility(column, row, value);
  togglePeriodsVisibility(column, row, value);
  toggleMetricTableRowsVisibility(column, row, value);
  setCustomDates(column, row, value);
}

function initNamedRanges(){
  return {
      isInit: true,
      current_period_length: spreadSheet.getRangeByName(appScriptConsts.CURRENT_PERIOD_LENGTH),
      current_period_unit: spreadSheet.getRangeByName(appScriptConsts.CURRENT_PERIOD_UNIT),
      current_ended_length_ago: spreadSheet.getRangeByName(appScriptConsts.CURRENT_ENDED_LENGTH_AGO),
      current_end_unit: spreadSheet.getRangeByName(appScriptConsts.CURRENT_END_UNIT),
      past_period_length: spreadSheet.getRangeByName(appScriptConsts.PAST_PERIOD_LENGTH),
      past_period_unit: spreadSheet.getRangeByName(appScriptConsts.PAST_PERIOD_UNIT),
      past_ended_length_ago: spreadSheet.getRangeByName(appScriptConsts.PAST_ENDED_LENGTH_AGO),
      past_end_unit: spreadSheet.getRangeByName(appScriptConsts.PAST_END_UNIT),
      avg_type_dropdown: spreadSheet.getRangeByName(appScriptConsts.AVG_TYPE_DROPDOWN),
      avg_weekday_past: spreadSheet.getRangeByName(appScriptConsts.AVG_WEEKDAY_PAST),
      avg_weekly_current: spreadSheet.getRangeByName(appScriptConsts.AVG_WEEKLY_CURRENT),
      avg_daily_current: spreadSheet.getRangeByName(appScriptConsts.AVG_DAILY_CURRENT),
      avg_daily_past: spreadSheet.getRangeByName(appScriptConsts.AVG_DAILY_PAST),
      avg_weekly_past: spreadSheet.getRangeByName(appScriptConsts.AVG_WEEKLY_PAST),
      today_vs_yes: spreadSheet.getRangeByName(appScriptConsts.AVG_TODAY_VS_YESTERDAY),
      last_hour_vs_before: spreadSheet.getRangeByName(appScriptConsts.AVG_LAST_HOUR_VS_BEFORE),
      data_aggregation: spreadSheet.getRangeByName(appScriptConsts.DATA_AGGREGATION),
      all_campaign_for_accounts : spreadSheet.getRangeByName(appScriptConsts.ALL_CAMPAIGN_FOR_ACCOUNTS)
    }
}

function togglePeriodsVisibility(column, row, value) {
  const avgWeekdayPast = appScriptNamedRange.avg_weekday_past;
  const todayVsYesterday = appScriptNamedRange.today_vs_yes;
  const avg_type_dropdown = appScriptNamedRange.avg_type_dropdown;

  if ((column == avg_type_dropdown.getColumn()) && (row == avg_type_dropdown.getRow())) {
    switch (value) {
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
        {
          sheet.hideRows(avgWeekdayPast.getRow() - 1, 8);
          sheet.showRows(appScriptNamedRange.last_hour_vs_before.getRow(), 2);

          sheet.hideRows(current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 6);
          sheet.showRows(todayVsYesterday.getRow(), 2);
          sheet.hideRows(appScriptNamedRange.last_hour_vs_before.getRow(), 2);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
        {
          sheet.showRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 2);
          appScriptNamedRange.avg_weekday_past.setValue("0");
          sheet.hideRows(appScriptNamedRange.avg_daily_current.getRow(), 8);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 2);
          sheet.showRows(appScriptNamedRange.avg_daily_current.getRow(), 2);
          appScriptNamedRange.avg_daily_current.setValue("0");
          appScriptNamedRange.avg_daily_past.setValue("0");
          sheet.hideRows(appScriptNamedRange.avg_weekly_current.getRow(), 6);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_WEEKLY:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 4);
          sheet.showRows(appScriptNamedRange.avg_weekly_current.getRow(), 2);
          appScriptNamedRange.avg_weekly_current.setValue("0");
          appScriptNamedRange.avg_weekly_past.setValue("0");
          sheet.hideRows(appScriptNamedRange.today_vs_yes.getRow(), 4);

          sheet.hideRows(current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_CUSTOM:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 10);

          sheet.showRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
    }
  }
}

function toggleMetricTableRowsVisibility(column, row, value) {
  const expectedCall = appScriptNamedRange.avg_type_dropdown;
  if ((column == expectedCall.getColumn()) && (row == expectedCall.getRow())) {
    switch (value) {
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 8);
          sheet.showRows(appScriptNamedRange.last_hour_vs_before.getRow(), 2);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 6);
          sheet.showRows(appScriptNamedRange.today_vs_yes.getRow(), 2);
          sheet.hideRows(appScriptNamedRange.last_hour_vs_before.getRow(), 2);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
        {
          sheet.showRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 2);
          appScriptNamedRange.avg_weekday_past.setValue("0");
          sheet.hideRows(appScriptNamedRange.avg_daily_current.getRow(), 8);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY:
        {

          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 2);
          sheet.showRows(appScriptNamedRange.avg_daily_current.getRow(), 2);
          appScriptNamedRange.avg_daily_current.setValue("0");
          appScriptNamedRange.avg_daily_past.setValue("0");
          sheet.hideRows(appScriptNamedRange.avg_weekly_current.getRow(), 6);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_WEEKLY:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 4);
          sheet.showRows(appScriptNamedRange.avg_weekly_current.getRow(), 2);
          appScriptNamedRange.avg_weekly_current.setValue("0");
          appScriptNamedRange.avg_weekly_past.setValue("0");
          sheet.hideRows(appScriptNamedRange.today_vs_yes.getRow(), 4);

          sheet.hideRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
      case AVG_TYPE.AVG_TYPE_CUSTOM:
        {
          sheet.hideRows(appScriptNamedRange.avg_weekday_past.getRow() - 1, 10);

          sheet.showRows(appScriptNamedRange.current_period_length.getRow(), 9);
          break;
        }
    }
  }
}

function setCustomDatesForPastVsToday(pastEndedLengthAgo) {
  appScriptNamedRange.current_period_length.setValue(1);
  appScriptNamedRange.current_period_unit.setValue("Days");
  appScriptNamedRange.current_ended_length_ago.setValue(0);
  appScriptNamedRange.current_end_unit.setValue("Days");

  appScriptNamedRange.past_period_length.setValue(1);
  appScriptNamedRange.past_period_unit.setValue("Days");
  appScriptNamedRange.past_ended_length_ago.setValue(pastEndedLengthAgo);
  appScriptNamedRange.past_end_unit.setValue("Days");
}

function toggleAggregationVisibility(column, row, value) {
  const expectedCell = appScriptNamedRange.data_aggregation;
  if ((column == expectedCell.getColumn()) && (row == expectedCell.getRow())) {
    const row = appScriptNamedRange.all_campaign_for_accounts.getRow();    
    switch (value) {
      case "Account":
        {
          sheet.hideRows(row, 1);
          break;
        }
      case "Campaign":
        {
          sheet.showRows(row, 1);
          break;
        }
    }
  }
}


function setCustomDates(column, row, value) {
  //===========today
  const expectedCell = appScriptNamedRange.last_hour_vs_before;
  if ((column == expectedCell.getColumn()) && (row == expectedCell.getRow())) {
    setCustomDatesForPastVsToday(0);
  }

  //===========today vs. yesterday
  else if ((column == appScriptNamedRange.today_vs_yes.getColumn()) && (row == appScriptNamedRange.today_vs_yes.getRow())) {
    setCustomDatesForPastVsToday(1);
  }

  //===========avg_weekday
  else if ((column == appScriptNamedRange.avg_weekday_past.getColumn()) && (row == appScriptNamedRange.avg_weekday_past.getRow())) {
    appScriptNamedRange.current_period_length.setValue(1);
    appScriptNamedRange.current_period_unit.setValue("Days");
    appScriptNamedRange.current_ended_length_ago.setValue(0);
    appScriptNamedRange.current_end_unit.setValue("Days");

    appScriptNamedRange.past_period_length.setValue(Number(value) * 7);
    appScriptNamedRange.past_period_unit.setValue("Days");
    appScriptNamedRange.past_ended_length_ago.setValue(1);
    appScriptNamedRange.past_end_unit.setValue("Days");
  }

  //===========avg_daily
  else if ((column == appScriptNamedRange.avg_daily_current.getColumn()) && (row == appScriptNamedRange.avg_daily_current.getRow())) {
    appScriptNamedRange.current_period_length.setValue(value);
    appScriptNamedRange.current_period_unit.setValue("Days");
    appScriptNamedRange.current_ended_length_ago.setValue(1);
    appScriptNamedRange.current_end_unit.setValue("Days");

    appScriptNamedRange.past_ended_length_ago.setValue(Number(value));
    appScriptNamedRange.past_end_unit.setValue("Days");
  }
  else if ((column == appScriptNamedRange.avg_daily_past.getColumn()) && (row == appScriptNamedRange.avg_daily_past.getRow())) {
    appScriptNamedRange.past_period_length.setValue(value);
    let unit = appScriptNamedRange.past_period_unit.getValue();
    // Check if the current value is "Weeks" before changing it
    if (unit === "Weeks") {
      appScriptNamedRange.past_period_unit.setValue("Days");
    }
  }

  //===========avg_weekly
  else if ((column == appScriptNamedRange.avg_weekly_current.getColumn()) && (row == appScriptNamedRange.avg_weekly_current.getRow())) {
    appScriptNamedRange.current_period_length.setValue(value);
    appScriptNamedRange.current_period_unit.setValue("Weeks");
    appScriptNamedRange.current_ended_length_ago.setValue(1);
    appScriptNamedRange.current_end_unit.setValue("Days");

    appScriptNamedRange.past_ended_length_ago.setValue(7 * Number(value));
    appScriptNamedRange.past_end_unit.setValue("Days");
  }
  else if ((column == appScriptNamedRange.avg_weekly_past.getColumn()) && (row == appScriptNamedRange.avg_weekly_past.getRow())) {
    appScriptNamedRange.past_period_length.setValue(value);
    appScriptNamedRange.past_period_unit.setValue("Weeks");
  }
}

