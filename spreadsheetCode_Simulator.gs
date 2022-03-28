function updateInput(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.getRangeByName(CONST.CURRENT_RANGE_DAY).setValue(spreadsheet.getRangeByName(CONST.SIMULATOR_CURRENT_DAY).getValue());
  spreadsheet.getRangeByName(CONST.CURRENT_RANGE_PERIOD).setValue(spreadsheet.getRangeByName(CONST.SIMULATOR_CURRENT_PERIOD).getValue());

    spreadsheet.getRangeByName(CONST.PAST_RANGE_DAY).setValue(spreadsheet.getRangeByName(CONST.SIMULATOR_PAST_DAY).getValue());
  spreadsheet.getRangeByName(CONST.PAST_RANGE_PERIOD).setValue(spreadsheet.getRangeByName(CONST.SIMULATOR_PAST_PERIOD).getValue());
}
