/**
 * @name Campaign Anomaly Detector
 *
 * @fileoverview The Campaign Anomaly Detector alerts the advertiser whenever
 * one or more accounts/campaigns in a group of advertiser accounts under an MCC account
 * are suddenly behaving too differently from what's historically observed. See
 * https://developers.google.com/google-ads/scripts/docs/solutions/adsmanagerapp-account-anomaly-detector
 * for more details.
 *
 * @author: eladb@google.com
 * Date: 2022-5-26
**/


//USER_TO_DO: Fill in url
var SPREADSHEET_URL = '----';
var spreadsheet;
var resultsSheet;
var CONST = {

  FIRST_DATA_ROW: 7,
  FIRST_DATA_COLUMN: 1,
  MCC_CHILD_ACCOUNT_LIMIT: 50,
  TOTAL_DATA_COLUMNS: 38,

  ACCOUNT_IDS: "account_ids",
  EXCLUDED_ACCOUNT_IDS: "excluded_account_ids",

  ACCOUNT_LABELS: "account_labels",

  CAMPAIGN_IDS: "campaign_ids",
  EXCLUDED_CAMPAIGN_IDS: "excluded_campaign_ids",

  CAMPAIGN_LABELS: "campaign_labels",
  RESULTS_SHEET: "Results",


  CURRENT_RANGE_DATES: "current_range_dates",
  PAST_RANGE_DATES: "past_range_dates",
  LENGTH_UNIT: "length_unit",

  PAST_RANGE_PERIOD: 'past_range_period',
  PAST_RANGE_DAY: 'past_range_day',

  CURRENT_RANGE_PERIOD: 'current_range_period',
  CURRENT_RANGE_DAY: 'current_range_day',

  SIMULATOR_CURRENT_END: "simulator_current_end",
  SIMULATOR_CURRENT_START: "simulator_current_start",
  SIMULATOR_PAST_START: "simulator_past_start",
  SIMULATOR_PAST_END: "simulator_past_end",
  SIMULATOR_CURRENT_DAY: "simulator_current_day",
  SIMULATOR_CURRENT_PERIOD: "simulator_current_period",
  SIMULATOR_PAST_DAY: "simulator_past_day",
  SIMULATOR_PAST_PERIOD: "simulator_past_period",

  DATE: "date",
  ACCOUNT_ID: "account_id",

  EMAILS: "emails",

  COLOR_HIGH: "color_high",
  COLOR_LOW: "color_low",

  DEVIDE_COST_BY: 10000000, //10,000,000
  DEVIDE_IMPRESSIONS_BY: 10,
  DEVIDE_CLICKS_BY: 10

};

var STATS = {
  'LEVEL': "stats_level",
  'NumOfColumns': 9,
  'Id':
    { 'Column': 1, },

  //Native metrics
  'impressions':
    { 'Column': 6, 'GAQL_name': 'metrics.impressions', 'Named_range': 'impressions', 'Email_text': 'Impressions' },
  'clicks':
    { 'Column': 10, 'GAQL_name': 'metrics.clicks', 'Named_range': 'clicks', 'Email_text': 'Clicks' },
  'conversions':
    { 'Column': 14, 'GAQL_name': 'metrics.conversions', 'Named_range': 'conversions', 'Email_text': 'Conversions' },
  'cost_micros':
    { 'Column': 18, 'GAQL_name': 'metrics.cost_micros', 'Named_range': 'cost_micros', 'Email_text': 'Cost' },

  //Manually calculated
  'cost_per_conversion':
    { 'Column': 22, 'GAQL_name': 'cost_per_conversion', 'Named_range': 'cost_per_conversion', 'Email_text': 'CPA' },
  'ctr':
    { 'Column': 26, 'GAQL_name': 'ctr', 'Named_range': 'ctr', 'Email_text': 'CTR' },
  'conversions_from_interactions_rate':
    { 'Column': 30, 'GAQL_name': 'conversions_from_interactions_rate', 'Named_range': 'conversions_from_interactions_rate', 'Email_text': 'CVR' },
  'average_cpc':
    { 'Column': 34, 'GAQL_name': 'average_cpc', 'Named_range': 'average_cpc', 'Email_text': 'CPC' },
  'roi':
    { 'Column': 38, 'GAQL_name': 'roi', 'Named_range': 'roi', 'Email_text': 'ROI' },
  'all_conversions_value':
    { 'GAQL_name': 'metrics.all_conversions_value' }
};

// CPC - average_cpc
// CTR- ctr
// CVR- conversions_from_interactions_rate
// CPA- cost_per_conversion

/**
 * Configuration to be used for running reports.
 */
var REPORTING_OPTIONS = {
  apiVersion: 'v10'
};

function main() {
  var alertTextForMcc = [];
  mccManager = mccManager();

  spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  resultsSheet = spreadsheet.getSheetByName(CONST.RESULTS_SHEET);
  spreadsheet.setSpreadsheetTimeZone(AdsApp.currentAccount().getTimeZone());
  SheetUtil.cleanStats();
  SheetUtil.setupData(spreadsheet);
  var fullScanAccountIds = SheetUtil.getFullScanAccountIds();
  var campaignIds = SheetUtil.getCampaignIds();

  var allSubAccountsObj = (fullScanAccountIds[0].toUpperCase() != "ALL" && campaignIds == "") ? mccManager.getAccountsForIds(fullScanAccountIds) : mccManager.getAllSubAccounts();
  for (var index in allSubAccountsObj) {
    var currentAccount = allSubAccountsObj[index];
    Logger.log('Switching to account ' + currentAccount.getCustomerId());
    AdsManagerApp.select(currentAccount)
    var newStartingRow = Math.max(CONST.FIRST_DATA_ROW, resultsSheet.getLastRow() + 1);
    alertTextForMcc.push(generateAlertForSingleAccount(currentAccount, spreadsheet, newStartingRow));
  }

  sendEmail(mccManager.mccAccount(), alertTextForMcc, spreadsheet);
}

function getExistingAccountLabels(labels, account) {
  var accountLabelSelector = account.labels();
  var relevantLabels = [];
  for (i in labels) {
    var condition = "Name CONTAINS '" + labels[i] + "'";
    var iterator = accountLabelSelector.withCondition(condition).get();
    if (iterator.hasNext()) {
      iterator.next();
      relevantLabels.push(labels[i]);
    }
  }
  return relevantLabels;
}

function rowsToCampaignDict(campaigns, response, value) {
  rows = response.rows();
  while (rows.hasNext()) {
    var row = rows.next();
    var campaignId = row["campaign.id"];
    if (!campaigns[campaignId]) {
      campaigns[campaignId] = [];
    }
    campaigns[campaignId].push(row[value]);
  }
  return campaigns;
}

/**
 * For each of Impressions, Clicks, Conversions, and Cost, check to see if the
 * values are out of range. If they are, and no alert has been set in the
 * spreadsheet, then 1) Add text to the email, and 2) Add coloring to the cells
 * corresponding to the statistic.
 *
 * @return {string} the next piece of the alert text to include in the email.
 */
function generateAlertForSingleAccount(currentAccount, spreadsheet, currentRowToWrite) {
  var resultsSheet = spreadsheet.getSheetByName(CONST.RESULTS_SHEET);
  var currentStatsDict = {};
  var pastStatsDict = {};
  var fullScanAccountIds = SheetUtil.getFullScanAccountIds();
  var excludedFullAccountIds = SheetUtil.getExcludedFullAccountIds();
  var alertTextForAccount = [];
  var relevantLabelsForCurrentAccount = getExistingAccountLabels(SheetUtil.getAccountLabelsFilter(), currentAccount);
  Logger.log("current account matches these monitored account labels = " + relevantLabelsForCurrentAccount);

  function _fillDictsWithFullAccountScanStats(currentAccount) {
    var currentAccountId = currentAccount.getCustomerId();
    /** Is currenly scanned account match id filter*/
    if ((fullScanAccountIds.length > 0 && fullScanAccountIds[0].toUpperCase() == "ALL")
      || fullScanAccountIds.includes(currentAccountId) ||
      (SheetUtil.getAccountLabelsFilter()[0].length > 0 && relevantLabelsForCurrentAccount[0].length > 0)) {
      if (excludedFullAccountIds.includes(currentAccountId)) {
        return;
      }
      Logger.log("Full scan for account " + currentAccountId);
      var currentFullAccount = SheetUtil.getCurrentQueryForFullAccount();
      var currentFullAccount = AdsApp.report(currentFullAccount, REPORTING_OPTIONS);
      var pastFullAccount = AdsApp.report(SheetUtil.getPastQueryForFullAccount(), REPORTING_OPTIONS);
      populateStatsDict(currentStatsDict, currentFullAccount.rows());
      populateStatsDict(pastStatsDict, pastFullAccount.rows());
    }
  }

  function _getCampaignDictForCurrentAccount(currentAccount) {
    //Campiagn Dict, not dependednt on the chosen accounts
    //We iterate all the account anyways
    var campaignDictForCurrentAccount = {};
    var campaignIds = SheetUtil.getCampaignIds();
    var excludedCampaignIds = SheetUtil.getExcludedCampaignIds();
    var campaignLabels = SheetUtil.getCampaignLabels();
    var excludedClause = excludedCampaignIds == "" ? "" : "AND campaign.id NOT IN (" + excludedCampaignIds + ")";
    if (campaignIds.length > 0) {
      if (campaignIds.toUpperCase().includes("ALL_ENABLED_FOR_ACCOUNT_")) {
        var forAccountId = campaignIds.toUpperCase().replace("ALL_ENABLED_FOR_ACCOUNT_", "");
        if (currentAccount.getCustomerId() == forAccountId) {
          var query = "SELECT campaign.id, campaign.name FROM campaign";
          Logger.log("campaigns for current account query =" + query);
          campaignDictForCurrentAccount = rowsToCampaignDict(campaignDictForCurrentAccount, AdsApp.report(query, REPORTING_OPTIONS), "campaign.id");
        }
      } else if (campaignIds != "") {
        var query = "SELECT campaign.id, campaign.name FROM campaign WHERE campaign.id IN (" + campaignIds + ") " + excludedClause;
        Logger.log("campaignIds query =" + query);
        campaignDictForCurrentAccount = rowsToCampaignDict(campaignDictForCurrentAccount, AdsApp.report(query, REPORTING_OPTIONS), "campaign.id");
      }
    }
    if (campaignLabels.length > 0) {
      var query = "SELECT campaign.id, campaign.name, label.name FROM campaign_label WHERE label.name IN (" + campaignLabels + ") " + excludedClause;
      Logger.log("campaignLabels query =" + query);
      campaignDictForCurrentAccount = rowsToCampaignDict(campaignDictForCurrentAccount, AdsApp.report(query, REPORTING_OPTIONS), "label.name");
    }
    return campaignDictForCurrentAccount;
  }

  function _populateStatDictsWithCampaignsData(campaignDictForCurrentAccount) {
    var campaignIdsForCurrentAccount = Object.keys(campaignDictForCurrentAccount);
    Logger.log("Reporting for campaign ids =" + JSON.stringify(campaignIdsForCurrentAccount));
    if (campaignIdsForCurrentAccount.length > 0) {
      var currentQuery = SheetUtil.getCurrentQueryForCampaigns().replace("__CAMPAIGN_IDS__", campaignIdsForCurrentAccount);
      var currentSpecificCampaigns = AdsApp.report(currentQuery, REPORTING_OPTIONS);
      var pastQuery = SheetUtil.getPastQueryForCampaigns().replace("__CAMPAIGN_IDS__", campaignIdsForCurrentAccount);
      var pastSpecificCampaigns = AdsApp.report(pastQuery, REPORTING_OPTIONS);
      populateStatsDict(currentStatsDict, currentSpecificCampaigns.rows());
      populateStatsDict(pastStatsDict, pastSpecificCampaigns.rows());
    }
  }

  _fillDictsWithFullAccountScanStats(currentAccount);
  var campaignDictForCurrentAccount = _getCampaignDictForCurrentAccount(currentAccount);
  _populateStatDictsWithCampaignsData(campaignDictForCurrentAccount);

  function _updateStatDicts(currentStatsForCurrentAccount, pastStatsForCurrentAccount) {
    var nameParts = accountOrCampaignId.split(':');
    var campaignId = nameParts.length > 2 ? nameParts[2].toString().trim() : "";
    var campaignName = nameParts.length > 3 ? nameParts[3].toString().trim() : "";
    relevantLabelsForCurrentAccount = campaignId == "" ? relevantLabelsForCurrentAccount : campaignDictForCurrentAccount[campaignId];
    var atLeastOneMetricIsAboveMin = false;

    /** 
    * Finalize the data
    * Write results to "results tab" 
    **/
    var rowData =
      [
        JSON.stringify(relevantLabelsForCurrentAccount),
        currentAccount.getCustomerId(),
        mccManager.getCurrentAccountName(currentAccount),
        campaignId,
        campaignName];

    function setAvgToBuldingBlockStats(statName, currentStats, pastStats) {
      setAvgToBuldingBlockStatsDividedBy(statName, currentStats, pastStats, 1);
    }

    function setAvgToBuldingBlockStatsDividedBy(statName, currentStats, pastStats, divideBy) {
      currentStats[statName] /= (SheetUtil.currentTimeWindowUnits() * divideBy);
      pastStats[statName] /= (SheetUtil.pastTimeWindowUnits() * divideBy);
      current = currentStats[statName];
      past = pastStats[statName];
    }

    setAvgToBuldingBlockStatsDividedBy(STATS.impressions.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount, CONST.DEVIDE_IMPRESSIONS_BY);
    rowData = rowData.concat([formattingFuncZeroDigits(current), formattingFuncZeroDigits(past), formattingFuncZeroDigits(current - past), formattingFuncTwoDigits(current / past)]);

    setAvgToBuldingBlockStatsDividedBy(STATS.clicks.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount, CONST.DEVIDE_CLICKS_BY);
    rowData = rowData.concat([formattingFuncZeroDigits(current), formattingFuncZeroDigits(past), formattingFuncZeroDigits(current - past), formattingFuncTwoDigits(current / past)]);

    setAvgToBuldingBlockStats(STATS.conversions.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount);
    rowData = rowData.concat([formattingFuncTwoDigits(current), formattingFuncTwoDigits(past), formattingFuncTwoDigits(current - past), formattingFuncTwoDigits(current / past)]);

    setAvgToBuldingBlockStatsDividedBy(STATS.cost_micros.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount, CONST.DEVIDE_COST_BY);
    rowData = rowData.concat([formattingFuncTwoDigits(current), formattingFuncTwoDigits(past), formattingFuncTwoDigits(current - past), formattingFuncTwoDigits(current / past)]);

    setAvgToBuldingBlockStats(STATS.all_conversions_value.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount);

    function _setManuallyCalculatedRatioStas(statName, numerator, denominator, currentStats, pastStats) {
      _manualCalc(statName, numerator, denominator, currentStats, pastStats, 1);
    }
    function _setManualPriceStats(statName, numerator, denominator, currentStats, pastStats) {
      _manualCalc(statName, numerator, denominator, currentStats, pastStats, 1);
    }
    function _manualCalc(statName, numerator, denominator, currentStats, pastStats, devideBy) {
      var denomValue = currentStats[denominator];
      currentStats[statName] = denomValue == 0 ? -1 : currentStats[numerator] / currentStats[denominator] / devideBy;
      pastStats[statName] = denomValue == 0 ? -1 : pastStats[numerator] / pastStats[denominator] / devideBy;
      atLeastOneMetricIsAboveMin = atLeastOneMetricIsAboveMin || isAboveMinimalDelta(statName, currentStats[statName] - pastStats[statName]);
      current = currentStats[statName];
      past = pastStats[statName];
    }

    //CPA = sum(cost)/sum(conv) daily avg
    _setManualPriceStats(STATS.cost_per_conversion.Named_range, STATS.cost_micros.GAQL_name, STATS.conversions.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount);
    rowData = rowData.concat([current, past, (current - past), formattingFuncTwoDigits(current / past)]);

    //CTR = sum(clicks)/sum(imps)  daily avg
    _setManuallyCalculatedRatioStas(STATS.ctr.Named_range, STATS.clicks.GAQL_name, STATS.impressions.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount);
    rowData = rowData.concat([formattingFuncThreeDigits(current), formattingFuncThreeDigits(past), formattingFuncThreeDigits(current - past), formattingFuncTwoDigits(current / past)]);

    //CVR = sum(conv)/sum(clicks) daily avg
    _setManuallyCalculatedRatioStas(STATS.conversions_from_interactions_rate.Named_range, STATS.conversions.GAQL_name, STATS.clicks.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount);
    rowData = rowData.concat([formattingFuncThreeDigits(current), formattingFuncThreeDigits(past), formattingFuncThreeDigits(current - past), formattingFuncTwoDigits(current / past)]);

    //CPC = sum(cost)/sum(clicks) daily avg
    _setManualPriceStats(STATS.average_cpc.Named_range, STATS.cost_micros.GAQL_name, STATS.clicks.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount);
    rowData = rowData.concat([formattingFuncTwoDigits(current), formattingFuncTwoDigits(past), formattingFuncTwoDigits(current - past), formattingFuncTwoDigits(current / past)]);

    //ROI = conversion value / Cost 
    _setManuallyCalculatedRatioStas(STATS.roi.Named_range, STATS.all_conversions_value.GAQL_name, STATS.cost_micros.GAQL_name, currentStatsForCurrentAccount, pastStatsForCurrentAccount);
    rowData = rowData.concat([formattingFuncTwoDigits(current), formattingFuncTwoDigits(past), formattingFuncTwoDigits(current - past), formattingFuncTwoDigits(current / past)]);

    return atLeastOneMetricIsAboveMin ? rowData : [];
  }

  //Read all values from dict
  var accountOrCampaignIdsForCurrentCid = Object.keys(currentStatsDict);
  for (i in accountOrCampaignIdsForCurrentCid) {
    var accountOrCampaignId = accountOrCampaignIdsForCurrentCid[i];

    //Entity doesn't exist in the present
    if (!currentStatsDict[accountOrCampaignId]) {
      Logger.log("ERROR: currentStats are not defined for accountIdOrCampaignId = " + accountOrCampaignId);
      currentStatsDict[accountOrCampaignId] = new ZeroMetrics().zeroLine;
    }
    //Entity didn't exist in the past
    if (!pastStatsDict[accountOrCampaignId]) {
      Logger.log("ERROR: pastStatsDict are not defined for accountIdOrCampaignId = " + accountOrCampaignId);
      pastStatsDict[accountOrCampaignId] = new ZeroMetrics().zeroLine;
    }
    var isLastRowForCurrentAccount = (accountOrCampaignIdsForCurrentCid.length - 1 == i);
    var currentStatsForCurrentAccount = currentStatsDict[accountOrCampaignId];
    var pastStatsForCurrentAccount = pastStatsDict[accountOrCampaignId];
    var rowData = _updateStatDicts(currentStatsForCurrentAccount, pastStatsForCurrentAccount,);
    if (rowData.length == 0) {     //All noise cancelling thresholds failed.
    Logger.log("i "+i);
      continue;
    }
    resultsSheet.getRange(currentRowToWrite, CONST.FIRST_DATA_COLUMN, 1, (5 + 4 * 9)).setValues([rowData]);
    //Prepare email for scanned accounts
    var alertTextForCampaign = generateEmailTextForCampiagn(accountOrCampaignId, currentRowToWrite, currentStatsForCurrentAccount, pastStatsForCurrentAccount);
    var isHighlighted = alertTextForCampaign.includes("<td>");
    Logger.log("i "+i);
    if (isHighlighted) {
      alertTextForAccount.push(alertTextForCampaign);
          currentRowToWrite++;
    } else {
        resultsSheet.deleteRow(currentRowToWrite);
      }
    }
  return alertTextForAccount;
}

function colorResultsCell(cell, color) {
  resultsSheet.getRange(cell[0], cell[1], 1, 1).setBackground(color);
}

function isAboveMinimalDelta(statName, actualDelta) {
  var minimalDelta = spreadsheet.getRangeByName(statName + "_ignore").getValue();
  minimalDelta = minimalDelta == "" ? -1 : minimalDelta;
  var upper = spreadsheet.getRangeByName(statName + "_high").getValue();
  var lower = spreadsheet.getRangeByName(statName + "_low").getValue();
  var emptyMetric = (minimalDelta == "" && upper == "" && lower == "");
  var isMin = !emptyMetric && (Math.abs(actualDelta) > Math.abs(minimalDelta));
  return isMin;
}

//Add to future email body content
//Email formatting
function highlightThresholdsInTrixAndEmail(statNames, thresholdType, currentStats, pastStats, campaignCellInTrix) {

  var currencyCode = AdsApp.currentAccount().getCurrencyCode();
  currencyCode = (currencyCode == "USD") ? "$" : currencyCode;
  currencyCode = currencyCode + " ";

  var statsName = statNames.GAQL_name;
  var currentThresholdName = statNames.Named_range + "_" + thresholdType;
  var thresholds = SheetUtil.thresholds();
  if (!thresholds[currentThresholdName]) {
    return "";
  }
  var pastAvgValue = pastStats[statsName];
  var currentAvgValue = currentStats[statsName];
  var thresholdValue = (pastAvgValue) * thresholds[currentThresholdName];
  var numericDelta = currentAvgValue - pastAvgValue;
  var percentageDelta = (currentAvgValue / pastAvgValue) * 100;

  var currentValueStr = currentAvgValue;
  var pastAvgValueStr = pastAvgValue;
  var numericDeltaStr = numericDelta;
  var percentageDeltaStr = percentageDelta;

  if (currentThresholdName.includes("impressions") || currentThresholdName.includes("clicks")
    || currentThresholdName.includes("conversions") || currentThresholdName.includes("roi")) {
    currentValueStr = Math.round(currentAvgValue);
    pastAvgValueStr = Math.round(pastAvgValue);
    numericDeltaStr = Math.round(numericDelta);
    percentageDeltaStr = formatNums(percentageDelta, 2) + "%";
  }
  else if (currentThresholdName.includes("cost_micros")) {
    currentValueStr = currencyCode + Math.round(currentAvgValue);
    pastAvgValueStr = currencyCode + Math.round(pastAvgValue);
    numericDeltaStr = currencyCode + Math.round(numericDelta);
    percentageDeltaStr = formatNums(percentageDelta, 2) + "%";
  }
  //daily avg
  else if (currentThresholdName.includes("cost_per_conversion")) {
    currentValueStr = currencyCode + formatNums(currentAvgValue, 2);
    pastAvgValueStr = currencyCode + formatNums(pastAvgValue, 2);
    numericDeltaStr = currencyCode + formatNums(numericDelta, 2);
    percentageDeltaStr = formatNums(percentageDelta, 2) + "%";
  }
  //daily avg
  else if (currentThresholdName.includes("ctr") || currentThresholdName.includes("conversions_from_interactions_rate")) {
    currentValueStr = formatNums(currentAvgValue, 3);
    pastAvgValueStr = formatNums(pastAvgValue, 3);
    numericDeltaStr = formatNums(numericDelta, 3);
    percentageDeltaStr = formatNums(percentageDelta, 3) + "%";
  }
  //daily avg
  else if (currentThresholdName.includes("average_cpc")) {
    currentValueStr = currencyCode + formatNums(currentAvgValue, 1);
    pastAvgValueStr = currencyCode + formatNums(pastAvgValue, 1);
    numericDeltaStr = currencyCode + formatNums(numericDelta, 1);
    percentageDeltaStr = formatNums(percentageDelta, 2) + "%";
  }

  currentValueStr = addNumberCommas(currentValueStr);
  pastAvgValueStr = addNumberCommas(pastAvgValueStr);
  numericDeltaStr = addNumberCommas(numericDeltaStr);

  var isHighlighted = false;
  //Email text
  if (currentAvgValue == -1 || thresholdValue == -1) {
    // nothing
  }
  else if (thresholdType.includes("high") && currentAvgValue >= thresholdValue) {
    isHighlighted = true;
    colorResultsCell(campaignCellInTrix, spreadsheet.getRangeByName(CONST.COLOR_HIGH).getValue());
  }
  else if (thresholdType.includes("low") && currentAvgValue <= thresholdValue) {
    isHighlighted = true;
    colorResultsCell(campaignCellInTrix, spreadsheet.getRangeByName(CONST.COLOR_LOW).getValue());
  }
  return isHighlighted ?
    "<tr><td>" + statNames.Email_text + "</td><td>" + currentValueStr + "</td><td>" + pastAvgValueStr + "</td><td>" + numericDeltaStr + "</td><td>" + percentageDeltaStr + "</td></tr>" : "";
}


//** Email body row */
function generateEmailTextForCampiagn(accountIdOrCampaignId, campaignRowInTrix, currentStats, pastStats) {

  var alertTextForCampaign = ['<br>Anomalies for: ' + accountIdOrCampaignId +
    '<br><table style="width:50%;border:1px solid black;"><tr style="border:1px solid black;"><th style="text-align:left;border:1px solid black;">Metric</th><th style="text-align:left;border:1px solid black;">Current</th><th style="text-align:left;border:1px solid black;">Past</th> <th style="text-align:left;border:1px solid black;">Δ</th> <th style="text-align:left;border:1px solid black;">Δ%</th></tr>'];

  var cell = [campaignRowInTrix, STATS.impressions.Column];
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.impressions, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.impressions, "low", currentStats, pastStats, cell));

  cell[1] = STATS.clicks.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.clicks, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.clicks, "low", currentStats, pastStats, cell));

  cell[1] = STATS.conversions.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.conversions, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.conversions, "low", currentStats, pastStats, cell));

  cell[1] = STATS.cost_micros.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.cost_micros, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.cost_micros, "low", currentStats, pastStats, cell));

  cell[1] = STATS.cost_per_conversion.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.cost_per_conversion, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.cost_per_conversion, "low", currentStats, pastStats, cell));

  cell[1] = STATS.ctr.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.ctr, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.ctr, "low", currentStats, pastStats, cell));

  cell[1] = STATS.conversions_from_interactions_rate.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.conversions_from_interactions_rate, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.conversions_from_interactions_rate, "low", currentStats, pastStats, cell));

  cell[1] = STATS.average_cpc.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.average_cpc, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.average_cpc, "low", currentStats, pastStats, cell));

  cell[1] = STATS.roi.Column;
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.roi, "high", currentStats, pastStats, cell));
  alertTextForCampaign.push(highlightThresholdsInTrixAndEmail(STATS.roi, "low", currentStats, pastStats, cell));

  var oneLineCampaignText = alertTextForCampaign.join('');
  return (oneLineCampaignText.includes('<td>')) ? oneLineCampaignText + "</table>" : "";
}


var formattingFuncZeroDigits = function (x) { return formatNums(x, 0); };
var formattingFuncTwoDigits = function (x) { return formatNums(x, 2); };
var formattingFuncThreeDigits = function (x) { return formatNums(x, 3); };

function formatNums(num, toFixed) {
  toFixed = toFixed || 1;
  return String(num.toFixed(toFixed));
}

function addNumberCommas(numb) {
  var str = numb.toString().split(".");
  str[0] = str[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return str.join(".");
}


var SheetUtil = (function () {
  //PUBLIC VARS
  var thresholds = {};
  var currentQueryForFullAccount = '';
  var pastQueryForFullAccount = '';
  var currentQueryForCampaigns = '';
  var currentQueryForCampaigns = '';

  var accountToCampaignsMap = {};
  var fullAccountIds = [];
  var excludedAccountIds = [];
  var accountLabels = [];
  var campaignIds = "";
  var excludedCampaignIds = "";
  var campaignLabels = "";

  var currentTimeWindowUnits = 1;
  var currentTimeWindowDays = 1;
  var pastTimeWindowUnit = 1;
  var pastTimeWindowDays = 1;


  var cleanResults = function cleanResults() {
    var lastRow = resultsSheet.getLastRow();
    var howMany = lastRow - CONST.FIRST_DATA_ROW + 1;
    if (howMany > 0) {
      resultsSheet.deleteRows(CONST.FIRST_DATA_ROW, howMany);
    }
    lastRow = resultsSheet.getMaxRows();
    if (lastRow < 350) {
      resultsSheet.insertRows(lastRow, 1000);
    }
  }

  var setupData = function setupData(spreadsheet) {
    spreadsheet.getRangeByName(CONST.DATE).setValue(new Date());
    spreadsheet.getRangeByName(CONST.ACCOUNT_ID).setValue(
      mccManager.mccAccount().getCustomerId());

    var readThresholdFor = function (field) {
      thresholds[field + "_high"] = parseField(spreadsheet.
        getRangeByName(field + "_high").getValue());
      thresholds[field + "_low"] = parseField(spreadsheet.
        getRangeByName(field + "_low").getValue());
    };

    readThresholdFor(STATS.impressions.Named_range);
    readThresholdFor(STATS.clicks.Named_range);
    readThresholdFor(STATS.conversions.Named_range);
    readThresholdFor(STATS.cost_micros.Named_range);
    readThresholdFor(STATS.ctr.Named_range);
    readThresholdFor(STATS.conversions_from_interactions_rate.Named_range);
    readThresholdFor(STATS.average_cpc.Named_range);
    readThresholdFor(STATS.cost_per_conversion.Named_range);
    readThresholdFor(STATS.roi.Named_range);

    var convertToDays = function (avgType) {
      var base = 1;
      if (avgType.includes("week")) {
        base = 7;
      }
      if (avgType.includes("14")) {
        base = 14;
      }
      if (avgType.includes("30")) {
        base = 30;
      }
      return base;
    };

    var lengthUnit = spreadsheet.getRangeByName(CONST.LENGTH_UNIT).getValue();
    var lengthUnitDays = convertToDays(lengthUnit);

    var currentStartGoBack = spreadsheet.getRangeByName(CONST.CURRENT_RANGE_DAY).getValue();
    currentTimeWindowUnits = spreadsheet.getRangeByName(CONST.CURRENT_RANGE_PERIOD).getValue();
    currentTimeWindowDays = currentTimeWindowUnits * lengthUnitDays;
    var pastStartGoBack = spreadsheet.getRangeByName(CONST.PAST_RANGE_DAY).getValue();
    pastTimeWindowUnit = spreadsheet.getRangeByName(CONST.PAST_RANGE_PERIOD).getValue();
    pastTimeWindowDays = pastTimeWindowUnit * lengthUnitDays;

    var currentRangeEndDate = getDateStringForMinusDays(currentStartGoBack);
    var currentRangeStartDate = getDateStringForMinusDays(currentStartGoBack + currentTimeWindowDays - 1);
    var pastRangeEndDate = getDateStringForMinusDays(pastStartGoBack);
    var pastRangeStartDate = getDateStringForMinusDays(pastStartGoBack + pastTimeWindowDays - 1);


    var gAdsQueryFields = 'metrics.clicks, metrics.impressions, metrics.conversions, metrics.cost_micros, metrics.all_conversions_value';
    gAdsAndManualHeaders = gAdsQueryFields + ', ctr, conversions_from_interactions_rate, average_cpc, cost_per_conversion';
    gAdsAndManualHeaders = gAdsAndManualHeaders.replace(/\s/g, '').split(',');

    var baseQuery = 'SELECT ' + gAdsQueryFields +
      ' FROM customer WHERE segments.date BETWEEN';

    currentQueryForFullAccount = baseQuery + ' \"' + currentRangeStartDate.query_date + '\" AND \"' +
      currentRangeEndDate.query_date + "\"";

    pastQueryForFullAccount = baseQuery + ' \"' + pastRangeStartDate.query_date + '\" AND \"' +
      pastRangeEndDate.query_date + "\"";

    var specificCampaignsQuery = 'SELECT campaign.id, campaign.name, ' + gAdsQueryFields +
      ' FROM campaign WHERE campaign.id IN ( __CAMPAIGN_IDS__ ) AND segments.date BETWEEN';

    currentQueryForCampaigns = specificCampaignsQuery + ' \"' + currentRangeStartDate.query_date + '\" AND \"' +
      currentRangeEndDate.query_date + "\"";

    pastQueryForCampaigns = specificCampaignsQuery + ' \"' + pastRangeStartDate.query_date + '\" AND \"' +
      pastRangeEndDate.query_date + "\"";
    Logger.log("currentQueryForFullAccount:" + currentQueryForFullAccount);
    Logger.log("pastQueryForFullAccount:" + pastQueryForFullAccount);

    spreadsheet.getRangeByName(CONST.CURRENT_RANGE_DATES).
      setValue(currentRangeStartDate.sheet_date + "-" + currentRangeEndDate.sheet_date +
        " (" + currentTimeWindowUnits + " * " + spreadsheet.getRangeByName(CONST.LENGTH_UNIT).getValue() + ")");
    spreadsheet.getRangeByName(CONST.PAST_RANGE_DATES).
      setValue(pastRangeStartDate.sheet_date + "-" + pastRangeEndDate.sheet_date
        + " (" + pastTimeWindowUnit + " * " + spreadsheet.getRangeByName(CONST.LENGTH_UNIT).getValue() + ")");

    campaignIds = spreadsheet.getRangeByName(CONST.CAMPAIGN_IDS).getValue();
    excludedCampaignIds = spreadsheet.getRangeByName(CONST.EXCLUDED_CAMPAIGN_IDS).getValue();
    campaignLabels = spreadsheet.getRangeByName(CONST.CAMPAIGN_LABELS).getValue();

    accountLabels = spreadsheet.getRangeByName(CONST.ACCOUNT_LABELS).getValue().trim().replace(/"/g, "").split(",").map(function (item) { return item.toString().trim() });
    excludedAccountIds = spreadsheet.getRangeByName(CONST.EXCLUDED_ACCOUNT_IDS).getValue().trim().replace(/"/g, "").split(",").map(function (item) { return item.toString().trim() });
    fullAccountIds = spreadsheet.getRangeByName(CONST.ACCOUNT_IDS).getValue().trim().replace(/"/g, "").split(",").map(function (item) { return item.toString().trim() });
  };


  //Insure these are singletons
  var getThresholds = function () { return thresholds; };
  var getCurrentTimeWindowUnits = function () { return currentTimeWindowUnits; };
  var getCurrentTimeWindowDays = function () { return currentTimeWindowDays; };
  var getPastTimeWindowUnits = function () { return pastTimeWindowUnit; };
  var getPastTimeWindowDays = function () { return pastTimeWindowDays; };
  var getFullAccountIds = function () { return fullAccountIds; };
  var getExcludedFullAccountIds = function () { return excludedAccountIds; };
  var getAccountToCampaignsMap = function () { return accountToCampaignsMap; };
  var getPastQueryForFullAccount = function () { return pastQueryForFullAccount; };
  var getCurrentQueryForFullAccount = function () { return currentQueryForFullAccount; };
  var getPastQueryForCampaigns = function () { return pastQueryForCampaigns; };
  var getCurrentQueryForCampaigns = function () { return currentQueryForCampaigns; };
  var getCampaignIds = function () { return campaignIds; };
  var getExcludedCampaignIds = function () { return excludedCampaignIds; };
  var getCampaignLabels = function () { return campaignLabels; };
  var getAccountLabels = function () { return accountLabels; };

  // The SheetUtil public interface.
  return {
    setupData: setupData,
    thresholds: getThresholds,
    currentTimeWindowUnits: getCurrentTimeWindowUnits,
    currentTimeWindowDays: getCurrentTimeWindowDays,
    pastTimeWindowUnits: getPastTimeWindowUnits,
    pastTimeWindowDays: getPastTimeWindowDays,
    getFullScanAccountIds: getFullAccountIds,
    getExcludedFullAccountIds: getExcludedFullAccountIds,
    getAccountToCampaignsMap: getAccountToCampaignsMap,
    getPastQueryForFullAccount: getPastQueryForFullAccount,
    getCurrentQueryForFullAccount: getCurrentQueryForFullAccount,
    getPastQueryForCampaigns: getPastQueryForCampaigns,
    getCurrentQueryForCampaigns: getCurrentQueryForCampaigns,
    cleanStats: cleanResults,
    getCampaignIds: getCampaignIds,
    getExcludedCampaignIds: getExcludedCampaignIds,
    getCampaignLabels: getCampaignLabels,
    getAccountLabelsFilter: getAccountLabels
  };
})();

function sendEmail(account, alertTextForMcc, spreadsheet) {
  var bodyText = '';
  alertTextForMcc.forEach(function (accountAlert) {
    // When zero alerts, this is an empty array, which we don't want to add.
    if (accountAlert == undefined || accountAlert.length == 0) { return }
    bodyText += accountAlert.join('<br>');
  });
  bodyText = bodyText.trim();
  var email = spreadsheet.getRangeByName(CONST.EMAILS).getValue();
  if (bodyText.includes("td") && email && email.length > 0) {
    Logger.log('Sending Email');
    MailApp.sendEmail(
      {
        to: email,
        subject: '[Anomaly Alert] MCC Account [' + account.getCustomerId() + ']',
        htmlBody: "New period: " + spreadsheet.getRangeByName(CONST.CURRENT_RANGE_DATES).getValue() + "<br>" +
          "Past period: " + spreadsheet.getRangeByName(CONST.PAST_RANGE_DATES).getValue() + "<br>" +
          '<br>' +
          bodyText +
          '<br><br>' +
          'Log into Google Ads and take a look: ' +
          'ads.google.com\n\nAlerts dashboard: ' +
          SPREADSHEET_URL
      });
  }
  else {
    Logger.log('No alerts triggered. No email being sent.');
  }
}

function toFloat(value) {
  if (!value) return 0;
  value = value.toString().replace(/,/g, '');
  return parseFloat(value);
}

function parseField(value) {
  if (value == 'No alert' || value === "") {
    return null;
  } else {
    return toFloat(value);
  }
}

function populateStatsDict(results, rows) {
  //Init with zeros in case there are no rows
  var accountIdntifier = AdsApp.currentAccount().getCustomerId() + " : " + mccManager.getCurrentAccountName(AdsApp.currentAccount());
  while (rows.hasNext()) {
    var row = rows.next();
    var id = row["campaign.id"] == undefined ?
      accountIdntifier :
      accountIdntifier + " : " + row["campaign.id"] + " : " + row["campaign.name"];
    results[id] = accumulateToReportRows(row, results[id]);
  }
}

//Inner class
function ZeroMetrics() {
  this.zeroLine = (function () {
    var zeros = {};
    gAdsAndManualHeaders.forEach(function (key) {
      zeros[key] = 0;
    });
    return zeros;
  })();
}

//EACH ROW
function accumulateToReportRows(accForEntity, row) {
  if (row == undefined) {
    row = new ZeroMetrics().zeroLine;
  }
  if (accForEntity == undefined) {
    accForEntity = new ZeroMetrics().zeroLine;
  }
  //Same as gAdsQueryFields
  accForEntity[STATS.clicks.GAQL_name] += toFloat(row[STATS.clicks.GAQL_name]);
  accForEntity[STATS.impressions.GAQL_name] += toFloat(row[STATS.impressions.GAQL_name]);
  accForEntity[STATS.conversions.GAQL_name] += toFloat(row[STATS.conversions.GAQL_name]);
  accForEntity[STATS.cost_micros.GAQL_name] += toFloat(row[STATS.cost_micros.GAQL_name]);
  accForEntity[STATS.all_conversions_value.GAQL_name] += toFloat(row[STATS.all_conversions_value.GAQL_name]);
  return accForEntity;
}

/**
 * Produces a formatted string representing a date in the past of a given date.
 */
function getDateStringForMinusDays(numDays) {
  var expectedDate = new Date(new Date().setDate(new Date().getDate() - numDays));
  return { 'query_date': getDateStringInTimeZone('yyyy-MM-dd', expectedDate), 'sheet_date': getDateStringInTimeZone('dd/MM/YY', expectedDate) };
}


/**
 * Produces a formatted string representing a given date in a given time zone.
 *
 * @param {string} format A format specifier for the string to be produced.
 * @param {date} date A date object. Defaults to the current date.
 * @param {string} timeZone A time zone. Defaults to the account's time zone.
 * @return {string} A formatted string of the given date in the given time zone.
 */
function getDateStringInTimeZone(format, date, timeZone) {
  date = date || new Date();
  timeZone = timeZone || AdsApp.currentAccount().getTimeZone();
  return Utilities.formatDate(date, timeZone, format);
}


/**
 * Module that deals with fetching and iterating through multiple accounts.
 *
 * @return {object} callable functions corresponding to the available
 * actions. Specifically, it currently supports next, current, mccAccount.
 */
var mccManager = function () {
  var mccAccount;

  //Private one-time init function.
  var init = function () {
    mccAccount = AdsApp.currentAccount(); // save the mccAccount    
  };


  var getCurrentAccountName = function (account) {
    var accountName = "";
    try {
      accountName = account.getName();
    }
    catch (exc) { }
    return accountName;
  }

  var getAllSubAccounts = function () {
    var accounts = [];
    accountIterator = AdsManagerApp.accounts().get();
    while (accountIterator.hasNext()) {
      var currentAccount = accountIterator.next();
      accounts.push(currentAccount);
    }
    Logger.log("MCC has " + accounts.length + " child accounts");
    return accounts;
  };

  var getAccountsForLabels = function (labels) {
    var accounts = [];
    labels = labels.split(",");
    for (i in labels) {
      var accountIterator = AdsManagerApp.accounts().withCondition("LabelNames CONTAINS '" + labels[i] + "'").get();
      while (accountIterator.hasNext()) {
        var currentAccount = accountIterator.next();
        accounts.push(currentAccount);
      }
    }
    return accounts;
  };


  var getAccountsForIds = function (ids) {
    var accounts = [];
    var accountIterator = AdsManagerApp.accounts().withIds(ids).get();
    while (accountIterator.hasNext()) {
      var currentAccount = accountIterator.next();
      accounts.push(currentAccount);
    }
    return accounts;
  };

  /**
   * Returns the original MCC account.
   *
   * @return {AdsApp.Account} The original account that was selected.
   */
  var getMccAccount = function () {
    return mccAccount;
  };

  // Set up internal variables; called only once, here.
  init();

  // Expose the external interface.
  return {
    mccAccount: getMccAccount,
    getAccountsForLabels: getAccountsForLabels,
    getAllSubAccounts: getAllSubAccounts,
    getCurrentAccountName: getCurrentAccountName,
    getAccountsForIds: getAccountsForIds
  };
};/*