const weekdays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/** ============= Cad Result ==================== */
/**
 * A metric's performance
 *
 */
class MetricResult {
  constructor() {
    this.metricAlertDirection = undefined;
    this.past = 0;
    this.current = 0;
    this.changeAbs = 0;
    this.changePercent = 0;
  }
}

/**
 * Metric Types
 *
 */
const MetricTypes = {
  impressions: {
    email: 'Impressions',
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true
  },
  clicks: {
    email: 'Clicks',
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true
  },
  all_conversions: {
    email: 'All Conversions',
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true
  },
  conversions: {
    email: 'conversions',
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true
  },
  cost_micros: {
    email: 'Cost',
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true
  },
  cost_per_all_conversions: {
    email: 'CPA all',
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true
  },
  cost_per_conversion: {
    email: 'CPA',
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true
  },
  ctr: {
    email: 'CTR',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
    isPercent: true
  },
  all_conversions_from_interactions_rate: {
    email: 'CVR all',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },
  conversions_from_interactions_rate: {
    email: 'CVR',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },

  average_cpc: {
    email: 'CPC',
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },

  roas_all: {
    email: 'ROAS all',
    isMonitored: true,
    shouldBeRounded: false,
    isWriteToSheet: true
  },
  roas: {
    email: 'ROAS',
    isMonitored: true,
    shouldBeRounded: false,
    isWriteToSheet: true
  },

  all_conversions_value:
  {
    email: 'All conversions value',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },

  conversions_value: {
    email: 'Conversions value',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },


  search_click_share: {
    email: 'Search click share',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },
  average_cpm: {
    email: 'Avg CPM',
    isMicro: true,
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },

  average_cpv: {
    email: 'Avg CPV',
    isMicro: true,
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },

  video_view_rate: {
    email: 'Video View Rate',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },
};

/**
 * Comparison results of one entity to monitor
 */
class CadSingleResult {
  constructor() {
    this.relevant_label = undefined;
    this.account = { id: undefined, name: undefined };
    this.campaign = { id: undefined, name: undefined };

    this.isTriggerAlert = false;
    this.allMetricsComparisons = {};

    let currencyCode =  AdsApp.currentAccount().getCurrencyCode();
    currencyCode = (currencyCode === 'USD') ? '$' : currencyCode;
    this.currencyCode = currencyCode + ' ';
  }

  /**
   * Fills the metrics map
   *
   * @param {string} id Entity id
   * @param {!Object} pastStats past stats map
   * @param {!Object} currentStats current stats map
   * @param {!CadConfig} cadConfig CAD config
   */
  fillMetricResults(id, pastStats, currentStats, cadConfig) {
    // Columns
    let monitoredMetrics = [];
    Object.keys(MetricTypes).forEach(function (key, index) {
      if (MetricTypes[key].isMonitored) {
        monitoredMetrics.push(key);
      }
    });

    for (let metric of monitoredMetrics) {
      let gAdsMetric = `metrics.${metric}`;
      this.allMetricsComparisons[gAdsMetric] = new MetricResult();
      let comparisonResult = this.allMetricsComparisons[gAdsMetric];

      currentStats[id] = (currentStats[id]) ? (currentStats[id]) : {};
      pastStats[id] = (pastStats[id]) ? (pastStats[id]) : {};


      if (MetricTypes[metric].isMicro) {
        currentStats[id][gAdsMetric] /= 1e6;
        pastStats[id][gAdsMetric] = pastStats[id][gAdsMetric]? pastStats[id][gAdsMetric] /= 1e6: 0;
      }
      if (MetricTypes[metric].isCumulative) {
        comparisonResult.past = pastStats[id][gAdsMetric] ? pastStats[id][gAdsMetric] /
          cadConfig.dividePastBy :
          0;
        comparisonResult.current = currentStats[id][gAdsMetric] /
          cadConfig.divideCurrentBy;

      } else if (gAdsMetric.includes('roas_all')) {
        comparisonResult.past = pastStats[id][gAdsMetric] ?
          pastStats[id]['metrics.all_conversions_value'] /
          pastStats[id]['metrics.cost_micros'] :
          0;
        comparisonResult.current =
          currentStats[id]['metrics.all_conversions_value'] /
          currentStats[id]['metrics.cost_micros'];

      } else if (gAdsMetric.includes('roas')) {
        comparisonResult.past = pastStats[id][gAdsMetric] ?
          pastStats[id]['metrics.conversions_value'] /
          pastStats[id]['metrics.cost_micros'] :
          0;
        comparisonResult.current = currentStats[id]['metrics.conversions_value'] /
          currentStats[id]['metrics.cost_micros'];
      }

      //Only after done using conversions_value for roas calculation.
      else if (gAdsMetric.includes('conversions_value')) {
        comparisonResult.past = pastStats[id][gAdsMetric] ? pastStats[id][gAdsMetric] /
          cadConfig.dividePastBy :
          0;
        comparisonResult.current = currentStats[id][gAdsMetric] /
          cadConfig.divideCurrentBy;

      }
      else {
        comparisonResult.past = pastStats[id][gAdsMetric] ? pastStats[id][gAdsMetric] : 0;
        comparisonResult.current =
          (currentStats[id]) ? currentStats[id][gAdsMetric] : 0;
      }
      if (MetricTypes[metric].isPercent) {
        comparisonResult.past *= 100;
        comparisonResult.current *= 100;
      }

      comparisonResult.changeAbs = comparisonResult.current - comparisonResult.past;
      comparisonResult.changePercent = (comparisonResult.current / comparisonResult.past) * 100 - 100;
      comparisonResult.isAboveHigh =
        comparisonResult.changePercent >= cadConfig.thresholds[`${metric}_high`] * 100;
      comparisonResult.isBelowLow =
        comparisonResult.changePercent <= cadConfig.thresholds[`${metric}_low`] * 100;

      let ignoreAbs = cadConfig.thresholds[`${metric}_ignore`];
      if (!ignoreAbs || Math.abs(comparisonResult.changeAbs) >= ignoreAbs) {
        if (comparisonResult.isAboveHigh) {
          comparisonResult.metricAlertDirection = 'up';
        } else if (comparisonResult.isBelowLow) {
          comparisonResult.metricAlertDirection = 'down';
        } else if (cadConfig.showOnlyAnomalies) {
          comparisonResult.changeAbs = " - ";
          comparisonResult.changePercent = " - ";
        }
      }
      this.isTriggerAlert = this.isTriggerAlert ||
        (comparisonResult.metricAlertDirection != undefined);
    }
  }


  class MetricStrings{
    constructor() {
      this.currentValueStr = '';
      this.pastAvgValueStr = '';
      this.numericDeltaStr = '';
      this.percentageDeltaStr = '';
    }
    constructor(currentValueStr, pastAvgValueStr, numericDeltaStr, percentageDeltaStr) {
      this.currentValueStr = currentValueStr;
      this.pastAvgValueStr = pastAvgValueStr;
      this.numericDeltaStr = numericDeltaStr;
      this.percentageDeltaStr = percentageDeltaStr;
    }
  }

  /**
   * Truncate decimal digits
   */
  truncateDecimalDigits() {
    for (let metricType in this.allMetricsComparisons) {
      let metricTypeName = metricType.split(".")[1];
      if (!MetricTypes[metricTypeName].isWriteToSheet) continue;
      if (!this.allMetricsComparisons[metricType].metricAlertDirection) continue;

      let currentAvgValue = this.allMetricsComparisons[metricType].current;
      let pastAvgValue = this.allMetricsComparisons[metricType].past;
      let numericDelta = this.allMetricsComparisons[metricType].changeAbs;
      let percentageDelta = this.allMetricsComparisons[metricType].changePercent;
      let metricAlertDirection =
        this.allMetricsComparisons[metricType].metricAlertDirection;

      let metricStrings = new MetricStrings();

      if (MetricTypes[metricTypeName].shouldBeRounded) {
        metricStrings = this.toRoundedMetricString(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta);
      }
      // A stat already avg for period
      else if (
        metricType.includes('cost_per_conversion') ||
        metricType.includes('cost_micros')) {
        metricStrings = this.toMetricStrings(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta,
          this.currencyCode + ' ', '', 2);
      }
      // A stat already avg for period
      else if (
        metricType.includes('ctr') ||
        metricType.includes('roas')) {
        metricStrings = this.toMetricStrings(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta, '',
          '%', 1);
      }
      // A stat already avg for period
      else {
        metricStrings = this.toMetricStrings(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta, '',
          '', 1);        
      }

      if (metricAlertDirection) {
        metricStrings.percentageDeltaStr = metricAlertDirection.includes('up') ?
          `⇪⇪ ${metricStrings.percentageDeltaStr}` :
          `⇩⇩ ${metricStrings.percentageDeltaStr}`;
      }      

      this.allMetricsComparisons[metricType].current = metricStrings.currentValueStr;
      this.allMetricsComparisons[metricType].past = metricStrings.pastAvgValueStr;
      this.allMetricsComparisons[metricType].changeAbs = metricStrings.numericDeltaStr;
      this.allMetricsComparisons[metricType].changePercent = metricStrings.percentageDeltaStr;
    }
  }

  /**
   * Returns a string for an email format
   *
   * @return {string} an email string format
   */
  toEmailFormat() {
    if (!(this.isTriggerAlert)) return '';

    let campaignHeader =
      (this.campaign.id) ? `: ${this.campaign.id} ${this.campaign.name}` : '';
    let alertTextForEntity = [
      `<br>Anomalies for: ${this.account.id} ${this.account.name} ${campaignHeader}
      <br>(Only relevant metrics. For all metrics see the end of the email)
     <br><table style="width:50%;border:1px solid black;"><tr style="border:1px solid black;"><th style="text-align:left;border:1px solid black;">Metric</th><th style="text-align:left;border:1px solid black;">Current</th><th style="text-align:left;border:1px solid black;">Past</th> <th style="text-align:left;border:1px solid black;">Δ</th> <th style="text-align:left;border:1px solid black;">Δ%</th></tr>`
    ];

    for (let metricType in this.allMetricsComparisons) {
      let metricType_name = metricType.split(".")[1];
      if (!MetricTypes[metricType_name].isWriteToSheet) continue;
      if (!this.allMetricsComparisons[metricType].metricAlertDirection) continue;

      const toStringFormatter = ToStringFormatter.getInstance();
      const currentValueStr = toStringFormatter.addNumberCommas(this.allMetricsComparisons[metricType].current);
      const pastAvgValueStr = toStringFormatter.addNumberCommas(this.allMetricsComparisons[metricType].past);
      const numericDeltaStr = toStringFormatter.addNumberCommas(this.allMetricsComparisons[metricType].changeAbs);
      const percentageDeltaStr = toStringFormatter.addNumberCommas(this.allMetricsComparisons[metricType].changePercent);

      alertTextForEntity.push(`<tr><td> ${metricType} </td><td> ${currentValueStr} </td><td> ${pastAvgValueStr} </td><td>
${numericDeltaStr} </td><td style="color: ${this.getDeltaColorPercentage(percentageDeltaStr)};"> ${percentageDeltaStr} </td></tr>`);
    }
    return `${alertTextForEntity.join('')} </table>`;
  };


  // Function to get the color based on positive/negative percentage values
  getDeltaColorPercentage(percentage) {
    if (percentage.includes("⇪")) {
      return 'green';
    } else if (percentage.includes("⇩")) {
      return 'red'; // If the value is zero
    } else {
      return 'white';
    }
  }

  /**
   * @param {number} currentAvgValue current value (avg)
   * @param {number} pastAvgValue  past value (avg)
   * @param {number} numericDelta  numeric delta
   *@param {number} percentageDelta percentage delta
   *@return {!Object} a map of string formats
   */
  toRoundedMetricString(currentAvgValue, pastAvgValue, numericDelta, percentageDelta) {
    return new MetricStrings(Math.round(currentAvgValue), Math.round(pastAvgValue), Math.round(numericDelta),Math.round(percentageDelta) + '%')
    };
  }


  /**
   * @param {number} currentAvgValue current value (avg)
   * @param {number} pastAvgValue  past value (avg)
   * @param {number} numericDelta  numeric delta
   *@param {number} percentageDelta percentage delta
   * @param {string} prefix prefix each string
   * @param {string} postfix postfix each string
   * @param {string} decimalDigits decimal digits accuricy
   *@return {!Object} a map of string formats
   */
  toMetricStrings(
    currentAvgValue, pastAvgValue, numericDelta, percentageDelta, prefix,
    postfix, decimalDigits) {
    const toStringFormatter = ToStringFormatter.getInstance();
    return {
      currentValueStr: prefix +
        toStringFormatter.formatNumWithDigitAfterPeriod(
          currentAvgValue, decimalDigits) +
        postfix,
      pastAvgValueStr: prefix +
        toStringFormatter.formatNumWithDigitAfterPeriod(
          pastAvgValue, decimalDigits) +
        postfix,
      numericDeltaStr: prefix +
        toStringFormatter.formatNumWithDigitAfterPeriod(
          numericDelta, decimalDigits) +
        postfix,
      percentageDeltaStr: toStringFormatter.formatNumWithDigitAfterPeriod(
        percentageDelta, decimalDigits) +
        '%'
    };
  }

  /**
   * Returns a string for a sheet format
   *
   * @return {!Array<?>} a string for a sheet format
   */
  toSheetFormat() {
    let rowData = [
      this.relevant_label,
      this.account.id,
      this.account.name,
      this.campaign.id,
      this.campaign.name,
    ];

    console.log(`this.metricComparisonResults = ${JSON.stringify(this.allMetricsComparisons)}`);
    for (let metricType in this.allMetricsComparisons) {
      let metricName = metricType.split(".")[1];
      if (!MetricTypes[metricName].isMonitored) continue;

      rowData = rowData.concat([
        this.allMetricsComparisons[metricType].current, this.allMetricsComparisons[metricType].past,
        this.allMetricsComparisons[metricType].changeAbs,
        this.allMetricsComparisons[metricType].changePercent
      ]);
    }
    return [rowData];
  };
}

/** ============= Cad Config ==================== */


/**
 * The user input from the sheet
 */
class CadConfig {
  /**
   * A constructor
   */
  constructor() {
    this.mcc = undefined;
    this.users = [];
    this.avgType = "Daily Avg";
    this.showOnlyAnomalies = false;

    this.hourSegmentsWhereClause = { current: "", past: "" };
    this.hourSegmentInSelect = "";

    this.divideCurrentBy = 1;
    this.dividePastBy = 1;

    this.lookbackInUnits = {
      current_period_length: undefined,
      current_ended_length_ago: undefined,
      past_period_length: undefined,
      past_ended_length_ago: undefined
    };

    //after code calculation:
    this.lookbackInDays = {
      current_period_length: {},
      current_ended_length_ago: {},
      past_period_length: {},
      past_ended_length_ago: {},
    };

    this.lookbackDates = {
      current_range_start_date: {},
      current_range_end_date: {},
      past_range_start_date: {},
      past_range_end_date: {},
    };
    this.accounts = {
      account_ids: [],
      account_excluded_ids: [],
      account_labels: []
    };
    this.campaigns = {
      campaign_ids: "",
      campaign_excluded_ids: "",
      campaign_labels: "",
      all_campaigns_for_accounts: []
    };
    // A map of "metric-type -> metric-thresholds"
    this.thresholds = {};
  }
}

/**
 * A metric's thresholds
 *
 */
class MetricThreshold {
  constructor() {
    this.metricType = undefind;
    this.high = undefind;
    this.low = undefind;
    this.ignore = undefind;
  }
}


/** ============= main ==================== */
/**
 * 
 **/
function main() {

  const sheetUtils = SheetUtils.getInstance();
  sheetUtils.validateInput();
  const cadConfig = sheetUtils.readInput();

  // process request
  let cadResults = getResultsForAllRelevantEntitiesUnderMCC(cadConfig);
  cadResults.forEach((item) => {
    item.truncateDecimalDigits();
  });

  // write to spreadsheet
  SheetUtils.getInstance().clearResults();
  SheetUtils.getInstance().writeMetaData(cadConfig);
  SheetUtils.getInstance().writeResults(cadResults);

  // send update by email
  MailHandler.getInstance().sendEmail(cadConfig, cadResults);
}


/** ============= Sheet utils  ==================== */
/**
 * @fileoverview Description of this file.
 */

const NamedRanges = {
  RESULTS_SHEET_NAME: 'results',
  EMAILS: 'emails',
  MCC_ID: 'MCC_ID',
  SHOW_ONLY_ANOMALIES: 'SHOW_ONLY_ANOMALIES',

  CURRENT_PERIOD_UNIT: 'current_period_unit',
  CURRENT_END_UNIT: 'current_end_unit',

  PAST_END_UNIT: 'past_end_unit',

  AVG_TYPE: 'AVG_TYPE',
  AVG_TYPE_ERROR: 'AVG_TYPE_ERROR',
  RESULTS_CURRENT_RANGE_DATES: 'results_current_range_dates',
  RESULTS_PAST_RANGE_DATES: 'results_past_range_dates',
  FIRST_DATA_ROW: 7,
  FIRST_DATA_COLUMN: 1,
  ENTITIY_IDS: "entity_ids",
  ENTITY_LABELS: "entity_labels",
  ENTITY_EXCLUDED_IDS: "entity_excluded_ids",
  ALL_CAMPAIGNS_FOR_ACCOUNTS: "all_campaigns_for_accounts",

  DATA_AGGREGATION: "data_aggregation",
};

const AVG_TYPE = {
  AVG_TYPE_DAILY_WEEKDAYS: "Daily Avg. Same Weekday as today ---> (Today) vs. (a few instances from past weeks)----> both: midnight till last data hour.",
  AVG_TYPE_DAILY_TODAY_VS_YESTERDAY: "Daily Avg. (Today) vs. (Yesterday)  -----> both: midnight till last data hour.",
  AVG_TYPE_HOURLY_TODAY: "Hourly Avg. Inside Today ------> (Last data hour) vs. (midnight till that hout)",
  AVG_TYPE_DAILY: "Daily Avg. Full days.",
  AVG_TYPE_WEEKLY: "Weekly (last 7 days) Avg",
  AVG_TYPE_CUSTOM: "Custom"
};


const TimeFrameUnits = {
  'Days': 1,
  'Weeks': 7
};


/**
 * Time utils methods
 */
class TimeUtils {

  /**
   * The singleton instance
   * @return {!TimeUtils} The singleton instance
   */
  static getInstance() {
    if (!this.instance) {
      this.instance = new TimeUtils();
    }
    return this.instance;
  }


  constructor() {
    this.TIMEZONE = AdsApp.currentAccount().getTimeZone();
    this.HOURS_BACK = 3;
  }
  /**
   * Get last queryable hour
   */
  getLastQueryableHourMinusHours(minusHours) {
    const NOW = new Date();
    const past = new Date(NOW.getTime() - (this.HOURS_BACK + minusHours) * 3600 * 1000);
    const pastHourStr = Utilities.formatDate(past, this.TIMEZONE, 'HH');
    return {
      hourInt: parseInt(pastHourStr, 10),
      query_date: ToStringFormatter.getInstance().getDateStringInTimeZone(past, 'yyyy-MM-dd'),
      sheet_date: ToStringFormatter.getInstance().getDateStringInTimeZone(past, 'dd/MM/YY'),
      weekday: `AND segments.day_of_week = '${weekdays[past.getDay()].toUpperCase()}'`,
      hourWhereClauseEqual: `AND segments.hour = ${pastHourStr}`,
      hourWhereClauseSmaller: `AND segments.hour < ${pastHourStr}`
    };
  }
}

// const timeUtils = new TimeUtils();

/**
 * Input sheet representation
 */
class SheetUtils {
  constructor() {
    let tmpSpreadsheet = SpreadsheetApp.openByUrl(CONFIG.spreadsheet_url);
    this.mySpreadsheet = tmpSpreadsheet;
    this.resultsSheet = tmpSpreadsheet.getSheetByName(NamedRanges.RESULTS_SHEET_NAME);
    this.monitoredMetrics = [];
  }


  getMonitoredMetrics() {
    const monitoredMetrics = [];
    if (this.monitoredMetrics && this.monitoredMetrics.length > 0) return this.monitoredMetrics;

    Object.keys(MetricTypes).forEach(function (key, index) {
      if (MetricTypes[key].isMonitored) {
        monitoredMetrics.push(key);
      }
    });
    this.monitoredMetrics = monitoredMetrics;
    return monitoredMetrics;
  }


  /**
   * The singleton instance
   * @return {!SheetUtils} The singleton instance
   */
  static getInstance() {
    if (!this.instance) {
      this.instance = new SheetUtils();
    }
    return this.instance;
  }

  /**
   * Validate sheet input
   */
  validateInput() {
    const mySpreadsheet = this.mySpreadsheet;
    const avgType = mySpreadsheet.getRangeByName(NamedRanges.AVG_TYPE).getValue();
    const currentEndUnitRange = mySpreadsheet.getRangeByName(NamedRanges.CURRENT_END_UNIT);
    if (avgType.includes("Weekday")) {
      currentEndUnitRange.setValue("Weeks");
    }
    if (avgType.includes("Weekly")) {
      currentEndUnitRange.setValue("Weeks");
    }
  }

  /**
   * Read input into CADConfig object
   * @return {!CadConfig} a filled cad config object.
   */
  readInput() {
    const mySpreadsheet = this.mySpreadsheet;
    let cadConfig = new CadConfig();
    cadConfig.users = mySpreadsheet.getRangeByName(NamedRanges.EMAILS).getValue();
    cadConfig.avgType = mySpreadsheet.getRangeByName(NamedRanges.AVG_TYPE).getValue();
    cadConfig.showOnlyAnomalies = mySpreadsheet.getRangeByName(NamedRanges.SHOW_ONLY_ANOMALIES).getValue();

    const lastQueryableHour = TimeUtils.getInstance().getLastQueryableHourMinusHours(0);
    cadConfig = this.fillLookbackDates(cadConfig);

    console.log(`cadConfig.avgType ==== ${JSON.stringify(cadConfig.avgType)}`);

    cadConfig = this.addSegmentsHourToQuery(cadConfig, lastQueryableHour);
    cadConfig = this.setDivideBy(cadConfig, lastQueryableHour);

    const entity_ids = mySpreadsheet.getRangeByName(NamedRanges.ENTITIY_IDS).getValue();
    const entity_labels = mySpreadsheet.getRangeByName(NamedRanges.ENTITY_LABELS).getValue();
    const entity_excluded_ids = mySpreadsheet.getRangeByName(NamedRanges.ENTITY_EXCLUDED_IDS).getValue();
    const data_aggregation = mySpreadsheet.getRangeByName(NamedRanges.DATA_AGGREGATION).getValue();

    if (data_aggregation === "Account") {
      cadConfig.accounts.account_ids = SheetUtils.getInstance().toArray(entity_ids);
      cadConfig.accounts.account_labels = SheetUtils.getInstance().toArray(entity_labels);
      cadConfig.accounts.excluded_account_ids = SheetUtils.getInstance().toArray(entity_excluded_ids);
    }

    if (data_aggregation === "Campaign") {
      cadConfig.campaigns.campaign_ids = entity_ids;
      cadConfig.campaigns.campaign_labels = entity_labels;
      cadConfig.campaigns.excluded_campaign_ids = entity_excluded_ids;
      cadConfig.campaigns.all_campaigns_for_accounts = SheetUtils.getInstance().toArray(mySpreadsheet.getRangeByName(NamedRanges.ALL_CAMPAIGNS_FOR_ACCOUNTS).getValue());
    }

    for (let metric of this.getMonitoredMetrics()) {
      cadConfig.thresholds[`${metric}_high`] =
        parseFloat(mySpreadsheet.getRangeByName(`${metric}_high`).getValue());
      cadConfig.thresholds[`${metric}_low`] =
        parseFloat(mySpreadsheet.getRangeByName(`${metric}_low`).getValue());
      cadConfig.thresholds[`${metric}_ignore`] =
        parseFloat(mySpreadsheet.getRangeByName(`${metric}_ignore`).getValue());
    }

    return cadConfig;
  }

  addSegmentsHourToQuery(cadConfig, other) {
    switch (cadConfig.avgType) {
      case AVG_TYPE.AVG_TYPE_DAILY:
      case AVG_TYPE.AVG_TYPE_WEEKLY:
      case AVG_TYPE.AVG_TYPE_CUSTOM:
        {
          cadConfig.hourSegmentInSelect = ``;
          cadConfig.hourSegmentsWhereClause = { current: "", past: "" };
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
        {
          cadConfig.hourSegmentInSelect = `, segments.date, segments.hour, segments.day_of_week `;
          cadConfig.hourSegmentsWhereClause = {
            current: `${other.hourWhereClauseSmaller} ${other.weekday}`,
            past: `${other.hourWhereClauseSmaller} ${other.weekday}`
          };
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
        {
          cadConfig.hourSegmentInSelect = `, segments.date, segments.hour `;
          cadConfig.hourSegmentsWhereClause = { current: other.hourWhereClauseSmaller, past: other.hourWhereClauseSmaller };
          break;
        }
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
        {
          cadConfig.hourSegmentInSelect = `,segments.date, segments.hour `;
          cadConfig.hourSegmentsWhereClause = { current: other.hourWhereClauseEqual, past: other.hourWhereClauseSmaller };
          break;
        }
    }
    return cadConfig;
  }

  setDivideBy(cadConfig) {
    switch (cadConfig.avgType) {
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
        {
          cadConfig.divideCurrentBy = 1;
          cadConfig.dividePastBy = cadConfig.lookbackInUnits.past_period_length;
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
        {
          cadConfig.divideCurrentBy = 1;
          cadConfig.dividePastBy = 1;
          break;
        }
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
        {
          cadConfig.divideCurrentBy = 1;

          const beforeLastQueryableHour = TimeUtils.getInstance().getLastQueryableHourMinusHours(1);
          cadConfig.dividePastBy = beforeLastQueryableHour.hourInt;
          break;
        }

      case AVG_TYPE.AVG_TYPE_DAILY:
      case AVG_TYPE.AVG_TYPE_WEEKLY:
      case AVG_TYPE.AVG_TYPE_CUSTOM:
        {
          cadConfig.divideCurrentBy = cadConfig.lookbackInUnits.current_period_length;
          cadConfig.dividePastBy = cadConfig.lookbackInUnits.past_period_length;
          break;
        }
    }
    return cadConfig;
  }

  fillLookbackDates(cadConfig) {
    const mySpreadsheet = this.mySpreadsheet;
    const lastQueryableHour = TimeUtils.getInstance().getLastQueryableHourMinusHours(0);

    switch (cadConfig.avgType) {
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
        {
          this.setSheetAndQueryDates(cadConfig.lookbackDates.current_range_start_date, lastQueryableHour);
          this.setSheetAndQueryDates(cadConfig.lookbackDates.current_range_end_date, lastQueryableHour);
          break;
        }
    }
    const currentAndPastPeriodUnit = TimeFrameUnits[mySpreadsheet.getRangeByName(NamedRanges.CURRENT_PERIOD_UNIT).getValue()] || 1;
    const currentEndUnit = TimeFrameUnits[mySpreadsheet.getRangeByName(NamedRanges.CURRENT_END_UNIT).getValue()] || 1;
    const pastEndUnit = TimeFrameUnits[mySpreadsheet.getRangeByName(NamedRanges.PAST_END_UNIT).getValue()] || 1;
    for (let el in cadConfig.lookbackInUnits) {
      cadConfig.lookbackInUnits[el] = mySpreadsheet.getRangeByName(el).getValue();
    }

    const toStringFormatter = ToStringFormatter.getInstance();
    let beforeLastQueryableHour = TimeUtils.getInstance().getLastQueryableHourMinusHours(1);

    switch (cadConfig.avgType) {
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
        {
          this.setSheetAndQueryDates(cadConfig.lookbackDates.past_range_start_date, beforeLastQueryableHour);
          this.setSheetAndQueryDates(cadConfig.lookbackDates.past_range_end_date, beforeLastQueryableHour);

          cadConfig.lookbackInDays.current_period_length = "partial 1";
          cadConfig.lookbackInDays.current_ended_length_ago = 0;
          cadConfig.lookbackInDays.past_period_length = "partial 1";
          cadConfig.lookbackInDays.past_ended_length_ago = 0;
          break;
        }

      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
        {
          beforeLastQueryableHour = TimeUtils.getInstance().getLastQueryableHourMinusHours(24);
          this.setSheetAndQueryDates(cadConfig.lookbackDates.past_range_start_date, beforeLastQueryableHour);
          this.setSheetAndQueryDates(cadConfig.lookbackDates.past_range_end_date, beforeLastQueryableHour);


          cadConfig.lookbackInDays.current_period_length = "partial 1";
          cadConfig.lookbackInDays.current_ended_length_ago = 0;
          cadConfig.lookbackInDays.past_period_length = "partial 1";
          cadConfig.lookbackInDays.past_ended_length_ago = 1;
          break;
        }
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
        {
          cadConfig.lookbackInDays.current_period_length = 1;
          cadConfig.lookbackInDays.current_ended_length_ago = 0;

          cadConfig.lookbackInDays.past_period_length =
            parseFloat(cadConfig.lookbackInUnits.past_period_length) * currentAndPastPeriodUnit;
          cadConfig.lookbackInDays.past_ended_length_ago =
            (parseFloat(cadConfig.lookbackInUnits.past_ended_length_ago) * pastEndUnit);

          cadConfig.lookbackDates.current_range_start_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.current_ended_length_ago - 1 +
              cadConfig.lookbackInDays.current_period_length);

          cadConfig.lookbackDates.current_range_end_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.current_ended_length_ago);

          cadConfig.lookbackDates.past_range_start_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.past_ended_length_ago - 1 +
              cadConfig.lookbackInDays.past_period_length);

          cadConfig.lookbackDates.past_range_end_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.past_ended_length_ago);

          break;
        }

      case AVG_TYPE.AVG_TYPE_DAILY:
      case AVG_TYPE.AVG_TYPE_WEEKLY:
      case AVG_TYPE.AVG_TYPE_CUSTOM:
        {
          cadConfig.lookbackInDays.current_period_length =
            parseFloat(cadConfig.lookbackInUnits.current_period_length) * currentAndPastPeriodUnit;
          cadConfig.lookbackInDays.current_ended_length_ago =
            (parseFloat(cadConfig.lookbackInUnits.current_ended_length_ago) * currentEndUnit);
          cadConfig.lookbackInDays.past_period_length =
            parseFloat(cadConfig.lookbackInUnits.past_period_length) * currentAndPastPeriodUnit;
          cadConfig.lookbackInDays.past_ended_length_ago =
            (parseFloat(cadConfig.lookbackInUnits.past_ended_length_ago) * pastEndUnit);

          cadConfig.lookbackDates.current_range_start_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.current_ended_length_ago +
              cadConfig.lookbackInDays.current_period_length);

          cadConfig.lookbackDates.current_range_end_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.current_ended_length_ago + 1);

          cadConfig.lookbackDates.past_range_start_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.past_ended_length_ago +
              cadConfig.lookbackInDays.past_period_length);

          cadConfig.lookbackDates.past_range_end_date =
            toStringFormatter.getStringForMinusDaysAgo(
              cadConfig.lookbackInDays.past_ended_length_ago + 1);
          break;
        }
    }
    return cadConfig;
  }


  setSheetAndQueryDates(cadConfigDate, other) {
    cadConfigDate.query_date = other.query_date;
    cadConfigDate.sheet_date = other.sheet_date;
  }

  /**
   * Write meta data to the sheet
   */
  writeMetaData(cadConfig) {
    this.mySpreadsheet.getRangeByName(NamedRanges.MCC_ID).setValue(cadConfig.mcc);
    this.setDatesInResultsSheet(cadConfig);
  }

  /**
   * Sets date ranges header in results sheet
   */
  setDatesInResultsSheet(cadConfig) {
    this.mySpreadsheet.getRangeByName(NamedRanges.RESULTS_CURRENT_RANGE_DATES)
      .setValue(
        `${cadConfig.lookbackDates.current_range_start_date.sheet_date} - ${cadConfig.lookbackDates.current_range_end_date.sheet_date} (${cadConfig.lookbackInDays.current_period_length} days)`);

    this.mySpreadsheet.getRangeByName(NamedRanges.RESULTS_PAST_RANGE_DATES)
      .setValue(
        `${cadConfig.lookbackDates.past_range_start_date.sheet_date} - ${cadConfig.lookbackDates.past_range_end_date.sheet_date} (${cadConfig.lookbackInDays.past_period_length} days)`);
  }


  /**
   * write CAD results to results sheet
   * @param {Array<!CadSingleResult>s} cadResults CAD results
   */
  writeResults(cadResults) {
    const newStartingRow =
      Math.max(NamedRanges.FIRST_DATA_ROW, this.resultsSheet.getLastRow() + 1);

    let newRows = [];
    cadResults.forEach(function (row) {
      newRows = newRows.concat(row.toSheetFormat());
    });


    console.log(`newRows = ${JSON.stringify(newRows)}`)
    if (newRows.length) {
      const metricListLength = Object.keys(this.getMonitoredMetrics()).length;
      this.resultsSheet
        .getRange(
          newStartingRow, NamedRanges.FIRST_DATA_COLUMN, newRows.length,
          (5 + 4 * metricListLength))
        .setValues(newRows);
    }
  }


  /**
   * delete all results from the results-sheet
   */
  clearResults() {
    const lastRow = this.resultsSheet.getLastRow();
    const howMany = lastRow - NamedRanges.FIRST_DATA_ROW + 1;
    if (howMany > 0) {
      this.resultsSheet.deleteRows(NamedRanges.FIRST_DATA_ROW, howMany);
    }
    let maxRows = this.resultsSheet.getMaxRows();
    if (maxRows < 350) {
      this.resultsSheet.insertRows(maxRows, 1000);
    }
  }


  getValueForNamedRange(name) {
    return this.mySpreadsheet.getRangeByName(name).getValue();
  }

  /**
   * Generates a string list from a value in a sheet named range.
   *
   * @param {string} value The value to convert.
   * @return {!Array<?>} Two variartions for a string format of that date.
   */
  toArray(value) {

    value = (value == "" || value == undefined) ?
      [] :
      value
        .replace(/"/g, '')
        .replace(/\s/g, '')
        .split(',');

    return value;
  }
}

/** ============= Google Ads Account Selector ==================== */

/**
 * Google Ads Account Selector
 */
class GoogleAdsAccountSelector {
  /**
   * GoogleAdsAccountSelector Constructor
   *
   * @param {!Account} mccAccount The MCC account.
   */
  constructor(mccAccount) {
    this.mccAccount = mccAccount;
  }

  /**
   * Returns all child accounts
   * @return {!Array<?>} account list
   */
  getAllSubAccounts() {
    const accountIterator = AdsManagerApp.accounts().get();
    let accounts = this.iterate(accountIterator);
    Logger.log('MCC has ' + accounts.length + ' child accounts');
    return accounts;
  };

  /**
   * Returns an account list for the id list.
   *
   * @param {!Array<?>} ids Id list
   * @return {!Array<?>} account list
   */
  getAccountObjectsForIds(ids) {
    Logger.log('accountIterator=' + JSON.stringify(ids));
    let accountIterator = AdsManagerApp.accounts().withIds(ids).get();
    return this.iterate(accountIterator);
  };

  /**
   * Gets an account list for the id list.
   *
   * @param {!AccountIterator} accountIterator An account iterator
   * @return {!Array<?>} account list
   */
  iterate(accountIterator) {
    let accounts = {};
    while (accountIterator.hasNext()) {
      let currentAccount = accountIterator.next();
      Logger.log('iterate(accountIterator) ' + currentAccount.getCustomerId());
      accounts[currentAccount.getCustomerId()] = currentAccount;
    }
    return accounts;
  };

  /**
   * Returns the account to traverse.
   * @param {!CadConfig} cadConfig CAD config
   * @return {!Set<?>} Account set to traverse
   */
  getAccountsToTraverse(cadConfig) {
    let accountsObjects = {};

    if (CONFIG.is_debug_log) {
      console.log('getAccountsToTraverse= ' + JSON.stringify(cadConfig));
    }

    //because the campaign can be anywhere - we need to scan all the child accounts anyhow.
    if (cadConfig.campaigns.campaign_ids != '') {
      Logger.log(JSON.stringify(cadConfig.campaigns.campaign_ids));
      return this.getAllSubAccounts();
    }
    if (cadConfig.accounts.account_ids.length > 0) {
      if (cadConfig.accounts.account_ids[0].toUpperCase() == 'ALL') {
        return this.getAllSubAccounts();
      } else {
        Object.assign(accountsObjects, this.getAccountObjectsForIds(cadConfig.accounts.account_ids));
      }
    }
    if (cadConfig.accounts.account_labels.length > 0 ||
      cadConfig.campaigns.campaign_labels.length > 0) {
      return this.getAllSubAccounts();
    }
    if (cadConfig.campaigns.all_campaigns_for_accounts != '') {
      const all_campaigns_for_accounts =
        this.getAccountObjectsForIds(cadConfig.campaigns.all_campaigns_for_accounts);
      Object.assign(accountsObjects, all_campaigns_for_accounts);
    }
    return accountsObjects;
  }


  /**
   * Returns an Array with the intersection between the monitored labels and the
   * account's labels.
   * @param {!Object} currentAccount Current account.
   * @param {!Object} cadConfig CAD config.
   * @return {!Array<?>} The monitored account labels which are relevant for the
   *     current account.
   */
  getRelevantLabelsForAccount(currentAccount, cadConfig) {
    let inputLabels = cadConfig.accounts.account_labels;
    const relevantLabels = [];
    for (const label of currentAccount.labels().get()) {
      if (label.getName() in inputLabels) relevantLabels.push(label.getName());
    }
    return relevantLabels;
  }
}



/** ============= Google Ads Campaign Selector ==================== */


/**
 * Google Ads Campaign Selector
 */
class GoogleAdsCampaignSelector {
  /**
   * Helper: generates a relevant campaign map for account
   * @param {!Object} currentAccount The current account.
   * @param {!Object} cadConfig CAD config.
   * @return {!Object} The relevant campaign map for the current account. Key:
   *     campaign-Id --> label name is exists
   */
  reportToRelevantCampaignMap(currentAccount, cadConfig) {
    let selectCampaignQuery;
    let campaignMapForCurrentAccount = {};
    const campaignIds = cadConfig.campaigns.campaign_ids;
    const all_campaigns_for_accounts =
      cadConfig.campaigns.all_campaigns_for_accounts;
    const campaignLabels = cadConfig.campaigns.campaign_labels;
    const excludedCampaignIds = cadConfig.campaigns.excluded_campaign_ids;
    const excludedClause = (!excludedCampaignIds || excludedCampaignIds == '') ?
      '' :
      `AND campaign.id NOT IN ('${excludedCampaignIds}' )`;

    if (all_campaigns_for_accounts != '') {
      //'ALL_ENABLED_FOR_ACCOUNT_'
      let forAccountIds = cadConfig.campaigns.all_campaigns_for_accounts;
      if (CONFIG.is_debug_log) {
        Logger.log(`cadConfig.campaigns.all_campaigns_for_accounts =  ${JSON.stringify(cadConfig.campaigns.all_campaigns_for_accounts)}`);
      }
      if (forAccountIds.includes(currentAccount.getCustomerId())) {
        selectCampaignQuery = `SELECT campaign.id, campaign.name FROM campaign`;
        if (CONFIG.is_debug_log) {
          Logger.log(
            `campaigns for current account's query = ${selectCampaignQuery}`);
        }
        campaignMapForCurrentAccount = this.campaignReportToKeyMap(
          AdsApp.report(selectCampaignQuery, CONFIG.reporting_options),
          campaignMapForCurrentAccount);
      }
    }
    if (campaignIds !== '') {
      selectCampaignQuery =
        `SELECT campaign.id, campaign.name FROM campaign WHERE campaign.id IN (${campaignIds})  ${excludedClause}`;

      if (CONFIG.is_debug_log) {
        Logger.log(`campaignIds query = ${selectCampaignQuery}`);
      }
      campaignMapForCurrentAccount = this.campaignReportToKeyMap(
        AdsApp.report(selectCampaignQuery, CONFIG.reporting_options),
        campaignMapForCurrentAccount);
    }
    // Adding campaigns by labels
    if (campaignLabels && campaignLabels.length > 0) {
      selectCampaignQuery =
        `SELECT campaign.id, campaign.name, label.name FROM campaign_label WHERE label.name IN (${campaignLabels}) ${excludedClause}`;

      if (CONFIG.is_debug_log) {
        Logger.log(`campaignLabels query = ${selectCampaignQuery}`);
      }
      campaignMapForCurrentAccount = this.campaignReportToKeyMap(
        AdsApp.report(selectCampaignQuery, CONFIG.reporting_options),
        campaignMapForCurrentAccount);
    }
    return campaignMapForCurrentAccount;
  }


  /**
   * Populates a hash-map from "campaign-id" to label name if exists
   * @param {!Object} searchResults Query results.
   * @param {!Object} campaignsMap a hash-map with "campaign-id" keys
   * @return {!Object} a hash-map with "campaign-id" keys
   */
  campaignReportToKeyMap(searchResults, campaignsMap) {
    if (!(campaignsMap instanceof Object)) {
      throw 'rowsToCampaignMap expect campaignsMap to be a map';
    }
    for (const row of searchResults.rows()) {
      const campaignId = row['campaign.id'];


      if (!campaignsMap[campaignId]) {
        campaignsMap[campaignId] = [];
      }
      if (row['label.name']) {
        campaignsMap[campaignId].push(row['label.name']);
      }
    }
    return campaignsMap;
  }
}


/** ============= Google Ads Reporter ==================== */

let _gads_query_fields = [];

/**
 * returns a string with all the GAQL relevant query fields 
 * @return {string} a string with all the GAQL relevant query fields 
 */
function getGadsQueryFields() {
  if (_gads_query_fields.length == 0) {

    Object.keys(MetricTypes).forEach(function (key, index) {
      if (MetricTypes[key].isFromGoogleAds) {
        _gads_query_fields.push(`metrics.${key}`);
      }
    });
    _gads_query_fields = _gads_query_fields.join(',');
  }
  return _gads_query_fields;
}

/**
 * Get results for relevant entities under MCC
 * @param {!Object} cadConfig CAD user input.
 * @return {Array<!CadSingleResult>} CAD monitoring results
 */
function getResultsForAllRelevantEntitiesUnderMCC(cadConfig) {
  let mccAccount = AdsApp.currentAccount();
  cadConfig.mcc = mccAccount.getCustomerId();
  const gAdsAccountSelector = new GoogleAdsAccountSelector(mccAccount);

  let allEntitiesAllMetricsComparisons = [];

  // Get accounts to traverse
  const accountsToTraverse =
    gAdsAccountSelector.getAccountsToTraverse(cadConfig);

  for (const currentAccount of Object.values(accountsToTraverse)) {
    console.log(`Switching to account ${currentAccount.getCustomerId()}`);

    AdsManagerApp.select(currentAccount);
    allEntitiesAllMetricsComparisons = allEntitiesAllMetricsComparisons.concat(
      getResultsForRelevantEntitiesUnderAccount(gAdsAccountSelector, currentAccount, cadConfig));
  }

  return allEntitiesAllMetricsComparisons;
}

/**
 * Get results for relevant entities under account
 * @param {!Object} gAdsAccountSelector Google ads account selector
 * @param {!Object} currentAccount The current account.
 * @param {!Object} cadConfig CAD user input.
 * @return {Array<!CadSingleResult>} CAD monitoring results
 */
function getResultsForRelevantEntitiesUnderAccount(gAdsAccountSelector, currentAccount, cadConfig) {
  let cadResults = [];
  cadResults = cadResults.concat(
    aggAccountReportToCadResults(gAdsAccountSelector, currentAccount, cadConfig));

  const gAdsCampaignSelector = new GoogleAdsCampaignSelector();
  const relevantCampaignMap =
    gAdsCampaignSelector.reportToRelevantCampaignMap(currentAccount, cadConfig);
  cadResults = cadResults.concat(campaignReportToCadResults(relevantCampaignMap, cadConfig));

  if (CONFIG.is_debug_log) {
    Logger.log(
      'getResultsForRelevantEntitiesUnderAccount= ' +
      JSON.stringify(cadResults.rows));
  }
  return cadResults;
}


/**
 * Aggregated account report to cad monitoring results
 * @param {!Object} gAdsAccountSelector Google ads account selector
 * @param {!Object} currentAccount The current account.
 * @param {!Object} cadConfig CAD user input.
 * @return {!Object} CAD monitoring results
 */
function aggAccountReportToCadResults(gAdsAccountSelector, currentAccount, cadConfig) {
  const excludedAggAccountIds = cadConfig.accounts.excluded_account_ids;
  let cadResults = [];

  const relevantLabelsForCurrentAccount =
    gAdsAccountSelector.getRelevantLabelsForAccount(
      currentAccount, cadConfig);

  // Does currently scanned account satisfy any input id filter
  if (!isCurrentAccountSatisfyCadConfig(
    currentAccount, cadConfig, relevantLabelsForCurrentAccount)) {
    return cadResults;
  }

  if (excludedAggAccountIds.includes(currentAccount.getCustomerId())) {
    return cadResults;
  }

  // Customer query
  const getGadsCustomerQueryFields = removeElementFromStringList("metrics.search_click_share", getGadsQueryFields());
  console.log(`getGadsCustomerQueryFields for CUSTOMER = ${getGadsCustomerQueryFields}`)
  let baseQuery = `SELECT customer.descriptive_name, ${getGadsCustomerQueryFields} ${cadConfig.hourSegmentInSelect} FROM customer WHERE segments.date BETWEEN`;


  let currentQuery = `${baseQuery} "${cadConfig.lookbackDates.current_range_start_date.query_date}" AND "${cadConfig.lookbackDates.current_range_end_date.query_date}" 
  ${cadConfig.hourSegmentsWhereClause.current}`;
  let pastQuery = `${baseQuery} "${cadConfig.lookbackDates.past_range_start_date.query_date}" AND "${cadConfig.lookbackDates.past_range_end_date.query_date}" ${cadConfig.hourSegmentsWhereClause.past}`;

  let currentStats = storeReportByEntityId(
    AdsApp.report(currentQuery, CONFIG.reporting_options), cadConfig);

  let pastStats =
    storeReportByEntityId(AdsApp.report(pastQuery, CONFIG.reporting_options), cadConfig);

  if (CONFIG.is_debug_log) {
    Logger.log('currentStats accounts= ' + JSON.stringify(currentStats));
    Logger.log('currentQuery accounts= ' + JSON.stringify(currentQuery));
    Logger.log('pastQuery accounts= ' + JSON.stringify(pastQuery));
    Logger.log('pastStats accounts= ' + JSON.stringify(pastStats));
  }

  // single Row
  for (id in currentStats) {
    let cadSingleResult = new CadSingleResult();
    cadSingleResult.relevant_label = relevantLabelsForCurrentAccount;

    cadSingleResult.account.id = AdsApp.currentAccount().getCustomerId();
    cadSingleResult.account.name =
      currentStats[id]['customer.descriptive_name'];


    cadSingleResult.campaign.id = '';
    cadSingleResult.campaign.name = '';

    cadSingleResult.fillMetricResults(id, pastStats, currentStats, cadConfig);

    if (CONFIG.is_debug_log) {
      Logger.log(
        'aggAccountReportToCadResults scanned= ' + JSON.stringify(cadSingleResult));
    }
    if (cadSingleResult.isTriggerAlert) {

      if (CONFIG.is_debug_log) {
        Logger.log(
          'aggAccountReportToCadResults cadSingleResult triggers alert= ' + JSON.stringify(cadSingleResult));
      }
      cadResults.push(cadSingleResult);
    }
  }

  if (CONFIG.is_debug_log) {
    Logger.log('aggAccountReportToCadResults results= ' + JSON.stringify(cadResults));
  }
  return cadResults;
}

function removeElementFromStringList(elementName, elementString) {
  var elementsArray = elementString.split(",");
  var indexToRemove = elementsArray.indexOf(elementName);
  // Use the splice() method to remove the element at the specified index
  if (indexToRemove > -1) {
    elementsArray.splice(indexToRemove, 1);
  }
  return elementsArray.join(",");
}

/**
 * Does current account satisfy CAD config
 * @param {!Object} currentAccount The current account.
 * @param {!Object} cadConfig CAD user input.
 * @param {!Object} relevantLabelsForCurrentAccount relevant labels for current
 *     account.
 * @return {!Object} CAD monitoring results
 */
function isCurrentAccountSatisfyCadConfig(
  currentAccount, cadConfig, relevantLabelsForCurrentAccount) {
  const aggAccountIdList = cadConfig.accounts.account_ids;
  const currentAccountId = currentAccount.getCustomerId();

  if (CONFIG.is_debug_log) {
    Logger.log(`Current account matches these monitored account labels = ${relevantLabelsForCurrentAccount}`);
  }

  return (
    (aggAccountIdList.length > 0 &&
      aggAccountIdList[0].toUpperCase() === 'ALL') ||

    aggAccountIdList.includes(currentAccountId) ||

    (relevantLabelsForCurrentAccount.length > 0));
}

/**
 * campaign report to cad monitoring results
 * @param {!Object} campaignIdToLabelForCurrentAccount The current account.
 * @param {!Object} cadConfig CAD user input.
 * @return {Array<!CadSingleResult>} CAD monitoring results
 */
function campaignReportToCadResults(
  campaignIdToLabelForCurrentAccount, cadConfig) {
  let cadResults = [];

  let campaignIds = Object.keys(campaignIdToLabelForCurrentAccount);

  if (!campaignIds.length) {
    return cadResults;
  }

  campaignIds = campaignIds.join(",");
  if (CONFIG.is_debug_log) {
    Logger.log(`Reporting for campaign ids = ${JSON.stringify(campaignIds)}`);
  }

  // We expect current searchResults to contain DISABLED campaigns as well
  // thus > past results
  //  Campaign query
  let specificCampaignsQuery =
    `SELECT customer.descriptive_name, campaign.id, campaign.name, ${getGadsQueryFields()} ${cadConfig.hourSegmentInSelect} FROM campaign WHERE campaign.id IN (${campaignIds}) AND segments.date BETWEEN`;

  switch (cadConfig.avgType) {
    case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
    case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
    case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
      {
        specificCampaignsQuery = removeElementFromStringList("metrics.search_click_share", specificCampaignsQuery);
        break;
      };
  }

  let currentQuery = `${specificCampaignsQuery} "${cadConfig.lookbackDates.current_range_start_date.query_date}" AND "${cadConfig.lookbackDates.current_range_end_date.query_date}" ${cadConfig.hourSegmentsWhereClause.current}`;
  let pastQuery = `${specificCampaignsQuery} "${cadConfig.lookbackDates.past_range_start_date.query_date}" AND "${cadConfig.lookbackDates.past_range_end_date.query_date}"  ${cadConfig.hourSegmentsWhereClause.past}`;

  if (CONFIG.is_debug_log) {
    Logger.log(
      'currentQuery campaigns= ' +
      JSON.stringify(currentQuery));
    Logger.log(
      'pastQuery campaigns= ' +
      JSON.stringify(pastQuery));
  }

  let pastStats = storeReportByEntityId(AdsApp.report(pastQuery, CONFIG.reporting_options), cadConfig);
  let currentStats = storeReportByEntityId(AdsApp.report(currentQuery, CONFIG.reporting_options), cadConfig);

  if (CONFIG.is_debug_log) {
    Logger.log(
      'currentStats= ' +
      JSON.stringify(currentStats));
    Logger.log(
      'pastStats= ' +
      JSON.stringify(pastStats));
  }

  // Row
  for (id in currentStats) {
    let cadSingleResult = new CadSingleResult();
    cadSingleResult.relevant_label =
      campaignIdToLabelForCurrentAccount[id];
    cadSingleResult.account.id = AdsApp.currentAccount().getCustomerId();
    cadSingleResult.account.name =
      currentStats[id]['customer.descriptive_name'];
    cadSingleResult.campaign.id = currentStats[id]['campaign.id'];
    cadSingleResult.campaign.name = currentStats[id]['campaign.name'];
    cadSingleResult.fillMetricResults(id, pastStats, currentStats, cadConfig);
    if (cadSingleResult.isTriggerAlert) {
      cadResults.push(cadSingleResult);
    }
  }
  return cadResults;
}


/**
 * GAds report to metric stats map
 * @param {!Object} searchResults GAds search results
 * @param {!Object} cadConfig CAD config
 * @return {!Object} metric stats map.
 */
function storeReportByEntityId(searchResults, cadConfig) {
  let statsMap = {};
  for (const row of searchResults.rows()) {
    let id = (row["campaign.id"]) ? row["campaign.id"] :
      AdsApp.currentAccount().getCustomerId();
    Logger.log("storeReportByEntityId for id = " + id + " row=" + JSON.stringify(row));

    if (!statsMap[id]) {
      statsMap[id] = row;
      if (row["campaign.id"]) {
        switch (cadConfig.avgType) {
          case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
          case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
          case AVG_TYPE.AVG_TYPE_HOURLY_TODAY:
            {
              statsMap[id]["metrics.search_click_share"] = 0;
              break;
            };
        }
      }
    }
    else {
      statsMap[id] = sumRows(row, statsMap[id]);
    }
  }
  return statsMap;
}

function sumRows(row1, row2) {
  const result = {};
  for (var key in row1) {
    if (row1.hasOwnProperty(key) && row2.hasOwnProperty(key)) {
      var val1 = row1[key];
      var val2 = row2[key];
      if (key.toLocaleLowerCase().includes("id") || key.toLocaleLowerCase().includes("name")) {
        result[key] = val1;
      }
      else if (isNumber(val1) && isNumber(val2)) {
        result[key] = Number(val1) + Number(val2);
      } else if (typeof val1 === 'string' || typeof val2 === 'string') {
        result[key] = 'cannot sum string values';
      }
    }
  }
  return result;
}

function isNumber(x) {
  return !isNaN(x);
}



/** ============= Mail Handler ==================== */
/**
 * Mail handler class
 */
class MailHandler {
  /**
   * The singleton instance
   * @return {!MailHandler} The singleton instance
   */
  static getInstance() {
    if (!this.instance) {
      this.instance = new MailHandler();
    }
    return this.instance;
  }

  /**
   * Sends an email
   *
   * @param {!Object} cadConfig CAD user input.
   * @param {!Object} cadResults CAD monitoring results
   */
  sendEmail(cadConfig, cadResults) {
    const emails = cadConfig.users;

    const triggeringResults = cadResults.filter(row => row.isTriggerAlert);

    if (triggeringResults.length <= 0) {
      Logger.log('No alerts have been triggered. No email has been sent.');
      return;
    }

    if (!(emails && emails.length > 0)) {
      Logger.log('There are alerts triggered. But no email to send to.');
      return;
    }

    let bodyText = [];
    cadResults.forEach(item => bodyText.push(item.toEmailFormat()));
    bodyText = bodyText.join('<br>');


    MailApp.sendEmail({
      to: emails,
      subject: `[Anomaly Alert] MCC Account [${cadConfig.mcc}]`,
      htmlBody: `
          Avg Type:
          ${cadConfig.avgType}
          <br><br>
          New period:
          ${cadConfig.lookbackDates.current_range_start_date.sheet_date} - ${cadConfig.lookbackDates.current_range_end_date.sheet_date} (${cadConfig.lookbackInDays.current_period_length} days)<br>
          Past period:
          ${cadConfig.lookbackDates.past_range_start_date.sheet_date} - ${cadConfig.lookbackDates.past_range_end_date.sheet_date} (${cadConfig.lookbackInDays.past_period_length} days)
          <br><br>
          ${bodyText}
          <br><br>
          Log into Google Ads and take a look.
          Notifications dashboard and full results:  ${CONFIG.spreadsheet_url}`
    });
    Logger.log('Email sent');
  }
}


/** ============= toString formatter ==================== */

/**
 * A custom string formatter class
 */
class ToStringFormatter {
  /**
   * Gets an instance of the singleton
   *
   * @return {!ToStringFormatter} An instance of the singleton
   */
  static getInstance() {
    if (!this.instance) {
      this.instance = new ToStringFormatter();
    }
    return this.instance;
  }

  /**
   * Produces an object with two strings formatted in yyyy-MM-dd and dd/MM/YY.
   *
   * @param {!int} numDays How many days to subtract.
   * @return {!Object} An object with two strings formatted in yyyy-MM-dd and
   *     dd/MM/YY
   */
  getStringForMinusDaysAgo(numDays) {
    const expectedDate =
      new Date(new Date().setDate(new Date().getDate() - numDays));
    return {
      'query_date': this.getDateStringInTimeZone(expectedDate, 'yyyy-MM-dd'),
      'sheet_date': this.getDateStringInTimeZone(expectedDate, 'dd/MM/YY')
    };
  }

  /**
   * Produces a formatted string representing a given date in a given time zone.
   *
 
   * @param {?Date=} date A date object. Defaults to the current date.
   * @param {string} format A format specifier for the string to be produced.
   * @return {string} A date string representation.
   */
  getDateStringInTimeZone(date, format) {
    return Utilities.formatDate(
      date, AdsApp.currentAccount().getTimeZone(), format);
  }


  /**
   * Produces a string representation for a number with decimal digits.
   *
   * @param {number} num the number to format.
   * @param {number} digits digits after the period
   * @return {string} A string representation with decimal digits.
   */
  formatNumWithDigitAfterPeriod(num, digits) {
    return String(num.toFixed(digits));
  }


  /**
   * Produces a string representation for a number with a thousands comma.
   *
   * @param {number} num the number to format.
   * @return {string} A string representation with a thousands comma.
   */
  addNumberCommas(num) {
    if (isNaN(num)) return num;
    let str = num.toString().split('.');
    str[0] = str[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return str.join('.');
  }
}