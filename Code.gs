/** ============= Cad Result ==================== */

/**
 * A metric's performance
 *
 */
class MetricResult {
  constructor() {
    this.metricType = undefined;
    this.isWriteToSheet = true;
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
  cost_micros: {
    email: 'Cost',
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
  average_cpc: {
    email: 'CPC',
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
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

  all_conversions_value:
    { email: 'All conversions value', isFromGoogleAds: true },

  conversions_value: { email: 'Conversions value', isFromGoogleAds: true },


  search_click_share: {
    email: 'Search click share',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },
  average_cpm: {
    email: 'Avg CVM',
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
    email: 'View rate',
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true
  },

  roas_all: {
    email: 'ROAS all',
    isMonitored: true,
    shouldBeRounded: true,
    isWriteToSheet: true
  },
  roas: {
    email: 'ROAS',
    isMonitored: true,
    shouldBeRounded: true,
    isWriteToSheet: true
  },
};

/**
 * An entity representation to monitor
 */
class CadSingleResult {
  constructor() {
    this.relevant_label = undefined;
    this.account = { id: undefined, name: undefined };
    this.campaign = { id: undefined, name: undefined };

    this.isTriggerAlert = false;
    this.allMetricsComparisons = {};

    let currencyCode = AdsApp.currentAccount().getCurrencyCode();
    currencyCode = (currencyCode === 'USD') ? '$' : currencyCode;
    this.currencyCode = currencyCode + ' ';
  }



  /**
   * Returns a string for an email format
   *
   * @param {string} id Entity id
   * @param {!Object} pastStats past stats map
   * @param {!Object} currentStats current stats map
   * @param {!Object} cadConfig CAD config
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
      this.allMetricsComparisons[gAdsMetric] = {};
      let comparisonResult = this.allMetricsComparisons[gAdsMetric];

      if (MetricTypes[metric].isMicro) {
        currentStats[id][gAdsMetric] /= 1e6;
        pastStats[id][gAdsMetric] /= 1e6;
      }
      if (MetricTypes[metric].isCumulative) {
        comparisonResult.past = (pastStats[id]) ? pastStats[id][gAdsMetric] /
          cadConfig.lookbackInDays.past_range_period :
          0;
        comparisonResult.current = currentStats[id][gAdsMetric] /
          cadConfig.lookbackInDays.current_range_period;

      } else if (gAdsMetric.includes('roas_all')) {
        comparisonResult.past = (pastStats[id]) ?
          pastStats[id]['metrics.all_conversions_value'] /
          pastStats[id]['metrics.cost_micros'] :
          0;
        comparisonResult.current =
          currentStats[id]['metrics.all_conversions_value'] /
          currentStats[id]['metrics.cost_micros'];

      } else if (gAdsMetric.includes('roas')) {
        comparisonResult.past = (pastStats[id]) ?
          pastStats[id]['metrics.conversions_value'] /
          pastStats[id]['metrics.cost_micros'] :
          0;
        comparisonResult.current = currentStats[id]['metrics.conversions_value'] /
          currentStats[id]['metrics.cost_micros'];
      }

      else {
        comparisonResult.past = (pastStats[id]) ? pastStats[id][gAdsMetric] : 0;
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
        }
      }

      if (CONFIG.is_debug_log) {
        Logger.log(`=====comparisonResult= ` + JSON.stringify(comparisonResult));
        Logger.log(`=====this.metricComparisonResults= ` + JSON.stringify(this.allMetricsComparisons));
        Logger.log(`ignoreAbs= ` + JSON.stringify(ignoreAbs));
        Logger.log(
          `comparisonResult.metricAlertDirection= ` +
          JSON.stringify(comparisonResult.metricAlertDirection));
      }

      this.isTriggerAlert = this.isTriggerAlert ||
        (comparisonResult.metricAlertDirection != undefined);
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
     <br><table style="width:50%;border:1px solid black;"><tr style="border:1px solid black;"><th style="text-align:left;border:1px solid black;">Metric</th><th style="text-align:left;border:1px solid black;">Current</th><th style="text-align:left;border:1px solid black;">Past</th> <th style="text-align:left;border:1px solid black;">Δ</th> <th style="text-align:left;border:1px solid black;">Δ%</th></tr>`
    ];
    MetricTypes

    for (let metricType in this.allMetricsComparisons) {
      let metricType_name = metricType.split(".")[1];
      if (!MetricTypes[metricType_name].isWriteToSheet) continue;
      if (!this.allMetricsComparisons[metricType].metricAlertDirection) continue;


      let currentAvgValue = this.allMetricsComparisons[metricType].current;
      let pastAvgValue = this.allMetricsComparisons[metricType].past;
      let numericDelta = this.allMetricsComparisons[metricType].changeAbs;
      let percentageDelta = this.allMetricsComparisons[metricType].changePercent;
      let metricAlertDirection =
        this.allMetricsComparisons[metricType].metricAlertDirection;

      let currentValueStr = '';
      let pastAvgValueStr = '';
      let numericDeltaStr = '';
      let percentageDeltaStr = '';

      let metricTypeName = metricType.split('.')[1];
      if (MetricTypes[metricTypeName].shouldBeRounded) {
        const __ret = this.toRoundedMetricString(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta);
        currentValueStr = __ret.currentValueStr;
        pastAvgValueStr = __ret.pastAvgValueStr;
        numericDeltaStr = __ret.numericDeltaStr;
        percentageDeltaStr = __ret.percentageDeltaStr;

      }
      // A stat already avg for period
      else if (
        metricType.includes('cost_per_conversion') ||
        metricType.includes('cost_micros')) {
        const __ret = this.toMetricStrings(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta,
          this.currencyCode + ' ', '', 2);
        currentValueStr = __ret.currentValueStr;
        pastAvgValueStr = __ret.pastAvgValueStr;
        numericDeltaStr = __ret.numericDeltaStr;
        percentageDeltaStr = __ret.percentageDeltaStr;
      }
      // A stat already avg for period
      else if (metricType.includes('ctr')) {
        const __ret = this.toMetricStrings(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta, '',
          '%', 1);
        currentValueStr = __ret.currentValueStr;
        pastAvgValueStr = __ret.pastAvgValueStr;
        numericDeltaStr = __ret.numericDeltaStr;
        percentageDeltaStr = __ret.percentageDeltaStr;
      }
      // A stat already avg for period
      else if (
        metricType.includes('conversions_from_interactions_rate') ||
        metricType.includes('search_click_share') ||
        metricType.includes('video_view_rate') ||
        metricType.includes('average_cpc') ||
        metricType.includes('average_cpm') ||
        metricType.includes('average_cpv')) {
        const __ret = this.toMetricStrings(
          currentAvgValue, pastAvgValue, numericDelta, percentageDelta, '',
          '', 1);
        currentValueStr = __ret.currentValueStr;
        pastAvgValueStr = __ret.pastAvgValueStr;
        numericDeltaStr = __ret.numericDeltaStr;
        percentageDeltaStr = __ret.percentageDeltaStr;
      }

      const toStringFormatter = ToStringFormatter.getInstance();
      currentValueStr = toStringFormatter.addNumberCommas(currentValueStr);
      pastAvgValueStr = toStringFormatter.addNumberCommas(pastAvgValueStr);
      numericDeltaStr = toStringFormatter.addNumberCommas(numericDeltaStr);

      if (metricAlertDirection) {
        percentageDeltaStr = metricAlertDirection.includes('up') ?
          `⇪⇪ ${percentageDeltaStr}` :
          `⇩⇩ ${percentageDeltaStr}`;
      }


      alertTextForEntity.push(`<tr><td> ${metricType} </td><td> ${currentValueStr} </td><td> ${pastAvgValueStr} </td><td>
${numericDeltaStr} </td><td> ${percentageDeltaStr} </td></tr>`);
    }
    return `${alertTextForEntity.join('')} </table>`;
  };



  /**
   * @param {number} currentAvgValue current value (avg)
   * @param {number} pastAvgValue  past value (avg)
   * @param {number} numericDelta  numeric delta
   *@param {number} percentageDelta percentage delta
   *@return {!Object} a map of string formats
   */
  toRoundedMetricString(
    currentAvgValue, pastAvgValue, numericDelta, percentageDelta) {
    return {
      currentValueStr: Math.round(currentAvgValue),
      pastAvgValueStr: Math.round(pastAvgValue),
      numericDeltaStr: Math.round(numericDelta),
      percentageDeltaStr: Math.round(percentageDelta) + '%'
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

      let currentAvgValue = this.allMetricsComparisons[metricType].current;
      let pastAvgValue = this.allMetricsComparisons[metricType].past;
      let numericDelta = this.allMetricsComparisons[metricType].changeAbs;
      let percentageDelta = this.allMetricsComparisons[metricType].changePercent;
      let metricAlertDirection =
        this.allMetricsComparisons[metricType].metricAlertDirection;

      let postfix = '';
      if (metricType.includes('ctr')) {
        postfix = '%';
      }

      rowData = rowData.concat([
        currentAvgValue + postfix, pastAvgValue + postfix,
        numericDelta + postfix,
        this.addTriggerDirectionSign(metricAlertDirection, percentageDelta)
      ]);
    }
    return [rowData];
  };

  /**
   * Returns a string for a sheet format
   *
   * @param {string} metricAlertDirection is the alert for up/down anomaly
   * @param {string} metricValue the metrics value
   * @return {!Array<?>} a string for a sheet format
   */
  addTriggerDirectionSign(metricAlertDirection, metricValue) {
    if (metricAlertDirection) {
      metricValue = metricAlertDirection.includes('up') ? `⇪⇪ ${metricValue}` :
        `⇩⇩ ${metricValue}`;
    }
    return metricValue;
  }
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
    this.avg_type = "days";

    this.lookbackInUnits = {
      current_range_day: undefined,
      current_range_period: undefined,
      past_range_day: undefined,
      past_range_period: undefined
    };

//after code calculation:
    this.lookbackInDays = {
      current_range_day: undefined,
      current_range_period: undefined,
      past_range_day: undefined,
      past_range_period: undefined
    };

    this.lookbackDates = {
      current_range_start_date: undefined,
      current_range_end_date: undefined,
      past_range_start_date: undefined,
      past_range_end_date: undefined,
    };
    this.accounts = {
      account_ids: [],
      excluded_account_ids: [],
      account_labels: []
    };
    this.campaigns = {
      campaign_ids: [],
      all_campaigns_for_accounts: [],
      excluded_campaign_ids: [],
      campaign_labels: []
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

  const sheet = SheetUtils.getInstance();
  sheet.validateInput();
  const cadConfig = sheet.readInput();

  // process request
  const cadResults = getResultsForAllRelevantEntitiesUnderMCC(cadConfig);


  //Logger.log(JSON.stringify(cadResults));

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

// import {CadConfig} from './cadClasses'
// import {MetricTypes, TimeFrameUnits} from './consts'


const CONSTS = {
  RESULTS_SHEET_NAME: 'results',
  EMAILS: 'emails',

  CURRENT_RANGE_UNIT: 'current_range_unit',
  
  CURRENT_END_UNIT: 'current_end_unit',
  PAST_END_UNIT: 'past_end_unit',
  
  AVG_TYPE : 'AVG_TYPE',
  AVG_TYPE_ERROR : 'AVG_TYPE_ERROR',
  RESULTS_CURRENT_RNAGE_DATES: 'results_current_range_dates',
  RESULTS_PAST_RNAGE_DATES: 'results_past_range_dates',
  FIRST_DATA_ROW: 7,
  FIRST_DATA_COLUMN: 1,
};


const TimeFrameUnits = {
  'days': 1,
  'weeks': 7
};

const AvgTypes = {
  'current weekday (00:00 till now)': 1,
  'days': 1,
  'weeks': 7
};


/**
 * Input sheet representation
 */
class SheetUtils {
  constructor() {
    let my_shpreadsheet = SpreadsheetApp.openByUrl(CONFIG.spreadsheet_url);
    this.my_shpreadsheet = my_shpreadsheet;
    this.resultsSheet =
      my_shpreadsheet.getSheetByName(CONSTS.RESULTS_SHEET_NAME);
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
    const sheet = this.my_shpreadsheet;
    
    const avgType =sheet.getRangeByName(CONSTS.AVG_TYPE).getValue();
    const currentEndUnitRange = sheet.getRangeByName(CONSTS.CURRENT_END_UNIT);


    if (avgType.includes("weekday")){
      currentEndUnitRange.setValue("weeks");
    }
    if (avgType.includes("weeks")){
      currentEndUnitRange.setValue("weeks");
    } 
  }

  /**
   * Read input into CADConfig object
   * @return {!CadConfig} a filled cad config object.
   */
  readInput() {
    const sheet = this.my_shpreadsheet;
    let cadConfig = new CadConfig();

    cadConfig.users = sheet.getRangeByName(CONSTS.EMAILS).getValue();

    // get lookback window
    const lookbackUnitsFirstFilter = TimeFrameUnits[sheet.getRangeByName(CONSTS.CURRENT_RANGE_UNIT).getValue()] || 1;
    const currentRangeEndedUnit = TimeFrameUnits[sheet.getRangeByName(CONSTS.CURRENT_END_UNIT).getValue()] || 1;
    const pastRangeEndedUnit = TimeFrameUnits[sheet.getRangeByName(CONSTS.PAST_END_UNIT).getValue()] || 1;


    for (let el in cadConfig.lookbackInUnits) {
      cadConfig.lookbackInUnits[el] = sheet.getRangeByName(el).getValue();
    }

    const toStringFormatter = ToStringFormatter.getInstance();

    cadConfig.lookbackInDays.current_range_day =
      parseFloat(cadConfig.lookbackInUnits.current_range_day) * lookbackUnitsFirstFilter;
    cadConfig.lookbackInDays.current_range_period =
      (parseFloat(cadConfig.lookbackInUnits.current_range_period) * currentRangeEndedUnit);
    cadConfig.lookbackInDays.past_range_day =
      parseFloat(cadConfig.lookbackInUnits.past_range_day) * lookbackUnitsFirstFilter;
    cadConfig.lookbackInDays.past_range_period =
      (parseFloat(cadConfig.lookbackInUnits.past_range_period) * pastRangeEndedUnit);

    cadConfig.lookbackDates.current_range_start_date =
      toStringFormatter.getDateStringForMinusDays(
        cadConfig.lookbackInDays.current_range_period - 1 +
        cadConfig.lookbackInDays.current_range_day);

    cadConfig.lookbackDates.current_range_end_date =
      toStringFormatter.getDateStringForMinusDays(
        cadConfig.lookbackInDays.current_range_day);

    cadConfig.lookbackDates.past_range_start_date =
      toStringFormatter.getDateStringForMinusDays(
        cadConfig.lookbackInDays.past_range_period - 1 +
        cadConfig.lookbackInDays.past_range_day);

    cadConfig.lookbackDates.past_range_end_date =
      toStringFormatter.getDateStringForMinusDays(
        cadConfig.lookbackInDays.past_range_day);


    // get accounts config
    for (let el in cadConfig.accounts) {
      cadConfig.accounts[el] = this.toArray(sheet.getRangeByName(el).getValue());
    }

    // get campaigns config
    for (let el in cadConfig.campaigns) {
      cadConfig.campaigns[el] = sheet.getRangeByName(el).getValue();
    }

    for (let metric of this.getMonitoredMetrics()) {
      cadConfig.thresholds[`${metric}_high`] =
        parseFloat(sheet.getRangeByName(`${metric}_high`).getValue());
      cadConfig.thresholds[`${metric}_low`] =
        parseFloat(sheet.getRangeByName(`${metric}_low`).getValue());
      cadConfig.thresholds[`${metric}_ignore`] =
        parseFloat(sheet.getRangeByName(`${metric}_ignore`).getValue());
    }

    return cadConfig;
  }


  /**
   * Write meta data to the sheet
   */
  writeMetaData(cadConfig) {
    this.my_shpreadsheet.getRangeByName(CONSTS.RESULTS_CURRENT_RNAGE_DATES)
      .setValue(
        `${cadConfig.lookbackDates.current_range_start_date.sheet_date} - ${cadConfig.lookbackDates.current_range_end_date.sheet_date} (${cadConfig.lookbackInDays.current_range_period} days)`);

    this.my_shpreadsheet.getRangeByName(CONSTS.RESULTS_PAST_RNAGE_DATES)
      .setValue(
        `${cadConfig.lookbackDates.past_range_start_date.sheet_date} - ${cadConfig.lookbackDates.past_range_end_date.sheet_date} (${cadConfig.lookbackInDays.past_range_period} days)`);
  }

  /**
   * write CAD results to results sheet
   * @param {!CadResults} cadResults CAD results
   */
  writeResults(cadResults) {
    const newStartingRow =
      Math.max(CONSTS.FIRST_DATA_ROW, this.resultsSheet.getLastRow() + 1);

    let newRows = [];
    cadResults.forEach(function (row) {
      newRows = newRows.concat(row.toSheetFormat());
    });


    console.log(`newRows = ${JSON.stringify(newRows)}`)
    if (newRows.length) {
      const metricListLength = Object.keys(this.getMonitoredMetrics()).length;
      this.resultsSheet
        .getRange(
          newStartingRow, CONSTS.FIRST_DATA_COLUMN, newRows.length,
          (5 + 4 * metricListLength))
        .setValues(newRows);
    }
  }


  /**
   * delete all results from the results-sheet
   */
  clearResults() {
    const lastRow = this.resultsSheet.getLastRow();
    const howMany = lastRow - CONSTS.FIRST_DATA_ROW + 1;
    if (howMany > 0) {
      this.resultsSheet.deleteRows(CONSTS.FIRST_DATA_ROW, howMany);
    }
    let maxRows = this.resultsSheet.getMaxRows();
    if (maxRows < 350) {
      this.resultsSheet.insertRows(maxRows, 1000);
    }
  }


  /**
   * Sets date ranges header in results sheet
   *
   * @param {!Object} currentRangeStartDate range start date.
   * @param {!Object} currentRangeEndDate range end date.
   * @param {!Object} pastRangeStartDate range start date.
   * @param {!Object} pastRangeEndDate range end date.
   *
   */
  setDatesInResultsSheet(
    currentRangeStartDate, currentRangeEndDate, pastRangeStartDate,
    pastRangeEndDate) {
    this.MY_SPREADSHEET.getRangeByName(CONSTS.CURRENT_RANGE_DATES)
      .setValue(`${currentRangeStartDate.sheet_date} - ${currentRangeEndDate.sheet_date} (${this.current_range_period} * ${this.getValueForNamedRange(CONSTS.LENGTH_UNIT)} )`);
    this.MY_SPREADSHEET.getRangeByName(CONSTS.PAST_RANGE_DATES)
      .setValue(`${pastRangeStartDate.sheet_date} - ${pastRangeEndDate.sheet_date} (${this.past_range_period} * ${this.getValueForNamedRange(CONSTS.LENGTH_UNIT)} )`);
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

// export default

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
  getAccountsForIds(ids) {
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
    let accounts = {};

    if (CONFIG.is_debug_log) {
      console.log('getAccountsToTraverse= ' + JSON.stringify(cadConfig));
    }

    if (cadConfig.campaigns.campaign_ids != '') {
      Logger.log(JSON.stringify(cadConfig.campaigns.campaign_ids));
      return this.getAllSubAccounts();
    }
    if (cadConfig.accounts.account_ids.length > 0) {
      if (cadConfig.accounts.account_ids[0].toUpperCase() == 'ALL') {
        return this.getAllSubAccounts();
      } else {
        Object.assign(accounts, this.getAccountsForIds(cadConfig.accounts.account_ids));
      }
    }
    if (cadConfig.accounts.account_labels.length > 0 ||
      cadConfig.campaigns.campaign_labels.length > 0) {
      return this.getAllSubAccounts();
    }
    if (cadConfig.campaigns.all_campaigns_for_accounts != '') {
      const all_campaigns_for_accounts =
        this.getAccountsForIds(SheetUtils.getInstance().toArray(
          cadConfig.campaigns.all_campaigns_for_accounts));
      Object.assign(accounts, all_campaigns_for_accounts);
    }
    return accounts;
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
    const campaignLabels = cadConfig.campaigns.campaignLabels;
    const excludedCampaignIds = cadConfig.campaigns.excludedCampaignIds;
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
        `SELECT campaign.id, campaign.name, label.name FROM campaign_label WHERE label.name IN ( ${campaignLabels}) ${excludedClause}`;

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
 * @return {!Object} CAD monitoring results
 */
function getResultsForAllRelevantEntitiesUnderMCC(cadConfig) {
  let mccAccount = AdsApp.currentAccount();
  cadConfig.mcc = mccAccount.getCustomerId();
  const gAdsAccountSelector = new GoogleAdsAccountSelector(mccAccount);

  let allEntitiesAllMetricsComparisons = [];

  // Get accounts to traverse
  const accountsToTraverse =
    gAdsAccountSelector.getAccountsToTraverse(cadConfig);


  for (const currentAccount of Object.keys(accountsToTraverse)) {
    console.log(`====Switching to account ${currentAccount}`);
  }

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
 * @return {!Object} CAD monitoring results
 */
function getResultsForRelevantEntitiesUnderAccount(gAdsAccountSelector, currentAccount, cadConfig) {
  let cadResults = [];
  cadResults = cadResults.concat(
    aggAccountReportToCadResults(gAdsAccountSelector, currentAccount, cadConfig));

  console.log(`1054 cadResults==== ${JSON.stringify(cadResults)}`);

  const gAdsCampaignSelector = new GoogleAdsCampaignSelector();
  const relevantCampaignMap =
    gAdsCampaignSelector.reportToRelevantCampaignMap(
      currentAccount, cadConfig);
  cadResults = cadResults.concat(campaignReportToCadResults(relevantCampaignMap, cadConfig));

  console.log(`1062 cadResults==== ${JSON.stringify(cadResults)}`);

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
  let baseQuery = `SELECT customer.descriptive_name, ${getGadsCustomerQueryFields} FROM customer WHERE segments.date BETWEEN`;

  let currentQuery = `${baseQuery} "${cadConfig.lookbackDates.current_range_start_date.query_date}" AND "${cadConfig.lookbackDates.current_range_end_date.query_date}"`;
  let pastQuery = `${baseQuery} "${cadConfig.lookbackDates.past_range_start_date.query_date}" AND "${cadConfig.lookbackDates.past_range_end_date.query_date}"`;

  let currentStats = storeReportByEntityId(
    AdsApp.report(currentQuery, CONFIG.reporting_options));

  let pastStats =
    storeReportByEntityId(AdsApp.report(pastQuery, CONFIG.reporting_options));


  if (CONFIG.is_debug_log) {
    Logger.log(
      'currentQuery= ' +
      JSON.stringify(currentQuery));
    Logger.log(
      'pastQuery= ' +
      JSON.stringify(pastQuery));
    Logger.log(
      'currentStats= ' +
      JSON.stringify(currentStats));
    Logger.log(
      'pastStats= ' +
      JSON.stringify(pastStats));
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
          'aggAccountReportToCadResults added= ' + JSON.stringify(cadSingleResult));
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
 * @return {!Object} CAD monitoring results
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
  const specificCampaignsQuery =
    `SELECT customer.descriptive_name, campaign.id, campaign.name, ${getGadsQueryFields()} FROM campaign WHERE campaign.id IN (${campaignIds}) AND segments.date BETWEEN`;


  let currentQuery = `${specificCampaignsQuery} "${cadConfig.lookbackDates.current_range_start_date.query_date}" AND "${cadConfig.lookbackDates.current_range_end_date.query_date}"`;
  let pastQuery = `${specificCampaignsQuery} "${cadConfig.lookbackDates.past_range_start_date.query_date}" AND "${cadConfig.lookbackDates.past_range_end_date.query_date}"`;


  let pastStats =
    storeReportByEntityId(AdsApp.report(pastQuery, CONFIG.reporting_options));
  let currentStats = storeReportByEntityId(
    AdsApp.report(currentQuery, CONFIG.reporting_options));



  if (CONFIG.is_debug_log) {
    Logger.log(
      'currentQuery campaigns= ' +
      JSON.stringify(currentQuery));
    Logger.log(
      'pastQuery campaigns= ' +
      JSON.stringify(pastQuery));
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
 * @return {!Object} metric stats map.
 */
function storeReportByEntityId(searchResults) {
  let statsMap = {};
  for (const row of searchResults.rows()) {

    let id = (row["campaign.id"]) ? row["campaign.id"] :
      AdsApp.currentAccount().getCustomerId();

    //Logger.log("storeReportByEntityId for id = " + id + " row=" + JSON.stringify(row));
    statsMap[id] = row;
  }
  return statsMap;
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
      htmlBody: `New period:
          ${cadConfig.lookbackDates.current_range_start_date.sheet_date} - ${cadConfig.lookbackDates.current_range_end_date.sheet_date} (${cadConfig.lookbackInDays.current_range_period} days)<br>
          Past period:
          ${cadConfig.lookbackDates.past_range_start_date.sheet_date} - ${cadConfig.lookbackDates.past_range_end_date.sheet_date} (${cadConfig.lookbackInDays.past_range_period} days)
          <br><br>
          ${bodyText}
          <br><br>
          Log into Google Ads and take a look:
          notifications dashboard:  ${CONFIG.spreadsheet_url}`
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
  getDateStringForMinusDays(numDays) {
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
    // digits = digits || 1;
    return String(num.toFixed(digits));
  }


  /**
   * Produces a string representation for a number with a thousands comma.
   *
   * @param {number} num the number to format.
   * @return {string} A string representation with a thousands comma.
   */
  addNumberCommas(num) {
    let str = num.toString().split('.');
    str[0] = str[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return str.join('.');
  }


  /**
   * Parse the string into a float or zero.
   *
   * @param {string} strValue A string representing a float.
   * @return {number} The numeric value for that float.
   */
  floatOrZero(strValue) {
    if (!strValue) return 0;
    return parseFloat(strValue);
  }

  /**
   * Parse the string into a float or null.
   *
   * @param {string} strValue A string representing a float.
   * @return {number} The numeric value for that float.
   */
  floatOrUndefined(strValue) {
    if (strValue === 'No alert' || strValue === '') return undefined;
    return parseFloat(strValue);
  }
}