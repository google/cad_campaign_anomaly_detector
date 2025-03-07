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

const LOG_PAST_DEVIDER = "==========past===========";

const weekdays = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];
const CONSTS = {
  ALL: "ALL",
};

const NamedRanges = {
  RESULTS_SHEET_NAME: "results",
  EMAILS: "emails",
  MCC_ID: "MCC_ID",
  SHOW_ONLY_ANOMALIES: "SHOW_ONLY_ANOMALIES",

  CURRENT_PERIOD_UNIT: "current_period_unit",
  CURRENT_END_UNIT: "current_end_unit",

  PAST_END_UNIT: "past_end_unit",

  AVG_TYPE: "AVG_TYPE",
  AVG_TYPE_ERROR: "AVG_TYPE_ERROR",
  RESULTS_CURRENT_RANGE_DATES: "results_current_range_dates",
  RESULTS_PAST_RANGE_DATES: "results_past_range_dates",
  FIRST_DATA_ROW: 7,
  FIRST_DATA_COLUMN: 1,
  ENTITIY_IDS: "entity_ids",
  ENTITY_LABELS: "entity_labels",
  ENTITY_EXCLUDED_IDS: "entity_excluded_ids",
  ALL_CAMPAIGNS_FOR_ACCOUNTS: "all_campaigns_for_accounts",

  DATA_AGGREGATION: "data_aggregation",
  SPLIT_BY_NETWORK: "SPLIT_BY_NETWORK",
};

const AVG_TYPE = {
  AVG_TYPE_DAILY_WEEKDAYS:
    "Daily Avg. Same Weekday as today ---> (Today) vs. (a few instances from past weeks)----> both: midnight till last data hour.",
  AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
    "Daily Avg. (Today) vs. (Yesterday)  -----> both: midnight till last data hour.",
  AVG_TYPE_HOURLY_TODAY:
    "Hourly Avg. Inside Today ------> (Last data hour) vs. (midnight till that hour)",
  AVG_TYPE_DAILY: "Daily Avg. Full days.",
  AVG_TYPE_WEEKLY: "Weekly (last 7 days) Avg",
  AVG_TYPE_CUSTOM: "Custom",
};

const TimeFrameUnits = {
  Days: 1,
  Weeks: 7,
};

const EntityType = {
  Account: "account",
  Campaign: "campaign",
  AdGroup: "ad_group",
};

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
    email: "Impressions",
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },
  clicks: {
    email: "Clicks",
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },
  all_conversions: {
    email: "All Conversions",
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },
  conversions: {
    email: "conversions",
    isMonitored: true,
    shouldBeRounded: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },
  cost_micros: {
    email: "Cost",
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },
  cost_per_all_conversions: {
    email: "CPA all",
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },
  cost_per_conversion: {
    email: "CPA",
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },
  ctr: {
    email: "CTR",
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
    isPercent: true,
  },
  all_conversions_from_interactions_rate: {
    email: "CVR all",
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
  },
  conversions_from_interactions_rate: {
    email: "CVR",
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
  },

  average_cpc: {
    email: "CPC",
    isMonitored: true,
    isMicro: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
  },

  roas_all: {
    email: "ROAS all",
    isMonitored: true,
    shouldBeRounded: false,
    isWriteToSheet: true,
  },
  roas: {
    email: "ROAS",
    isMonitored: true,
    shouldBeRounded: false,
    isWriteToSheet: true,
  },

  all_conversions_value: {
    email: "All conversions value",
    isMonitored: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },

  conversions_value: {
    email: "Conversions value",
    isMonitored: true,
    isFromGoogleAds: true,
    isCumulative: true,
    isWriteToSheet: true,
  },

  search_click_share: {
    email: "Search click share",
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
  },
  average_cpm: {
    email: "Avg CPM",
    isMicro: true,
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
  },

  average_cpv: {
    email: "Avg CPV",
    isMicro: true,
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
  },

  video_view_rate: {
    email: "Video View Rate",
    isMonitored: true,
    isFromGoogleAds: true,
    isWriteToSheet: true,
  },
};

/**
 * Comparison results of one entity to monitor
 */
class CadResultForEntity {
  constructor() {
    this.relevant_label = undefined;
    this.account = { id: undefined, name: undefined };
    this.campaign = { id: undefined, name: undefined };
    this.adGroup = { id: undefined, name: undefined };

    this.isTriggerAlert = false;
    this.allMetricsComparisons = {};

    let currencyCode = AdsApp.currentAccount().getCurrencyCode();
    currencyCode = currencyCode === "USD" ? "$" : currencyCode;
    this.currencyCode = currencyCode + " ";
  }


  /**
   * Fills the metrics map with comparison results.
   *
   * @param {string} id Entity id
   * @param {!Object} pastStats Past stats map
   * @param {!Object} currentStats Current stats map
   * @param {!CadConfig} cadConfig CAD config
   */
  fillMetricComparisonResults(id, currentStats, pastStats, cadConfig) {
    let monitoredMetrics = this.getMonitoredMetrics();

    for (let metric of monitoredMetrics) {
      let gAdsMetric = `metrics.${metric}`;
      this.allMetricsComparisons[gAdsMetric] = new MetricResult();
      let comparisonResultForSpecificMetric = this.allMetricsComparisons[gAdsMetric];

      this.handleNumericMetrics(metric, id, currentStats, pastStats);
      const isSelfCalc = this.handleSelfCalcMetrics(metric, id, comparisonResultForSpecificMetric, currentStats, pastStats, cadConfig);
      if (!isSelfCalc) {
        if (MetricTypes[metric].isCumulative) {
          this.handleRegularCumulativeMetrics(metric, id, comparisonResultForSpecificMetric, currentStats, pastStats, cadConfig);
        }
        else {
          this.handleNonCumulativeMetrics(metric, id, comparisonResultForSpecificMetric, currentStats, pastStats, cadConfig);
        }
      }
      this.fillDirections(metric, comparisonResultForSpecificMetric, cadConfig);
      this.updateTriggerAlert(comparisonResultForSpecificMetric);


      Logger.log(`this.allMetricsComparisons[${gAdsMetric}].past == ${this.allMetricsComparisons[gAdsMetric].past}`);
      Logger.log(`this.allMetricsComparisons[${gAdsMetric}].current == ${this.allMetricsComparisons[gAdsMetric].current}`);

    }
  }

  /**
     * Gets an array of monitored metric keys.
     *
     * @returns {Array<string>} Monitored metric keys
     */
  getMonitoredMetrics() {
    let monitoredMetrics = [];
    Object.keys(MetricTypes).forEach(function (key, index) {
      if (MetricTypes[key].isMonitored) {
        monitoredMetrics.push(key);
      }
    });
    return monitoredMetrics;
  }

  /**
   * Handles numeric metrics adjustments.
   *
   * @param {string} metric Metric key
   * @param {string} id Entity id
   * @param {!Object} currentStats Current stats map
   * @param {!Object} pastStats Past stats map
   */
  handleNumericMetrics(metric, id, currentStats, pastStats) {
    let gAdsMetric = `metrics.${metric}`;
    currentStats[id] = currentStats[id] || {};
    pastStats[id] = pastStats[id] || {};

    if (MetricTypes[metric].isMicro) {
      currentStats[id][gAdsMetric] /= 1e6;
      pastStats[id][gAdsMetric] = pastStats[id][gAdsMetric] ? pastStats[id][gAdsMetric] / 1e6 : 0;
    }

    if (MetricTypes[metric].isPercent) {
      currentStats[id][gAdsMetric] *= 100;
      pastStats[id][gAdsMetric] = pastStats[id][gAdsMetric] ? pastStats[id][gAdsMetric] * 100 : 0;
    }
  }

  /**
   * Handles regular cumulative metrics adjustments.
   *
   * @param {string} metric Metric key
   * @param {string} id Entity id
   * @param {!MetricResult} comparisonResultForSpecificMetric Comparison result object
   * @param {!Object} currentStats Current stats map
   * @param {!Object} pastStats Past stats map
   * @param {!CadConfig} cadConfig CAD config
   */
  handleRegularCumulativeMetrics(metric, id, comparisonResultForSpecificMetric, currentStats, pastStats, cadConfig) {
    let gAdsMetric = `metrics.${metric}`;
    currentStats[id] = currentStats[id] || {};
    pastStats[id] = pastStats[id] || {};

    comparisonResultForSpecificMetric.past = pastStats[id][gAdsMetric] ? pastStats[id][gAdsMetric] / cadConfig.dividePastBy : 0;
    comparisonResultForSpecificMetric.current = currentStats[id][gAdsMetric] / cadConfig.divideCurrentBy;

  }

  /**
   * Handles special metrics adjustments like ROAS.
   *
   * @param {string} metric Metric key
   * @param {string} id Entity id
   * @param {!MetricResult} comparisonResultForSpecificMetric Comparison result object
   * @param {!Object} currentStats Current stats map
   * @param {!Object} pastStats Past stats map
   */
  handleSelfCalcMetrics(metric, id, comparisonResultForSpecificMetric, currentStats, pastStats) {
    let gAdsMetric = `metrics.${metric}`;
    currentStats[id] = currentStats[id] || {};
    pastStats[id] = pastStats[id] || {};

    let isSpecialMetric = false;
    if (gAdsMetric.includes("roas_all")) {
      comparisonResultForSpecificMetric.past = pastStats[id][gAdsMetric] ? this.safeDivide(pastStats[id], "metrics.all_conversions_value", "metrics.cost_micros") : 0;
      comparisonResultForSpecificMetric.current = this.safeDivide(currentStats[id], "metrics.all_conversions_value", "metrics.cost_micros");
      isSpecialMetric = true;
    } else if (gAdsMetric.includes("roas")) {
      comparisonResultForSpecificMetric.past = pastStats[id][gAdsMetric] ? this.safeDivide(pastStats[id], "metrics.conversions_value", "metrics.cost_micros") : 0;
      comparisonResultForSpecificMetric.current = this.safeDivide(currentStats[id], "metrics.conversions_value", "metrics.cost_micros");
      isSpecialMetric = true;
    }
    return isSpecialMetric;
  }

  /**
   * Handles non-cumulative metrics adjustments.
   *
   * @param {string} metric Metric key
   * @param {string} id Entity id
   * @param {!MetricResult} comparisonResultForSpecificMetric Comparison result object
   * @param {!Object} currentStats Current stats map
   * @param {!Object} pastStats Past stats map
   * @param {!CadConfig} cadConfig CAD config
   */
  handleNonCumulativeMetrics(metric, id, comparisonResultForSpecificMetric, currentStats, pastStats, cadConfig) {
    let gAdsMetric = `metrics.${metric}`;
    currentStats[id] = currentStats[id] || {};
    pastStats[id] = pastStats[id] || {};

    comparisonResultForSpecificMetric.past = pastStats[id][gAdsMetric] || 0;
    comparisonResultForSpecificMetric.current = currentStats[id][gAdsMetric] || 0;
  }

  /**
   * Fills the direction fields of the comparison result.
   *
   * @param {string} metric Metric key
   * @param {!MetricResult} comparisonResult Comparison result object
   * @param {!CadConfig} cadConfig CAD config
   */
  fillDirections(metric, comparisonResult, cadConfig) {
    const past = comparisonResult.past;
    const current = comparisonResult.current;
    comparisonResult.changeAbs = current - past;
    comparisonResult.changePercent = this.calculatePercentageChange(past, current);

    comparisonResult.isAboveHigh = cadConfig.thresholds[`${metric}_high`] > 0 && comparisonResult.changePercent >= cadConfig.thresholds[`${metric}_high`] * 100;
    comparisonResult.isBelowLow = cadConfig.thresholds[`${metric}_low`] < 0 && comparisonResult.changePercent <= cadConfig.thresholds[`${metric}_low`] * 100;

    let ignoreAbs = cadConfig.thresholds[`${metric}_ignore`];
    if (!ignoreAbs || past >= ignoreAbs) {
      if (comparisonResult.isAboveHigh) {
        comparisonResult.metricAlertDirection = "up";
      } else if (comparisonResult.isBelowLow) {
        comparisonResult.metricAlertDirection = "down";
      } else if (cadConfig.showOnlyAnomalies) {
        comparisonResult.changeAbs = " - ";
        comparisonResult.changePercent = " - ";
      }
    }
  }

  /**
   * Updates the trigger alert flag based on comparison result.
   *
   * @param {!MetricResult} comparisonResult Comparison result object
   */
  updateTriggerAlert(comparisonResult) {
    this.isTriggerAlert = this.isTriggerAlert || comparisonResult.metricAlertDirection !== undefined;
  }



  safeDivide(mapName, nomenator, denomenator) {
    if (mapName[denomenator] === 0) return 0;
    return mapName[nomenator] / mapName[denomenator];
  }

  calculatePercentageChange(past, current) {
    if (past === 0 && current === 0) {
      return 0;
    } else if (past === 0) {
      return 100;
    } else if (current === 0) {
      return -100;
    } else {
      return (current / past) * 100 - 100;
    }
  }

  /**
   * Truncate decimal digits
   */
  truncateDecimalDigits() {
    for (let metricType in this.allMetricsComparisons) {
      let metricTypeName = metricType.split(".")[1];

      let currentAvgValue = this.allMetricsComparisons[metricType].current;
      let pastAvgValue = this.allMetricsComparisons[metricType].past;
      let numericDelta = this.allMetricsComparisons[metricType].changeAbs;
      let percentageDelta =
        this.allMetricsComparisons[metricType].changePercent;
      let metricAlertDirection =
        this.allMetricsComparisons[metricType].metricAlertDirection;

      let metricStrings = new MetricStrings();

      if (MetricTypes[metricTypeName].shouldBeRounded) {
        metricStrings = this.toRoundedMetricString(
          currentAvgValue,
          pastAvgValue,
          numericDelta,
          percentageDelta
        );
      } else if (
        metricType.includes("cost_per_conversion") ||
        metricType.includes("cost_micros")
      ) {
        metricStrings = this.toMetricStrings(
          currentAvgValue,
          pastAvgValue,
          numericDelta,
          percentageDelta,
          this.currencyCode + " ",
          "",
          2
        );
      } else if (metricType.includes("ctr") || metricType.includes("roas")) {
        metricStrings = this.toMetricStrings(
          currentAvgValue,
          pastAvgValue,
          numericDelta,
          percentageDelta,
          "",
          "%",
          1
        );
      } else {
        metricStrings = this.toMetricStrings(
          currentAvgValue,
          pastAvgValue,
          numericDelta,
          percentageDelta,
          "",
          "",
          1
        );
      }

      if (metricAlertDirection) {
        metricStrings.percentageDeltaStr = metricAlertDirection.includes("up")
          ? `⇪⇪ ${metricStrings.percentageDeltaStr}`
          : `⇩⇩ ${metricStrings.percentageDeltaStr}`;
      }

      this.allMetricsComparisons[metricType].current =
        metricStrings.currentValueStr;
      this.allMetricsComparisons[metricType].past =
        metricStrings.pastAvgValueStr;
      this.allMetricsComparisons[metricType].changeAbs =
        metricStrings.numericDeltaStr;
      this.allMetricsComparisons[metricType].changePercent =
        metricStrings.percentageDeltaStr;
    }
  }

  /**
   * Returns a string for an email format
   *
   * @return {string} an email string format
   */
  toEmailFormat() {
    if (!this.isTriggerAlert) return "";

    let adGtoupHeader = this.adGroup.id
      ? `: ${this.adGroup.id} ${this.adGroup.name}`
      : "";

    let campaignHeader = this.campaign.id
      ? `: ${this.campaign.id} ${this.campaign.name}`
      : "";

    let alertTextForEntity = [
      `<br>Anomalies for: ${this.account.id} ${this.account.name} ${campaignHeader} ${adGtoupHeader}
      <br>(Only relevant metrics. For all metrics see the end of the email)
     <br><table style="width:50%;border:1px solid black;"><tr style="border:1px solid black;"><th style="text-align:left;border:1px solid black;">Metric</th><th style="text-align:left;border:1px solid black;">Current</th><th style="text-align:left;border:1px solid black;">Past</th> <th style="text-align:left;border:1px solid black;">Δ</th> <th style="text-align:left;border:1px solid black;">Δ%</th></tr>`,
    ];

    for (let metricType in this.allMetricsComparisons) {
      let metricType_name = metricType.split(".")[1];
      if (!MetricTypes[metricType_name].isWriteToSheet) continue;
      if (!this.allMetricsComparisons[metricType].metricAlertDirection)
        continue;

      const toStringFormatter = ToStringFormatter.getInstance();
      const currentValueStr = toStringFormatter.addNumberCommas(
        this.allMetricsComparisons[metricType].current
      );
      const pastAvgValueStr = toStringFormatter.addNumberCommas(
        this.allMetricsComparisons[metricType].past
      );
      const numericDeltaStr = toStringFormatter.addNumberCommas(
        this.allMetricsComparisons[metricType].changeAbs
      );
      const percentageDeltaStr = toStringFormatter.addNumberCommas(
        this.allMetricsComparisons[metricType].changePercent
      );

      alertTextForEntity.push(`<tr><td> ${metricType} </td><td> ${currentValueStr} </td><td> ${pastAvgValueStr} </td><td>
${numericDeltaStr} </td><td style="color: ${this.getDeltaColorPercentage(
        percentageDeltaStr
      )};"> ${percentageDeltaStr} </td></tr>`);
    }
    return `${alertTextForEntity.join("")} </table>`;
  }

  // Function to get the color based on positive/negative percentage values
  getDeltaColorPercentage(percentage) {
    if (percentage.includes("⇪")) {
      return "green";
    } else if (percentage.includes("⇩")) {
      return "red"; // If the value is zero
    } else {
      return "white";
    }
  }

  /**
   * @param {number} currentAvgValue current value (avg)
   * @param {number} pastAvgValue  past value (avg)
   * @param {number} numericDelta  numeric delta
   *@param {number} percentageDelta percentage delta
   *@return {!Object} a map of string formats
   */
  toRoundedMetricString(
    currentAvgValue,
    pastAvgValue,
    numericDelta,
    percentageDelta
  ) {
    return new MetricStrings(
      Math.round(currentAvgValue),
      Math.round(pastAvgValue),
      Math.round(numericDelta),
      Math.round(percentageDelta) + "%"
    );
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
    currentAvgValue,
    pastAvgValue,
    numericDelta,
    percentageDelta,
    prefix,
    postfix,
    decimalDigits
  ) {
    // Assign zero to arguments which are null or undefined
    currentAvgValue =
      currentAvgValue !== null && currentAvgValue !== undefined
        ? currentAvgValue
        : 0;
    pastAvgValue =
      pastAvgValue !== null && pastAvgValue !== undefined ? pastAvgValue : 0;
    numericDelta =
      numericDelta !== null && numericDelta !== undefined ? numericDelta : 0;
    percentageDelta =
      percentageDelta !== null && percentageDelta !== undefined
        ? percentageDelta
        : 0;
    decimalDigits =
      decimalDigits !== null && decimalDigits !== undefined ? decimalDigits : 0;

    const toStringFormatter = ToStringFormatter.getInstance();
    return {
      currentValueStr:
        prefix +
        toStringFormatter.formatNumWithDigitAfterPeriod(
          currentAvgValue,
          decimalDigits
        ) +
        postfix,
      pastAvgValueStr:
        prefix +
        toStringFormatter.formatNumWithDigitAfterPeriod(
          pastAvgValue,
          decimalDigits
        ) +
        postfix,
      numericDeltaStr:
        prefix +
        toStringFormatter.formatNumWithDigitAfterPeriod(
          numericDelta,
          decimalDigits
        ) +
        postfix,
      percentageDeltaStr:
        toStringFormatter.formatNumWithDigitAfterPeriod(
          percentageDelta,
          decimalDigits
        ) + "%",
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
      this.adGroup.id,
      this.adGroup.name,
    ];

    console.log(
      `this.metricComparisonResults = ${JSON.stringify(
        this.allMetricsComparisons
      )}`
    );
    for (let metricType in this.allMetricsComparisons) {
      let metricName = metricType.split(".")[1];
      if (!MetricTypes[metricName].isMonitored) continue;

      rowData = rowData.concat([
        this.allMetricsComparisons[metricType].current,
        this.allMetricsComparisons[metricType].past,
        this.allMetricsComparisons[metricType].changeAbs,
        this.allMetricsComparisons[metricType].changePercent,
      ]);
    }
    return [rowData];
  }
}

/**
 * Metric strings
 */
class MetricStrings {
  constructor(
    currentValueStr = "",
    pastAvgValueStr = "",
    numericDeltaStr = "",
    percentageDeltaStr = ""
  ) {
    this.currentValueStr = currentValueStr;
    this.pastAvgValueStr = pastAvgValueStr;
    this.numericDeltaStr = numericDeltaStr;
    this.percentageDeltaStr = percentageDeltaStr;
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
      past_ended_length_ago: undefined,
    };

    //after code calculation:
    this.lookbackInDays = {
      current_period_length: {},
      current_period_length_text: {},
      current_ended_length_ago: {},

      past_period_length: {},
      past_period_length_text: {},
      past_ended_length_ago: {},
    };

    this.lookbackDates = {
      current_range_start_date: {},
      current_range_end_date: {},
      past_range_start_date: {},
      past_range_end_date: {},
    };
    this.accounts = {
      ids: [],
      excluded_ids: [],
      labels: [],
    };
    this.campaigns = {
      ids: "",
      excluded_ids: "",
      labels: "",
      all_under_parents: [],
    };
    this.adGroup = {
      ids: "",
      excluded_ids: "",
      labels: "",
      all_under_parents: [],
    };

    this.splitByNetworkQuery = ""

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
    const lastQueryableDate = new Date(
      NOW.getTime() - (this.HOURS_BACK + minusHours) * 3600 * 1000
    );
    const formattedDate = Utilities.formatDate(
      lastQueryableDate,
      this.TIMEZONE,
      "yyyy-MM-dd HH:mm:ss"
    );
    const pastHourStr = formattedDate.split(" ")[1].split(":")[0];

    if (CONFIG.is_debug_log) {
      Logger.log("getLastQueryableHourMinusHours(minusHours)= " + minusHours);
      Logger.log("formattedDate = " + formattedDate);
    }


    return {
      hourInt: parseInt(pastHourStr, 10),
      query_date: ToStringFormatter.getInstance().getDateStringInTimeZone(
        lastQueryableDate,
        "yyyy-MM-dd"
      ),
      sheet_date: ToStringFormatter.getInstance().getDateStringInTimeZone(
        lastQueryableDate,
        "dd/MM/YY"
      ),
      weekday: `AND segments.day_of_week = '${weekdays[
        lastQueryableDate.getUTCDay()
      ].toUpperCase()}'`,
      hourWhereClauseEqual: `AND segments.hour = ${pastHourStr}`,
      hourWhereClauseSmaller: `AND segments.hour < ${pastHourStr}`,
    };
  }
}

/**
 * Input sheet representation
 */
class SheetUtils {
  constructor() {
    let tmpSpreadsheet = SpreadsheetApp.openByUrl(CONFIG.spreadsheet_url);
    this.mySpreadsheet = tmpSpreadsheet;
    this.resultsSheet = tmpSpreadsheet.getSheetByName(
      NamedRanges.RESULTS_SHEET_NAME
    );
    this.monitoredMetrics = [];
  }

  getMonitoredMetrics() {
    const monitoredMetrics = [];
    if (this.monitoredMetrics && this.monitoredMetrics.length > 0)
      return this.monitoredMetrics;

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
   * Read input into CADConfig object
   * @return {!CadConfig} a filled cad config object.
   */
  readInput() {
    const mySpreadsheet = this.mySpreadsheet;
    let cadConfig = new CadConfig();
    cadConfig.users = mySpreadsheet
      .getRangeByName(NamedRanges.EMAILS)
      .getValue();
    cadConfig.avgType = mySpreadsheet
      .getRangeByName(NamedRanges.AVG_TYPE)
      .getValue();
    cadConfig.showOnlyAnomalies = mySpreadsheet
      .getRangeByName(NamedRanges.SHOW_ONLY_ANOMALIES)
      .getValue();
    cadConfig.splitByNetworkQuery = mySpreadsheet
      .getRangeByName(NamedRanges.SPLIT_BY_NETWORK)
      .getValue()? ", segments.ad_network_type" : "";


    const lastQueryableHour =
      TimeUtils.getInstance().getLastQueryableHourMinusHours(0);
    cadConfig = this.fillLookbackDatesForTextNotForDivision(cadConfig);

    console.log(`cadConfig.avgType ==== ${JSON.stringify(cadConfig.avgType)}`);

    cadConfig = this.addSegmentsHourToQuery(cadConfig, lastQueryableHour);
    cadConfig = this.setDivideBy(cadConfig, lastQueryableHour);

    const entity_ids = mySpreadsheet
      .getRangeByName(NamedRanges.ENTITIY_IDS)
      .getValue();
    const entity_labels = mySpreadsheet
      .getRangeByName(NamedRanges.ENTITY_LABELS)
      .getValue();
    const entity_excluded_ids = mySpreadsheet
      .getRangeByName(NamedRanges.ENTITY_EXCLUDED_IDS)
      .getValue();
    const data_aggregation = mySpreadsheet
      .getRangeByName(NamedRanges.DATA_AGGREGATION)
      .getValue();

    switch (data_aggregation) {
      case "Account":
        cadConfig.accounts.ids = SheetUtils.getInstance().toArray(entity_ids);
        cadConfig.accounts.labels =
          SheetUtils.getInstance().toArray(entity_labels);
        cadConfig.accounts.excluded_ids =
          SheetUtils.getInstance().toArray(entity_excluded_ids);
        break;

      case "Campaign":
        cadConfig.campaigns.ids = this.addQuotesIfNeeded(entity_ids);
        cadConfig.campaigns.labels = this.addQuotesIfNeeded(entity_labels);
        cadConfig.campaigns.excluded_ids =
          this.addQuotesIfNeeded(entity_excluded_ids);
        cadConfig.campaigns.all_under_parents =
          SheetUtils.getInstance().toArray(
            mySpreadsheet
              .getRangeByName(NamedRanges.ALL_CAMPAIGNS_FOR_ACCOUNTS)
              .getValue()
          );
        break;

      default:
      case "Ad Group":
        cadConfig.adGroup.ids = entity_ids;
        cadConfig.adGroup.labels = entity_labels;
        cadConfig.adGroup.excluded_ids = entity_excluded_ids;
        cadConfig.adGroup.all_under_parents = SheetUtils.getInstance().toArray(
          mySpreadsheet
            .getRangeByName(NamedRanges.ALL_CAMPAIGNS_FOR_ACCOUNTS)
            .getValue()
        );
        console.log(
          `Ad Group read cadConfig.adGroup==== ${JSON.stringify(
            cadConfig.adGroup
          )}`
        );
        break;
    }

    for (let metric of this.getMonitoredMetrics()) {
      cadConfig.thresholds[`${metric}_high`] = parseFloat(
        mySpreadsheet.getRangeByName(`${metric}_high`).getValue()
      );
      let thresholdValue = parseFloat(mySpreadsheet.getRangeByName(`${metric}_low`).getValue());
      cadConfig.thresholds[`${metric}_low`] = thresholdValue >= 0 ? -1 * thresholdValue : thresholdValue;

      cadConfig.thresholds[`${metric}_ignore`] = parseFloat(
        mySpreadsheet.getRangeByName(`${metric}_ignore`).getValue()
      );
    }
    return cadConfig;
  }

  addQuotesIfNeeded(inputString) {
    inputString = String(inputString);
    if (!inputString) return inputString;
    let items = inputString.split(",");
    let output = [];
    for (let i = 0; i < items.length; i++) {
      let item = items[i].trim(); // Remove leading/trailing whitespace
      if (!item.startsWith('"')) {
        item = '"' + item;
      }
      if (!item.endsWith('"')) {
        item = item + '"';
      }
      output.push(item);
    }
    return output.join(",");
  }

  addSegmentsHourToQuery(cadConfig, lastQueryableHour) {
    switch (cadConfig.avgType) {
      case AVG_TYPE.AVG_TYPE_DAILY:
      case AVG_TYPE.AVG_TYPE_WEEKLY:
      case AVG_TYPE.AVG_TYPE_CUSTOM: {
        cadConfig.hourSegmentInSelect = ``;
        cadConfig.hourSegmentsWhereClause = { current: "", past: "" };
        break;
      }
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS: {
        cadConfig.hourSegmentInSelect = `, segments.date, segments.hour, segments.day_of_week `;
        cadConfig.hourSegmentsWhereClause = {
          current: `${lastQueryableHour.hourWhereClauseSmaller} ${lastQueryableHour.weekday}`,
          past: `${lastQueryableHour.hourWhereClauseSmaller} ${lastQueryableHour.weekday}`,
        };
        break;
      }
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY: {
        cadConfig.hourSegmentInSelect = `, segments.date, segments.hour `;
        cadConfig.hourSegmentsWhereClause = {
          current: lastQueryableHour.hourWhereClauseSmaller,
          past: lastQueryableHour.hourWhereClauseSmaller,
        };
        break;
      }
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY: {
        cadConfig.hourSegmentInSelect = `,segments.date, segments.hour `;
        cadConfig.hourSegmentsWhereClause = {
          current: lastQueryableHour.hourWhereClauseEqual,
          past: lastQueryableHour.hourWhereClauseSmaller,
        };
        break;
      }
    }
    return cadConfig;
  }

  setDivideBy(cadConfig) {
    switch (cadConfig.avgType) {
      //Daily avg - sum hours, but no need to divide by #hours
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS: {
        cadConfig.divideCurrentBy = 1;
        cadConfig.dividePastBy =
          cadConfig.lookbackInUnits.past_period_length / 7;
        break;
      }
      //Daily avg - sum hours, but no need to divide by #hours
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY: {
        cadConfig.divideCurrentBy = 1;
        cadConfig.dividePastBy = 1;
        break;
      }
      //Divide by # hours - hourly avg
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY: {
        cadConfig.divideCurrentBy = 1;

        const beforeLastQueryableHour =
          TimeUtils.getInstance().getLastQueryableHourMinusHours(1);
        cadConfig.dividePastBy = beforeLastQueryableHour.hourInt;
        break;
      }

      case AVG_TYPE.AVG_TYPE_DAILY:
      case AVG_TYPE.AVG_TYPE_WEEKLY:
      case AVG_TYPE.AVG_TYPE_CUSTOM: {
        cadConfig.divideCurrentBy =
          cadConfig.lookbackInUnits.current_period_length;
        cadConfig.dividePastBy = cadConfig.lookbackInUnits.past_period_length;
        break;
      }
    }
    return cadConfig;
  }

  fillLookbackDatesForTextNotForDivision(cadConfig) {
    const mySpreadsheet = this.mySpreadsheet;
    const lastQueryableHour =
      TimeUtils.getInstance().getLastQueryableHourMinusHours(0);

    switch (cadConfig.avgType) {
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY: {
        this.setSheetAndQueryDates(
          cadConfig.lookbackDates.current_range_start_date,
          lastQueryableHour
        );
        this.setSheetAndQueryDates(
          cadConfig.lookbackDates.current_range_end_date,
          lastQueryableHour
        );
        break;
      }
    }
    const currentAndPastPeriodUnit =
      TimeFrameUnits[
      mySpreadsheet.getRangeByName(NamedRanges.CURRENT_PERIOD_UNIT).getValue()
      ] || 1;
    const currentEndUnit =
      TimeFrameUnits[
      mySpreadsheet.getRangeByName(NamedRanges.CURRENT_END_UNIT).getValue()
      ] || 1;
    const pastEndUnit =
      TimeFrameUnits[
      mySpreadsheet.getRangeByName(NamedRanges.PAST_END_UNIT).getValue()
      ] || 1;
    for (let el in cadConfig.lookbackInUnits) {
      cadConfig.lookbackInUnits[el] = mySpreadsheet
        .getRangeByName(el)
        .getValue();
    }

    const toStringFormatter = ToStringFormatter.getInstance();
    let beforeLastQueryableHour =
      TimeUtils.getInstance().getLastQueryableHourMinusHours(1);

    switch (cadConfig.avgType) {
      case AVG_TYPE.AVG_TYPE_HOURLY_TODAY: {
        this.setSheetAndQueryDates(
          cadConfig.lookbackDates.past_range_start_date,
          beforeLastQueryableHour
        );
        this.setSheetAndQueryDates(
          cadConfig.lookbackDates.past_range_end_date,
          beforeLastQueryableHour
        );
        cadConfig.lookbackInDays.current_period_length = 1;
        cadConfig.lookbackInDays.current_period_length_text =
          `${beforeLastQueryableHour.hourWhereClauseEqual} for partial 1`.replace(
            "AND ",
            ""
          );
        cadConfig.lookbackInDays.current_ended_length_ago = 0;
        cadConfig.lookbackInDays.past_period_length = parseFloat(cadConfig.lookbackInUnits.past_period_length);
        cadConfig.lookbackInDays.past_period_length_text =
          `${beforeLastQueryableHour.hourWhereClauseSmaller} for partial 1`.replace(
            "AND ",
            ""
          );
        cadConfig.lookbackInDays.past_ended_length_ago = 0;
        break;
      }

      case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY: {
        beforeLastQueryableHour =
          TimeUtils.getInstance().getLastQueryableHourMinusHours(24);
        this.setSheetAndQueryDates(
          cadConfig.lookbackDates.past_range_start_date,
          beforeLastQueryableHour
        );
        this.setSheetAndQueryDates(
          cadConfig.lookbackDates.past_range_end_date,
          beforeLastQueryableHour
        );
        cadConfig.lookbackInDays.current_period_length = 1;
        cadConfig.lookbackInDays.current_period_length_text =
          `${beforeLastQueryableHour.hourWhereClauseSmaller} for partial 1`.replace(
            "AND ",
            ""
          );
        cadConfig.lookbackInDays.current_ended_length_ago = 0;
        cadConfig.lookbackInDays.past_period_length = 1;
        cadConfig.lookbackInDays.past_period_length_text =
          `${beforeLastQueryableHour.hourWhereClauseSmaller} for partial 1`.replace(
            "AND ",
            ""
          );
        cadConfig.lookbackInDays.past_ended_length_ago = 1;
        break;
      }
      case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS: {
        cadConfig.lookbackInDays.current_period_length = 1;
        cadConfig.lookbackInDays.current_period_length_text =
          `${beforeLastQueryableHour.hourWhereClauseSmaller} for partial 1`.replace(
            "AND ",
            ""
          );
        cadConfig.lookbackInDays.current_ended_length_ago = 0;

        cadConfig.lookbackInDays.past_period_length =
          parseFloat(cadConfig.lookbackInUnits.past_period_length) *
          currentAndPastPeriodUnit;
        cadConfig.lookbackInDays.past_period_length_text =
         `${beforeLastQueryableHour.hourWhereClauseSmaller} for partial 1`.replace(
            "AND ",
            ""
          );

        cadConfig.lookbackInDays.past_ended_length_ago =
          parseFloat(cadConfig.lookbackInUnits.past_ended_length_ago) *
          pastEndUnit;

        cadConfig.lookbackDates.current_range_start_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.current_ended_length_ago -
            1 +
            cadConfig.lookbackInDays.current_period_length
          );

        cadConfig.lookbackDates.current_range_end_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.current_ended_length_ago
          );

        cadConfig.lookbackDates.past_range_start_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.past_ended_length_ago -
            1 +
            cadConfig.lookbackInDays.past_period_length
          );

        cadConfig.lookbackDates.past_range_end_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.past_ended_length_ago
          );

        break;
      }

      case AVG_TYPE.AVG_TYPE_DAILY:
      case AVG_TYPE.AVG_TYPE_WEEKLY:
      case AVG_TYPE.AVG_TYPE_CUSTOM: {
        cadConfig.lookbackInDays.current_period_length =
          parseFloat(cadConfig.lookbackInUnits.current_period_length) *
          currentAndPastPeriodUnit;
        cadConfig.lookbackInDays.current_ended_length_ago =
          parseFloat(cadConfig.lookbackInUnits.current_ended_length_ago) *
          currentEndUnit;
        cadConfig.lookbackInDays.past_period_length =
          parseFloat(cadConfig.lookbackInUnits.past_period_length) *
          currentAndPastPeriodUnit;
        cadConfig.lookbackInDays.past_ended_length_ago =
          parseFloat(cadConfig.lookbackInUnits.past_ended_length_ago) *
          pastEndUnit;

        cadConfig.lookbackDates.current_range_start_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.current_ended_length_ago +
            cadConfig.lookbackInDays.current_period_length
          );

        cadConfig.lookbackDates.current_range_end_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.current_ended_length_ago + 1
          );

        cadConfig.lookbackDates.past_range_start_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.past_ended_length_ago +
            cadConfig.lookbackInDays.past_period_length
          );

        cadConfig.lookbackDates.past_range_end_date =
          toStringFormatter.getStringForMinusDaysAgo(
            cadConfig.lookbackInDays.past_ended_length_ago + 1
          );
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
    this.mySpreadsheet
      .getRangeByName(NamedRanges.MCC_ID)
      .setValue(cadConfig.mcc);
    this.setDatesInResultsSheet(cadConfig);
  }

  /**
   * Sets date ranges header in results sheet
   */
  setDatesInResultsSheet(cadConfig) {
    this.mySpreadsheet
      .getRangeByName(NamedRanges.RESULTS_CURRENT_RANGE_DATES)
      .setValue(
        `${cadConfig.lookbackDates.current_range_start_date.sheet_date} - ${cadConfig.lookbackDates.current_range_end_date.sheet_date} (${cadConfig.lookbackInDays.current_period_length_text} days)`
      );

    this.mySpreadsheet
      .getRangeByName(NamedRanges.RESULTS_PAST_RANGE_DATES)
      .setValue(
        `${cadConfig.lookbackDates.past_range_start_date.sheet_date} - ${cadConfig.lookbackDates.past_range_end_date.sheet_date} (${cadConfig.lookbackInDays.past_period_length_text} days)`
      );
  }

  /**
   * write CAD results to results sheet
   * @param {Array<!CadResultForEntity>s} cadResults CAD results
   */
  writeResults(cadResults) {
    const newStartingRow = Math.max(
      NamedRanges.FIRST_DATA_ROW,
      this.resultsSheet.getLastRow() + 1
    );

    let newRows = [];
    cadResults.forEach(function (row) {
      newRows = newRows.concat(row.toSheetFormat());
    });
    console.log(`newRows = ${JSON.stringify(newRows)}`);
    if (newRows.length) {
      const metricListLength = Object.keys(this.getMonitoredMetrics()).length;
      this.resultsSheet
        .getRange(
          newStartingRow,
          NamedRanges.FIRST_DATA_COLUMN,
          newRows.length,
          7 + 4 * metricListLength
        )
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
    value =
      value == "" || value == undefined
        ? []
        : value.replace(/"/g, "").replace(/\s/g, "").split(",");

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
    Logger.log("MCC has " + accounts.length + " child accounts");
    return accounts;
  }

  /**
   * Returns an account list for the id list.
   *
   * @param {!Array<?>} ids Id list
   * @return {!Array<?>} account list
   */
  getAccountObjectsForIds(ids) {
    Logger.log("accountIterator=" + JSON.stringify(ids));
    let accountIterator = AdsManagerApp.accounts().withIds(ids).get();
    return this.iterate(accountIterator);
  }

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
      Logger.log("iterate(accountIterator) " + currentAccount.getCustomerId());
      accounts[currentAccount.getCustomerId()] = currentAccount;
    }
    return accounts;
  }

  /**
   * Returns the account to traverse.
   * @param {!CadConfig} cadConfig CAD config
   * @return {!Set<?>} Account set to traverse
   */
  getAccountsToTraverse(cadConfig) {
    let accountsObjects = {};

    if (CONFIG.is_debug_log) {
      console.log(
        "getAccountsToTraverse   cadConfig= " + JSON.stringify(cadConfig)
      );
    }

    //because the campaign can be anywhere - we need to scan all the child accounts anyhow.
    if (cadConfig.campaigns.ids != "" || cadConfig.adGroup.ids != "") {
      console.log(
        "getAccountsToTraverse   cadConfig.campaigns.ids= " +
        JSON.stringify(cadConfig.campaigns.ids)
      );
      console.log(
        "getAccountsToTraverse   cadConfig.adGroup.ids= " +
        JSON.stringify(cadConfig.adGroup.ids)
      );
      return this.getAllSubAccounts();
    }
    if (cadConfig.accounts.ids.length > 0) {
      if (cadConfig.accounts.ids[0].toUpperCase() == CONSTS.ALL) {
        return this.getAllSubAccounts();
      } else {
        Object.assign(
          accountsObjects,
          this.getAccountObjectsForIds(cadConfig.accounts.ids)
        );
      }
    }
    if (
      cadConfig.accounts.labels.length > 0 ||
      cadConfig.accounts.labels.length > 0 ||
      cadConfig.adGroup.labels.length > 0 ||
      cadConfig.adGroup.all_under_parents != ""
    ) {
      return this.getAllSubAccounts();
    }
    if (cadConfig.campaigns.all_under_parents != "") {
      Object.assign(
        accountsObjects,
        this.getAccountObjectsForIds(cadConfig.campaigns.all_under_parents)
      );
    }

    if (cadConfig.adGroup.all_under_parents != "") {
      Object.assign(
        accountsObjects,
        this.getAccountObjectsForIds(cadConfig.adGroup.all_under_parents)
      );
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
    let inputLabels = cadConfig.accounts.labels;
    const relevantLabels = [];
    for (const label of currentAccount.labels().get()) {
      if (label.getName() in inputLabels) {
        relevantLabels.push(label.getName());
      }
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
    const campaignIds = cadConfig.campaigns.ids;
    let forAccountIds = cadConfig.campaigns.all_under_parents;
    const campaignLabels = cadConfig.campaigns.campaign_labels;
    const excludedCampaignIds = cadConfig.campaigns.excluded_ids;
    const excludedClause =
      !excludedCampaignIds || excludedCampaignIds == ""
        ? ""
        : `AND campaign.id NOT IN (${excludedCampaignIds})`;

    if (forAccountIds != "") {
      if (CONFIG.is_debug_log) {
        Logger.log(
          `cadConfig.campaigns.all_under_parents =  ${JSON.stringify(
            cadConfig.campaigns.all_under_parents
          )}`
        );
      }
      if (forAccountIds.includes(currentAccount.getCustomerId())) {
        selectCampaignQuery = `SELECT campaign.id, campaign.name FROM campaign`;
        if (CONFIG.is_debug_log) {
          Logger.log(
            `campaigns for current account's query = ${selectCampaignQuery}`
          );
        }
        campaignMapForCurrentAccount = mapIdsToLabelNames(
          AdsApp.report(selectCampaignQuery, CONFIG.reporting_options),
          "campaign",
          campaignMapForCurrentAccount
        );
      }
    }
    if (campaignIds !== "") {
      selectCampaignQuery = `SELECT campaign.id, campaign.name  FROM campaign WHERE campaign.id IN (${campaignIds})  ${excludedClause}`;

      if (CONFIG.is_debug_log) {
        Logger.log(`campaignIds query = ${selectCampaignQuery}`);
      }
      campaignMapForCurrentAccount = mapIdsToLabelNames(
        AdsApp.report(selectCampaignQuery, CONFIG.reporting_options),
        "campaign",
        campaignMapForCurrentAccount
      );
    }
    // Adding campaigns by labels
    if (campaignLabels && campaignLabels.length > 0) {
      selectCampaignQuery = `SELECT campaign.id, campaign.name, label.name FROM campaign_label WHERE label.name IN (${campaignLabels}) ${excludedClause}`;

      if (CONFIG.is_debug_log) {
        Logger.log(`campaignLabels query = ${selectCampaignQuery}`);
      }
      campaignMapForCurrentAccount = mapIdsToLabelNames(
        AdsApp.report(selectCampaignQuery, CONFIG.reporting_options),
        "campaign",
        campaignMapForCurrentAccount
      );
    }
    return campaignMapForCurrentAccount;
  }
}

/**
 * Google Ads Ad-Grpoup Selector
 */
class AdGroupSelector {
  /**
   * Helper: generates a relevant ad group map for account
   * @param {!Object} cadConfig CAD config.
   * @return {!Object} The relevant adGroup map for the current account. Key:
   *     campaign-Id --> label name is exists
   */
  reportToRelevantAdGroupsMap(cadConfig) {
    let selectAdGroupsQuery;
    let adGroupMapForCurrentAccount = {};

    const adGroupIds = cadConfig.adGroup.ids;
    const underParentIds = cadConfig.adGroup.all_under_parents;
    const adGroupLabels = cadConfig.adGroup.labels;
    const excludedAdGroupIds = cadConfig.adGroup.excluded_ids;
    const excludedClause =
      !excludedAdGroupIds || excludedAdGroupIds == ""
        ? ""
        : `AND ad_group.id NOT IN (${excludedAdGroupIds})`;

    if (underParentIds != "") {
      selectAdGroupsQuery = `SELECT ad_group.id, ad_group.name FROM ad_group`;
      if (CONFIG.is_debug_log) {
        Logger.log(
          `adGroups for current account's query = ${selectAdGroupsQuery}`
        );
      }
      adGroupMapForCurrentAccount = mapIdsToLabelNames(
        AdsApp.report(selectAdGroupsQuery, CONFIG.reporting_options),
        "ad_group",
        adGroupMapForCurrentAccount
      );
    }
    if (adGroupIds !== "") {
      selectAdGroupsQuery = `SELECT ad_group.id, ad_group.name FROM ad_group WHERE ad_group.id IN (${adGroupIds})  ${excludedClause}`;

      if (CONFIG.is_debug_log) {
        Logger.log(`adGroup query = ${selectAdGroupsQuery}`);
      }
      adGroupMapForCurrentAccount = mapIdsToLabelNames(
        AdsApp.report(selectAdGroupsQuery, CONFIG.reporting_options),
        "ad_group",
        adGroupMapForCurrentAccount
      );
    }
    // Adding campaigns by labels
    if (adGroupLabels && adGroupLabels.length > 0) {
      selectAdGroupsQuery = `SELECT ad_group.id, ad_group.name, label.name FROM ad_group_label WHERE label.name IN (${adGroupLabels}) ${excludedClause}`;

      if (CONFIG.is_debug_log) {
        Logger.log(`adGroup Labels query = ${selectAdGroupsQuery}`);
      }
      adGroupMapForCurrentAccount = mapIdsToLabelNames(
        AdsApp.report(selectAdGroupsQuery, CONFIG.reporting_options),
        "ad_group",
        adGroupMapForCurrentAccount
      );
    }
    return adGroupMapForCurrentAccount;
  }
}

/**
 * Populates a hash-map from "campaign-id" to label names if exists
 * @param {!Object} searchResults Query results.
 * @param {!Object} entityMap a hash-map with "campaign-id" keys
 * @return {!Object} a hash-map with "campaign-id" keys
 */
function mapIdsToLabelNames(searchResults, entityLevel, entityMap) {
  if (!(entityMap instanceof Object)) {
    throw "reportToKeyMap expect entityMap to be a map";
  }
  for (const row of searchResults.rows()) {
    const levelId =
      entityLevel == "campaign" ? row["campaign.id"] : row["ad_group.id"];

    if (!entityMap[levelId]) {
      entityMap[levelId] = [];
    }
    if (row["label.name"]) {
      entityMap[levelId].push(row["label.name"]);
    }
  }
  return entityMap;
}

/** ============= Google Ads Reporter ==================== */

let _gads_query_fields_as_string = "";
let _gads_query_fields_as_array = [];

/**
 * returns a string with all the GAQL relevant query fields
 * @return {string} a string with all the GAQL relevant query fields
 */
function getGadsQueryFieldsAsString() {
  if (_gads_query_fields_as_array.length == 0) {
    Object.keys(MetricTypes).forEach(function (key) {
      if (MetricTypes[key].isFromGoogleAds) {
        _gads_query_fields_as_array.push(`metrics.${key}`);
      }
    });
    _gads_query_fields_as_string = _gads_query_fields_as_array.join(",");
  }
  return _gads_query_fields_as_string;
}

function getGadsQueryFieldsAsArray() {
  if (_gads_query_fields_as_array.length == 0) {
    getGadsQueryFieldsAsString();
  }
  return _gads_query_fields_as_array;
}

/**
 * Get results for relevant entities under MCC
 * @param {!Object} cadConfig CAD user input.
 * @return {Array<!CadResultForEntity>} CAD monitoring results
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
      getResultsForRelevantEntitiesUnderAccount(
        gAdsAccountSelector,
        currentAccount,
        cadConfig
      )
    );
  }

  return allEntitiesAllMetricsComparisons;
}

/**
 * Get results for relevant entities under account
 * @param {!Object} gAdsAccountSelector Google ads account selector
 * @param {!Object} currentAccount The current account.
 * @param {!Object} cadConfig CAD user input.
 * @return {Array<!CadResultForEntity>} CAD monitoring results
 */
function getResultsForRelevantEntitiesUnderAccount(
  gAdsAccountSelector,
  currentAccount,
  cadConfig
) {
  let cadResults = [];
  cadResults = cadResults.concat(
    aggAccountReportToCadResults(gAdsAccountSelector, currentAccount, cadConfig)
  );

  const gAdsCampaignSelector = new GoogleAdsCampaignSelector();
  const relevantCampaignIdsMap =
    gAdsCampaignSelector.reportToRelevantCampaignMap(currentAccount, cadConfig);
  cadResults = cadResults.concat(
    campaignReportToCadResults(relevantCampaignIdsMap, cadConfig)
  );

  const adGroupSelector = new AdGroupSelector();
  const relevantAdGroupIdsMap =
    adGroupSelector.reportToRelevantAdGroupsMap(cadConfig);
  cadResults = cadResults.concat(
    adGroupReportToCadResults(relevantAdGroupIdsMap, cadConfig)
  );

  if (CONFIG.is_debug_log) {
    Logger.log(
      "getResultsForRelevantEntitiesUnderAccount returned cadResults of " +
      JSON.stringify(cadResults)
    );
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
function aggAccountReportToCadResults(
  gAdsAccountSelector,
  currentAccount,
  cadConfig
) {
  const excludedAggAccountIds = cadConfig.accounts.excluded_ids;
  let cadResults = [];

  const relevantLabelsForCurrentAccount =
    gAdsAccountSelector.getRelevantLabelsForAccount(currentAccount, cadConfig);

  // Does currently scanned account satisfy any id selector
  if (
    !isCurrentAccountSatisfyCadConfig(
      currentAccount,
      cadConfig,
      relevantLabelsForCurrentAccount
    )
  ) {
    return cadResults;
  }

  if (excludedAggAccountIds.includes(currentAccount.getCustomerId())) {
    return cadResults;
  }

  // Customer query
  const getGadsCustomerQueryFields = removeElementFromStringList(
    "metrics.search_click_share",
    getGadsQueryFieldsAsString()
  );
  console.log(
    `getGadsCustomerQueryFields for CUSTOMER = ${getGadsCustomerQueryFields}`
  );
  let baseQuery = `SELECT customer.descriptive_name, ${getGadsCustomerQueryFields} ${cadConfig.hourSegmentInSelect}${cadConfig.splitByNetworkQuery} FROM customer WHERE segments.date BETWEEN`;

  let currentQuery = `${baseQuery} "${cadConfig.lookbackDates.current_range_start_date.query_date}" AND "${cadConfig.lookbackDates.current_range_end_date.query_date}"
  ${cadConfig.hourSegmentsWhereClause.current}`;
  let pastQuery = `${baseQuery} "${cadConfig.lookbackDates.past_range_start_date.query_date}" AND "${cadConfig.lookbackDates.past_range_end_date.query_date}" ${cadConfig.hourSegmentsWhereClause.past}`;


  if (CONFIG.is_debug_log) {
    Logger.log("currentQuery accounts= " + JSON.stringify(currentQuery));
    Logger.log("pastQuery accounts= " + JSON.stringify(pastQuery));
  }



  let currentStats = storeReportByEntityId(
    EntityType.Account,
    AdsApp.report(currentQuery, CONFIG.reporting_options)
  );
  Logger.log(LOG_PAST_DEVIDER);

  let pastStats = storeReportByEntityId(
    EntityType.Account,
    AdsApp.report(pastQuery, CONFIG.reporting_options)
  );

  if (CONFIG.is_debug_log) {
    Logger.log("currentStats accounts= " + JSON.stringify(currentStats));
    Logger.log("pastStats accounts= " + JSON.stringify(pastStats));
  }



  // single Row
  for (id in currentStats) {
    let cadResultForEntity = new CadResultForEntity();
    cadResultForEntity.relevant_label = relevantLabelsForCurrentAccount;

    cadResultForEntity.account.id = AdsApp.currentAccount().getCustomerId();
    cadResultForEntity.account.name =
      currentStats[id]["customer.descriptive_name"];

    cadResultForEntity.campaign.id = "";
    cadResultForEntity.campaign.name = "";
    cadResultForEntity.adGroup.id = "";
    cadResultForEntity.adGroup.name = "";

    cadResultForEntity.fillMetricComparisonResults(id, currentStats, pastStats, cadConfig);

    if (CONFIG.is_debug_log) {
      Logger.log(
        "aggAccountReportToCadResults. cadSingleResult= " +
        JSON.stringify(cadResultForEntity)
      );
    }
    if (cadResultForEntity.isTriggerAlert) {
      if (CONFIG.is_debug_log) {
        Logger.log(
          "aggAccountReportToCadResults triggers an alert");
      }
    }
    cadResults.push(cadResultForEntity);
  }

  if (CONFIG.is_debug_log) {
    Logger.log(
      "aggAccountReportToCadResults all-results= " + JSON.stringify(cadResults)
    );
  }
  return cadResults;
}

function removeElementFromStringList(elementName, elementString) {
  let elementsArray = elementString.split(",");
  let indexToRemove = elementsArray.indexOf(elementName);
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
  currentAccount,
  cadConfig,
  relevantLabelsForCurrentAccount
) {
  const aggAccountIdList = cadConfig.accounts.ids;
  const currentAccountId = currentAccount.getCustomerId();

  if (CONFIG.is_debug_log) {
    Logger.log(
      `Current account matches these monitored account labels = ${relevantLabelsForCurrentAccount}`
    );
  }

  return (
    (aggAccountIdList.length > 0 &&
      aggAccountIdList[0].toUpperCase() === CONSTS.ALL) ||
    aggAccountIdList.includes(currentAccountId) ||
    relevantLabelsForCurrentAccount.length > 0
  );
}

/**
 * campaign report to cad monitoring results
 * @param {!Object} entitieIdsForCurrentAccount The current account.
 * @param {!Object} cadConfig CAD user input.
 * @return {Array<!CadResultForEntity>} CAD monitoring results
 */
function campaignReportToCadResults(entitieIdsForCurrentAccount, cadConfig) {
  let cadResults = [];
  let entityIdsArr = Object.keys(entitieIdsForCurrentAccount);
  if (!entityIdsArr.length) {
    return cadResults;
  }
  let entityIds = entityIdsArr.join(",");
  if (CONFIG.is_debug_log) {
    Logger.log(`Reporting for ids = ${JSON.stringify(entityIds)}`);
  }

  // We expect current searchResults to contain DISABLED campaigns as well
  // thus > past results
  //  Campaign query
  let specificCampaignsQuery = `SELECT customer.descriptive_name, campaign.id, campaign.name, ${getGadsQueryFieldsAsString()} ${cadConfig.hourSegmentInSelect
    }${cadConfig.splitByNetworkQuery} FROM campaign WHERE campaign.id IN (${entityIds}) AND segments.date BETWEEN`;

  switch (cadConfig.avgType) {
    case AVG_TYPE.AVG_TYPE_DAILY_WEEKDAYS:
    case AVG_TYPE.AVG_TYPE_DAILY_TODAY_VS_YESTERDAY:
    case AVG_TYPE.AVG_TYPE_HOURLY_TODAY: {
      specificCampaignsQuery = removeElementFromStringList(
        "metrics.search_click_share",
        specificCampaignsQuery
      );
      break;
    }
  }

  let currentQuery = `${specificCampaignsQuery} "${cadConfig.lookbackDates.current_range_start_date.query_date}" AND "${cadConfig.lookbackDates.current_range_end_date.query_date}" ${cadConfig.hourSegmentsWhereClause.current}`;
  let pastQuery = `${specificCampaignsQuery} "${cadConfig.lookbackDates.past_range_start_date.query_date}" AND "${cadConfig.lookbackDates.past_range_end_date.query_date}"  ${cadConfig.hourSegmentsWhereClause.past}`;

  if (CONFIG.is_debug_log) {
    Logger.log("currentQuery = " + JSON.stringify(currentQuery));
    Logger.log("pastQuery = " + JSON.stringify(pastQuery));
  }

  let currentStats = storeReportByEntityId(
    EntityType.Campaign,
    AdsApp.report(currentQuery, CONFIG.reporting_options)
  );
  Logger.log(LOG_PAST_DEVIDER);
  let pastStats = storeReportByEntityId(
    EntityType.Campaign,
    AdsApp.report(pastQuery, CONFIG.reporting_options)
  );

  if (CONFIG.is_debug_log) {
    Logger.log("currentStats= " + JSON.stringify(currentStats));
    Logger.log(LOG_PAST_DEVIDER);
    Logger.log("pastStats= " + JSON.stringify(pastStats));
  }

  // Row
  for (campaignId of entityIdsArr) {
    let cadSingleResult = new CadResultForEntity();

    cadSingleResult.relevant_label = entitieIdsForCurrentAccount[campaignId];
    cadSingleResult.account.id = AdsApp.currentAccount().getCustomerId();
    cadSingleResult.account.name =
      currentStats[campaignId]["customer.descriptive_name"];
    cadSingleResult.campaign.id = currentStats[campaignId]["campaign.id"];
    cadSingleResult.campaign.name = currentStats[campaignId]["campaign.name"];
    cadSingleResult.fillMetricComparisonResults(
      campaignId,
      currentStats,
      pastStats,
      cadConfig
    );
    if (cadSingleResult.isTriggerAlert) {
      cadResults.push(cadSingleResult);
    }
  }
  return cadResults;
}

/**
 * campaign report to cad monitoring results
 * @param {!Object} entitieIdsForCurrentAccount The current account.
 * @param {!Object} cadConfig CAD user input.
 * @return {Array<!CadResultForEntity>} CAD monitoring results
 */
function adGroupReportToCadResults(entitieIdsForCurrentAccount, cadConfig) {
  let cadResults = [];
  let entityIdsArr = Object.keys(entitieIdsForCurrentAccount);
  if (!entityIdsArr.length) {
    return cadResults;
  }
  let entityIds = entityIdsArr.join(",");
  if (CONFIG.is_debug_log) {
    Logger.log(`Reporting for ids = ${JSON.stringify(entityIds)}`);
  }

  // We expect current searchResults to contain DISABLED campaigns as well
  // thus > past results
  //  Campaign query
  let specificEntityQuery = `SELECT customer.descriptive_name, campaign.id, campaign.name, ad_group.id, ad_group.name, ${getGadsQueryFieldsAsString()} ${cadConfig.hourSegmentInSelect
    }${cadConfig.splitByNetworkQuery} FROM ad_group WHERE ad_group.id IN (${entityIds}) AND segments.date BETWEEN`;

  specificEntityQuery = removeElementFromStringList(
    "metrics.search_click_share",
    specificEntityQuery
  );

  let currentQuery = `${specificEntityQuery} "${cadConfig.lookbackDates.current_range_start_date.query_date}" AND "${cadConfig.lookbackDates.current_range_end_date.query_date}" ${cadConfig.hourSegmentsWhereClause.current}`;
  let pastQuery = `${specificEntityQuery} "${cadConfig.lookbackDates.past_range_start_date.query_date}" AND "${cadConfig.lookbackDates.past_range_end_date.query_date}"  ${cadConfig.hourSegmentsWhereClause.past}`;

  if (CONFIG.is_debug_log) {
    Logger.log("currentQuery = " + JSON.stringify(currentQuery));
    Logger.log(LOG_PAST_DEVIDER);
    Logger.log("pastQuery = " + JSON.stringify(pastQuery));
  }

  let currentStats = storeReportByEntityId(
    EntityType.AdGroup,
    AdsApp.report(currentQuery, CONFIG.reporting_options)
  );
  Logger.log(LOG_PAST_DEVIDER);
  let pastStats = storeReportByEntityId(
    EntityType.AdGroup,
    AdsApp.report(pastQuery, CONFIG.reporting_options)
  );

  if (CONFIG.is_debug_log) {
    Logger.log("currentStats= " + JSON.stringify(currentStats));
    Logger.log(LOG_PAST_DEVIDER);
    Logger.log("pastStats= " + JSON.stringify(pastStats));
  }

  // Row
  for (adGroupId of entityIdsArr) {
    let cadSingleResult = new CadResultForEntity();

    cadSingleResult.relevant_label = entitieIdsForCurrentAccount[adGroupId];
    cadSingleResult.account.id = AdsApp.currentAccount().getCustomerId();
    cadSingleResult.account.name =
      currentStats[adGroupId]["customer.descriptive_name"];
    cadSingleResult.campaign.id = currentStats[adGroupId]["campaign.id"];
    cadSingleResult.campaign.name = currentStats[adGroupId]["campaign.name"];
    cadSingleResult.adGroup.id = currentStats[adGroupId]["ad_group.id"];
    cadSingleResult.adGroup.name = currentStats[adGroupId]["ad_group.name"];
    cadSingleResult.fillMetricComparisonResults(
      adGroupId,
      currentStats,
      pastStats,
      cadConfig
    );
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
function storeReportByEntityId(entityStr, searchResults) {
  let statsMap = {};
  for (const row of searchResults.rows()) {
    let id;
    switch (entityStr) {
      case "ad_group": {
        id = row["ad_group.id"];
        break;
      }
      case "campaign": {
        id = row["campaign.id"];
        break;
      }
      default:
      case "account": {
        id = AdsApp.currentAccount().getCustomerId();
        break;
      }
    }
    fillDictWithZerosWhenMissingKeys(row, getGadsQueryFieldsAsArray());

    if (row["segments.adNetworkType"]) {
      id = `${id}_${row["segments.adNetworkType"]}`;
    }

    if (CONFIG.is_debug_log) {
      //Logger.log("Row from API for id = " + id + " row=" + JSON.stringify(row));
    }
    if (!statsMap[id] || statsMap[id] == {}) {
      statsMap[id] = row;
    } else {
      statsMap[id] = sumRows(statsMap[id], row);
    }
  }
  return statsMap;
}

function fillDictWithZerosWhenMissingKeys(originalDict, newKeysDict) {
  for (let key of newKeysDict) {
    if (!(key in originalDict)) {
      originalDict[key] = 0;
    }
  }
}

function sumRows(row1, row2) {
  let result = {};
  const allKeys = unifyKeys(row1, row2);
  // for-in for dictionary. for-of for list/array
  for (const gAdsKey of allKeys) {
    // Handle undefined values gracefully
    const value1 = row1[gAdsKey] === undefined ? 0 : row1[gAdsKey];
    const value2 = row2[gAdsKey] === undefined ? 0 : row2[gAdsKey];

    if (isNumericOrNumericString(value1) && isNumericOrNumericString(value2)) {
      const key = removeGadsPrefix(gAdsKey);
      if (key && MetricTypes[key].isCumulative) {
        result[gAdsKey] = parseFloat(value1) + parseFloat(value2); // No need for || 0 since undefined is 0 now
      } else {
        result[gAdsKey] = 0;
      }
    } else if (gAdsKey.toLocaleLowerCase().includes('id') || gAdsKey.toLocaleLowerCase().includes('name')
    || gAdsKey.toLocaleLowerCase().includes('network')) {
      result[gAdsKey] = value1;
    } else {
      result[gAdsKey] = `cannot sum value2 = ${value2}`;
    }
  }
  return result;
}

function isNumericOrNumericString(value) {
  return typeof value === 'number' || (typeof value === 'string' && !isNaN(parseFloat(value)));
}

function removeGadsPrefix(gAdsKey) {
  if (gAdsKey.includes("metrics.")) {
    key = gAdsKey.replace("metrics.", "");
    return key;
  }
  return false;
}

function unifyKeys(obj1, obj2) {
  return Array.from(new Set(Object.keys(obj1).concat(Object.keys(obj2))));
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

    const triggeringResults = cadResults.filter((row) => row.isTriggerAlert);

    if (triggeringResults.length <= 0) {
      Logger.log("No alerts have been triggered. No email has been sent.");
      return;
    }

    if (!(emails && emails.length > 0)) {
      Logger.log("There are alerts triggered. But no email to send to.");
      return;
    }

    let bodyText = [];
    cadResults.forEach((item) => bodyText.push(item.toEmailFormat()));
    bodyText = bodyText.join("<br>");

    MailApp.sendEmail({
      to: emails,
      subject: `[Anomaly Alert] MCC Account [${cadConfig.mcc}]`,
      htmlBody: `
          Avg Type:
          ${cadConfig.avgType}
          <br><br>
          New period:
          ${cadConfig.lookbackDates.current_range_start_date.sheet_date} - ${cadConfig.lookbackDates.current_range_end_date.sheet_date} (${cadConfig.lookbackInDays.current_period_length_text} days)<br>
          Past period:
          ${cadConfig.lookbackDates.past_range_start_date.sheet_date} - ${cadConfig.lookbackDates.past_range_end_date.sheet_date} (${cadConfig.lookbackInDays.past_period_length_text} days)
          <br><br>
          ${bodyText}
          <br><br>
          Log into Google Ads and take a look.
          Notifications dashboard and full results:  ${CONFIG.spreadsheet_url}`,
    });
    Logger.log("Email sent");
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
    const expectedDate = new Date(
      new Date().setDate(new Date().getDate() - numDays)
    );

    if (CONFIG.is_debug_log) {
      Logger.log("getStringForMinusDaysAgo(numDays)= " + numDays);
      Logger.log("this.getDateStringInTimeZone(expectedDate, `yyyy-MM-dd`)= " + this.getDateStringInTimeZone(expectedDate, "yyyy-MM-dd"));
    }

    return {
      query_date: this.getDateStringInTimeZone(expectedDate, "yyyy-MM-dd"),
      sheet_date: this.getDateStringInTimeZone(expectedDate, "dd/MM/YY"),
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
      date,
      AdsApp.currentAccount().getTimeZone(),
      format
    );
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
    let str = num.toString().split(".");
    str[0] = str[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return str.join(".");
  }
}