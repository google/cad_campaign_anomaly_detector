// Copyright 2022 Google LLC.
//
// This solution, including any related sample code or data, is made available
// on an “as is,” “as available,” and “with all faults” basis, solely for
// illustrative purposes, and without warranty or representation of any kind.
// This solution is experimental, unsupported and provided solely for your
// convenience. Your use of it is subject to your agreements with Google, as
// applicable, and may constitute a beta feature as defined under those
// agreements.
//
// To the extent that you make any data available to Google in connection with
// your use of the solution, you represent and warrant that you have all
// necessary and appropriate rights, consents and permissions to permit Google
// to use and process that data.  By using any portion of this solution, you
// acknowledge, assume and accept all risks, known and unknown, associated with
// its usage, including with respect to your deployment of any portion of this
// solution in your systems, or usage in connection with your business, if at
// all.

/**
 * @name Campaign Anomaly Detector for an MCC
 *
 * @fileoverview Alerts the advertisers whenever one or more selected
 *accounts/campaigns an MCC account are suddenly behaving too differently from
 *the average of a past compared timeframe. See
 *https://developers.google.com/google-ads/scripts/docs/solutions/adsmanagerapp-account-anomaly-detector
 *for more details.
 *
 * Template spreadsheet:
 * https://docs.google.com/spreadsheets/d/1iEGccfrcbTTinUdnNXzhZm4XbyDy51vpjVoH1Ky2esk/copy
 *
 *
 * Reminder:
 * CPA = sum(cost)/sum(conv)
 * CTR = sum(clicks)/sum(imps)
 * CVR = sum(conv)/sum(clicks)
 * CVA = all_conversions_from_interactions_rate
 * CPC = sum(cost)/sum(clicks)
 * ROI all = all conversion value / Cost  (no need to calculate avg)
 *
 * 
 *
 * @author eladb@google.com, idoshalit@google.com
 *
 * @version 1.0
 *
 **/