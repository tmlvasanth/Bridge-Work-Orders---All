import { ColumnMappingConfig, MatchedOutputRow, ParsedSheet, ReconciliationSummary } from '../types';
import { formatDateToMMDDYYYY } from './excelParser';

export const CONFIG_STORAGE_KEY = 'bridge_jms_reconcile_mapping_config_v1';

/**
 * Loads saved column mapping and configuration from browser localStorage.
 */
export function loadSavedConfig(): Partial<ColumnMappingConfig> | null {
  try {
    const raw = localStorage.getItem(CONFIG_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') {
      return parsed;
    }
  } catch (err) {
    console.warn('Failed to load saved column mapping config from storage', err);
  }
  return null;
}

/**
 * Saves current column mapping and configuration to browser localStorage.
 */
export function saveConfigToStorage(config: ColumnMappingConfig): void {
  try {
    localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(config));
  } catch (err) {
    console.warn('Failed to save column mapping config to storage', err);
  }
}

/**
 * Clears saved column mapping configuration from browser localStorage.
 */
export function clearSavedConfig(): void {
  try {
    localStorage.removeItem(CONFIG_STORAGE_KEY);
  } catch (err) {
    console.warn('Failed to clear column mapping config from storage', err);
  }
}

/**
 * Normalizes work order numbers for accurate matching.
 * Handles numbers, strings, leading/trailing whitespace, and common prefixes.
 */
export function normalizeWorkOrderKey(val: any): string {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  // Strip non-breaking spaces or tabs
  str = str.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
  // Case insensitive matching
  return str.toUpperCase();
}

/**
 * Clean Job Name by removing '| Bridge XBRL' if present.
 */
export function cleanJobName(rawTitle: any): string {
  if (!rawTitle) return '';
  let title = String(rawTitle).trim();

  // Pattern 1: '| Bridge XBRL' with various spacing or casing
  title = title.replace(/\s*\|\s*Bridge\s*XBRL(\b|$)/gi, '');

  // Pattern 2: '- Bridge XBRL' or 'Bridge XBRL' at the end
  title = title.replace(/\s*-\s*Bridge\s*XBRL(\b|$)/gi, '');
  title = title.replace(/\s*Bridge\s*XBRL$/gi, '');

  // Clean any dangling trailing pipe or hyphen
  title = title.replace(/[\s|\-]+$/, '').trim();

  return title;
}

/**
 * Milestone transformation:
 * - if 'XBRL Update - Final to Client' then keep only 'Final to Client'
 * - if 'XBRL Update - Non Standard' then keep only 'Non Standard'
 * - if 'XBRL Update - Standard' then keep only 'Standard'
 * - generalized fallback for 'XBRL Update - ...'
 */
export function transformMilestone(rawActivity: any): string {
  if (!rawActivity) return '';
  const activity = String(rawActivity).trim();

  // Specific exact checks requested
  if (/^XBRL\s*Update\s*-\s*Final\s*to\s*Client$/i.test(activity)) {
    return 'Final to Client';
  }
  if (/^XBRL\s*Update\s*-\s*Non\s*Standard$/i.test(activity)) {
    return 'Non Standard';
  }
  if (/^XBRL\s*Update\s*-\s*Standard$/i.test(activity)) {
    return 'Standard';
  }

  // Generalized pattern: XBRL Update - <Something>
  const match = activity.match(/^XBRL\s*Update\s*[-–—:]\s*(.+)$/i);
  if (match && match[1]) {
    return match[1].trim();
  }

  return activity;
}

/**
 * Clean Feedback Quality value:
 * Strips leading numbers and punctuation from FB Quality wordings as requested:
 * e.g., "4 - Outstanding" -> "Outstanding"
 *       "3 - Good" -> "Good"
 *       "2 - Fair" -> "Fair"
 *       "1 - Poor" -> "Poor"
 *       "4-Outstanding" -> "Outstanding"
 *       "4. Outstanding" -> "Outstanding"
 */
export function cleanFeedbackQuality(rawQuality: any): string {
  if (rawQuality === undefined || rawQuality === null) return '';
  const str = String(rawQuality).trim();
  if (!str) return '';

  // Remove leading digits followed by dash, colon, period, and any surrounding spaces
  const cleaned = str.replace(/^\d+\s*[-–—:.]\s*/, '').trim();
  return cleaned || str;
}

/**
 * Determine Form Type:
 * matched from Grabber Day to Day and if not available check under job type if available bring it.
 */
export function resolveFormType(
  grabberVal: any,
  jobTypeVal: any,
  fallbackVal?: any
): string {
  if (grabberVal !== undefined && grabberVal !== null && String(grabberVal).trim() !== '') {
    return String(grabberVal).trim();
  }
  if (jobTypeVal !== undefined && jobTypeVal !== null && String(jobTypeVal).trim() !== '') {
    return String(jobTypeVal).trim();
  }
  if (fallbackVal !== undefined && fallbackVal !== null && String(fallbackVal).trim() !== '') {
    return String(fallbackVal).trim();
  }
  return '';
}

/**
 * Exact canonical default mapping profile from user image specification:
 * Bridge: Work Order Number, Work Order Title, Activity, Due Date
 * JMS: Work Order No, X Track ID, Grabber Day to Day, Job Type
 * Production Day to Day sheet: Work Order No, X Track ID
 */
export const DEFAULT_CANONICAL_CONFIG: ColumnMappingConfig = {
  bridgeWorkOrderCol: 'Work Order',
  bridgeTrackIdCol: 'X Track ID',
  bridgeTitleCol: 'Work Order Title',
  bridgeActivityCol: 'Activity',
  bridgeDueDateCol: 'Due Date',
  bridgeJobTypeCol: 'Job Type',
  bridgeGrabberCol: 'Grabber Day to Day',
  bridgeTeamStaticValue: 'Bridge',

  // Feedback input columns from screenshot
  bridgeXbrlFeedbackCommentsCol: 'XBRL Feedback Comments',
  bridgeXbrlFeedbackQualityCol: 'XBRL Feedback Quality',
  bridgeXbrlFeedbackStatusCol: 'XBRL Feedback Status',
  bridgeXbrlFeedbackReviewedCol: 'XBRL Feedback Reviewed',
  bridgeXbrlFeedbackValidityCol: 'XBRL Feedback Validity',

  samplingQualityPrefill: 'STP Feedback',

  jmsWorkOrderCol: 'Work Order No',
  jmsTrackIdCol: 'X Track ID',
  jmsGrabberCol: 'Grabber Day to Day',
  jmsJobTypeCol: 'Job Type',
  jmsFormTypeCol: 'Form Type',

  jmsProductionSheet: 'Production Day to Day',
  jmsProductionWorkOrderCol: 'Work Order No',
  jmsProductionTrackIdCol: 'X Track ID',
};

/**
 * Normalizes a header or text string: lowercases, removes non-alphanumeric punctuation (except spaces),
 * collapses multiple spaces, and trims.
 */
export function normalizeHeaderString(h: string): string {
  if (!h) return '';
  return h
    .toLowerCase()
    .replace(/[\r\n\t_/#\-\.]+/g, ' ')
    .replace(/[^\w\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Robust column finder that prioritizes exact canonical matches before partial substring matches,
 * while respecting exclusion rules so titles/names/dates are never confused with numbers/IDs.
 */
export function findBestColumnMatch(
  availableHeaders: string[],
  targetPreferences: string[],
  excludePatterns: string[] = []
): string {
  if (!availableHeaders || availableHeaders.length === 0) return '';

  const normalizedAvailable = availableHeaders.map((h) => ({
    original: h,
    normalized: normalizeHeaderString(h),
  }));

  // Pass 1: Exact normalized match (ignoring case, spaces, symbols)
  for (const target of targetPreferences) {
    const normTarget = normalizeHeaderString(target);
    const exact = normalizedAvailable.find((item) => item.normalized === normTarget);
    if (exact) return exact.original;
  }

  // Pass 2: Contains match with exclusion filtering
  for (const target of targetPreferences) {
    const normTarget = normalizeHeaderString(target);
    for (const item of normalizedAvailable) {
      const isExcluded = excludePatterns.some((exc) => item.normalized.includes(exc.toLowerCase()));
      if (isExcluded) continue;

      if (
        item.normalized.includes(normTarget) ||
        normTarget.includes(item.normalized) ||
        item.original.toLowerCase().includes(target.toLowerCase())
      ) {
        return item.original;
      }
    }
  }

  return '';
}

/**
 * Auto-detect column headers based on user preferences and smart keywords.
 */
export function autoDetectColumns(
  bridgeHeaders: string[],
  jmsHeaders: string[],
  jmsSheets?: Record<string, ParsedSheet>,
  savedPreferences?: Partial<ColumnMappingConfig>
): ColumnMappingConfig {
  // Bridge columns
  const bridgeWorkOrderCol =
    (savedPreferences?.bridgeWorkOrderCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeWorkOrderCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      [
        'Work Order Number',
        'Work Order No',
        'Work Order',
        'WO Number',
        'WO#',
        'JMS Job Number/Work Order No',
        'Job Number',
        'Job No',
        'Order No',
      ],
      ['title', 'name', 'desc', 'activity', 'date', 'status', 'type', 'cycle', 'reviewer', 'ver']
    ) ||
    bridgeHeaders[0] ||
    '';

  const bridgeTitleCol =
    (savedPreferences?.bridgeTitleCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeTitleCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      [
        'Work Order Title',
        'Job Name',
        'Work Order Name',
        'Title',
        'Client Name',
        'Company Name',
        'Order Title',
        'Description',
      ],
      ['number', 'no', 'num', 'id', 'date', 'due', 'activity', 'type', 'track', 'cycle', 'reviewer', 'status']
    ) ||
    '';

  const bridgeActivityCol =
    (savedPreferences?.bridgeActivityCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeActivityCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      ['Activity', 'Milestone', 'Task', 'Activity Name', 'Stage', 'Status'],
      ['date', 'title', 'number', 'id']
    ) ||
    '';

  const bridgeDueDateCol =
    (savedPreferences?.bridgeDueDateCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeDueDateCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      ['Due Date', 'Delivered Date', 'Delivery Date', 'Target Date', 'Due_Date', 'Deadline'],
      ['create', 'created', 'start', 'update', 'modified']
    ) ||
    '';

  const bridgeJobTypeCol =
    (savedPreferences?.bridgeJobTypeCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeJobTypeCol])) ||
    findBestColumnMatch(bridgeHeaders, ['Job Type', 'Form Type', 'Type', 'Filing Type'], ['number', 'title', 'date']) ||
    '';

  const bridgeGrabberCol =
    (savedPreferences?.bridgeGrabberCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeGrabberCol])) ||
    findBestColumnMatch(bridgeHeaders, ['Grabber Day to Day', 'Grabber', 'Form Type', 'Day to Day']) ||
    '';

  // Input Columns from Feedback Screenshot (Work Order, Work Order Title, Due Date, Activity, Job Type + Feedback columns)
  const bridgeXbrlFeedbackQualityCol =
    (savedPreferences?.bridgeXbrlFeedbackQualityCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeXbrlFeedbackQualityCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      [
        'XBRL Feedback Quality',
        'Feedback Quality',
        'XBRL Quality',
        'FB Quality',
        'JMS Cycle No / FB Quality',
        'Cycle No / FB Quality',
        'FB_Quality',
      ],
      ['status', 'comment', 'review', 'valid']
    ) ||
    '';

  const bridgeXbrlFeedbackCommentsCol =
    (savedPreferences?.bridgeXbrlFeedbackCommentsCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeXbrlFeedbackCommentsCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      [
        'XBRL Feedback Comments',
        'Feedback Comments',
        'XBRL Comments',
        'Feedback Comment',
        'FB Comments',
        'Comments',
      ],
      ['status', 'quality', 'review', 'valid']
    ) ||
    '';

  const bridgeXbrlFeedbackStatusCol =
    (savedPreferences?.bridgeXbrlFeedbackStatusCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeXbrlFeedbackStatusCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      [
        'XBRL Feedback Status',
        'Feedback Status',
        'XBRL Status',
        'FB Status',
      ],
      ['quality', 'comment', 'review', 'valid']
    ) ||
    '';

  const bridgeXbrlFeedbackReviewedCol =
    (savedPreferences?.bridgeXbrlFeedbackReviewedCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeXbrlFeedbackReviewedCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      [
        'XBRL Feedback Reviewed',
        'Feedback Reviewed',
        'XBRL Reviewed',
        'FB Reviewed',
        'Reviewed',
      ],
      ['quality', 'comment', 'status', 'valid']
    ) ||
    '';

  const bridgeXbrlFeedbackValidityCol =
    (savedPreferences?.bridgeXbrlFeedbackValidityCol &&
      findBestColumnMatch(bridgeHeaders, [savedPreferences.bridgeXbrlFeedbackValidityCol])) ||
    findBestColumnMatch(
      bridgeHeaders,
      [
        'XBRL Feedback Validity',
        'Feedback Validity',
        'XBRL Validity',
        'FB Validity',
        'Validity',
      ],
      ['quality', 'comment', 'status', 'review']
    ) ||
    '';

  // JMS Main Sheet Columns
  const jmsWorkOrderCol =
    (savedPreferences?.jmsWorkOrderCol &&
      findBestColumnMatch(jmsHeaders, [savedPreferences.jmsWorkOrderCol])) ||
    findBestColumnMatch(
      jmsHeaders,
      [
        'Work Order No',
        'Work Order Number',
        'JMS Job Number',
        'Work Order',
        'Job Number',
        'JMS Job No',
        'Job No',
        'WO No',
        'WO#',
        'JMS Job Number/Work Order No',
      ],
      ['title', 'name', 'desc', 'date', 'status', 'type', 'cycle', 'reviewer', 'ver', 'track']
    ) ||
    jmsHeaders[0] ||
    '';

  const jmsTrackIdCol =
    (savedPreferences?.jmsTrackIdCol &&
      findBestColumnMatch(jmsHeaders, [savedPreferences.jmsTrackIdCol])) ||
    findBestColumnMatch(
      jmsHeaders,
      ['X Track ID', 'X-Track ID', 'X Track', 'Track ID', 'X_Track_JobID', 'Tracking ID', 'XTrackID'],
      ['work', 'order', 'title', 'date', 'type']
    ) ||
    '';

  const jmsGrabberCol =
    (savedPreferences?.jmsGrabberCol &&
      findBestColumnMatch(jmsHeaders, [savedPreferences.jmsGrabberCol])) ||
    findBestColumnMatch(
      jmsHeaders,
      ['Grabber Day to Day', 'Grabber Form', 'Grabber', 'Form Type', 'Day to Day', 'Form_Type'],
      ['production']
    ) ||
    '';

  const jmsJobTypeCol =
    (savedPreferences?.jmsJobTypeCol &&
      findBestColumnMatch(jmsHeaders, [savedPreferences.jmsJobTypeCol])) ||
    findBestColumnMatch(
      jmsHeaders,
      ['Job Type', 'Job_Type', 'Type', 'Form Type'],
      ['track', 'order', 'number']
    ) ||
    '';

  const jmsFormTypeCol =
    (savedPreferences?.jmsFormTypeCol &&
      findBestColumnMatch(jmsHeaders, [savedPreferences.jmsFormTypeCol])) ||
    findBestColumnMatch(jmsHeaders, ['Form Type', 'Form_Type', 'Sec Form', 'Form']) ||
    '';

  // Auto-detect "Production Day to Day" sheet in JMS workbook
  let jmsProductionSheet = '';
  let jmsProductionWorkOrderCol = '';
  let jmsProductionTrackIdCol = '';

  if (jmsSheets) {
    const sheetNames = Object.keys(jmsSheets);
    const preferredSheet = savedPreferences?.jmsProductionSheet;
    const prodSheetName =
      (preferredSheet && sheetNames.find((s) => normalizeHeaderString(s) === normalizeHeaderString(preferredSheet))) ||
      sheetNames.find((s) => /production.*day.*to.*day/i.test(s)) ||
      sheetNames.find((s) => /production/i.test(s));

    if (prodSheetName && jmsSheets[prodSheetName]) {
      jmsProductionSheet = prodSheetName;
      const prodHeaders = jmsSheets[prodSheetName].headers;
      jmsProductionWorkOrderCol =
        (savedPreferences?.jmsProductionWorkOrderCol &&
          findBestColumnMatch(prodHeaders, [savedPreferences.jmsProductionWorkOrderCol])) ||
        findBestColumnMatch(
          prodHeaders,
          ['Work Order No', 'Work Order Number', 'Work Order', 'Job Number', 'WO No', 'WO#'],
          ['title', 'track', 'status', 'shift']
        ) ||
        prodHeaders[0] ||
        '';

      jmsProductionTrackIdCol =
        (savedPreferences?.jmsProductionTrackIdCol &&
          findBestColumnMatch(prodHeaders, [savedPreferences.jmsProductionTrackIdCol])) ||
        findBestColumnMatch(
          prodHeaders,
          ['X Track ID', 'X-Track ID', 'X Track', 'Track ID', 'Tracking ID', 'X_Track_JobID'],
          ['order', 'work', 'status']
        ) ||
        '';
    }
  }

  return {
    bridgeWorkOrderCol,
    bridgeTrackIdCol: savedPreferences?.bridgeTrackIdCol || 'X Track ID',
    bridgeTitleCol,
    bridgeActivityCol,
    bridgeDueDateCol,
    bridgeJobTypeCol,
    bridgeGrabberCol,
    bridgeTeamStaticValue: savedPreferences?.bridgeTeamStaticValue || 'Bridge',
    bridgeXbrlFeedbackCommentsCol,
    bridgeXbrlFeedbackQualityCol,
    bridgeXbrlFeedbackStatusCol,
    bridgeXbrlFeedbackReviewedCol,
    bridgeXbrlFeedbackValidityCol,
    samplingQualityPrefill: savedPreferences?.samplingQualityPrefill || 'STP Feedback',
    jmsWorkOrderCol,
    jmsTrackIdCol,
    jmsGrabberCol,
    jmsJobTypeCol,
    jmsFormTypeCol,
    jmsProductionSheet,
    jmsProductionWorkOrderCol,
    jmsProductionTrackIdCol,
  };
}

/**
 * Performs reconciliation between Bridge work orders and JMS/SaaS workbook.
 * Also checks the Production Day to Day sheet in the JMS/SaaS workbook for X Track ID.
 */
export function reconcileWorkOrders(
  bridgeData: Record<string, any>[],
  jmsData: Record<string, any>[],
  config: ColumnMappingConfig,
  jmsSheets?: Record<string, ParsedSheet>
): {
  matchedRows: MatchedOutputRow[];
  unmatchedBridgeRows: MatchedOutputRow[];
  unmatchedJmsRows: MatchedOutputRow[];
  allRows: MatchedOutputRow[];
  summary: ReconciliationSummary;
} {
  // Index main JMS data by normalized work order key
  const jmsMap = new Map<string, { row: Record<string, any>; originalIndex: number }>();
  const jmsMatchedKeys = new Set<string>();

  jmsData.forEach((row, idx) => {
    const rawWo = row[config.jmsWorkOrderCol];
    const key = normalizeWorkOrderKey(rawWo);
    if (key && !jmsMap.has(key)) {
      jmsMap.set(key, { row, originalIndex: idx });
    }
  });

  // Index Production Day to Day sheet if configured
  const productionMap = new Map<string, Record<string, any>>();
  if (config.jmsProductionSheet && jmsSheets && jmsSheets[config.jmsProductionSheet]) {
    const prodRows = jmsSheets[config.jmsProductionSheet].data || [];
    prodRows.forEach((row) => {
      const rawWo = row[config.jmsProductionWorkOrderCol];
      const key = normalizeWorkOrderKey(rawWo);
      if (key && !productionMap.has(key)) {
        productionMap.set(key, row);
      }
    });
  }

  const matchedRows: MatchedOutputRow[] = [];
  const unmatchedBridgeRows: MatchedOutputRow[] = [];

  bridgeData.forEach((bridgeRow, idx) => {
    const rawWo = bridgeRow[config.bridgeWorkOrderCol];
    const key = normalizeWorkOrderKey(rawWo);

    if (!key) return; // Skip completely empty work orders

    const jmsMatch = jmsMap.get(key);
    const prodMatch = productionMap.get(key);

    const rawTitle = bridgeRow[config.bridgeTitleCol];
    const cleanedJobName = cleanJobName(rawTitle);

    const rawActivity = bridgeRow[config.bridgeActivityCol];
    const milestone = transformMilestone(rawActivity);

    const rawDueDate = bridgeRow[config.bridgeDueDateCol];
    const { formatted: deliveredDate, month } = formatDateToMMDDYYYY(rawDueDate);

    // Form Type resolution:
    // "matched from Grabber Day to Day and if not available check under job type if available bring it"
    let grabberVal = '';
    let jobTypeVal = '';
    let formTypeVal = '';

    if (jmsMatch) {
      grabberVal = jmsMatch.row[config.jmsGrabberCol] || bridgeRow[config.bridgeGrabberCol] || '';
      jobTypeVal = jmsMatch.row[config.jmsJobTypeCol] || bridgeRow[config.bridgeJobTypeCol] || '';
      formTypeVal = jmsMatch.row[config.jmsFormTypeCol] || '';
    } else {
      grabberVal = bridgeRow[config.bridgeGrabberCol] || '';
      jobTypeVal = bridgeRow[config.bridgeJobTypeCol] || '';
    }

    const formType = resolveFormType(grabberVal, jobTypeVal, formTypeVal);

    // Resolve X Track ID:
    // 1. Check primary JMS sheet
    // 2. Also check and take for the respective work order number if available in the Production Day to Day sheet
    let xTrackId = '';
    let xTrackIdSource: 'main_sheet' | 'production_sheet' | 'none' = 'none';

    if (jmsMatch && config.jmsTrackIdCol) {
      const mainVal = jmsMatch.row[config.jmsTrackIdCol];
      if (mainVal !== undefined && mainVal !== null && String(mainVal).trim() !== '') {
        xTrackId = String(mainVal).trim();
        xTrackIdSource = 'main_sheet';
      }
    }

    // Check Production Day to Day sheet if main was blank or not found
    if (prodMatch && config.jmsProductionTrackIdCol) {
      const prodVal = prodMatch[config.jmsProductionTrackIdCol];
      if (prodVal !== undefined && prodVal !== null && String(prodVal).trim() !== '') {
        // If main was empty, take from production day to day sheet!
        if (!xTrackId) {
          xTrackId = String(prodVal).trim();
          xTrackIdSource = 'production_sheet';
        }
      }
    }

    const isMatch = Boolean(jmsMatch || prodMatch);

    // Extract XBRL Feedback values from Bridge row
    const rawFbQuality =
      (config.bridgeXbrlFeedbackQualityCol && bridgeRow[config.bridgeXbrlFeedbackQualityCol] !== undefined && bridgeRow[config.bridgeXbrlFeedbackQualityCol] !== null)
        ? String(bridgeRow[config.bridgeXbrlFeedbackQualityCol]).trim()
        : '';
    const fbQuality = cleanFeedbackQuality(rawFbQuality);

    const fbComments =
      (config.bridgeXbrlFeedbackCommentsCol && bridgeRow[config.bridgeXbrlFeedbackCommentsCol] !== undefined && bridgeRow[config.bridgeXbrlFeedbackCommentsCol] !== null)
        ? String(bridgeRow[config.bridgeXbrlFeedbackCommentsCol]).trim()
        : '';

    const fbStatus =
      (config.bridgeXbrlFeedbackStatusCol && bridgeRow[config.bridgeXbrlFeedbackStatusCol] !== undefined && bridgeRow[config.bridgeXbrlFeedbackStatusCol] !== null)
        ? String(bridgeRow[config.bridgeXbrlFeedbackStatusCol]).trim()
        : '';

    const fbReviewed =
      (config.bridgeXbrlFeedbackReviewedCol && bridgeRow[config.bridgeXbrlFeedbackReviewedCol] !== undefined && bridgeRow[config.bridgeXbrlFeedbackReviewedCol] !== null)
        ? String(bridgeRow[config.bridgeXbrlFeedbackReviewedCol]).trim()
        : '';

    const fbValidity =
      (config.bridgeXbrlFeedbackValidityCol && bridgeRow[config.bridgeXbrlFeedbackValidityCol] !== undefined && bridgeRow[config.bridgeXbrlFeedbackValidityCol] !== null)
        ? String(bridgeRow[config.bridgeXbrlFeedbackValidityCol]).trim()
        : '';

    // Wherever "JMS Cycle No / FB Quality" is filled in the output, "Sampling/Quality" is filled as "STP Feedback"
    const hasFbQuality = Boolean(fbQuality && fbQuality.length > 0);
    const rowSamplingQuality = hasFbQuality ? (config.samplingQualityPrefill || 'STP Feedback') : '';

    if (isMatch) {
      jmsMatchedKeys.add(key);

      const outputRow: MatchedOutputRow = {
        id: `matched-${idx}-${key}`,
        team: 'Bridge',
        workOrderNo: String(rawWo).trim(),
        jmsCycleNoFbQuality: fbQuality, // Mapped from "XBRL Feedback Quality"
        xorVer: '',
        xTrackId: xTrackId,
        jobName: cleanedJobName,
        formType: formType,
        filingPeriod: '',
        milestone: milestone,
        jmsTeam: 'Bridge',
        samplingQuality: rowSamplingQuality, // "STP Feedback" whenever JMS Cycle No / FB Quality is filled
        consultantName: '',
        secondReviewerName: '',
        deliveredDate: deliveredDate,
        month: month,
        xbrlFeedbackComments: fbComments,
        xbrlFeedbackQuality: fbQuality,
        xbrlFeedbackStatus: fbStatus,
        xbrlFeedbackReviewed: fbReviewed,
        xbrlFeedbackValidity: fbValidity,
        rawBridgeRow: bridgeRow,
        rawJmsRow: jmsMatch?.row,
        rawProdRow: prodMatch,
        xTrackIdSource: xTrackIdSource,
        matchKey: key,
        matchStatus: 'matched',
      };

      matchedRows.push(outputRow);
    } else {
      // Unmatched Bridge record
      const outputRow: MatchedOutputRow = {
        id: `unmatched-bridge-${idx}-${key}`,
        team: 'Bridge',
        workOrderNo: String(rawWo).trim(),
        jmsCycleNoFbQuality: fbQuality, // Mapped from "XBRL Feedback Quality"
        xorVer: '',
        xTrackId: '',
        jobName: cleanedJobName,
        formType: formType,
        filingPeriod: '',
        milestone: milestone,
        jmsTeam: 'Bridge',
        samplingQuality: rowSamplingQuality, // "STP Feedback" whenever JMS Cycle No / FB Quality is filled
        consultantName: '',
        secondReviewerName: '',
        deliveredDate: deliveredDate,
        month: month,
        xbrlFeedbackComments: fbComments,
        xbrlFeedbackQuality: fbQuality,
        xbrlFeedbackStatus: fbStatus,
        xbrlFeedbackReviewed: fbReviewed,
        xbrlFeedbackValidity: fbValidity,
        rawBridgeRow: bridgeRow,
        matchKey: key,
        matchStatus: 'unmatched_bridge',
      };

      unmatchedBridgeRows.push(outputRow);
    }
  });

  // Collect unmatched JMS rows
  const unmatchedJmsRows: MatchedOutputRow[] = [];

  jmsData.forEach((jmsRow, idx) => {
    const rawWo = jmsRow[config.jmsWorkOrderCol];
    const key = normalizeWorkOrderKey(rawWo);

    if (key && !jmsMatchedKeys.has(key)) {
      let xTrackId = jmsRow[config.jmsTrackIdCol] ? String(jmsRow[config.jmsTrackIdCol]).trim() : '';
      let xTrackIdSource: 'main_sheet' | 'production_sheet' | 'none' = xTrackId ? 'main_sheet' : 'none';

      const prodMatch = productionMap.get(key);
      if (!xTrackId && prodMatch && config.jmsProductionTrackIdCol) {
        const prodVal = prodMatch[config.jmsProductionTrackIdCol];
        if (prodVal !== undefined && prodVal !== null && String(prodVal).trim() !== '') {
          xTrackId = String(prodVal).trim();
          xTrackIdSource = 'production_sheet';
        }
      }

      const grabberVal = jmsRow[config.jmsGrabberCol] || '';
      const jobTypeVal = jmsRow[config.jmsJobTypeCol] || '';
      const formTypeVal = jmsRow[config.jmsFormTypeCol] || '';
      const formType = resolveFormType(grabberVal, jobTypeVal, formTypeVal);

      unmatchedJmsRows.push({
        id: `unmatched-jms-${idx}-${key}`,
        team: 'Bridge',
        workOrderNo: String(rawWo).trim(),
        jmsCycleNoFbQuality: '',
        xorVer: '',
        xTrackId: xTrackId,
        jobName: '',
        formType: formType,
        filingPeriod: '',
        milestone: '',
        jmsTeam: 'Bridge',
        samplingQuality: '', // Blank since JMS Cycle No / FB Quality is not filled
        consultantName: '',
        secondReviewerName: '',
        deliveredDate: '',
        month: '',
        rawJmsRow: jmsRow,
        rawProdRow: prodMatch,
        xTrackIdSource: xTrackIdSource,
        matchKey: key,
        matchStatus: 'unmatched_jms',
      });
    }
  });

  const totalBridgeRows = bridgeData.filter((r) => normalizeWorkOrderKey(r[config.bridgeWorkOrderCol])).length;
  const totalJmsRows = jmsData.filter((r) => normalizeWorkOrderKey(r[config.jmsWorkOrderCol])).length;
  const matchRatePercent = totalBridgeRows > 0 ? Math.round((matchedRows.length / totalBridgeRows) * 100) : 0;

  const productionCrossMatchedCount = matchedRows.filter((r) => r.xTrackIdSource === 'production_sheet').length;

  const summary: ReconciliationSummary = {
    totalBridgeRows,
    totalJmsRows,
    matchedCount: matchedRows.length,
    unmatchedBridgeCount: unmatchedBridgeRows.length,
    unmatchedJmsCount: unmatchedJmsRows.length,
    matchRatePercent,
    productionCrossMatchedCount,
  };

  const allRows = [...matchedRows, ...unmatchedBridgeRows, ...unmatchedJmsRows];

  return {
    matchedRows,
    unmatchedBridgeRows,
    unmatchedJmsRows,
    allRows,
    summary,
  };
}
