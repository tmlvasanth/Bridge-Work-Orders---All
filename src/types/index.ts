export interface ParsedSheet {
  sheetName: string;
  headers: string[];
  data: Record<string, any>[];
  rowCount: number;
}

export interface UploadedFileState {
  file: File | null;
  fileName: string;
  sheetNames: string[];
  selectedSheet: string;
  sheets: Record<string, ParsedSheet>;
  headers: string[];
  data: Record<string, any>[];
  loading: boolean;
  error?: string;
}

export interface ColumnMappingConfig {
  // Bridge Base Excel Columns
  bridgeWorkOrderCol: string;
  bridgeTrackIdCol: string;
  bridgeTitleCol: string;
  bridgeActivityCol: string;
  bridgeDueDateCol: string;
  bridgeJobTypeCol: string;
  bridgeGrabberCol: string;
  bridgeTeamStaticValue: string; // cell B1 or default "Bridge"

  // Bridge Feedback Input Columns (from Base / Bridge Excel)
  bridgeXbrlFeedbackCommentsCol: string;
  bridgeXbrlFeedbackQualityCol: string;
  bridgeXbrlFeedbackStatusCol: string;
  bridgeXbrlFeedbackReviewedCol: string;
  bridgeXbrlFeedbackValidityCol: string;

  // Static prefill configuration
  samplingQualityPrefill?: string; // Default: "STP Feedback"

  // JMS Tracker Columns
  jmsWorkOrderCol: string;
  jmsTrackIdCol: string;
  jmsGrabberCol: string;
  jmsJobTypeCol: string;
  jmsFormTypeCol: string;

  // JMS Production Day to Day sheet cross-reference
  jmsProductionSheet: string;
  jmsProductionWorkOrderCol: string;
  jmsProductionTrackIdCol: string;
}

export interface MatchedOutputRow {
  id: string;
  team: string; // Base Excel cell B1 ("Bridge")
  workOrderNo: string; // JMS Job Number/Work Order No
  jmsCycleNoFbQuality: string; // Mapped from "XBRL Feedback Quality"
  xorVer: string; // blank
  xTrackId: string; // matched from JMS / Production Day to Day sheet / Bridge
  jobName: string; // Work Order Title (| Bridge XBRL removed)
  formType: string; // matched from Grabber Day to Day, or Job Type
  filingPeriod: string; // blank
  milestone: string; // Activity transformed (e.g. "Final to Client")
  jmsTeam: string; // "Bridge"
  samplingQuality: string; // Prefilled with "STP Feedback" for all applicable records
  consultantName: string; // blank
  secondReviewerName: string; // blank
  deliveredDate: string; // Due Date formatted as MM/DD/YYYY
  month: string; // month name derived from Delivered Date

  // Feedback audit and provenance fields
  xbrlFeedbackComments?: string;
  xbrlFeedbackQuality?: string;
  xbrlFeedbackStatus?: string;
  xbrlFeedbackReviewed?: string;
  xbrlFeedbackValidity?: string;

  // Secondary report fields if present
  secondaryReportType?: string;
  productionConsultantName?: string;
  internalReviewerName?: string;
  externalReviewerName?: string;
  deiErrors?: number;
  elementErrors?: number;
  valueErrors?: number;
  untaggedErrors?: number;
  l1L3Errors?: number;
  reverseSignErrors?: number;
  alternateValuesErrors?: number;
  dateContextErrors?: number;
  memberContextErrors?: number;
  scalePrecisionErrors?: number;
  taxonomyStructureInvalidSpi?: number;
  incorrectMappingAttribute?: number;
  judgmentalErrors?: number;
  validationErrors?: number;
  otherErrors?: number;
  labelErrors?: number;
  totalErrors?: number;
  rawSecondaryRow?: Record<string, any>;

  // Audit and provenance metadata
  rawBridgeRow?: Record<string, any>;
  rawJmsRow?: Record<string, any>;
  rawProdRow?: Record<string, any>;
  xTrackIdSource?: 'main_sheet' | 'production_sheet' | 'bridge_file' | 'none';
  matchKey: string;
  matchStatus: 'matched' | 'unmatched_bridge' | 'unmatched_jms';
}

export interface ReconciliationSummary {
  totalBridgeRows: number;
  totalJmsRows: number;
  matchedCount: number;
  unmatchedBridgeCount: number;
  unmatchedJmsCount: number;
  matchRatePercent: number;
  productionCrossMatchedCount: number;
}

export type ReconciliationFilterType =
  | 'matched'
  | 'unmatched_bridge'
  | 'matched_and_bridge_unmatched'
  | 'unmatched_jms'
  | 'all';
