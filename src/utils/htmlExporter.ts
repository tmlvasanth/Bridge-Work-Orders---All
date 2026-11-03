import { MatchedOutputRow } from '../types';
import {
  ColumnDefinition,
  STANDARD_15_COLUMNS,
  XBRL_FEEDBACK_COLUMNS,
  QC_IMAGE_COLUMNS,
  ALL_COLUMNS,
} from './excelExporter';
import {
  SAMPLE_BRIDGE_DATA,
  SAMPLE_JMS_DATA,
  SAMPLE_PRODUCTION_DAY_TO_DAY_DATA,
  SCREENSHOT_EXAMPLE_WORK_ORDERS,
  SCREENSHOT_CORE_5_WORK_ORDERS,
  SCREENSHOT_TABLE1_WORK_ORDERS,
} from '../data/sampleData';

export const COLUMN_HEADER_COLORS: Record<string, { bg: string; text: string; border: string }> = {
  'Team': { bg: '#F1F5F9', text: '#1E293B', border: '#CBD5E1' },
  'JMS Job Number/ Work Order No': { bg: '#EFF6FF', text: '#1E40AF', border: '#93C5FD' },
  'JMS Job Number/Work Order No': { bg: '#EFF6FF', text: '#1E40AF', border: '#93C5FD' },
  'JMS Cycle No / FB Quality': { bg: '#F0FDFA', text: '#0F766E', border: '#99F6E4' },
  'XOR Ver': { bg: '#F8FAFC', text: '#475569', border: '#E2E8F0' },
  'X Track ID': { bg: '#F0F9FF', text: '#0369A1', border: '#7DD3FC' },
  'Job Name': { bg: '#F0FDF4', text: '#166534', border: '#86EFAC' },
  'Month': { bg: '#FFF1F2', text: '#9F1239', border: '#FECDD3' },
  'Form Type': { bg: '#FFFBEB', text: '#92400E', border: '#FCD34D' },
  'Filing Period': { bg: '#F8FAFC', text: '#475569', border: '#E2E8F0' },
  'Milestone': { bg: '#FAF5FF', text: '#6B21A8', border: '#D8B4FE' },
  'JMS Team': { bg: '#F1F5F9', text: '#1E293B', border: '#CBD5E1' },
  'Sampling/Quality': { bg: '#EEF2FF', text: '#3730A3', border: '#C7D2FE' },
  'Consultant Name': { bg: '#F8FAFC', text: '#475569', border: '#E2E8F0' },
  '2nd Reviewer Name': { bg: '#F8FAFC', text: '#475569', border: '#E2E8F0' },
  'Delivered Date': { bg: '#FFF1F2', text: '#9F1239', border: '#FDA4AF' },

  // Feedback Columns
  'XBRL Feedback Quality': { bg: '#F0FDFA', text: '#0F766E', border: '#99F6E4' },
  'XBRL Feedback Status': { bg: '#F0FDF4', text: '#166534', border: '#86EFAC' },
  'XBRL Feedback Reviewed': { bg: '#FDF4FF', text: '#86198F', border: '#F0ABFC' },
  'XBRL Feedback Validity': { bg: '#EFF6FF', text: '#1E40AF', border: '#93C5FD' },
  'XBRL Feedback Comments': { bg: '#FEF3C7', text: '#92400E', border: '#FCD34D' },

  // Secondary Report Details
  'Production Consultant Name': { bg: '#FEF3C7', text: '#92400E', border: '#FCD34D' },
  'Internal Reviewer name': { bg: '#F3E8FF', text: '#6B21A8', border: '#D8B4FE' },
  'External Reviewer Name': { bg: '#EDE9FE', text: '#5B21B6', border: '#C4B5FD' },
  'DEI Errors': { bg: '#FEF2F2', text: '#991B1B', border: '#FCA5A5' },
  'Element Errors': { bg: '#FEF2F2', text: '#991B1B', border: '#FCA5A5' },
  'Value Errors': { bg: '#FEF2F2', text: '#991B1B', border: '#FCA5A5' },
  'Untagged Errors': { bg: '#FEF2F2', text: '#991B1B', border: '#FCA5A5' },
  'L1-L3 Errors': { bg: '#FEF2F2', text: '#991B1B', border: '#FCA5A5' },
  'Reverse Sign Errors': { bg: '#E0F2FE', text: '#0369A1', border: '#7DD3FC' },
  'Alternate Values Errors': { bg: '#E0F2FE', text: '#0369A1', border: '#7DD3FC' },
  'Date Context Errors': { bg: '#FFF7ED', text: '#C2410C', border: '#FDBA74' },
  'Member Context Errors': { bg: '#FFF7ED', text: '#C2410C', border: '#FDBA74' },
  'Scale/Precision Errors': { bg: '#FFF7ED', text: '#C2410C', border: '#FDBA74' },
  'Taxonomy/Structure/Invalid Spi': { bg: '#FDF4FF', text: '#86198F', border: '#F0ABFC' },
  'Incorrect Mapping | Attribute': { bg: '#E0E7FF', text: '#3730A3', border: '#A5B4FC' },
  'Judgmental Errors': { bg: '#FEE2E2', text: '#B91C1C', border: '#F87171' },
  'Validation Errors': { bg: '#FEE2E2', text: '#991B1B', border: '#EF4444' },
  'Other Errors': { bg: '#FEF3C7', text: '#B45309', border: '#FBBF24' },
  'Label Errors': { bg: '#F3E8FF', text: '#7E22CE', border: '#C084FC' },
  'Total Errors': { bg: '#FEE2E2', text: '#991B1B', border: '#EF4444' },
};

export interface StandaloneExportInput {
  matchedRows?: MatchedOutputRow[];
  unmatchedBridgeRows?: MatchedOutputRow[];
  unmatchedJmsRows?: MatchedOutputRow[];
  allRows?: MatchedOutputRow[];
  initialFilter?: 'matched' | 'unmatched_bridge' | 'matched_and_bridge_unmatched' | 'unmatched_jms' | 'all' | string;
  initialView?: 'standard' | 'feedback' | 'qc' | 'all';
  highlightedWorkOrders?: string[];
  bridgeFileName?: string;
  bridgeSheetName?: string;
  jmsFileName?: string;
  jmsSheetName?: string;
  rawBridgeData?: Record<string, any>[];
  rawJmsData?: Record<string, any>[];
  rawProdData?: Record<string, any>[];
}

/**
 * Exports current reconciled rows to a standalone, offline interactive HTML application
 * perfectly matching the live dashboard layout, KPI statistics, rules banner, and controls.
 */
export function exportToStandaloneHTML(
  data: MatchedOutputRow[] | StandaloneExportInput,
  viewLabel = 'Reconciled Records',
  filename = 'Bridge_Work_Orders_Reconciled.html',
  defaultCols: ColumnDefinition[] = STANDARD_15_COLUMNS
): void {
  // Normalize input data
  let matchedRows: MatchedOutputRow[] = [];
  let unmatchedBridgeRows: MatchedOutputRow[] = [];
  let unmatchedJmsRows: MatchedOutputRow[] = [];
  let allRows: MatchedOutputRow[] = [];
  let initialFilter = 'matched';
  let initialView = 'standard';
  let initialHighlightedWos: string[] = [];
  let bridgeFileName = 'Bridge work order - All.xlsx';
  let bridgeSheetName = 'Table 1';
  let jmsFileName = 'Saas_Job_Status_Report.xlsx';
  let jmsSheetName = 'Grabber Day to Day';

  let rawBridge = SAMPLE_BRIDGE_DATA;
  let rawJms = SAMPLE_JMS_DATA;
  let rawProd = SAMPLE_PRODUCTION_DAY_TO_DAY_DATA;

  if (Array.isArray(data)) {
    matchedRows = data.filter((r) => r.matchStatus === 'matched');
    unmatchedBridgeRows = data.filter((r) => r.matchStatus === 'unmatched_bridge');
    unmatchedJmsRows = data.filter((r) => r.matchStatus === 'unmatched_jms');
    allRows = data;
    if (matchedRows.length === 0 && data.length > 0) {
      matchedRows = data;
    }
  } else {
    matchedRows = data.matchedRows || [];
    unmatchedBridgeRows = data.unmatchedBridgeRows || [];
    unmatchedJmsRows = data.unmatchedJmsRows || [];
    allRows = data.allRows || [...matchedRows, ...unmatchedBridgeRows, ...unmatchedJmsRows];
    if (data.initialFilter) initialFilter = data.initialFilter;
    if (data.initialView) initialView = data.initialView;
    if (data.highlightedWorkOrders) initialHighlightedWos = data.highlightedWorkOrders;
    if (data.bridgeFileName) bridgeFileName = data.bridgeFileName;
    if (data.bridgeSheetName) bridgeSheetName = data.bridgeSheetName;
    if (data.jmsFileName) jmsFileName = data.jmsFileName;
    if (data.jmsSheetName) jmsSheetName = data.jmsSheetName;
    if (data.rawBridgeData) rawBridge = data.rawBridgeData as any;
    if (data.rawJmsData) rawJms = data.rawJmsData as any;
    if (data.rawProdData) rawProd = data.rawProdData as any;
  }

  // If matched has 0 but Bridge unmatched has records, default to combined or unmatched so the table isn't blank
  if (matchedRows.length === 0 && unmatchedBridgeRows.length > 0 && initialFilter === 'matched') {
    initialFilter = 'matched_and_bridge_unmatched';
  }

  const combinedMatchedAndBridgeUnmatched = [...matchedRows, ...unmatchedBridgeRows];
  const totalBridge = matchedRows.length + unmatchedBridgeRows.length;
  const matchRate = totalBridge > 0 ? Math.round((matchedRows.length / totalBridge) * 100) : 0;

  // Pre-serialize dataset to JSON safe for script tag
  const datasetJson = JSON.stringify({
    matched: matchedRows,
    unmatched_bridge: unmatchedBridgeRows,
    matched_and_bridge_unmatched: combinedMatchedAndBridgeUnmatched,
    unmatched_jms: unmatchedJmsRows,
    all: allRows,
  }).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');

  const highlightedWosJson = JSON.stringify(initialHighlightedWos).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');

  const sampleDataJson = JSON.stringify({
    bridge: rawBridge,
    jms: rawJms,
    prod: rawProd,
    core5: SCREENSHOT_CORE_5_WORK_ORDERS,
    table1: SCREENSHOT_TABLE1_WORK_ORDERS,
    all15: SCREENSHOT_EXAMPLE_WORK_ORDERS,
  }).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');

  const viewDefinitionsJson = JSON.stringify({
    standard: STANDARD_15_COLUMNS,
    feedback: XBRL_FEEDBACK_COLUMNS,
    qc: QC_IMAGE_COLUMNS,
    all: ALL_COLUMNS,
  }).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');

  const headerColorsJson = JSON.stringify(COLUMN_HEADER_COLORS).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');

  const dateFormatted = new Date().toLocaleDateString('en-US', {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  const fullHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Bridge Work Orders - Reconciled (Offline Dashboard)</title>
  <!-- Embedded SheetJS for 100% offline parsing, reconciliation, and Excel export -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
  <style>
    :root {
      --bg: #F8FAFC;
      --card-bg: #FFFFFF;
      --text-main: #0F172A;
      --text-muted: #64748B;
      --border-color: #CBD5E1;
      --primary: #0284C7;
      --primary-dark: #0369A1;
      --emerald: #059669;
      --amber: #D97706;
      --amber-bg: #FEF3C7;
      --indigo: #4F46E5;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background-color: var(--bg);
      color: var(--text-main);
      padding: 0;
      margin: 0;
      font-size: 13px;
      line-height: 1.4;
    }

    /* Top Navigation Header */
    .top-header {
      background: #FFFFFF;
      border-bottom: 1px solid #E2E8F0;
      position: sticky;
      top: 0;
      z-index: 40;
      box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .header-inner {
      max-width: 1440px;
      margin: 0 auto;
      padding: 10px 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }
    .brand-group {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .brand-icon {
      width: 34px;
      height: 34px;
      border-radius: 8px;
      background: #0F172A;
      color: #FFFFFF;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 800;
      font-size: 14px;
    }
    .brand-title {
      font-size: 15px;
      font-weight: 800;
      color: #0F172A;
      letter-spacing: -0.01em;
    }
    .brand-subtitle {
      font-size: 11.5px;
      color: #64748B;
    }
    .header-badge {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      background: #F0FDF4;
      border: 1px solid #BBF7D0;
      color: #166534;
      font-size: 11px;
      font-weight: 700;
      padding: 3px 8px;
      border-radius: 6px;
    }

    /* Container */
    .main-container {
      max-width: 1440px;
      margin: 0 auto;
      padding: 16px 20px 40px 20px;
    }

    /* Intro Banner */
    .intro-banner {
      background: #FFFFFF;
      border: 1px solid #CBD5E1;
      border-radius: 12px;
      padding: 14px 18px;
      margin-bottom: 16px;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }

    /* KPI Summary Stats Cards */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    .kpi-card {
      background: #FFFFFF;
      border: 2px solid #E2E8F0;
      border-radius: 12px;
      padding: 12px 14px;
      cursor: pointer;
      transition: all 0.15s ease;
      text-align: left;
    }
    .kpi-card:hover {
      border-color: #0284C7;
      transform: translateY(-1px);
      box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }
    .kpi-card.active {
      background: #0F172A;
      color: #FFFFFF;
      border-color: #0F172A;
      box-shadow: 0 4px 10px rgba(15, 23, 42, 0.15);
    }
    .kpi-card-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 4px;
    }
    .kpi-title {
      font-size: 10.5px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }
    .kpi-card.active .kpi-title { color: #93C5FD; }
    .kpi-count-row {
      display: flex;
      align-items: baseline;
      gap: 6px;
      margin: 4px 0 2px 0;
    }
    .kpi-number {
      font-size: 22px;
      font-weight: 800;
      font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
    }
    .kpi-desc {
      font-size: 11px;
      color: #64748B;
    }
    .kpi-card.active .kpi-desc { color: #CBD5E1; }

    /* Specification Rules Banner */
    .rules-card {
      background: #F0FDF4;
      border: 1px solid #BBF7D0;
      border-radius: 10px;
      padding: 10px 14px;
      margin-bottom: 16px;
      font-size: 11.5px;
      color: #166534;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }
    .rules-list {
      display: flex;
      flex-wrap: wrap;
      gap: 14px;
      align-items: center;
    }
    .rule-item {
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }

    /* Input Options Section */
    .input-options-box {
      background: #FFFBEB;
      border: 1px solid #FDE68A;
      border-radius: 12px;
      padding: 14px 16px;
      margin-bottom: 16px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    .input-options-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
      gap: 14px;
      margin-top: 10px;
    }
    .input-subcard {
      background: #FFFFFF;
      border: 1px solid #FDE68A;
      border-radius: 8px;
      padding: 12px;
    }
    .input-subcard-title {
      font-size: 11px;
      font-weight: 800;
      color: #92400E;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      margin-bottom: 8px;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    /* Buttons */
    .btn {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 5px 12px;
      font-size: 11.5px;
      font-weight: 600;
      border-radius: 6px;
      border: 1px solid #CBD5E1;
      background: #FFFFFF;
      color: #334155;
      cursor: pointer;
      transition: all 0.15s ease;
      white-space: nowrap;
    }
    .btn:hover {
      background: #F1F5F9;
      color: #0F172A;
      border-color: #94A3B8;
    }
    .btn-primary {
      background: #0284C7;
      color: #FFFFFF;
      border-color: #0369A1;
    }
    .btn-primary:hover {
      background: #0369A1;
    }
    .btn-emerald {
      background: #059669;
      color: #FFFFFF;
      border-color: #047857;
      font-weight: 700;
    }
    .btn-emerald:hover { background: #047857; }
    .btn-teal {
      background: #F0FDFA;
      color: #0F766E;
      border-color: #99F6E4;
      font-weight: 700;
    }
    .btn-teal:hover { background: #CCFBF1; }
    .btn-amber {
      background: #F59E0B;
      color: #451A03;
      border-color: #D97706;
      font-weight: 700;
    }
    .btn-amber:hover { background: #D97706; color: #FFF; }
    .batch-btn {
      padding: 3px 8px;
      font-size: 11px;
      font-weight: 600;
      background: #FEF3C7;
      color: #78350F;
      border: 1px solid #FCD34D;
      border-radius: 4px;
      cursor: pointer;
    }
    .batch-btn:hover {
      background: #FDE68A;
    }

    /* Main Table Container */
    .table-card {
      background: #FFFFFF;
      border: 2px solid #CBD5E1;
      border-radius: 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.04);
      overflow: hidden;
    }

    /* Filter Tabs Bar */
    .filter-tabs-bar {
      padding: 8px 14px;
      background: #F8FAFC;
      border-bottom: 1px solid #E2E8F0;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }
    .tabs-group {
      display: flex;
      flex-wrap: wrap;
      gap: 4px;
      background: #E2E8F0;
      padding: 3px;
      border-radius: 8px;
    }
    .tab-btn {
      border: none;
      background: transparent;
      padding: 6px 12px;
      font-size: 11.5px;
      font-weight: 700;
      color: #475569;
      border-radius: 6px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.15s ease;
    }
    .tab-btn:hover { color: #0F172A; }
    .tab-btn.active {
      background: #FFFFFF;
      color: #0F172A;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .tab-badge {
      font-size: 10.5px;
      font-weight: 800;
      padding: 1px 6px;
      border-radius: 10px;
      background: #E2E8F0;
      color: #334155;
      font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
    }
    .tab-btn.active .tab-badge {
      background: #EFF6FF;
      color: #1D4ED8;
    }

    /* Report Views Bar */
    .view-switcher-bar {
      padding: 8px 14px;
      background: #F1F5F9;
      border-bottom: 1px solid #CBD5E1;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }
    .view-pills {
      display: inline-flex;
      background: #E2E8F0;
      padding: 2px;
      border-radius: 6px;
      gap: 2px;
    }
    .view-pill {
      border: none;
      background: transparent;
      padding: 4px 10px;
      font-size: 11px;
      font-weight: 700;
      color: #475569;
      border-radius: 5px;
      cursor: pointer;
      transition: all 0.15s ease;
    }
    .view-pill:hover { color: #0F172A; }
    .view-pill.active {
      background: #FFFFFF;
      color: #0F172A;
      box-shadow: 0 1px 2px rgba(0,0,0,0.08);
    }

    /* Toolbar */
    .toolbar {
      padding: 10px 14px;
      background: #FFFFFF;
      border-bottom: 1px solid #E2E8F0;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }
    .search-box {
      position: relative;
      min-width: 280px;
      max-width: 400px;
      flex: 1;
    }
    .search-input {
      width: 100%;
      padding: 6px 28px 6px 10px;
      font-size: 12px;
      border: 1px solid #CBD5E1;
      border-radius: 6px;
      outline: none;
    }
    .search-input:focus {
      border-color: #0284C7;
      box-shadow: 0 0 0 2px rgba(2, 132, 199, 0.15);
    }
    .clear-search-btn {
      position: absolute;
      right: 8px;
      top: 50%;
      transform: translateY(-50%);
      background: none;
      border: none;
      color: #94A3B8;
      cursor: pointer;
      font-size: 14px;
      display: none;
    }

    /* Table Scroll Viewport */
    .table-container {
      width: 100%;
      overflow-x: auto;
      max-height: 650px;
      position: relative;
    }
    table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 12px;
      text-align: left;
    }
    th {
      position: sticky;
      top: 0;
      z-index: 20;
      font-size: 11px;
      font-weight: 700;
      padding: 8px 10px;
      white-space: nowrap;
      cursor: pointer;
      user-select: none;
      border-right: 1px solid #E2E8F0;
    }
    th:hover { opacity: 0.9; }
    th .sort-arrow { font-size: 9px; margin-left: 4px; color: #64748B; }
    td {
      padding: 7px 10px;
      border-bottom: 1px solid #E2E8F0;
      border-right: 1px solid #F1F5F9;
      white-space: nowrap;
      color: #1E293B;
    }
    tr:nth-child(even) td { background-color: #F8FAFC; }
    tr:hover td { background-color: #F1F5F9; }

    /* Sticky Columns */
    th.sticky-col-idx, td.sticky-col-idx {
      position: sticky;
      left: 0;
      z-index: 21;
      background-color: #FFFFFF;
      width: 44px;
      text-align: center;
      border-right: 1px solid #CBD5E1;
    }
    th.sticky-col-wo, td.sticky-col-wo {
      position: sticky;
      left: 44px;
      z-index: 21;
      background-color: #FFFFFF;
      font-weight: 700;
      font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
      border-right: 2px solid #CBD5E1;
    }
    tr:nth-child(even) td.sticky-col-idx,
    tr:nth-child(even) td.sticky-col-wo {
      background-color: #F8FAFC;
    }
    tr:hover td.sticky-col-idx,
    tr:hover td.sticky-col-wo {
      background-color: #F1F5F9;
    }

    /* Yellow Highlight Row */
    tr.row-highlighted td {
      background-color: #FEF08A !important;
      color: #451A03 !important;
      border-bottom-color: #FCD34D !important;
    }
    tr.row-highlighted td.sticky-col-idx,
    tr.row-highlighted td.sticky-col-wo {
      background-color: #FDE047 !important;
      color: #78350F !important;
    }
    .badge-highlighted {
      background-color: #F59E0B;
      color: #451A03;
      font-size: 9px;
      font-weight: 800;
      padding: 1px 4px;
      border-radius: 4px;
      margin-left: 6px;
      display: inline-block;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }

    /* Custom Badges */
    .badge {
      display: inline-block;
      padding: 2px 7px;
      font-size: 11px;
      font-weight: 700;
      border-radius: 12px;
      white-space: nowrap;
    }
    .badge-teal { background: #CCFBF1; color: #0F766E; }
    .badge-indigo { background: #E0E7FF; color: #3730A3; }
    .badge-amber { background: #FEF3C7; color: #92400E; }

    /* Compact mode */
    table.compact th, table.compact td {
      padding: 4px 6px;
      font-size: 11px;
    }

    /* Footer & Pagination */
    .footer-bar {
      padding: 10px 14px;
      background: #F8FAFC;
      border-top: 1px solid #E2E8F0;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      font-size: 11.5px;
      color: #64748B;
    }
    .page-btn {
      padding: 3px 8px;
      font-size: 11.5px;
      border: 1px solid #CBD5E1;
      background: #FFFFFF;
      border-radius: 4px;
      cursor: pointer;
    }
    .page-btn.active {
      background: #0F172A;
      color: #FFFFFF;
      border-color: #0F172A;
      font-weight: 700;
    }

    /* Column dropdown */
    .dropdown { position: relative; display: inline-block; }
    .dropdown-menu {
      display: none;
      position: absolute;
      right: 0;
      top: 100%;
      background: #FFFFFF;
      border: 1px solid #CBD5E1;
      border-radius: 8px;
      padding: 8px;
      min-width: 240px;
      max-height: 380px;
      overflow-y: auto;
      box-shadow: 0 4px 12px rgba(0,0,0,0.12);
      z-index: 50;
    }
    .dropdown.open .dropdown-menu { display: block; }
    .column-check-item {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 4px 6px;
      font-size: 11px;
      cursor: pointer;
      border-radius: 4px;
    }
    .column-check-item:hover { background: #F1F5F9; }

    /* Modals */
    .modal-overlay {
      display: none;
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(15, 23, 42, 0.5);
      backdrop-filter: blur(2px);
      z-index: 100;
      align-items: center;
      justify-content: center;
    }
    .modal-overlay.open { display: flex; }
    .modal-content {
      background: #FFFFFF;
      border-radius: 12px;
      width: 90%;
      max-width: 600px;
      max-height: 85vh;
      overflow: hidden;
      box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1);
      display: flex;
      flex-col: column;
    }
    .modal-header {
      padding: 12px 18px;
      background: #0F172A;
      color: #FFFFFF;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    .modal-close {
      background: none;
      border: none;
      color: #CBD5E1;
      font-size: 20px;
      cursor: pointer;
    }
    .modal-body {
      padding: 16px;
      overflow-y: auto;
    }

    @media print {
      .top-header, .intro-banner, .input-options-box, .filter-tabs-bar, .view-switcher-bar, .toolbar, .footer-bar {
        display: none !important;
      }
      .table-container { max-height: none !important; overflow: visible !important; }
      th, td { border: 1px solid #CBD5E1 !important; }
    }
  </style>
</head>
<body>

  <!-- Top Navigation Header matching Live Dashboard -->
  <header class="top-header">
    <div class="header-inner">
      <div class="brand-group">
        <div class="brand-icon">X</div>
        <div>
          <div class="brand-title">Bridge Work Orders - All</div>
          <div class="brand-subtitle">Reconciliation &amp; Formatting Tool (Offline Self-Contained App)</div>
        </div>
      </div>
      <div style="display: flex; align-items: center; gap: 8px;">
        <span class="header-badge">&#10003; Offline Ready</span>
        <button type="button" class="btn btn-emerald" id="btn-top-download-xlsx" title="Download Excel (.xlsx) of current view">
          Download Excel (.xlsx)
        </button>
        <button type="button" class="btn" id="btn-top-reload-sample">
          Reload Demo Data
        </button>
      </div>
    </div>
  </header>

  <main class="main-container">

    <!-- Intro Banner -->
    <div class="intro-banner">
      <div>
        <div style="display: flex; align-items: center; gap: 8px;">
          <h2 style="font-size: 14px; font-weight: 800; color: #0F172A;">Bridge Work Orders &amp; JMS SaaS Tracker Reconciliation</h2>
          <span style="font-size: 10px; font-weight: 800; background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; padding: 2px 6px; border-radius: 4px;">
            16 Mapped Columns (incl. XBRL Feedback Comments)
          </span>
        </div>
        <p style="font-size: 11.5px; color: #64748B; margin-top: 3px;">
          Compare Bridge Work Orders and JMS/SaaS workbooks to match orders with automated transformations,
          Production Day to Day cross-referencing, yellow Work Order highlighting, and instant Excel export.
        </p>
      </div>

      <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
        <!-- Matched + Bridge Unmatched Combined Excel Download -->
        <button type="button" class="btn btn-teal" id="btn-top-export-matched-bridge-xlsx" title="Download Excel (.xlsx) combining Matched and Bridge Unmatched work orders (100% of Base Bridge Table 1)">
          <span>Download Matched + Bridge Unmatched (.xlsx)</span>
          <span style="font-family: monospace; font-size: 11px; background: #CCFBF1; color: #115E59; padding: 1px 5px; border-radius: 6px;" id="badge-top-combined-count">${totalBridge}</span>
        </button>

        <button type="button" class="btn btn-emerald" id="btn-intro-export-xlsx" title="Download Excel (.xlsx)">
          Download Excel
        </button>
      </div>
    </div>

    <!-- KPI Summary Statistics Filter Cards matching Live Dashboard -->
    <div class="kpi-grid">
      <!-- 1. Matched -->
      <div class="kpi-card" data-filter="matched" id="kpi-card-matched">
        <div class="kpi-card-header">
          <span class="kpi-title" style="color: #059669;">Matched Work Orders</span>
          <span style="width: 8px; height: 8px; border-radius: 50%; background: #10B981;"></span>
        </div>
        <div class="kpi-count-row">
          <span class="kpi-number" id="kpi-matched-count">${matchedRows.length}</span>
          <span class="kpi-desc" id="kpi-match-rate">(${matchRate}%)</span>
        </div>
        <div class="kpi-desc">Fully reconciled between Bridge &amp; JMS</div>
      </div>

      <!-- 2. Bridge Unmatched -->
      <div class="kpi-card" data-filter="unmatched_bridge" id="kpi-card-unmatched-bridge">
        <div class="kpi-card-header">
          <span class="kpi-title" style="color: #D97706;">Bridge Unmatched</span>
          <span style="width: 8px; height: 8px; border-radius: 50%; background: #F59E0B;"></span>
        </div>
        <div class="kpi-count-row">
          <span class="kpi-number" id="kpi-unmatched-bridge-count">${unmatchedBridgeRows.length}</span>
          <span class="kpi-desc" id="kpi-bridge-total-desc">of ${totalBridge} Bridge</span>
        </div>
        <div class="kpi-desc">In Bridge Table 1 but missing in JMS</div>
      </div>

      <!-- 3. Matched + Bridge Unmatched Combined -->
      <div class="kpi-card" data-filter="matched_and_bridge_unmatched" id="kpi-card-matched-bridge-unmatched">
        <div class="kpi-card-header">
          <span class="kpi-title" style="color: #0F766E;">Matched + Bridge Unmatched</span>
          <span style="width: 8px; height: 8px; border-radius: 50%; background: #14B8A6;"></span>
        </div>
        <div class="kpi-count-row">
          <span class="kpi-number" id="kpi-matched-bridge-count">${combinedMatchedAndBridgeUnmatched.length}</span>
          <span class="kpi-desc">(100% Base Bridge)</span>
        </div>
        <div class="kpi-desc">Complete Bridge Table 1 order universe</div>
      </div>

      <!-- 4. JMS Unmatched -->
      <div class="kpi-card" data-filter="unmatched_jms" id="kpi-card-unmatched-jms">
        <div class="kpi-card-header">
          <span class="kpi-title" style="color: #6366F1;">JMS Unmatched</span>
          <span style="width: 8px; height: 8px; border-radius: 50%; background: #818CF8;"></span>
        </div>
        <div class="kpi-count-row">
          <span class="kpi-number" id="kpi-unmatched-jms-count">${unmatchedJmsRows.length}</span>
        </div>
        <div class="kpi-desc">In JMS Tracker without Bridge match</div>
      </div>

      <!-- 5. All Records -->
      <div class="kpi-card" data-filter="all" id="kpi-card-all">
        <div class="kpi-card-header">
          <span class="kpi-title" style="color: #475569;">All Records</span>
          <span style="width: 8px; height: 8px; border-radius: 50%; background: #94A3B8;"></span>
        </div>
        <div class="kpi-count-row">
          <span class="kpi-number" id="kpi-all-count">${allRows.length}</span>
        </div>
        <div class="kpi-desc">Combined universe of all work orders</div>
      </div>
    </div>

    <!-- Active Specification Rules Banner -->
    <div class="rules-card">
      <div class="rules-list">
        <strong style="color: #065F46;">Active Specification Rules:</strong>
        <span class="rule-item">&#10003; Sheet: <strong>Table 1</strong> (Sheet1 strictly ignored)</span>
        <span class="rule-item">&#10003; FB Quality: <strong>Cleaned</strong> (leading numbers like "4 - " removed)</span>
        <span class="rule-item">&#10003; Sampling Quality: <strong>"STP Feedback"</strong> prefill</span>
        <span class="rule-item">&#10003; Form Type: Grabber Day to Day / Job Type</span>
        <span class="rule-item">&#10003; Final Column: <strong>XBRL Feedback Comments</strong></span>
      </div>
      <button type="button" class="btn" id="btn-toggle-input-options" style="font-size: 11px; padding: 2px 8px;">
        Toggle Input Controls
      </button>
    </div>

    <!-- Collapsible Interactive Input Options (Work Orders & File Uploads) -->
    <div class="input-options-box" id="standalone-input-options-container">
      <div style="display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #FDE68A; padding-bottom: 8px; margin-bottom: 8px;">
        <div style="display: flex; align-items: center; gap: 8px;">
          <strong style="color: #78350F; font-size: 12px; text-transform: uppercase;">Interactive Input Options (Files &amp; Work Orders)</strong>
          <span style="font-size: 10px; font-weight: 700; background: #FEF3C7; color: #92400E; border: 1px solid #FCD34D; padding: 1px 6px; border-radius: 10px;">Offline Engine Ready</span>
        </div>
        <div id="reconcile-status" style="font-size: 11px; color: #047857; font-weight: 700;">Ready</div>
      </div>

      <div class="input-options-grid">
        <!-- 1. Work Order Highlight Input Option -->
        <div class="input-subcard">
          <div class="input-subcard-title">
            <span>Work Order Input Option (Highlight in Yellow)</span>
            <button type="button" id="btn-open-modal-from-bar" class="batch-btn" style="background: #FFFFFF;">Paste / Modal...</button>
          </div>

          <div style="display: flex; align-items: center; gap: 5px; flex-wrap: wrap; margin-bottom: 8px;">
            <span style="font-size: 11px; color: #78350F; font-weight: 600;">Load Screenshot IDs:</span>
            <button type="button" id="btn-quick-core5" class="batch-btn">+ Core 5 IDs</button>
            <button type="button" id="btn-quick-table1" class="batch-btn">+ Table 1 (20 IDs)</button>
            <button type="button" id="btn-quick-all15" class="batch-btn">+ All Screenshot (15 IDs)</button>
          </div>

          <div style="display: flex; align-items: center; gap: 6px;">
            <input type="text" id="inline-wo-input" placeholder="Type or paste Work Order IDs (e.g. WO-26-003172)..."
                   style="flex: 1; font-size: 11.5px; padding: 6px 8px; border: 1px solid #CBD5E1; border-radius: 6px; font-family: monospace;" />
            <button type="button" id="btn-add-inline-wo" class="btn btn-amber">+ Highlight ID</button>
            <button type="button" id="btn-clear-inline-wos" class="btn" style="color: #DC2626;">Clear</button>
          </div>

          <div id="inline-highlight-status" style="margin-top: 6px; font-size: 11px; color: #78350F; display: none; align-items: center; justify-content: space-between;">
            <span><strong id="inline-highlight-count">0</strong> IDs highlighted in yellow</span>
            <label style="display: flex; align-items: center; gap: 4px; font-weight: 600; cursor: pointer;">
              <input type="checkbox" id="check-inline-only-highlighted" style="cursor: pointer;" />
              <span>Show Highlighted Only</span>
            </label>
          </div>
        </div>

        <!-- 2. File Upload Input Options -->
        <div class="input-subcard">
          <div class="input-subcard-title">
            <span>File Input Options (Upload &amp; Reconcile Offline)</span>
            <button type="button" id="btn-standalone-reload-sample" class="batch-btn" style="background: #FFFFFF;">Reset Demo Data</button>
          </div>

          <!-- Bridge File Input -->
          <div style="border: 1px dashed #CBD5E1; border-radius: 6px; padding: 8px 10px; background: #F8FAFC; margin-bottom: 6px; display: flex; align-items: center; justify-content: space-between; gap: 8px;">
            <div style="min-width: 0;">
              <strong style="color: #0F172A; display: block; font-size: 11.5px;" id="label-bridge-filename">1. Bridge: ${bridgeFileName}</strong>
              <span style="color: #0F766E; font-size: 10.5px; font-family: monospace;" id="label-bridge-sheet">Sheet: ${bridgeSheetName}</span>
            </div>
            <div style="display: flex; align-items: center; gap: 4px;">
              <select id="select-bridge-sheet" style="font-size: 11px; padding: 3px 6px; border: 1px solid #CBD5E1; border-radius: 4px; display: none;"></select>
              <input type="file" id="standalone-file-bridge" accept=".xlsx,.xls,.csv" style="display: none;" />
              <button type="button" class="btn" onclick="document.getElementById('standalone-file-bridge').click()" style="font-size: 11px; padding: 4px 8px;">Change File</button>
            </div>
          </div>

          <!-- JMS File Input -->
          <div style="border: 1px dashed #CBD5E1; border-radius: 6px; padding: 8px 10px; background: #F8FAFC; display: flex; align-items: center; justify-content: space-between; gap: 8px;">
            <div style="min-width: 0;">
              <strong style="color: #0F172A; display: block; font-size: 11.5px;" id="label-jms-filename">2. JMS: ${jmsFileName}</strong>
              <span style="color: #0369A1; font-size: 10.5px; font-family: monospace;" id="label-jms-sheet">Sheet: ${jmsSheetName}</span>
            </div>
            <div style="display: flex; align-items: center; gap: 4px;">
              <select id="select-jms-sheet" style="font-size: 11px; padding: 3px 6px; border: 1px solid #CBD5E1; border-radius: 4px; display: none;"></select>
              <input type="file" id="standalone-file-jms" accept=".xlsx,.xls,.csv" style="display: none;" />
              <button type="button" class="btn" onclick="document.getElementById('standalone-file-jms').click()" style="font-size: 11px; padding: 4px 8px;">Change File</button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Main Results Table Card matching Live Dashboard -->
    <div class="table-card" id="results-card">

      <!-- Filter Tabs Bar (Matched / Bridge Unmatched / Matched+Bridge Unmatched / JMS Unmatched / All) -->
      <div class="filter-tabs-bar">
        <div class="tabs-group" id="filter-tabs">
          <button type="button" class="tab-btn" data-filter="matched">
            <span>Matched Records</span>
            <span class="tab-badge" id="badge-matched">${matchedRows.length}</span>
          </button>
          <button type="button" class="tab-btn" data-filter="unmatched_bridge">
            <span>Bridge Unmatched</span>
            <span class="tab-badge" id="badge-unmatched-bridge">${unmatchedBridgeRows.length}</span>
          </button>
          <button type="button" class="tab-btn" data-filter="matched_and_bridge_unmatched" title="Combined view of Matched and Bridge Unmatched work orders (100% of Base Bridge Table 1 records)">
            <span>Matched + Bridge Unmatched</span>
            <span class="tab-badge" id="badge-matched-bridge-unmatched">${combinedMatchedAndBridgeUnmatched.length}</span>
          </button>
          <button type="button" class="tab-btn" data-filter="unmatched_jms">
            <span>JMS Unmatched</span>
            <span class="tab-badge" id="badge-unmatched-jms">${unmatchedJmsRows.length}</span>
          </button>
          <button type="button" class="tab-btn" data-filter="all">
            <span>All Records</span>
            <span class="tab-badge" id="badge-all">${allRows.length}</span>
          </button>
        </div>

        <div style="font-size: 11.5px; color: #64748B;">
          Active Filter: <strong id="status-label" style="color: #0F172A;">Matched Records</strong>
        </div>
      </div>

      <!-- Report View Switcher (16 Headings / XBRL Feedback / QC Errors / All Columns) -->
      <div class="view-switcher-bar">
        <div style="display: flex; align-items: center; gap: 8px;">
          <span style="font-weight: 700; color: #334155; font-size: 11px;">Report View:</span>
          <div class="view-pills" id="view-pills">
            <button type="button" class="view-pill" data-view="standard">Standard Headings (16 Cols)</button>
            <button type="button" class="view-pill" data-view="feedback">XBRL Feedback Audit</button>
            <button type="button" class="view-pill" data-view="qc">QC &amp; Secondary Details (24 Cols)</button>
            <button type="button" class="view-pill" data-view="all">All Columns View</button>
          </div>
        </div>

        <div style="font-size: 11.5px; color: #64748B;">
          Showing <strong id="active-columns-count" style="color: #0F172A;">16</strong> columns
        </div>
      </div>

      <!-- Main Toolbar -->
      <div class="toolbar">
        <div class="search-box">
          <input type="text" id="search-input" class="search-input" placeholder="Search work order, X Track ID, reviewer, client..." />
          <button type="button" id="clear-search" class="clear-search-btn" title="Clear search">&times;</button>
        </div>

        <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap;">
          <!-- Highlight Work Orders Button -->
          <button type="button" class="btn btn-amber" id="btn-open-highlight-modal" title="Highlight specific Work Orders in yellow">
            <span>Highlight Work Orders</span>
            <span id="badge-highlight-count" style="display: none; background: #92400E; color: #FFF; font-size: 10px; font-weight: 700; padding: 1px 5px; border-radius: 10px;">0</span>
          </button>

          <label id="container-only-highlighted" style="display: none; align-items: center; gap: 4px; font-size: 11.5px; font-weight: 700; color: #78350F; background: #FFFBEB; padding: 4px 8px; border-radius: 6px; border: 1px solid #FDE68A; cursor: pointer;">
            <input type="checkbox" id="check-only-highlighted" style="cursor: pointer;">
            <span>Show Highlighted Only</span>
          </label>

          <!-- Compact Fit Toggle -->
          <button type="button" class="btn" id="btn-compact-toggle" title="Fit all columns to screen">
            <span>Fit to Screen</span>
          </button>

          <!-- Column Visibility Dropdown -->
          <div class="dropdown" id="dropdown-columns">
            <button type="button" class="btn" id="btn-columns-menu">
              <span>Columns...</span>
            </button>
            <div class="dropdown-menu" id="columns-menu-list"></div>
          </div>

          <!-- Copy TSV for Excel -->
          <button type="button" class="btn" id="btn-copy-tsv" title="Copy visible rows to clipboard as TSV format for instant Excel paste">
            <span>Copy TSV</span>
          </button>

          <!-- CSV Export -->
          <button type="button" class="btn" id="btn-export-csv" title="Download visible rows as CSV">
            <span>CSV</span>
          </button>

          <!-- Dedicated Matched + Bridge Unmatched Excel Export (.xlsx) -->
          <button type="button" class="btn btn-teal" id="btn-export-matched-bridge-xlsx" title="Download Excel (.xlsx) combining Matched and Bridge Unmatched work orders">
            <span>Download Matched + Bridge Unmatched (.xlsx)</span>
            <span style="font-family: monospace; font-size: 10.5px; background: #CCFBF1; color: #115E59; padding: 1px 6px; border-radius: 8px;" id="badge-table-combined-count">${totalBridge}</span>
          </button>

          <!-- Primary Excel Export (.xlsx) -->
          <button type="button" class="btn btn-emerald" id="btn-export-xlsx" title="Download Excel workbook (.xlsx)">
            <span>Download Excel (.xlsx)</span>
          </button>

          <!-- Print / PDF -->
          <button type="button" class="btn" id="btn-print" title="Print or save as PDF">
            <span>Print</span>
          </button>
        </div>
      </div>

      <!-- Table Scroll Viewport -->
      <div class="table-container">
        <table id="main-table">
          <thead id="table-head"></thead>
          <tbody id="table-body"></tbody>
        </table>
      </div>

      <!-- Footer Bar & Pagination -->
      <div class="footer-bar">
        <div style="display: flex; align-items: center; gap: 12px;">
          <span>Showing <strong id="row-range-start">1</strong> to <strong id="row-range-end">0</strong> of <strong id="row-total-count">0</strong> records</span>
          <div style="display: flex; align-items: center; gap: 4px;">
            <span>Per page:</span>
            <select id="page-size-select" style="font-size: 11px; padding: 2px 6px; border: 1px solid #CBD5E1; border-radius: 4px;">
              <option value="10">10</option>
              <option value="25" selected>25</option>
              <option value="50">50</option>
              <option value="100">100</option>
              <option value="99999">All</option>
            </select>
          </div>
        </div>

        <div class="pagination-controls" id="pagination-controls"></div>
      </div>
    </div>
  </main>

  <!-- Row Inspector Modal -->
  <div class="modal-overlay" id="inspect-modal">
    <div class="modal-content">
      <div class="modal-header">
        <div>
          <h3 id="modal-title" style="margin: 0; font-size: 14px;">Record Details</h3>
          <span style="font-size: 11px; color: #94A3B8;" id="modal-subtitle">Work Order Audit</span>
        </div>
        <button type="button" class="modal-close" id="modal-close-btn">&times;</button>
      </div>
      <div class="modal-body" id="modal-body"></div>
    </div>
  </div>

  <!-- Work Order Highlight Modal -->
  <div class="modal-overlay" id="highlight-modal">
    <div class="modal-content" style="max-width: 520px;">
      <div class="modal-header" style="background: linear-gradient(to right, #F59E0B, #EAB308); color: #FFF;">
        <div>
          <h3 style="color: #FFF; margin: 0; font-size: 14px;">Highlight Work Orders in Yellow</h3>
          <span style="font-size: 11px; color: #FEF3C7;">Enter Work Order IDs to highlight in this report</span>
        </div>
        <button type="button" class="modal-close" id="highlight-modal-close" style="color: #FFF;">&times;</button>
      </div>
      <div class="modal-body" style="padding: 16px;">
        <div style="background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 8px; padding: 12px; margin-bottom: 12px;">
          <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; gap: 8px; flex-wrap: wrap;">
            <div>
              <strong style="font-size: 11.5px; color: #92400E; display: block;">Load Screenshot IDs:</strong>
              <span style="font-size: 10px; color: #B45309;">Select a batch to load multiple Work Order IDs:</span>
            </div>
            <label style="font-size: 11px; font-weight: 600; color: #92400E; display: flex; align-items: center; gap: 4px; cursor: pointer;">
              <input type="checkbox" id="check-append-highlights" style="margin: 0;">
              <span>Append to existing list</span>
            </label>
          </div>
          <div style="display: flex; gap: 6px; flex-wrap: wrap;">
            <button type="button" class="batch-btn" id="btn-load-sample-ids">All Screenshot IDs (15 IDs)</button>
            <button type="button" class="batch-btn" id="btn-load-core-ids">Core 5 IDs</button>
            <button type="button" class="batch-btn" id="btn-load-table1-ids">Table 1 (20 IDs)</button>
          </div>
        </div>

        <label style="display: block; font-size: 11.5px; font-weight: 700; margin-bottom: 4px; color: #1E293B;">
          Paste Work Order IDs (one per line or separated by commas / tabs):
        </label>
        <textarea id="highlight-textarea" rows="7" style="width: 100%; font-family: monospace; font-size: 11.5px; padding: 8px; border: 1px solid #CBD5E1; border-radius: 6px; resize: vertical;" placeholder="WO-26-003172&#10;WO-26-003125&#10;WO-26-003276&#10;WO-26-003627&#10;WO-26-003535"></textarea>
      </div>
      <div style="padding: 12px 16px; background: #F8FAFC; border-top: 1px solid #E2E8F0; display: flex; align-items: center; justify-content: space-between;">
        <button type="button" class="btn" id="btn-clear-highlights" style="color: #DC2626;">Clear</button>
        <div style="display: flex; gap: 8px;">
          <button type="button" class="btn" id="btn-cancel-highlights">Cancel</button>
          <button type="button" class="btn btn-amber" id="btn-apply-highlights">Apply Highlight</button>
        </div>
      </div>
    </div>
  </div>

  <script>
    // Embedded datasets & configurations
    let DATASETS = ${datasetJson};
    const VIEW_DEFS = ${viewDefinitionsJson};
    const HEADER_COLORS = ${headerColorsJson};
    const INITIAL_HIGHLIGHTED = ${highlightedWosJson};
    const SAMPLE_RAW_DATA = ${sampleDataJson};

    let currentBridgeData = SAMPLE_RAW_DATA.bridge;
    let currentJmsData = SAMPLE_RAW_DATA.jms;
    let currentProdData = SAMPLE_RAW_DATA.prod;
    let currentBridgeWorkbook = null;
    let currentJmsWorkbook = null;

    function normalizeWo(val) {
      if (val === null || val === undefined) return '';
      let str = String(val).trim();
      str = str.replace(/[\\u200B-\\u200D\\uFEFF]/g, '').trim();
      return str.toUpperCase();
    }

    function cleanFbQuality(val) {
      if (val === null || val === undefined) return '';
      return String(val).replace(/^\\d+\\s*[-–—:.]\\s*/, '').trim();
    }

    // Application state
    let state = {
      filter: '${initialFilter}',
      view: '${initialView}',
      search: '',
      sortKey: 'workOrderNo',
      sortAsc: true,
      compact: false,
      pageSize: 25,
      page: 1,
      hiddenColumns: new Set(),
      inspectingRow: null,
      highlightedWos: new Set(INITIAL_HIGHLIGHTED.map(normalizeWo)),
      onlyHighlighted: false,
    };

    // DOM Elements
    const filterTabs = document.getElementById('filter-tabs');
    const viewPills = document.getElementById('view-pills');
    const searchInput = document.getElementById('search-input');
    const clearSearch = document.getElementById('clear-search');
    const statusLabel = document.getElementById('status-label');
    const activeColsCount = document.getElementById('active-columns-count');
    const tableHead = document.getElementById('table-head');
    const tableBody = document.getElementById('table-body');
    const mainTable = document.getElementById('main-table');
    const btnCompact = document.getElementById('btn-compact-toggle');
    const dropdownCols = document.getElementById('dropdown-columns');
    const btnColsMenu = document.getElementById('btn-columns-menu');
    const colsMenuList = document.getElementById('columns-menu-list');
    const btnCopyTsv = document.getElementById('btn-copy-tsv');
    const btnExportCsv = document.getElementById('btn-export-csv');
    const btnExportXlsx = document.getElementById('btn-export-xlsx');
    const btnIntroExportXlsx = document.getElementById('btn-intro-export-xlsx');
    const btnTopDownloadXlsx = document.getElementById('btn-top-download-xlsx');
    const btnPrint = document.getElementById('btn-print');
    const pageSizeSelect = document.getElementById('page-size-select');
    const paginationControls = document.getElementById('pagination-controls');
    const rowRangeStart = document.getElementById('row-range-start');
    const rowRangeEnd = document.getElementById('row-range-end');
    const rowTotalCount = document.getElementById('row-total-count');
    const inspectModal = document.getElementById('inspect-modal');
    const modalCloseBtn = document.getElementById('modal-close-btn');

    // Highlight controls
    const btnOpenHighlightModal = document.getElementById('btn-open-highlight-modal');
    const badgeHighlightCount = document.getElementById('badge-highlight-count');
    const containerOnlyHighlighted = document.getElementById('container-only-highlighted');
    const checkOnlyHighlighted = document.getElementById('check-only-highlighted');
    const highlightModal = document.getElementById('highlight-modal');
    const highlightModalClose = document.getElementById('highlight-modal-close');
    const highlightTextarea = document.getElementById('highlight-textarea');
    const btnLoadSampleIds = document.getElementById('btn-load-sample-ids');
    const btnLoadCoreIds = document.getElementById('btn-load-core-ids');
    const btnLoadTable1Ids = document.getElementById('btn-load-table1-ids');
    const checkAppendHighlights = document.getElementById('check-append-highlights');
    const btnClearHighlights = document.getElementById('btn-clear-highlights');
    const btnCancelHighlights = document.getElementById('btn-cancel-highlights');
    const btnApplyHighlights = document.getElementById('btn-apply-highlights');

    // Quick bar controls
    const btnQuickCore5 = document.getElementById('btn-quick-core5');
    const btnQuickTable1 = document.getElementById('btn-quick-table1');
    const btnQuickAll15 = document.getElementById('btn-quick-all15');
    const inlineWoInput = document.getElementById('inline-wo-input');
    const btnAddInlineWo = document.getElementById('btn-add-inline-wo');
    const btnClearInlineWos = document.getElementById('btn-clear-inline-wos');
    const checkInlineOnlyHighlighted = document.getElementById('check-inline-only-highlighted');
    const btnOpenModalFromBar = document.getElementById('btn-open-modal-from-bar');
    const btnToggleInputOptions = document.getElementById('btn-toggle-input-options');
    const inputOptionsContainer = document.getElementById('standalone-input-options-container');

    const ALL_SCREENSHOT_IDS = [
      'WO-26-003172', 'WO-26-003125', 'WO-26-003276', 'WO-26-003627', 'WO-26-003535',
      'WO-26-003361', 'WO-26-003537', 'WO-26-003558', 'WO-26-003623', 'WO-26-003311',
      'WO-26-003612', 'WO-26-003404', 'WO-26-003647', 'WO-26-003023', 'WO-26-003632'
    ];
    const CORE_5_IDS = [
      'WO-26-003172', 'WO-26-003125', 'WO-26-003276', 'WO-26-003627', 'WO-26-003535'
    ];
    const TABLE1_IDS = [
      'WO-26-003361', 'WO-26-003537', 'WO-26-003558', 'WO-26-003623', 'WO-26-003119',
      'WO-26-003226', 'WO-26-003336', 'WO-26-003346', 'WO-26-003395', 'WO-26-003405',
      'WO-26-003413', 'WO-26-003427', 'WO-26-003428', 'WO-26-003458', 'WO-26-003461',
      'WO-26-003311', 'WO-26-003612', 'WO-26-003647', 'WO-26-003023', 'WO-26-003632'
    ];

    // Get active column definitions
    function getActiveViewColumns() {
      const allCols = VIEW_DEFS[state.view] || VIEW_DEFS.standard;
      return allCols.filter(col => !state.hiddenColumns.has(col.key));
    }

    // Filter and sort rows
    function getProcessedRows() {
      let rows = DATASETS[state.filter] || [];

      // If active filter is empty, fallback to matched_and_bridge_unmatched or all
      if (rows.length === 0 && (DATASETS.matched_and_bridge_unmatched || []).length > 0 && state.filter === 'matched') {
        state.filter = 'matched_and_bridge_unmatched';
        rows = DATASETS[state.filter] || [];
      }

      // Highlight focus filter
      if (state.onlyHighlighted && state.highlightedWos.size > 0) {
        rows = rows.filter(r => state.highlightedWos.has(normalizeWo(r.workOrderNo)));
      }

      // Search filter
      const q = state.search.trim().toLowerCase();
      if (q) {
        rows = rows.filter(r => {
          return Object.values(r).some(val => {
            if (val === null || val === undefined) return false;
            return String(val).toLowerCase().includes(q);
          });
        });
      }

      // Sort
      if (state.sortKey) {
        rows = [...rows].sort((a, b) => {
          let va = a[state.sortKey];
          let vb = b[state.sortKey];
          if (va === undefined || va === null) va = '';
          if (vb === undefined || vb === null) vb = '';
          if (typeof va === 'number' && typeof vb === 'number') {
            return state.sortAsc ? va - vb : vb - va;
          }
          const sa = String(va).toLowerCase();
          const sb = String(vb).toLowerCase();
          return state.sortAsc ? sa.localeCompare(sb) : sb.localeCompare(sa);
        });
      }

      return rows;
    }

    // Render Table Header
    function renderTableHeader() {
      const cols = getActiveViewColumns();
      activeColsCount.innerText = cols.length;

      let html = '<tr>';
      html += '<th class="sticky-col-idx" style="border-bottom: 2px solid #CBD5E1;">#</th>';

      cols.forEach(col => {
        const color = HEADER_COLORS[col.header] || { bg: '#F8FAFC', text: '#1E293B', border: '#E2E8F0' };
        const isStickyWo = col.key === 'workOrderNo';
        const isSorted = state.sortKey === col.key;
        const arrow = isSorted ? (state.sortAsc ? ' ▲' : ' ▼') : '';

        const stickyClass = isStickyWo ? 'sticky-col-wo' : '';
        html += '<th class="' + stickyClass + '" style="background-color: ' + color.bg + '; color: ' + color.text + '; border-bottom: 2px solid ' + color.border + ';" onclick="handleSort(\\'' + col.key + '\\')">' +
          '<span>' + escapeHtml(col.header) + '</span>' +
          '<span class="sort-arrow">' + arrow + '</span>' +
        '</th>';
      });

      html += '<th style="background-color: #F1F5F9; color: #475569; border-bottom: 2px solid #CBD5E1; text-align: center; width: 60px;">Details</th>';
      html += '</tr>';

      tableHead.innerHTML = html;
    }

    // Render Table Body
    function renderTableBody(rows) {
      const cols = getActiveViewColumns();
      const startIdx = (state.page - 1) * state.pageSize;
      const endIdx = Math.min(startIdx + state.pageSize, rows.length);
      const pageRows = rows.slice(startIdx, endIdx);

      if (pageRows.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="' + (cols.length + 2) + '" style="text-align: center; padding: 36px; color: #64748B; font-weight: 600;">No records match your filters</td></tr>';
        rowRangeStart.innerText = '0';
        rowRangeEnd.innerText = '0';
        rowTotalCount.innerText = '0';
        return;
      }

      rowRangeStart.innerText = startIdx + 1;
      rowRangeEnd.innerText = endIdx;
      rowTotalCount.innerText = rows.length;

      let html = '';
      pageRows.forEach((row, i) => {
        const absoluteIndex = startIdx + i + 1;
        const isHighlighted = state.highlightedWos.has(normalizeWo(row.workOrderNo));
        const rowClass = isHighlighted ? 'row-highlighted' : '';
        html += '<tr class="' + rowClass + '">';
        html += '<td class="sticky-col-idx" style="color: ' + (isHighlighted ? '#78350F' : '#64748B') + '; font-family: monospace; font-weight: ' + (isHighlighted ? '700' : '400') + ';">' + absoluteIndex + '</td>';

        cols.forEach(col => {
          const val = row[col.key];
          const isStickyWo = col.key === 'workOrderNo';
          const stickyClass = isStickyWo ? 'sticky-col-wo' : '';
          let cellContent = val !== undefined && val !== null && String(val).trim() !== '' ? escapeHtml(val) : '<span style="color: #CBD5E1;">—</span>';

          if (col.key === 'workOrderNo' && isHighlighted) {
            cellContent = '<span style="color: #78350F; font-weight: 700;">' + escapeHtml(val) + '</span> <span class="badge-highlighted">Highlighted</span>';
          } else if ((col.key === 'jmsCycleNoFbQuality' || col.key === 'xbrlFeedbackQuality') && val) {
            cellContent = '<span class="badge badge-teal">' + escapeHtml(cleanFbQuality(val)) + '</span>';
          } else if (col.key === 'samplingQuality' && val) {
            cellContent = '<span class="badge badge-indigo">' + escapeHtml(val) + '</span>';
          } else if (col.key === 'formType' && val) {
            cellContent = '<span class="badge badge-amber">' + escapeHtml(val) + '</span>';
          } else if (col.key === 'xTrackId' && val) {
            cellContent = '<span style="color: #0284C7; font-weight: 700; font-family: monospace;">' + escapeHtml(val) + '</span>';
          } else if (col.key === 'month' && val) {
            cellContent = '<span style="color: #9F1239; font-weight: 600;">' + escapeHtml(val) + '</span>';
          } else if (typeof val === 'number') {
            cellContent = '<span style="font-family: monospace; text-align: right; display: block; ' + (val > 0 ? 'font-weight: 700; color: #DC2626;' : 'color: #94A3B8;') + '">' + val + '</span>';
          }

          html += '<td class="' + stickyClass + '">' + cellContent + '</td>';
        });

        // Details Button
        html += '<td style="text-align: center;"><button type="button" class="btn" style="padding: 2px 6px; font-size: 10.5px;" onclick="openInspector(' + (absoluteIndex - 1) + ')">View</button></td>';
        html += '</tr>';
      });

      tableBody.innerHTML = html;
    }

    // Render Pagination Controls
    function renderPagination(totalRows) {
      const totalPages = Math.ceil(totalRows / state.pageSize) || 1;
      let html = '';

      html += '<button type="button" class="page-btn" ' + (state.page <= 1 ? 'disabled' : '') + ' onclick="changePage(' + (state.page - 1) + ')">&laquo; Prev</button>';

      const maxBtns = 5;
      let startP = Math.max(1, state.page - 2);
      let endP = Math.min(totalPages, startP + maxBtns - 1);
      if (endP - startP < maxBtns - 1) {
        startP = Math.max(1, endP - maxBtns + 1);
      }

      for (let p = startP; p <= endP; p++) {
        html += '<button type="button" class="page-btn ' + (p === state.page ? 'active' : '') + '" onclick="changePage(' + p + ')">' + p + '</button>';
      }

      html += '<button type="button" class="page-btn" ' + (state.page >= totalPages ? 'disabled' : '') + ' onclick="changePage(' + (state.page + 1) + ')">Next &raquo;</button>';

      paginationControls.innerHTML = html;
    }

    // Render Columns Visibility Menu
    function renderColumnsMenu() {
      const allCols = VIEW_DEFS[state.view] || VIEW_DEFS.standard;
      let html = '<div style="font-size: 11px; font-weight: 700; color: #64748B; margin-bottom: 6px; padding: 0 4px;">Toggle Columns</div>';

      allCols.forEach(col => {
        const isChecked = !state.hiddenColumns.has(col.key);
        html += '<label class="column-check-item">' +
          '<input type="checkbox" ' + (isChecked ? 'checked' : '') + ' onchange="toggleColumnVisibility(\\'' + col.key + '\\')" />' +
          '<span>' + escapeHtml(col.header) + '</span>' +
        '</label>';
      });

      colsMenuList.innerHTML = html;
    }

    function updateHighlightUI() {
      const count = state.highlightedWos.size;
      const inlineStatus = document.getElementById('inline-highlight-status');
      const inlineCount = document.getElementById('inline-highlight-count');
      const inlineCheck = document.getElementById('check-inline-only-highlighted');

      if (count > 0) {
        badgeHighlightCount.style.display = 'inline-block';
        badgeHighlightCount.innerText = count;
        containerOnlyHighlighted.style.display = 'inline-flex';
        if (checkOnlyHighlighted) checkOnlyHighlighted.checked = state.onlyHighlighted;

        if (inlineStatus) inlineStatus.style.display = 'flex';
        if (inlineCount) inlineCount.innerText = count;
        if (inlineCheck) inlineCheck.checked = state.onlyHighlighted;
      } else {
        badgeHighlightCount.style.display = 'none';
        containerOnlyHighlighted.style.display = 'none';
        state.onlyHighlighted = false;
        if (checkOnlyHighlighted) checkOnlyHighlighted.checked = false;

        if (inlineStatus) inlineStatus.style.display = 'none';
        if (inlineCheck) inlineCheck.checked = false;
      }
    }

    // Update entire view
    function updateUI() {
      // Update Tab Buttons
      Array.from(filterTabs.querySelectorAll('.tab-btn')).forEach(btn => {
        btn.classList.toggle('active', btn.dataset.filter === state.filter);
      });

      // Update KPI Cards
      document.querySelectorAll('.kpi-card').forEach(card => {
        card.classList.toggle('active', card.dataset.filter === state.filter);
      });

      // Update View Pills
      Array.from(viewPills.querySelectorAll('.view-pill')).forEach(pill => {
        pill.classList.toggle('active', pill.dataset.view === state.view);
      });

      // Update status label
      const labels = {
        matched: 'Matched Records',
        unmatched_bridge: 'Bridge Unmatched Records',
        matched_and_bridge_unmatched: 'Matched + Bridge Unmatched (Base Bridge)',
        unmatched_jms: 'JMS Unmatched Records',
        all: 'All Reconciled Records',
      };
      statusLabel.innerText = labels[state.filter] || 'Matched Records';

      clearSearch.style.display = state.search ? 'block' : 'none';

      updateHighlightUI();
      renderTableHeader();
      renderColumnsMenu();

      const rows = getProcessedRows();
      renderTableBody(rows);
      renderPagination(rows.length);
    }

    // Event handlers
    window.handleSort = function(key) {
      if (state.sortKey === key) {
        state.sortAsc = !state.sortAsc;
      } else {
        state.sortKey = key;
        state.sortAsc = true;
      }
      state.page = 1;
      updateUI();
    };

    window.changePage = function(p) {
      state.page = p;
      updateUI();
    };

    window.toggleColumnVisibility = function(key) {
      if (state.hiddenColumns.has(key)) {
        state.hiddenColumns.delete(key);
      } else {
        state.hiddenColumns.add(key);
      }
      renderTableHeader();
      const rows = getProcessedRows();
      renderTableBody(rows);
      renderColumnsMenu();
    };

    window.openInspector = function(index) {
      const rows = getProcessedRows();
      const row = rows[index];
      if (!row) return;

      document.getElementById('modal-title').innerText = 'Audit Record: ' + (row.workOrderNo || 'N/A');
      document.getElementById('modal-subtitle').innerText = 'X Track ID: ' + (row.xTrackId || 'None') + ' · Status: ' + (row.matchStatus || 'matched');

      let bodyHtml = '<div style="background: #F0FDFA; border: 1px solid #99F6E4; border-radius: 8px; padding: 12px; margin-bottom: 12px;">' +
        '<h4 style="color: #0F766E; margin-bottom: 8px; font-size: 12px;">XBRL Feedback &amp; Quality Audit</h4>' +
        '<div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; font-size: 11.5px;">' +
          '<div><span style="color: #64748B; display: block;">JMS Cycle No / FB Quality:</span><strong style="color: #0F766E;">' + (row.jmsCycleNoFbQuality || 'None') + '</strong></div>' +
          '<div><span style="color: #64748B; display: block;">Sampling / Quality:</span><strong style="color: #3730A3;">' + (row.samplingQuality || 'STP Feedback') + '</strong></div>' +
          '<div><span style="color: #64748B; display: block;">FB Status:</span><span>' + (row.xbrlFeedbackStatus || 'Completed') + '</span></div>' +
          '<div><span style="color: #64748B; display: block;">FB Reviewed:</span><span>' + (row.xbrlFeedbackReviewed || 'Yes') + '</span></div>' +
        '</div>' +
        (row.xbrlFeedbackComments ? '<div style="margin-top: 8px; background: #FFF; padding: 8px; border-radius: 6px; border: 1px solid #99F6E4;"><span style="color: #64748B; font-size: 11px;">Comments:</span><p style="margin: 2px 0 0 0; font-style: italic;">&ldquo;' + escapeHtml(row.xbrlFeedbackComments) + '&rdquo;</p></div>' : '') +
      '</div>';

      bodyHtml += '<div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 12px;">' +
        '<h4 style="color: #0F172A; margin-bottom: 8px; font-size: 12px;">Canonical Fields</h4>' +
        '<div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; font-size: 11.5px;">' +
          '<div><span style="color: #64748B; display: block;">Work Order:</span><strong>' + (row.workOrderNo || '—') + '</strong></div>' +
          '<div><span style="color: #64748B; display: block;">X Track ID:</span><span style="color: #0284C7; font-family: monospace; font-weight: 700;">' + (row.xTrackId || '—') + '</span></div>' +
          '<div><span style="color: #64748B; display: block;">Job Name:</span><span>' + (row.jobName || '—') + '</span></div>' +
          '<div><span style="color: #64748B; display: block;">Form Type:</span><span>' + (row.formType || '—') + '</span></div>' +
          '<div><span style="color: #64748B; display: block;">Milestone:</span><span>' + (row.milestone || '—') + '</span></div>' +
          '<div><span style="color: #64748B; display: block;">Delivered Date:</span><span>' + (row.deliveredDate || '—') + '</span></div>' +
        '</div>' +
      '</div>';

      document.getElementById('modal-body').innerHTML = bodyHtml;
      inspectModal.classList.add('open');
    };

    // Filter tab clicks
    filterTabs.addEventListener('click', e => {
      const btn = e.target.closest('.tab-btn');
      if (!btn) return;
      state.filter = btn.dataset.filter;
      state.page = 1;
      updateUI();
    });

    // KPI Card clicks
    document.querySelectorAll('.kpi-card').forEach(card => {
      card.addEventListener('click', () => {
        state.filter = card.dataset.filter;
        state.page = 1;
        updateUI();
      });
    });

    // View switcher clicks
    viewPills.addEventListener('click', e => {
      const pill = e.target.closest('.view-pill');
      if (!pill) return;
      state.view = pill.dataset.view;
      state.hiddenColumns.clear();
      state.page = 1;
      updateUI();
    });

    // Search input
    searchInput.addEventListener('input', e => {
      state.search = e.target.value;
      state.page = 1;
      updateUI();
    });

    clearSearch.addEventListener('click', () => {
      searchInput.value = '';
      state.search = '';
      state.page = 1;
      updateUI();
    });

    // Compact mode toggle
    btnCompact.addEventListener('click', () => {
      state.compact = !state.compact;
      mainTable.classList.toggle('compact', state.compact);
      btnCompact.innerText = state.compact ? 'Standard View' : 'Fit to Screen';
    });

    // Dropdown columns toggle
    btnColsMenu.addEventListener('click', e => {
      e.stopPropagation();
      dropdownCols.classList.toggle('open');
    });

    document.addEventListener('click', e => {
      if (!dropdownCols.contains(e.target)) {
        dropdownCols.classList.remove('open');
      }
    });

    pageSizeSelect.addEventListener('change', e => {
      state.pageSize = Number(e.target.value);
      state.page = 1;
      updateUI();
    });

    modalCloseBtn.addEventListener('click', () => inspectModal.classList.remove('open'));
    inspectModal.addEventListener('click', e => {
      if (e.target === inspectModal) inspectModal.classList.remove('open');
    });

    // Copy TSV
    btnCopyTsv.addEventListener('click', () => {
      const rows = getProcessedRows();
      const cols = getActiveViewColumns();
      let tsv = cols.map(c => c.header).join('\\t') + '\\n';

      rows.forEach(r => {
        tsv += cols.map(c => {
          let v = r[c.key];
          if (c.key === 'jmsCycleNoFbQuality' || c.key === 'xbrlFeedbackQuality') v = cleanFbQuality(v);
          return v !== undefined && v !== null ? v : '';
        }).join('\\t') + '\\n';
      });

      navigator.clipboard.writeText(tsv).then(() => {
        const oldText = btnCopyTsv.innerText;
        btnCopyTsv.innerText = 'Copied!';
        setTimeout(() => { btnCopyTsv.innerText = oldText; }, 1800);
      });
    });

    // Download CSV
    btnExportCsv.addEventListener('click', () => {
      const rows = getProcessedRows();
      const cols = getActiveViewColumns();
      let csv = cols.map(c => '"' + c.header.replace(/"/g, '""') + '"').join(',') + '\\r\\n';

      rows.forEach(r => {
        csv += cols.map(c => {
          let v = r[c.key];
          if (c.key === 'jmsCycleNoFbQuality' || c.key === 'xbrlFeedbackQuality') v = cleanFbQuality(v);
          const str = v !== undefined && v !== null ? String(v) : '';
          return '"' + str.replace(/"/g, '""') + '"';
        }).join(',') + '\\r\\n';
      });

      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'Bridge_Work_Orders_' + state.filter + '_' + state.view + '.csv';
      link.click();
    });

    // Excel Export Function
    function exportRowsToExcel(targetRows, filterName) {
      const cols = getActiveViewColumns();

      if (window.XLSX) {
        const hasHighlights = state.highlightedWos && state.highlightedWos.size > 0;
        const exportData = targetRows.map(r => {
          const item = {};
          if (hasHighlights) {
            const isH = state.highlightedWos.has(normalizeWo(r.workOrderNo)) || state.highlightedWos.has(normalizeWo(r.jmsJobNo));
            item['Highlighted'] = isH ? 'YES (Highlighted)' : '';
          }
          cols.forEach(c => {
            let val = r[c.key];
            if (c.key === 'jmsCycleNoFbQuality' || c.key === 'xbrlFeedbackQuality') val = cleanFbQuality(val);
            item[c.header] = val !== undefined && val !== null ? val : '';
          });
          return item;
        });

        const ws = XLSX.utils.json_to_sheet(exportData);
        const colList = hasHighlights ? [{ header: 'Highlighted', width: 18 }, ...cols] : cols;
        ws['!cols'] = colList.map(c => ({ wch: Math.max(c.header.length + 4, c.width || 15) }));

        const totalCols = colList.length;
        const totalRowCount = targetRows.length;

        // Header Styling
        for (let c = 0; c < totalCols; c++) {
          const cellAddr = XLSX.utils.encode_cell({ r: 0, c: c });
          if (ws[cellAddr]) {
            ws[cellAddr].s = {
              fill: { patternType: 'solid', fgColor: { rgb: (hasHighlights && c === 0) ? 'F59E0B' : '0F172A' } },
              font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
              alignment: { vertical: 'center', horizontal: 'center' }
            };
          }
        }

        // Row Highlighting in Yellow
        for (let r = 0; r < totalRowCount; r++) {
          const rData = targetRows[r];
          const isH = hasHighlights && (state.highlightedWos.has(normalizeWo(rData.workOrderNo)) || state.highlightedWos.has(normalizeWo(rData.jmsJobNo)));
          const excelRow = r + 1;

          for (let c = 0; c < totalCols; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: excelRow, c: c });
            if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: '' };

            if (isH) {
              ws[cellAddr].s = {
                fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
                font: { name: 'Calibri', sz: 11, bold: c <= 2, color: { rgb: '000000' } },
                alignment: { vertical: 'center' }
              };
            } else {
              ws[cellAddr].s = {
                fill: { patternType: 'solid', fgColor: { rgb: r % 2 === 1 ? 'F9FAFB' : 'FFFFFF' } },
                font: { name: 'Calibri', sz: 11, color: { rgb: '111827' } },
                alignment: { vertical: 'center' }
              };
            }
          }
        }

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Reconciled Output');
        XLSX.writeFile(wb, 'Bridge_Work_Orders_' + filterName + '_' + state.view + '.xlsx');
      } else {
        btnExportCsv.click();
      }
    }

    btnExportXlsx.addEventListener('click', () => exportRowsToExcel(getProcessedRows(), state.filter));
    if (btnIntroExportXlsx) btnIntroExportXlsx.addEventListener('click', () => exportRowsToExcel(getProcessedRows(), state.filter));
    if (btnTopDownloadXlsx) btnTopDownloadXlsx.addEventListener('click', () => exportRowsToExcel(getProcessedRows(), state.filter));

    // Combined Matched + Bridge Unmatched Export
    function exportCombinedMatchedAndBridgeUnmatched() {
      const combined = DATASETS.matched_and_bridge_unmatched || [...(DATASETS.matched || []), ...(DATASETS.unmatched_bridge || [])];
      exportRowsToExcel(combined, 'Matched_and_Bridge_Unmatched');
    }
    const btnExportComb1 = document.getElementById('btn-export-matched-bridge-xlsx');
    if (btnExportComb1) btnExportComb1.addEventListener('click', exportCombinedMatchedAndBridgeUnmatched);
    const btnExportComb2 = document.getElementById('btn-top-export-matched-bridge-xlsx');
    if (btnExportComb2) btnExportComb2.addEventListener('click', exportCombinedMatchedAndBridgeUnmatched);

    btnPrint.addEventListener('click', () => window.print());

    // Highlight Modal Wiring
    btnOpenHighlightModal.addEventListener('click', () => {
      highlightTextarea.value = Array.from(state.highlightedWos).join('\\n');
      highlightModal.classList.add('open');
    });
    highlightModalClose.addEventListener('click', () => highlightModal.classList.remove('open'));
    btnCancelHighlights.addEventListener('click', () => highlightModal.classList.remove('open'));
    highlightModal.addEventListener('click', e => {
      if (e.target === highlightModal) highlightModal.classList.remove('open');
    });

    function populateIds(idsArray) {
      if (checkAppendHighlights && checkAppendHighlights.checked && highlightTextarea.value.trim().length > 0) {
        const existing = highlightTextarea.value.split(/[\\r\\n,;\\t]+/).map(s => s.trim()).filter(Boolean);
        const combined = Array.from(new Set([...existing, ...idsArray]));
        highlightTextarea.value = combined.join('\\n');
      } else {
        highlightTextarea.value = idsArray.join('\\n');
      }
    }

    if (btnLoadSampleIds) btnLoadSampleIds.addEventListener('click', () => populateIds(ALL_SCREENSHOT_IDS));
    if (btnLoadCoreIds) btnLoadCoreIds.addEventListener('click', () => populateIds(CORE_5_IDS));
    if (btnLoadTable1Ids) btnLoadTable1Ids.addEventListener('click', () => populateIds(TABLE1_IDS));

    btnClearHighlights.addEventListener('click', () => {
      state.highlightedWos.clear();
      highlightTextarea.value = '';
      updateHighlightUI();
      highlightModal.classList.remove('open');
      state.page = 1;
      updateUI();
    });

    btnApplyHighlights.addEventListener('click', () => {
      const text = highlightTextarea.value;
      const ids = text.split(/[\\r\\n,;\\t]+/).map(normalizeWo).filter(Boolean);
      state.highlightedWos = new Set(ids);
      updateHighlightUI();
      highlightModal.classList.remove('open');
      state.page = 1;
      updateUI();
    });

    checkOnlyHighlighted.addEventListener('change', e => {
      state.onlyHighlighted = e.target.checked;
      const inlineCheck = document.getElementById('check-inline-only-highlighted');
      if (inlineCheck) inlineCheck.checked = e.target.checked;
      state.page = 1;
      updateUI();
    });

    if (btnToggleInputOptions && inputOptionsContainer) {
      btnToggleInputOptions.addEventListener('click', () => {
        const isHidden = inputOptionsContainer.style.display === 'none';
        inputOptionsContainer.style.display = isHidden ? 'block' : 'none';
        btnToggleInputOptions.innerText = isHidden ? 'Hide Input Controls' : 'Toggle Input Controls';
      });
    }

    function addBatchHighlights(idsArray) {
      idsArray.forEach(id => {
        const n = normalizeWo(id);
        if (n) state.highlightedWos.add(n);
      });
      updateHighlightUI();
      state.page = 1;
      updateUI();
    }

    if (btnQuickCore5) btnQuickCore5.addEventListener('click', () => addBatchHighlights(CORE_5_IDS));
    if (btnQuickTable1) btnQuickTable1.addEventListener('click', () => addBatchHighlights(TABLE1_IDS));
    if (btnQuickAll15) btnQuickAll15.addEventListener('click', () => addBatchHighlights(ALL_SCREENSHOT_IDS));

    if (btnAddInlineWo && inlineWoInput) {
      const handleAddInline = () => {
        const val = inlineWoInput.value;
        if (!val.trim()) return;
        const ids = val.split(/[\\r\\n,;\\t]+/).map(s => s.trim()).filter(Boolean);
        addBatchHighlights(ids);
        inlineWoInput.value = '';
      };
      btnAddInlineWo.addEventListener('click', handleAddInline);
      inlineWoInput.addEventListener('keydown', e => {
        if (e.key === 'Enter') {
          e.preventDefault();
          handleAddInline();
        }
      });
    }

    if (btnClearInlineWos) {
      btnClearInlineWos.addEventListener('click', () => {
        state.highlightedWos.clear();
        state.onlyHighlighted = false;
        updateHighlightUI();
        state.page = 1;
        updateUI();
      });
    }

    if (checkInlineOnlyHighlighted) {
      checkInlineOnlyHighlighted.addEventListener('change', e => {
        state.onlyHighlighted = e.target.checked;
        if (checkOnlyHighlighted) checkOnlyHighlighted.checked = e.target.checked;
        state.page = 1;
        updateUI();
      });
    }

    if (btnOpenModalFromBar) {
      btnOpenModalFromBar.addEventListener('click', () => btnOpenHighlightModal.click());
    }

    // ==========================================
    // Smart Offline Excel Parser & Auto-Detector
    // ==========================================
    function parseSheetToObjects(worksheet) {
      if (!worksheet) return { headers: [], rows: [] };

      // Convert sheet to 2D array
      const rawData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        blankrows: false,
        raw: false,
      });

      if (!rawData || rawData.length === 0) return { headers: [], rows: [] };

      // Detect header row (first row with >= 2 non-empty string cells)
      let headerRowIndex = 0;
      for (let r = 0; r < Math.min(rawData.length, 5); r++) {
        const row = rawData[r];
        const nonEmpty = (row || []).filter(c => c !== null && c !== undefined && String(c).trim() !== '');
        if (nonEmpty.length >= 2) {
          headerRowIndex = r;
          break;
        }
      }

      const rawHeaders = rawData[headerRowIndex] || [];
      const headers = [];
      const counts = {};

      rawHeaders.forEach((h, idx) => {
        let str = h ? String(h).trim() : ('Column_' + (idx + 1));
        if (counts[str]) {
          counts[str]++;
          headers.push(str + '_' + counts[str]);
        } else {
          counts[str] = 1;
          headers.push(str);
        }
      });

      const rows = [];
      for (let r = headerRowIndex + 1; r < rawData.length; r++) {
        const row = rawData[r];
        if (!row || row.length === 0) continue;

        const obj = {};
        let hasData = false;

        headers.forEach((h, colIdx) => {
          const val = row[colIdx];
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            hasData = true;
            obj[h] = val;
          } else {
            obj[h] = '';
          }
        });

        if (hasData) rows.push(obj);
      }

      return { headers, rows };
    }

    function normalizeHeaderString(h) {
      if (!h) return '';
      return String(h)
        .toLowerCase()
        .replace(/[\\r\\n\\t_/#\\-\\.]+/g, ' ')
        .replace(/[^\\w\\s]/g, '')
        .replace(/\\s+/g, ' ')
        .trim();
    }

    function findBestColumnMatch(availableHeaders, targetPreferences, excludePatterns = []) {
      if (!availableHeaders || availableHeaders.length === 0) return '';

      const normalizedAvailable = availableHeaders.map(h => ({
        original: h,
        normalized: normalizeHeaderString(h),
      }));

      // Pass 1: Exact match
      for (const target of targetPreferences) {
        const normTarget = normalizeHeaderString(target);
        const exact = normalizedAvailable.find(item => item.normalized === normTarget);
        if (exact) return exact.original;
      }

      // Pass 2: Contains match
      for (const target of targetPreferences) {
        const normTarget = normalizeHeaderString(target);
        for (const item of normalizedAvailable) {
          const isExcluded = excludePatterns.some(exc => item.normalized.includes(exc.toLowerCase()));
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

    function cleanJobName(raw) {
      if (!raw) return '';
      let title = String(raw).trim();
      title = title.replace(/\\s*\\|\\s*Bridge\\s*XBRL(\\b|$)/gi, '');
      title = title.replace(/\\s*-\\s*Bridge\\s*XBRL(\\b|$)/gi, '');
      title = title.replace(/\\s*Bridge\\s*XBRL$/gi, '');
      return title.replace(/[\\s|\\-]+$/, '').trim();
    }

    function transformMilestone(raw) {
      if (!raw) return '';
      const activity = String(raw).trim();
      if (/^XBRL\\s*Update\\s*-\\s*Final\\s*to\\s*Client$/i.test(activity)) return 'Final to Client';
      if (/^XBRL\\s*Update\\s*-\\s*Non\\s*Standard$/i.test(activity)) return 'Non Standard';
      if (/^XBRL\\s*Update\\s*-\\s*Standard$/i.test(activity)) return 'Standard';
      const match = activity.match(/^XBRL\\s*Update\\s*[-–—:]\\s*(.+)$/i);
      if (match && match[1]) return match[1].trim();
      return activity;
    }

    function resolveFormType(grabber, jobType, fallback) {
      if (grabber && String(grabber).trim()) return String(grabber).trim();
      if (jobType && String(jobType).trim()) return String(jobType).trim();
      if (fallback && String(fallback).trim()) return String(fallback).trim();
      return '';
    }

    function excelDateToString(serial) {
      if (typeof serial === 'number' && serial > 1000) {
        const utcDays = Math.floor(serial - 25569);
        const date = new Date(utcDays * 86400 * 1000);
        const m = String(date.getUTCMonth() + 1).padStart(2, '0');
        const d = String(date.getUTCDate()).padStart(2, '0');
        const y = date.getUTCFullYear();
        return m + '/' + d + '/' + y;
      }
      return String(serial);
    }

    function extractMonth(dateStr) {
      if (!dateStr) return '';
      const parts = String(dateStr).split(/[\\/\\-]/);
      if (parts.length >= 2) {
        const m = parseInt(parts[0], 10);
        const monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        if (m >= 1 && m <= 12) return monthNames[m];
      }
      return '';
    }

    // ==========================================
    // Full In-Browser Reconciliation Function
    // ==========================================
    function reconcileOffline() {
      const statusEl = document.getElementById('reconcile-status');
      if (statusEl) statusEl.innerText = 'Reconciling...';

      const bridgeHeaders = currentBridgeData.length > 0 ? Object.keys(currentBridgeData[0]) : [];
      const jmsHeaders = currentJmsData.length > 0 ? Object.keys(currentJmsData[0]) : [];
      const prodHeaders = currentProdData.length > 0 ? Object.keys(currentProdData[0]) : [];

      // Auto-detect columns
      const bridgeWoCol = findBestColumnMatch(bridgeHeaders, [
        'Work Order', 'Work Order Number', 'Work Order No', 'WO Number', 'WO#', 'JMS Job Number/Work Order No', 'Job Number', 'Job No'
      ], ['title', 'name', 'activity', 'date', 'type']) || bridgeHeaders[0] || 'Work Order';

      const bridgeTitleCol = findBestColumnMatch(bridgeHeaders, [
        'Work Order Title', 'Job Name', 'Work Order Name', 'Title', 'Client Name', 'Company Name'
      ], ['number', 'no', 'id', 'date']) || 'Work Order Title';

      const bridgeActivityCol = findBestColumnMatch(bridgeHeaders, [
        'Activity', 'Milestone', 'Task', 'Activity Name', 'Stage'
      ]) || 'Activity';

      const bridgeDueDateCol = findBestColumnMatch(bridgeHeaders, [
        'Due Date', 'Delivered Date', 'Delivery Date', 'Target Date'
      ]) || 'Due Date';

      const bridgeFbCommentsCol = findBestColumnMatch(bridgeHeaders, [
        'XBRL Feedback Comments', 'Feedback Comments', 'XBRL Comments', 'Feedback Comment', 'Comments'
      ]) || 'XBRL Feedback Comments';

      const bridgeFbQualityCol = findBestColumnMatch(bridgeHeaders, [
        'XBRL Feedback Quality', 'Feedback Quality', 'FB Quality', 'JMS Cycle No / FB Quality'
      ]) || 'XBRL Feedback Quality';

      const jmsWoCol = findBestColumnMatch(jmsHeaders, [
        'JMS Job Number/Work Order No', 'Work Order No', 'Work Order Number', 'JMS Job Number', 'Work Order', 'Job Number', 'JMS Job No', 'Job No', 'WO No', 'WO#'
      ], ['title', 'name', 'date', 'type']) || jmsHeaders[0] || 'Work Order No';

      const jmsTrackIdCol = findBestColumnMatch(jmsHeaders, [
        'X Track ID', 'X-Track ID', 'X Track', 'Track ID', 'Tracking ID'
      ]) || 'X Track ID';

      const jmsGrabberCol = findBestColumnMatch(jmsHeaders, [
        'Grabber Day to Day', 'Grabber Form', 'Grabber', 'Form Type', 'Day to Day'
      ]) || 'Grabber Day to Day';

      const jmsJobTypeCol = findBestColumnMatch(jmsHeaders, [
        'Job Type', 'Form Type', 'Type'
      ]) || 'Job Type';

      const prodWoCol = findBestColumnMatch(prodHeaders, [
        'Work Order No', 'Work Order Number', 'Work Order', 'Job Number', 'WO No', 'WO#'
      ]) || prodHeaders[0] || 'Work Order No';

      const prodTrackIdCol = findBestColumnMatch(prodHeaders, [
        'X Track ID', 'X-Track ID', 'X Track', 'Track ID'
      ]) || 'X Track ID';

      // Index Production Day to Day
      const prodMap = new Map();
      (currentProdData || []).forEach(row => {
        const raw = row[prodWoCol] || row['Work Order No'] || row['Work Order'];
        const wo = normalizeWo(raw);
        if (wo && !prodMap.has(wo)) prodMap.set(wo, row);
      });

      // Index JMS Tracker
      const jmsMap = new Map();
      (currentJmsData || []).forEach((row, idx) => {
        const raw = row[jmsWoCol] || row['JMS Job Number/Work Order No'] || row['Work Order No'] || row['Job Number'];
        const wo = normalizeWo(raw);
        if (wo && !jmsMap.has(wo)) jmsMap.set(wo, { row, idx });
      });

      const matched = [];
      const unmatchedBridge = [];
      const matchedJmsKeys = new Set();

      (currentBridgeData || []).forEach(bRow => {
        const rawWo = bRow[bridgeWoCol] || bRow['Work Order'] || bRow['Work Order Number'];
        const wo = normalizeWo(rawWo);
        if (!wo) return;

        const jmsEntry = jmsMap.get(wo);
        const jmsMatch = jmsEntry ? jmsEntry.row : null;
        const prodMatch = prodMap.get(wo);

        const rawTitle = bRow[bridgeTitleCol] || bRow['Work Order Title'] || bRow['Job Name'] || '';
        const cleanTitle = cleanJobName(rawTitle);
        const milestone = transformMilestone(bRow[bridgeActivityCol] || bRow['Activity'] || '');
        const dateVal = bRow[bridgeDueDateCol] || bRow['Due Date'] || '';
        const formattedDate = typeof dateVal === 'number' ? excelDateToString(dateVal) : String(dateVal);
        const month = extractMonth(formattedDate);

        const xTrackId = (jmsMatch && jmsMatch[jmsTrackIdCol]) ||
                         (prodMatch && prodMatch[prodTrackIdCol]) ||
                         bRow['X Track ID'] || '';

        const grabberVal = (jmsMatch && jmsMatch[jmsGrabberCol]) || bRow['Grabber Day to Day'] || '';
        const jobTypeVal = (jmsMatch && jmsMatch[jmsJobTypeCol]) || bRow['Job Type'] || '';
        const formType = resolveFormType(grabberVal, jobTypeVal, '10-Q');

        const feedbackComments = bRow[bridgeFbCommentsCol] || (jmsMatch && jmsMatch['XBRL Feedback Comments']) || '';
        const rawFeedbackQuality = bRow[bridgeFbQualityCol] || (jmsMatch && jmsMatch['JMS Cycle No / FB Quality']) || '';
        const feedbackQuality = cleanFbQuality(rawFeedbackQuality);

        const outputRow = {
          team: 'Bridge',
          workOrderNo: rawWo,
          xTrackId: xTrackId,
          jobName: cleanTitle || rawTitle,
          formType: formType,
          milestone: milestone,
          deliveredDate: formattedDate,
          month: month,
          jmsCycleNoFbQuality: feedbackQuality,
          samplingQuality: 'STP Feedback',
          xbrlFeedbackStatus: bRow['XBRL Feedback Status'] || 'Completed',
          xbrlFeedbackReviewed: bRow['XBRL Feedback Reviewed'] || 'Yes',
          xbrlFeedbackValidity: bRow['XBRL Feedback Validity'] || 'Valid',
          xbrlFeedbackComments: feedbackComments,
          matchStatus: jmsMatch ? 'matched' : 'unmatched_bridge',
          rawBridge: bRow,
          rawJms: jmsMatch
        };

        if (jmsMatch) {
          matched.push(outputRow);
          matchedJmsKeys.add(wo);
        } else {
          unmatchedBridge.push(outputRow);
        }
      });

      const unmatchedJms = [];
      (currentJmsData || []).forEach(jRow => {
        const rawWo = jRow[jmsWoCol] || jRow['JMS Job Number/Work Order No'] || jRow['Work Order No'] || jRow['Job Number'];
        const wo = normalizeWo(rawWo);
        if (!wo || !matchedJmsKeys.has(wo)) {
          unmatchedJms.push({
            team: 'JMS',
            workOrderNo: rawWo || '—',
            xTrackId: jRow[jmsTrackIdCol] || '',
            jobName: cleanJobName(jRow['Job Name'] || jRow['Title'] || ''),
            formType: jRow[jmsGrabberCol] || jRow['Form Type'] || '',
            milestone: jRow['Milestone'] || '',
            deliveredDate: '',
            month: '',
            jmsCycleNoFbQuality: cleanFbQuality(jRow['FB Quality'] || jRow['JMS Cycle No / FB Quality'] || ''),
            samplingQuality: 'STP Feedback',
            xbrlFeedbackStatus: 'Completed',
            xbrlFeedbackReviewed: 'Yes',
            xbrlFeedbackValidity: 'Valid',
            xbrlFeedbackComments: jRow['XBRL Feedback Comments'] || '',
            matchStatus: 'unmatched_jms',
            rawJms: jRow
          });
        }
      });

      const combined = [...matched, ...unmatchedBridge];
      const all = [...matched, ...unmatchedBridge, ...unmatchedJms];

      DATASETS = {
        matched: matched,
        unmatched_bridge: unmatchedBridge,
        matched_and_bridge_unmatched: combined,
        unmatched_jms: unmatchedJms,
        all: all
      };

      // Update KPI Statistics & Badges
      const totalBridgeOrders = combined.length;
      const ratePercent = totalBridgeOrders > 0 ? Math.round((matched.length / totalBridgeOrders) * 100) : 0;

      const kpiMatched = document.getElementById('kpi-matched-count');
      if (kpiMatched) kpiMatched.innerText = matched.length;
      const kpiRate = document.getElementById('kpi-match-rate');
      if (kpiRate) kpiRate.innerText = '(' + ratePercent + '%)';

      const kpiUnmB = document.getElementById('kpi-unmatched-bridge-count');
      if (kpiUnmB) kpiUnmB.innerText = unmatchedBridge.length;
      const kpiDescB = document.getElementById('kpi-bridge-total-desc');
      if (kpiDescB) kpiDescB.innerText = 'of ' + totalBridgeOrders + ' Bridge';

      const kpiComb = document.getElementById('kpi-matched-bridge-count');
      if (kpiComb) kpiComb.innerText = combined.length;

      const kpiUnmJ = document.getElementById('kpi-unmatched-jms-count');
      if (kpiUnmJ) kpiUnmJ.innerText = unmatchedJms.length;

      const kpiAll = document.getElementById('kpi-all-count');
      if (kpiAll) kpiAll.innerText = all.length;

      // Update Tab Badges
      const bMatched = document.getElementById('badge-matched');
      if (bMatched) bMatched.innerText = matched.length;
      const bUnmB = document.getElementById('badge-unmatched-bridge');
      if (bUnmB) bUnmB.innerText = unmatchedBridge.length;
      const bComb = document.getElementById('badge-matched-bridge-unmatched');
      if (bComb) bComb.innerText = combined.length;
      const bUnmJ = document.getElementById('badge-unmatched-jms');
      if (bUnmJ) bUnmJ.innerText = unmatchedJms.length;
      const bAll = document.getElementById('badge-all');
      if (bAll) bAll.innerText = all.length;

      // Update Combined Export Button Badges
      const badgeTopComb = document.getElementById('badge-top-combined-count');
      if (badgeTopComb) badgeTopComb.innerText = combined.length;
      const badgeTableComb = document.getElementById('badge-table-combined-count');
      if (badgeTableComb) badgeTableComb.innerText = combined.length;

      if (statusEl) {
        statusEl.innerText = 'Reconciled: ' + matched.length + ' matched, ' + unmatchedBridge.length + ' Bridge unmat.';
      }

      // If matched has 0 but we have Bridge records, auto switch to combined view so user sees records immediately
      if (matched.length === 0 && combined.length > 0 && state.filter === 'matched') {
        state.filter = 'matched_and_bridge_unmatched';
      }

      state.page = 1;
      updateUI();
    }

    // Reset to demo data
    function resetToDemoData() {
      currentBridgeData = SAMPLE_RAW_DATA.bridge;
      currentJmsData = SAMPLE_RAW_DATA.jms;
      currentProdData = SAMPLE_RAW_DATA.prod;

      const bFileLabel = document.getElementById('label-bridge-filename');
      if (bFileLabel) bFileLabel.innerText = '1. Bridge: Bridge work order - All.xlsx';
      const bSheetLabel = document.getElementById('label-bridge-sheet');
      if (bSheetLabel) bSheetLabel.innerText = 'Sheet: Table 1';

      const jFileLabel = document.getElementById('label-jms-filename');
      if (jFileLabel) jFileLabel.innerText = '2. JMS: Saas_Job_Status_Report.xlsx';
      const jSheetLabel = document.getElementById('label-jms-sheet');
      if (jSheetLabel) jSheetLabel.innerText = 'Sheet: Grabber Day to Day';

      const bSelect = document.getElementById('select-bridge-sheet');
      if (bSelect) bSelect.style.display = 'none';
      const jSelect = document.getElementById('select-jms-sheet');
      if (jSelect) jSelect.style.display = 'none';

      reconcileOffline();
    }

    const btnTopReload = document.getElementById('btn-top-reload-sample');
    if (btnTopReload) btnTopReload.addEventListener('click', resetToDemoData);
    const btnReloadSample = document.getElementById('btn-standalone-reload-sample');
    if (btnReloadSample) btnReloadSample.addEventListener('click', resetToDemoData);

    // Standalone Bridge File upload
    const fileBridgeInput = document.getElementById('standalone-file-bridge');
    const selectBridgeSheet = document.getElementById('select-bridge-sheet');
    if (fileBridgeInput) {
      fileBridgeInput.addEventListener('change', e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(evt) {
          try {
            const data = new Uint8Array(evt.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            currentBridgeWorkbook = wb;

            const sheetNames = wb.SheetNames || [];
            // Prioritize sheet named 'Table 1' or 'Table1' and strictly ignore 'Sheet1'
            let chosenSheet = sheetNames.find(s => {
              const l = s.trim().toLowerCase();
              return l === 'table 1' || l === 'table1' || l === 'table_1';
            });
            if (!chosenSheet) {
              chosenSheet = sheetNames.find(s => !/^sheet\\s*1$/i.test(s.trim())) || sheetNames[0];
            }

            if (selectBridgeSheet && sheetNames.length > 1) {
              selectBridgeSheet.innerHTML = sheetNames.map(s => '<option value="' + escapeHtml(s) + '"' + (s === chosenSheet ? ' selected' : '') + '>' + escapeHtml(s) + '</option>').join('');
              selectBridgeSheet.style.display = 'inline-block';
            }

            const bFileLabel = document.getElementById('label-bridge-filename');
            if (bFileLabel) bFileLabel.innerText = '1. Bridge: ' + file.name;
            const bSheetLabel = document.getElementById('label-bridge-sheet');
            if (bSheetLabel) bSheetLabel.innerText = 'Sheet: ' + chosenSheet;

            const parsed = parseSheetToObjects(wb.Sheets[chosenSheet]);
            currentBridgeData = parsed.rows;
            reconcileOffline();
          } catch (err) {
            alert('Error parsing Bridge Excel file: ' + err.message);
          }
        };
        reader.readAsArrayBuffer(file);
      });
    }

    if (selectBridgeSheet) {
      selectBridgeSheet.addEventListener('change', e => {
        if (!currentBridgeWorkbook) return;
        const chosenSheet = e.target.value;
        const parsed = parseSheetToObjects(currentBridgeWorkbook.Sheets[chosenSheet]);
        currentBridgeData = parsed.rows;
        const bSheetLabel = document.getElementById('label-bridge-sheet');
        if (bSheetLabel) bSheetLabel.innerText = 'Sheet: ' + chosenSheet;
        reconcileOffline();
      });
    }

    // Standalone JMS File upload
    const fileJmsInput = document.getElementById('standalone-file-jms');
    const selectJmsSheet = document.getElementById('select-jms-sheet');
    if (fileJmsInput) {
      fileJmsInput.addEventListener('change', e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(evt) {
          try {
            const data = new Uint8Array(evt.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            currentJmsWorkbook = wb;

            const sheetNames = wb.SheetNames || [];
            let chosenSheet = sheetNames.find(s => {
              const l = s.trim().toLowerCase();
              return l.includes('grabber') || l.includes('main') || l.includes('tracker');
            }) || sheetNames[0];

            // Production Day to Day check
            const prodSheet = sheetNames.find(s => {
              const l = s.trim().toLowerCase();
              return l.includes('production') && l.includes('day');
            });
            if (prodSheet) {
              const parsedProd = parseSheetToObjects(wb.Sheets[prodSheet]);
              currentProdData = parsedProd.rows;
            }

            if (selectJmsSheet && sheetNames.length > 1) {
              selectJmsSheet.innerHTML = sheetNames.map(s => '<option value="' + escapeHtml(s) + '"' + (s === chosenSheet ? ' selected' : '') + '>' + escapeHtml(s) + '</option>').join('');
              selectJmsSheet.style.display = 'inline-block';
            }

            const jFileLabel = document.getElementById('label-jms-filename');
            if (jFileLabel) jFileLabel.innerText = '2. JMS: ' + file.name;
            const jSheetLabel = document.getElementById('label-jms-sheet');
            if (jSheetLabel) jSheetLabel.innerText = 'Sheet: ' + chosenSheet;

            const parsedJms = parseSheetToObjects(wb.Sheets[chosenSheet]);
            currentJmsData = parsedJms.rows;
            reconcileOffline();
          } catch (err) {
            alert('Error parsing JMS Excel file: ' + err.message);
          }
        };
        reader.readAsArrayBuffer(file);
      });
    }

    if (selectJmsSheet) {
      selectJmsSheet.addEventListener('change', e => {
        if (!currentJmsWorkbook) return;
        const chosenSheet = e.target.value;
        const parsed = parseSheetToObjects(currentJmsWorkbook.Sheets[chosenSheet]);
        currentJmsData = parsed.rows;
        const jSheetLabel = document.getElementById('label-jms-sheet');
        if (jSheetLabel) jSheetLabel.innerText = 'Sheet: ' + chosenSheet;
        reconcileOffline();
      });
    }

    function escapeHtml(text) {
      if (text === null || text === undefined) return '';
      return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    // Initialize UI
    updateUI();
  </script>
</body>
</html>`;

  const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
