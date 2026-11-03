/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { Header } from './components/Header';
import { FileUploadCard } from './components/FileUploadCard';
import { ColumnMappingDrawer } from './components/ColumnMappingDrawer';
import { SummaryStats } from './components/SummaryStats';
import { ReconciliationTable } from './components/ReconciliationTable';
import { RuleCard } from './components/RuleCard';
import { UploadedFileState, ColumnMappingConfig, ReconciliationFilterType } from './types';
import { parseExcelFile } from './utils/excelParser';
import {
  autoDetectColumns,
  reconcileWorkOrders,
  loadSavedConfig,
  saveConfigToStorage,
  clearSavedConfig,
  DEFAULT_CANONICAL_CONFIG,
} from './utils/reconciliation';
import { exportToExcel } from './utils/excelExporter';
import { exportToStandaloneHTML } from './utils/htmlExporter';
import { SAMPLE_BRIDGE_DATA, SAMPLE_JMS_DATA, SAMPLE_PRODUCTION_DAY_TO_DAY_DATA } from './data/sampleData';
import { FileSpreadsheet, Globe } from 'lucide-react';

const INITIAL_FILE_STATE: UploadedFileState = {
  file: null,
  fileName: '',
  sheetNames: [],
  selectedSheet: '',
  sheets: {},
  headers: [],
  data: [],
  loading: false,
};

const INITIAL_CONFIG: ColumnMappingConfig = {
  ...DEFAULT_CANONICAL_CONFIG,
};

export default function App() {
  const [bridgeState, setBridgeState] = useState<UploadedFileState>(INITIAL_FILE_STATE);
  const [jmsState, setJmsState] = useState<UploadedFileState>(INITIAL_FILE_STATE);

  // Initialize config with saved preferences or canonical default preferences
  const [config, setConfig] = useState<ColumnMappingConfig>(() => {
    const saved = loadSavedConfig();
    return {
      ...DEFAULT_CANONICAL_CONFIG,
      ...(saved || {}),
    };
  });

  const [activeFilter, setActiveFilter] = useState<ReconciliationFilterType>('matched');

  // Work Orders to highlight in yellow with persistence
  const [highlightedWorkOrders, setHighlightedWorkOrders] = useState<Set<string>>(() => {
    try {
      const saved = localStorage.getItem('bridge_highlighted_work_orders');
      if (saved) {
        const arr = JSON.parse(saved);
        if (Array.isArray(arr)) return new Set(arr);
      }
    } catch {}
    return new Set<string>();
  });

  const handleUpdateHighlightedWorkOrders = useCallback((wos: Set<string>) => {
    setHighlightedWorkOrders(wos);
    try {
      localStorage.setItem('bridge_highlighted_work_orders', JSON.stringify(Array.from(wos)));
    } catch {}
  }, []);

  // Update and persist config changes
  const handleConfigChange = useCallback((newConfig: ColumnMappingConfig) => {
    setConfig(newConfig);
    saveConfigToStorage(newConfig);
  }, []);

  // Helper to merge newly detected columns with any user-saved preferences
  const mergeRetainingSaved = useCallback(
    (detected: ColumnMappingConfig, currentConfig: ColumnMappingConfig): ColumnMappingConfig => {
      const saved = loadSavedConfig() || currentConfig;
      return {
        bridgeWorkOrderCol: saved.bridgeWorkOrderCol || detected.bridgeWorkOrderCol,
        bridgeTrackIdCol: saved.bridgeTrackIdCol || detected.bridgeTrackIdCol,
        bridgeTitleCol: saved.bridgeTitleCol || detected.bridgeTitleCol,
        bridgeActivityCol: saved.bridgeActivityCol || detected.bridgeActivityCol,
        bridgeDueDateCol: saved.bridgeDueDateCol || detected.bridgeDueDateCol,
        bridgeJobTypeCol: saved.bridgeJobTypeCol || detected.bridgeJobTypeCol,
        bridgeGrabberCol: saved.bridgeGrabberCol || detected.bridgeGrabberCol,
        bridgeTeamStaticValue: saved.bridgeTeamStaticValue || detected.bridgeTeamStaticValue || 'Bridge',
        bridgeXbrlFeedbackCommentsCol:
          saved.bridgeXbrlFeedbackCommentsCol || detected.bridgeXbrlFeedbackCommentsCol,
        bridgeXbrlFeedbackQualityCol:
          saved.bridgeXbrlFeedbackQualityCol || detected.bridgeXbrlFeedbackQualityCol,
        bridgeXbrlFeedbackStatusCol:
          saved.bridgeXbrlFeedbackStatusCol || detected.bridgeXbrlFeedbackStatusCol,
        bridgeXbrlFeedbackReviewedCol:
          saved.bridgeXbrlFeedbackReviewedCol || detected.bridgeXbrlFeedbackReviewedCol,
        bridgeXbrlFeedbackValidityCol:
          saved.bridgeXbrlFeedbackValidityCol || detected.bridgeXbrlFeedbackValidityCol,
        samplingQualityPrefill:
          saved.samplingQualityPrefill || detected.samplingQualityPrefill || 'STP Feedback',
        jmsWorkOrderCol: saved.jmsWorkOrderCol || detected.jmsWorkOrderCol,
        jmsTrackIdCol: saved.jmsTrackIdCol || detected.jmsTrackIdCol,
        jmsGrabberCol: saved.jmsGrabberCol || detected.jmsGrabberCol,
        jmsJobTypeCol: saved.jmsJobTypeCol || detected.jmsJobTypeCol,
        jmsFormTypeCol: saved.jmsFormTypeCol || detected.jmsFormTypeCol,
        jmsProductionSheet:
          saved.jmsProductionSheet !== undefined && saved.jmsProductionSheet !== ''
            ? saved.jmsProductionSheet
            : detected.jmsProductionSheet,
        jmsProductionWorkOrderCol:
          saved.jmsProductionWorkOrderCol || detected.jmsProductionWorkOrderCol,
        jmsProductionTrackIdCol:
          saved.jmsProductionTrackIdCol || detected.jmsProductionTrackIdCol,
      };
    },
    []
  );

  // Reset to freshly detected columns and clear saved storage
  const handleResetToAutoDetect = useCallback(() => {
    clearSavedConfig();
    const detected = autoDetectColumns(bridgeState.headers, jmsState.headers, jmsState.sheets);
    handleConfigChange(detected);
  }, [bridgeState.headers, jmsState.headers, jmsState.sheets, handleConfigChange]);

  // Load sample dataset
  const handleLoadSample = useCallback(() => {
    const bridgeHeaders = Object.keys(SAMPLE_BRIDGE_DATA[0]);
    const jmsHeaders = Object.keys(SAMPLE_JMS_DATA[0]);
    const prodHeaders = Object.keys(SAMPLE_PRODUCTION_DAY_TO_DAY_DATA[0]);

    setBridgeState({
      file: null,
      fileName: 'Bridge work order - All.xlsx',
      sheetNames: ['Table1'],
      selectedSheet: 'Table1',
      sheets: {
        'Table1': {
          sheetName: 'Table1',
          headers: bridgeHeaders,
          data: SAMPLE_BRIDGE_DATA,
          rowCount: SAMPLE_BRIDGE_DATA.length,
        },
      },
      headers: bridgeHeaders,
      data: SAMPLE_BRIDGE_DATA,
      loading: false,
    });

    const sampleJmsSheets = {
      'JMS Active Jobs': {
        sheetName: 'JMS Active Jobs',
        headers: jmsHeaders,
        data: SAMPLE_JMS_DATA,
        rowCount: SAMPLE_JMS_DATA.length,
      },
      'Production Day to Day': {
        sheetName: 'Production Day to Day',
        headers: prodHeaders,
        data: SAMPLE_PRODUCTION_DAY_TO_DAY_DATA,
        rowCount: SAMPLE_PRODUCTION_DAY_TO_DAY_DATA.length,
      },
    };

    setJmsState({
      file: null,
      fileName: 'Sample_JMS_SaaS_Workbook.xlsx',
      sheetNames: ['JMS Active Jobs', 'Production Day to Day'],
      selectedSheet: 'JMS Active Jobs',
      sheets: sampleJmsSheets,
      headers: jmsHeaders,
      data: SAMPLE_JMS_DATA,
      loading: false,
    });

    const detected = autoDetectColumns(bridgeHeaders, jmsHeaders, sampleJmsSheets, DEFAULT_CANONICAL_CONFIG);
    setConfig(detected);
    saveConfigToStorage(detected);
  }, []);

  // Handle Bridge File Upload
  const handleBridgeFileSelect = async (file: File) => {
    try {
      setBridgeState((prev) => ({ ...prev, loading: true, error: undefined }));
      const { sheetNames, sheets } = await parseExcelFile(file);

      // Rule: "consider only the excel sheet named table 1 from the 'Bridge work order - All' excel file and pls ignore the name sheet1"
      // Filter out any sheet named "Sheet1" or "Sheet 1" (case-insensitive)
      const nonSheet1Names = sheetNames.filter((s) => !/^sheet\s*1$/i.test(s.trim()));

      // Strictly check for sheet named "Table 1" or "Table1"
      const table1Sheet = sheetNames.find((s) => /^table\s*1$/i.test(s.trim()));

      let chosenSheet = '';
      let effectiveSheetNames: string[] = [];

      if (table1Sheet) {
        // Table 1 exists: strictly select it and ignore Sheet1
        chosenSheet = table1Sheet;
        effectiveSheetNames = [table1Sheet, ...nonSheet1Names.filter((s) => s !== table1Sheet)];
      } else {
        // Fallback: look for sheets containing table, bridge, work order, order
        const candidates = nonSheet1Names.length > 0 ? nonSheet1Names : sheetNames;
        for (const s of candidates) {
          if (/^table/i.test(s) || /bridge|work.*order|order|active/i.test(s)) {
            chosenSheet = s;
            break;
          }
        }
        chosenSheet = chosenSheet || candidates[0] || sheetNames[0] || '';
        effectiveSheetNames = candidates;
      }

      const sheetData = sheets[chosenSheet] || { headers: [], data: [] };

      setBridgeState({
        file,
        fileName: file.name,
        sheetNames: effectiveSheetNames,
        selectedSheet: chosenSheet,
        sheets,
        headers: sheetData.headers,
        data: sheetData.data,
        loading: false,
      });

      // Update mapping auto-detection, prioritizing canonical user profile
      setConfig((prevConfig) => {
        const detected = autoDetectColumns(sheetData.headers, jmsState.headers, jmsState.sheets, prevConfig);
        const updated: ColumnMappingConfig = {
          ...prevConfig,
          bridgeWorkOrderCol: detected.bridgeWorkOrderCol,
          bridgeTitleCol: detected.bridgeTitleCol,
          bridgeActivityCol: detected.bridgeActivityCol,
          bridgeDueDateCol: detected.bridgeDueDateCol,
          bridgeJobTypeCol: detected.bridgeJobTypeCol,
          bridgeGrabberCol: detected.bridgeGrabberCol,
          bridgeXbrlFeedbackCommentsCol: detected.bridgeXbrlFeedbackCommentsCol,
          bridgeXbrlFeedbackQualityCol: detected.bridgeXbrlFeedbackQualityCol,
          bridgeXbrlFeedbackStatusCol: detected.bridgeXbrlFeedbackStatusCol,
          bridgeXbrlFeedbackReviewedCol: detected.bridgeXbrlFeedbackReviewedCol,
          bridgeXbrlFeedbackValidityCol: detected.bridgeXbrlFeedbackValidityCol,
          samplingQualityPrefill: prevConfig.samplingQualityPrefill || 'STP Feedback',
        };
        saveConfigToStorage(updated);
        return updated;
      });
    } catch (err: any) {
      console.error('Error parsing Bridge file', err);
      setBridgeState((prev) => ({
        ...prev,
        loading: false,
        error: 'Failed to read file. Please ensure it is a valid .xlsx, .xls, or .csv file.',
      }));
    }
  };

  // Handle JMS / SaaS File Upload
  const handleJmsFileSelect = async (file: File) => {
    try {
      setJmsState((prev) => ({ ...prev, loading: true, error: undefined }));
      const { sheetNames, sheets } = await parseExcelFile(file);

      // Check if Production Day to Day sheet exists
      const prodSheetName =
        sheetNames.find((s) => /production.*day.*to.*day/i.test(s)) ||
        sheetNames.find((s) => /production/i.test(s)) ||
        '';

      // Pick main sheet: prefer sheets that are NOT Production Day to Day
      const nonProdSheets = sheetNames.filter((s) => s !== prodSheetName);
      let chosenSheet = nonProdSheets[0] || sheetNames[0] || '';

      for (const s of (nonProdSheets.length > 0 ? nonProdSheets : sheetNames)) {
        if (/grabber.*day.*to.*day|grabber|jms.*active|active.*job|jms|saas|active/i.test(s)) {
          chosenSheet = s;
          break;
        }
      }

      const sheetData = sheets[chosenSheet] || { headers: [], data: [] };

      setJmsState({
        file,
        fileName: file.name,
        sheetNames,
        selectedSheet: chosenSheet,
        sheets,
        headers: sheetData.headers,
        data: sheetData.data,
        loading: false,
      });

      // Update mapping auto-detection, prioritizing canonical user profile
      setConfig((prevConfig) => {
        const detected = autoDetectColumns(bridgeState.headers, sheetData.headers, sheets, prevConfig);
        const updated: ColumnMappingConfig = {
          ...prevConfig,
          jmsWorkOrderCol: detected.jmsWorkOrderCol,
          jmsTrackIdCol: detected.jmsTrackIdCol,
          jmsGrabberCol: detected.jmsGrabberCol,
          jmsJobTypeCol: detected.jmsJobTypeCol,
          jmsFormTypeCol: detected.jmsFormTypeCol,
          jmsProductionSheet: detected.jmsProductionSheet,
          jmsProductionWorkOrderCol: detected.jmsProductionWorkOrderCol,
          jmsProductionTrackIdCol: detected.jmsProductionTrackIdCol,
        };
        saveConfigToStorage(updated);
        return updated;
      });
    } catch (err: any) {
      console.error('Error parsing JMS file', err);
      setJmsState((prev) => ({
        ...prev,
        loading: false,
        error: 'Failed to read file. Please ensure it is a valid .xlsx, .xls, or .csv file.',
      }));
    }
  };

  // Handle Sheet selection change
  const handleBridgeSheetChange = (sheetName: string) => {
    const sheetData = bridgeState.sheets[sheetName];
    if (sheetData) {
      setBridgeState((prev) => ({
        ...prev,
        selectedSheet: sheetName,
        headers: sheetData.headers,
        data: sheetData.data,
      }));
      const detected = autoDetectColumns(sheetData.headers, jmsState.headers, jmsState.sheets);
      handleConfigChange(mergeRetainingSaved(detected, config));
    }
  };

  const handleJmsSheetChange = (sheetName: string) => {
    const sheetData = jmsState.sheets[sheetName];
    if (sheetData) {
      setJmsState((prev) => ({
        ...prev,
        selectedSheet: sheetName,
        headers: sheetData.headers,
        data: sheetData.data,
      }));
      const detected = autoDetectColumns(bridgeState.headers, sheetData.headers, jmsState.sheets);
      handleConfigChange(mergeRetainingSaved(detected, config));
    }
  };

  const handleReset = () => {
    setBridgeState(INITIAL_FILE_STATE);
    setJmsState(INITIAL_FILE_STATE);
  };

  // Auto-run on mount with sample data
  useEffect(() => {
    handleLoadSample();
  }, [handleLoadSample]);

  // Execute reconciliation engine whenever data, sheet records, or column configuration updates
  const reconciliation = useMemo(() => {
    if (!bridgeState.data.length && !jmsState.data.length) {
      return {
        matchedRows: [],
        unmatchedBridgeRows: [],
        unmatchedJmsRows: [],
        allRows: [],
        summary: {
          totalBridgeRows: 0,
          totalJmsRows: 0,
          matchedCount: 0,
          unmatchedBridgeCount: 0,
          unmatchedJmsCount: 0,
          matchRatePercent: 0,
          productionCrossMatchedCount: 0,
        },
      };
    }

    return reconcileWorkOrders(bridgeState.data, jmsState.data, config, jmsState.sheets);
  }, [bridgeState.data, jmsState.data, config, jmsState.sheets]);

  // Top banner Download Excel trigger
  const handleTopDownloadExcel = () => {
    let filterLabel = 'Matched';
    if (activeFilter === 'unmatched_bridge') filterLabel = 'Bridge_Unmatched';
    else if (activeFilter === 'matched_and_bridge_unmatched') filterLabel = 'Matched_and_Bridge_Unmatched';
    else if (activeFilter === 'unmatched_jms') filterLabel = 'JMS_Unmatched';
    else if (activeFilter === 'all') filterLabel = 'All_Records';

    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `Reconciled_Bridge_Work_Orders_${filterLabel}_${dateStr}.xlsx`;

    const targetRows =
      activeFilter === 'matched'
        ? reconciliation.matchedRows
        : activeFilter === 'unmatched_bridge'
        ? reconciliation.unmatchedBridgeRows
        : activeFilter === 'matched_and_bridge_unmatched'
        ? [...reconciliation.matchedRows, ...reconciliation.unmatchedBridgeRows]
        : activeFilter === 'unmatched_jms'
        ? reconciliation.unmatchedJmsRows
        : reconciliation.allRows;

    exportToExcel(targetRows, filename, undefined, highlightedWorkOrders);
  };

  // Dedicated one-click trigger to download Matched + Bridge Unmatched together
  const handleDownloadMatchedAndBridgeUnmatchedExcel = () => {
    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `Bridge_Work_Orders_Matched_and_Bridge_Unmatched_${dateStr}.xlsx`;
    const targetRows = [...reconciliation.matchedRows, ...reconciliation.unmatchedBridgeRows];
    exportToExcel(targetRows, filename, undefined, highlightedWorkOrders);
  };

  // Top banner Download Standalone HTML trigger
  const handleTopDownloadHTML = () => {
    let filterLabel = 'Matched Records';
    if (activeFilter === 'unmatched_bridge') filterLabel = 'Bridge Unmatched';
    else if (activeFilter === 'matched_and_bridge_unmatched') filterLabel = 'Matched & Bridge Unmatched';
    else if (activeFilter === 'unmatched_jms') filterLabel = 'JMS Unmatched';
    else if (activeFilter === 'all') filterLabel = 'All Records';

    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `Bridge_Work_Orders_${filterLabel.replace(/\s+/g, '_')}_${dateStr}.html`;

    exportToStandaloneHTML(
      {
        matchedRows: reconciliation.matchedRows,
        unmatchedBridgeRows: reconciliation.unmatchedBridgeRows,
        unmatchedJmsRows: reconciliation.unmatchedJmsRows,
        allRows: reconciliation.allRows,
        initialFilter: activeFilter,
        highlightedWorkOrders: Array.from(highlightedWorkOrders),
        bridgeFileName: bridgeState.fileName,
        bridgeSheetName: bridgeState.selectedSheet,
        jmsFileName: jmsState.fileName,
        jmsSheetName: jmsState.selectedSheet,
        rawBridgeData: bridgeState.data,
        rawJmsData: jmsState.data,
      },
      filterLabel,
      filename
    );
  };

  const hasLoadedData = bridgeState.data.length > 0 || jmsState.data.length > 0;

  return (
    <div className="min-h-screen bg-neutral-100/70 text-neutral-900 flex flex-col font-sans">
      <Header
        onLoadSample={handleLoadSample}
        onReset={handleReset}
        hasData={hasLoadedData}
      />

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* Intro Banner */}
        <div className="mb-6 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-5 rounded-xl border-2 border-neutral-300 shadow-xs">
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-base font-bold text-neutral-900 tracking-tight">
                Bridge Work Orders - All
              </h2>
              <span className="text-[10px] text-emerald-800 bg-emerald-50 border border-emerald-300 px-2 py-0.5 rounded font-bold">
                16 Mapped Headings (incl. Feedback Comments)
              </span>
            </div>
            <p className="text-xs text-neutral-600 mt-1 max-w-3xl leading-relaxed">
              Compare Bridge Work Orders and JMS/SaaS workbooks to match work orders with automated transformation rules,
              Production Day to Day cross-referencing, and export to Excel or Standalone HTML.
            </p>
          </div>

          <div className="flex items-center gap-2 shrink-0 flex-wrap">
            {reconciliation.matchedRows.length > 0 && (
              <>
                <button
                  type="button"
                  onClick={handleDownloadMatchedAndBridgeUnmatchedExcel}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-teal-900 bg-teal-100 hover:bg-teal-200 border border-teal-300 rounded-md transition-colors whitespace-nowrap shadow-xs"
                  title="Download combined Excel (.xlsx) of Matched and Bridge Unmatched work orders (all Base Bridge Table 1 rows)"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5 text-teal-700" />
                  <span>Download Matched + Bridge Unmatched (.xlsx)</span>
                  <span className="text-[11px] font-mono text-teal-800 font-semibold">
                    ({reconciliation.matchedRows.length + reconciliation.unmatchedBridgeRows.length})
                  </span>
                </button>

                <button
                  type="button"
                  onClick={handleTopDownloadHTML}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-indigo-900 bg-indigo-50 border border-indigo-300 hover:bg-indigo-100 rounded-md transition-colors whitespace-nowrap shadow-2xs"
                  title="Download standalone offline HTML report"
                >
                  <Globe className="w-3.5 h-3.5 text-indigo-600" />
                  <span>Download Standalone HTML</span>
                </button>

                <button
                  type="button"
                  onClick={handleTopDownloadExcel}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 rounded-md transition-colors whitespace-nowrap shadow-xs"
                  title="Download Excel (.xlsx) file of the current filtered view"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-200" />
                  <span>Download Excel</span>
                </button>
              </>
            )}

            <button
              type="button"
              onClick={handleLoadSample}
              className="px-3 py-1.5 text-xs font-medium text-neutral-700 bg-neutral-100 border border-neutral-300 hover:bg-neutral-200 rounded-md transition-colors whitespace-nowrap"
            >
              Reload Sample Data
            </button>
          </div>
        </div>

        {/* Upload Dropzones */}
        <div id="upload-section" className="grid grid-cols-1 md:grid-cols-2 gap-5 mb-6">
          <FileUploadCard
            title="1. Bridge Work Orders Excel"
            subtitle="Reads sheet 'Table 1' from 'Bridge work order - All' Excel (Sheet1 is automatically ignored)"
            expectedColumns={[
              'Work Order',
              'Work Order Title',
              'Due Date',
              'Job Type',
              'Activity',
              'XBRL Feedback Comments',
              'XBRL Feedback Quality',
              'XBRL Feedback Status',
              'XBRL Feedback Reviewed',
              'XBRL Feedback Validity',
            ]}
            fileState={bridgeState}
            onFileSelect={handleBridgeFileSelect}
            onSheetChange={handleBridgeSheetChange}
            onRemove={() => setBridgeState(INITIAL_FILE_STATE)}
          />

          <FileUploadCard
            title="2. JMS / SaaS Workbook"
            subtitle="Contains Main Tracker, X Track ID, Grabber Day to Day, and Production Day to Day sheet"
            expectedColumns={['Work Order / Job Number', 'X Track ID', 'Grabber Day to Day', 'Production Day to Day']}
            fileState={jmsState}
            onFileSelect={handleJmsFileSelect}
            onSheetChange={handleJmsSheetChange}
            onRemove={() => setJmsState(INITIAL_FILE_STATE)}
          />
        </div>

        {/* Column Mapping & Customization Drawer with persistent settings and reset */}
        <ColumnMappingDrawer
          bridgeHeaders={bridgeState.headers}
          jmsHeaders={jmsState.headers}
          jmsSheetNames={jmsState.sheetNames}
          jmsSheets={jmsState.sheets}
          config={config}
          onChangeConfig={handleConfigChange}
          onResetToAutoDetect={handleResetToAutoDetect}
        />

        {/* Specification Rules Drawer */}
        <RuleCard />

        {/* Summary Statistics */}
        <SummaryStats
          summary={reconciliation.summary}
          activeFilter={activeFilter}
          onSelectFilter={setActiveFilter}
        />

        {/* Main Results Table with prominent 'Download Excel' and 'Download Standalone HTML' */}
        <ReconciliationTable
          rows={
            activeFilter === 'matched'
              ? reconciliation.matchedRows
              : activeFilter === 'unmatched_bridge'
              ? reconciliation.unmatchedBridgeRows
              : activeFilter === 'matched_and_bridge_unmatched'
              ? [...reconciliation.matchedRows, ...reconciliation.unmatchedBridgeRows]
              : activeFilter === 'unmatched_jms'
              ? reconciliation.unmatchedJmsRows
              : reconciliation.allRows
          }
          allDatasets={{
            matchedRows: reconciliation.matchedRows,
            unmatchedBridgeRows: reconciliation.unmatchedBridgeRows,
            unmatchedJmsRows: reconciliation.unmatchedJmsRows,
            allRows: reconciliation.allRows,
          }}
          activeFilter={activeFilter}
          onSelectFilter={setActiveFilter}
          highlightedWorkOrders={highlightedWorkOrders}
          onUpdateHighlightedWorkOrders={handleUpdateHighlightedWorkOrders}
          bridgeFileState={bridgeState}
          jmsFileState={jmsState}
          onBridgeFileSelect={handleBridgeFileSelect}
          onJmsFileSelect={handleJmsFileSelect}
          onBridgeSheetChange={handleBridgeSheetChange}
          onJmsSheetChange={handleJmsSheetChange}
          onLoadSample={handleLoadSample}
        />
      </main>

      {/* Quiet Footer */}
      <footer className="w-full bg-white border-t border-neutral-200 py-4 mt-8">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-neutral-500">
          <span>Bridge Work Order &amp; JMS Reconciliation System</span>
          <div className="flex items-center gap-4">
            <span>Outputs 15 mapped headings</span>
            <span>·</span>
            <span>Persistent column mappings</span>
            <span>·</span>
            <span>Download Excel (.xlsx) via SheetJS</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
