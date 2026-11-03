import React, { useState, useMemo, useRef } from 'react';
import {
  Download,
  Copy,
  Check,
  Search,
  ChevronLeft,
  ChevronRight,
  Eye,
  FileSpreadsheet,
  X,
  Globe,
  SlidersHorizontal,
  Maximize2,
  Minimize2,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Sparkles,
  Upload,
  RefreshCw,
  FolderOpen,
  ChevronDown,
  Layers,
  FileText,
} from 'lucide-react';
import { MatchedOutputRow, ReconciliationFilterType, UploadedFileState } from '../types';
import {
  STANDARD_15_COLUMNS,
  XBRL_FEEDBACK_COLUMNS,
  QC_IMAGE_COLUMNS,
  ALL_COLUMNS,
  ColumnDefinition,
  exportToExcel,
  exportToCSV,
  copyTableToClipboard,
} from '../utils/excelExporter';
import { exportToStandaloneHTML } from '../utils/htmlExporter';
import { WorkOrderHighlightModal } from './WorkOrderHighlightModal';
import { normalizeWorkOrderKey, cleanFeedbackQuality } from '../utils/reconciliation';
import {
  SCREENSHOT_EXAMPLE_WORK_ORDERS,
  SCREENSHOT_CORE_5_WORK_ORDERS,
  SCREENSHOT_TABLE1_WORK_ORDERS,
} from '../data/sampleData';

interface ReconciliationTableProps {
  rows: MatchedOutputRow[];
  allDatasets?: {
    matchedRows: MatchedOutputRow[];
    unmatchedBridgeRows: MatchedOutputRow[];
    unmatchedJmsRows: MatchedOutputRow[];
    allRows: MatchedOutputRow[];
  };
  activeFilter: ReconciliationFilterType;
  onSelectFilter: (filter: ReconciliationFilterType) => void;
  highlightedWorkOrders?: Set<string>;
  onUpdateHighlightedWorkOrders?: (wos: Set<string>) => void;
  bridgeFileState?: UploadedFileState;
  jmsFileState?: UploadedFileState;
  onBridgeFileSelect?: (file: File) => void;
  onJmsFileSelect?: (file: File) => void;
  onBridgeSheetChange?: (sheetName: string) => void;
  onJmsSheetChange?: (sheetName: string) => void;
  onLoadSample?: () => void;
}

const COLUMN_STYLES: Record<string, { th: string; cell?: string }> = {
  // Base Details
  'Team': {
    th: 'bg-slate-100 text-slate-800 border-b-2 border-slate-300 font-semibold',
    cell: 'font-medium text-slate-900',
  },
  'JMS Job Number/ Work Order No': {
    th: 'bg-blue-100/80 text-blue-950 border-b-2 border-blue-400 font-bold',
    cell: 'font-mono tabular-nums font-bold text-blue-900',
  },
  'JMS Job Number/Work Order No': {
    th: 'bg-blue-100/80 text-blue-950 border-b-2 border-blue-400 font-bold',
    cell: 'font-mono tabular-nums font-bold text-blue-900',
  },
  'JMS Cycle No / FB Quality': {
    th: 'bg-teal-100/80 text-teal-950 border-b-2 border-teal-400 font-bold',
    cell: 'font-medium text-teal-950',
  },
  'XOR Ver': {
    th: 'bg-neutral-100/70 text-neutral-600 border-b-2 border-neutral-300 font-semibold',
    cell: 'text-neutral-400',
  },
  'X Track ID': {
    th: 'bg-sky-100/80 text-sky-950 border-b-2 border-sky-400 font-bold',
    cell: 'font-mono tabular-nums font-bold text-sky-900',
  },
  'Job Name': {
    th: 'bg-emerald-100/80 text-emerald-950 border-b-2 border-emerald-400 font-bold',
    cell: 'text-neutral-900 font-semibold',
  },
  'Form Type': {
    th: 'bg-amber-100/80 text-amber-950 border-b-2 border-amber-400 font-bold',
    cell: 'font-mono font-bold text-amber-900',
  },
  'Filing Period': {
    th: 'bg-neutral-100/70 text-neutral-600 border-b-2 border-neutral-300 font-semibold',
    cell: 'text-neutral-400',
  },
  'Milestone': {
    th: 'bg-purple-100/80 text-purple-950 border-b-2 border-purple-400 font-bold',
    cell: 'font-medium text-purple-950',
  },
  'JMS Team': {
    th: 'bg-slate-100 text-slate-800 border-b-2 border-slate-300 font-semibold',
    cell: 'font-medium text-slate-700',
  },
  'Sampling/Quality': {
    th: 'bg-indigo-100/80 text-indigo-950 border-b-2 border-indigo-400 font-bold',
    cell: 'font-semibold text-indigo-950',
  },
  'Consultant Name': {
    th: 'bg-neutral-100/70 text-neutral-600 border-b-2 border-neutral-300 font-semibold',
    cell: 'text-neutral-400',
  },
  '2nd Reviewer Name': {
    th: 'bg-neutral-100/70 text-neutral-600 border-b-2 border-neutral-300 font-semibold',
    cell: 'text-neutral-400',
  },
  'Delivered Date': {
    th: 'bg-rose-100/80 text-rose-950 border-b-2 border-rose-400 font-bold',
    cell: 'font-mono tabular-nums font-semibold text-rose-900',
  },
  'Month': {
    th: 'bg-rose-100/60 text-rose-950 border-b-2 border-rose-300 font-bold',
    cell: 'font-semibold text-rose-900',
  },

  // Feedback Audit Columns
  'XBRL Feedback Quality': {
    th: 'bg-teal-100 text-teal-950 border-b-2 border-teal-400 font-bold',
    cell: 'font-semibold text-teal-900',
  },
  'XBRL Feedback Status': {
    th: 'bg-emerald-100/80 text-emerald-950 border-b-2 border-emerald-400 font-semibold',
    cell: 'font-medium text-emerald-900',
  },
  'XBRL Feedback Reviewed': {
    th: 'bg-fuchsia-100/80 text-fuchsia-950 border-b-2 border-fuchsia-400 font-semibold',
    cell: 'font-medium text-fuchsia-900',
  },
  'XBRL Feedback Validity': {
    th: 'bg-blue-100/80 text-blue-950 border-b-2 border-blue-400 font-semibold',
    cell: 'font-medium text-blue-900',
  },
  'XBRL Feedback Comments': {
    th: 'bg-amber-100/70 text-amber-950 border-b-2 border-amber-400 font-semibold',
    cell: 'text-neutral-700 italic max-w-xs truncate',
  },

  // Secondary QC Reviewers
  'Production Consultant Name': {
    th: 'bg-amber-100/70 text-amber-950 border-b-2 border-amber-400 font-semibold',
    cell: 'text-neutral-800 font-medium',
  },
  'Internal Reviewer name': {
    th: 'bg-purple-100/70 text-purple-950 border-b-2 border-purple-400 font-semibold',
    cell: 'text-neutral-800 font-medium',
  },
  'External Reviewer Name': {
    th: 'bg-indigo-100/70 text-indigo-950 border-b-2 border-indigo-400 font-semibold',
    cell: 'text-neutral-800 font-medium',
  },

  // Error Categories
  'DEI Errors': {
    th: 'bg-red-50 text-red-950 border-b-2 border-red-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Element Errors': {
    th: 'bg-red-50 text-red-950 border-b-2 border-red-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Value Errors': {
    th: 'bg-red-50 text-red-950 border-b-2 border-red-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Untagged Errors': {
    th: 'bg-red-50 text-red-950 border-b-2 border-red-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'L1-L3 Errors': {
    th: 'bg-red-50 text-red-950 border-b-2 border-red-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Reverse Sign Errors': {
    th: 'bg-sky-100 text-sky-950 border-b-2 border-sky-400 font-bold',
    cell: 'font-mono text-right',
  },
  'Alternate Values Errors': {
    th: 'bg-sky-100 text-sky-950 border-b-2 border-sky-400 font-bold',
    cell: 'font-mono text-right',
  },
  'Date Context Errors': {
    th: 'bg-orange-50 text-orange-950 border-b-2 border-orange-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Member Context Errors': {
    th: 'bg-orange-50 text-orange-950 border-b-2 border-orange-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Scale/Precision Errors': {
    th: 'bg-orange-50 text-orange-950 border-b-2 border-orange-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Taxonomy/Structure/Invalid Spi': {
    th: 'bg-fuchsia-50 text-fuchsia-950 border-b-2 border-fuchsia-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Incorrect Mapping | Attribute': {
    th: 'bg-blue-100 text-blue-950 border-b-2 border-blue-400 font-bold',
    cell: 'font-mono text-right',
  },
  'Judgmental Errors': {
    th: 'bg-red-50 text-red-950 border-b-2 border-red-300 font-semibold',
    cell: 'font-mono text-right',
  },
  'Validation Errors': {
    th: 'bg-rose-100/70 text-rose-950 border-b-2 border-rose-400 font-bold',
    cell: 'font-mono text-right font-semibold',
  },
  'Other Errors': {
    th: 'bg-amber-100/70 text-amber-950 border-b-2 border-amber-400 font-bold',
    cell: 'font-mono text-right font-semibold',
  },
  'Label Errors': {
    th: 'bg-purple-100/70 text-purple-950 border-b-2 border-purple-400 font-bold',
    cell: 'font-mono text-right font-semibold',
  },
  'Total Errors': {
    th: 'bg-red-100 text-red-950 border-b-2 border-red-400 font-bold',
    cell: 'font-mono text-right font-bold text-red-700',
  },
};

export const ReconciliationTable: React.FC<ReconciliationTableProps> = ({
  rows,
  allDatasets,
  activeFilter,
  onSelectFilter,
  highlightedWorkOrders,
  onUpdateHighlightedWorkOrders,
  bridgeFileState,
  jmsFileState,
  onBridgeFileSelect,
  onJmsFileSelect,
  onBridgeSheetChange,
  onJmsSheetChange,
  onLoadSample,
}) => {
  const [viewMode, setViewMode] = useState<'standard' | 'feedback' | 'qc' | 'all'>('standard');
  const [searchQuery, setSearchQuery] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [copied, setCopied] = useState(false);
  const [compactMode, setCompactMode] = useState(false);
  const [showColumnsMenu, setShowColumnsMenu] = useState(false);
  const [hiddenColumns, setHiddenColumns] = useState<Set<string>>(new Set());
  const [sortKey, setSortKey] = useState<string>('workOrderNo');
  const [sortAsc, setSortAsc] = useState<boolean>(true);
  const [inspectingRow, setInspectingRow] = useState<MatchedOutputRow | null>(null);

  // Dedicated Work Order Highlighting State
  const [internalHighlightedWos, setInternalHighlightedWos] = useState<Set<string>>(new Set());
  const highlightedWos = highlightedWorkOrders ?? internalHighlightedWos;
  const setHighlightedWos = onUpdateHighlightedWorkOrders ?? setInternalHighlightedWos;

  const [isHighlightModalOpen, setIsHighlightModalOpen] = useState(false);
  const [onlyHighlighted, setOnlyHighlighted] = useState(false);
  const [inlineWoInput, setInlineWoInput] = useState('');
  const [showInputOptions, setShowInputOptions] = useState(true);

  const bridgeFileInputRef = useRef<HTMLInputElement>(null);
  const jmsFileInputRef = useRef<HTMLInputElement>(null);

  const handleAddInlineWo = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!inlineWoInput.trim()) return;
    const parts = inlineWoInput
      .split(/[\r\n,\t; ]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    if (parts.length === 0) return;
    const next = new Set(highlightedWos);
    parts.forEach((p) => next.add(normalizeWorkOrderKey(p)));
    setHighlightedWos(next);
    setInlineWoInput('');
  };

  const handleLoadBatch = (batch: string[]) => {
    const next = new Set(highlightedWos);
    batch.forEach((wo) => next.add(normalizeWorkOrderKey(wo)));
    setHighlightedWos(next);
  };

  const handleClearAllHighlights = () => {
    setHighlightedWos(new Set());
    setOnlyHighlighted(false);
  };

  // Available Work Orders for the modal
  const availableWorkOrdersInTable = useMemo(() => {
    return rows.map((r) => r.workOrderNo).filter(Boolean);
  }, [rows]);

  // Count of highlighted rows currently present in rows
  const highlightedCountInTable = useMemo(() => {
    if (highlightedWos.size === 0) return 0;
    return rows.filter((r) => highlightedWos.has(normalizeWorkOrderKey(r.workOrderNo))).length;
  }, [rows, highlightedWos]);

  // Active column set based on report view mode
  const baseColumns: ColumnDefinition[] = useMemo(() => {
    switch (viewMode) {
      case 'feedback':
        return XBRL_FEEDBACK_COLUMNS;
      case 'qc':
        return QC_IMAGE_COLUMNS;
      case 'all':
        return ALL_COLUMNS;
      case 'standard':
      default:
        return STANDARD_15_COLUMNS;
    }
  }, [viewMode]);

  // Filter columns based on user column-visibility toggle
  const activeColumns = useMemo(() => {
    return baseColumns.filter((c) => !hiddenColumns.has(c.key));
  }, [baseColumns, hiddenColumns]);

  // Toggle single column visibility
  const toggleColumn = (key: string) => {
    setHiddenColumns((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  // Sort handler
  const handleSort = (key: string) => {
    if (sortKey === key) {
      setSortAsc(!sortAsc);
    } else {
      setSortKey(key);
      setSortAsc(true);
    }
    setCurrentPage(1);
  };

  // Filter and sort rows
  const filteredAndSortedRows = useMemo(() => {
    let result = rows.filter((row) => {
      // Tab filter
      if (activeFilter === 'matched' && row.matchStatus !== 'matched') return false;
      if (activeFilter === 'unmatched_bridge' && row.matchStatus !== 'unmatched_bridge') return false;
      if (
        activeFilter === 'matched_and_bridge_unmatched' &&
        row.matchStatus !== 'matched' &&
        row.matchStatus !== 'unmatched_bridge'
      )
        return false;
      if (activeFilter === 'unmatched_jms' && row.matchStatus !== 'unmatched_jms') return false;

      // Work Order Focus Filter (if user wants to only see highlighted records)
      if (onlyHighlighted && highlightedWos.size > 0) {
        const norm = normalizeWorkOrderKey(row.workOrderNo);
        if (!highlightedWos.has(norm)) return false;
      }

      // Search query
      if (!searchQuery.trim()) return true;
      const q = searchQuery.toLowerCase();
      return (
        row.workOrderNo.toLowerCase().includes(q) ||
        row.jobName.toLowerCase().includes(q) ||
        row.xTrackId.toLowerCase().includes(q) ||
        row.formType.toLowerCase().includes(q) ||
        row.milestone.toLowerCase().includes(q) ||
        row.deliveredDate.toLowerCase().includes(q) ||
        row.month.toLowerCase().includes(q) ||
        (row.jmsCycleNoFbQuality && row.jmsCycleNoFbQuality.toLowerCase().includes(q)) ||
        (row.samplingQuality && row.samplingQuality.toLowerCase().includes(q)) ||
        (row.xbrlFeedbackStatus && row.xbrlFeedbackStatus.toLowerCase().includes(q)) ||
        (row.xbrlFeedbackComments && row.xbrlFeedbackComments.toLowerCase().includes(q)) ||
        (row.productionConsultantName && row.productionConsultantName.toLowerCase().includes(q)) ||
        (row.internalReviewerName && row.internalReviewerName.toLowerCase().includes(q)) ||
        (row.externalReviewerName && row.externalReviewerName.toLowerCase().includes(q))
      );
    });

    // Interactive sorting
    if (sortKey) {
      result = [...result].sort((a, b) => {
        let va = (a as any)[sortKey];
        let vb = (b as any)[sortKey];
        if (va === undefined || va === null) va = '';
        if (vb === undefined || vb === null) vb = '';
        if (typeof va === 'number' && typeof vb === 'number') {
          return sortAsc ? va - vb : vb - va;
        }
        const sa = String(va).toLowerCase();
        const sb = String(vb).toLowerCase();
        return sortAsc ? sa.localeCompare(sb) : sb.localeCompare(sa);
      });
    }

    return result;
  }, [rows, activeFilter, searchQuery, sortKey, sortAsc, onlyHighlighted, highlightedWos]);

  // Pagination calculation
  const totalPages = Math.max(1, Math.ceil(filteredAndSortedRows.length / pageSize));
  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredAndSortedRows.slice(start, start + pageSize);
  }, [filteredAndSortedRows, currentPage, pageSize]);

  const handleCopyClipboard = async () => {
    const success = await copyTableToClipboard(filteredAndSortedRows, activeColumns, highlightedWos);
    if (success) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const handleExportExcel = () => {
    let filterLabel = 'Matched';
    if (activeFilter === 'unmatched_bridge') filterLabel = 'Bridge_Unmatched';
    else if (activeFilter === 'matched_and_bridge_unmatched') filterLabel = 'Matched_and_Bridge_Unmatched';
    else if (activeFilter === 'unmatched_jms') filterLabel = 'JMS_Unmatched';
    else if (activeFilter === 'all') filterLabel = 'All_Records';

    const viewTag =
      viewMode === 'standard'
        ? '15_Headings'
        : viewMode === 'feedback'
        ? 'XBRL_Feedback'
        : viewMode === 'qc'
        ? 'QC_Details'
        : 'All_Columns';

    const searchSuffix = searchQuery.trim() ? '_Filtered' : '';
    const highlightSuffix = onlyHighlighted ? '_HighlightedOnly' : highlightedWos.size > 0 ? '_WithHighlights' : '';
    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `Bridge_Work_Orders_${viewTag}_${filterLabel}${searchSuffix}${highlightSuffix}_${dateStr}.xlsx`;
    exportToExcel(filteredAndSortedRows, filename, activeColumns, highlightedWos);
  };

  // Dedicated one-click trigger to download Matched + Bridge Unmatched together
  const handleExportMatchedAndBridgeUnmatchedExcel = () => {
    const combinedRows = allDatasets
      ? [...allDatasets.matchedRows, ...allDatasets.unmatchedBridgeRows]
      : rows.filter((r) => r.matchStatus === 'matched' || r.matchStatus === 'unmatched_bridge');

    const viewTag =
      viewMode === 'standard'
        ? '15_Headings'
        : viewMode === 'feedback'
        ? 'XBRL_Feedback'
        : viewMode === 'qc'
        ? 'QC_Details'
        : 'All_Columns';

    const highlightSuffix = highlightedWos.size > 0 ? '_WithHighlights' : '';
    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `Bridge_Work_Orders_${viewTag}_Matched_and_Bridge_Unmatched${highlightSuffix}_${dateStr}.xlsx`;
    exportToExcel(combinedRows, filename, activeColumns, highlightedWos);
  };

  const handleExportCSV = () => {
    let filterLabel = 'Matched';
    if (activeFilter === 'unmatched_bridge') filterLabel = 'Bridge_Unmatched';
    else if (activeFilter === 'matched_and_bridge_unmatched') filterLabel = 'Matched_and_Bridge_Unmatched';
    else if (activeFilter === 'unmatched_jms') filterLabel = 'JMS_Unmatched';
    else if (activeFilter === 'all') filterLabel = 'All_Records';

    const viewTag =
      viewMode === 'standard'
        ? '15_Headings'
        : viewMode === 'feedback'
        ? 'XBRL_Feedback'
        : viewMode === 'qc'
        ? 'QC_Details'
        : 'All_Columns';

    const searchSuffix = searchQuery.trim() ? '_Filtered' : '';
    const highlightSuffix = onlyHighlighted ? '_HighlightedOnly' : highlightedWos.size > 0 ? '_WithHighlights' : '';
    const dateStr = new Date().toISOString().slice(0, 10);
    exportToCSV(
      filteredAndSortedRows,
      `Bridge_Work_Orders_${viewTag}_${filterLabel}${searchSuffix}${highlightSuffix}_${dateStr}.csv`,
      activeColumns,
      highlightedWos
    );
  };

  const handleExportHTML = () => {
    let filterLabel = 'Matched Records';
    if (activeFilter === 'unmatched_bridge') filterLabel = 'Bridge Unmatched';
    else if (activeFilter === 'matched_and_bridge_unmatched') filterLabel = 'Matched & Bridge Unmatched';
    else if (activeFilter === 'unmatched_jms') filterLabel = 'JMS Unmatched';
    else if (activeFilter === 'all') filterLabel = 'All Records';

    const viewLabels = {
      standard: '16 Headings',
      feedback: 'XBRL Feedback',
      qc: 'QC Report',
      all: 'All Columns',
    };
    const viewTag = viewLabels[viewMode] || '16 Headings';
    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `Bridge_Work_Orders_${viewTag.replace(/\s+/g, '_')}_${filterLabel.replace(/\s+/g, '_')}_${dateStr}.html`;

    // Package all datasets into the standalone HTML file so that the downloaded HTML app has ALL options visible!
    exportToStandaloneHTML(
      {
        matchedRows: allDatasets?.matchedRows || rows.filter((r) => r.matchStatus === 'matched'),
        unmatchedBridgeRows: allDatasets?.unmatchedBridgeRows || rows.filter((r) => r.matchStatus === 'unmatched_bridge'),
        unmatchedJmsRows: allDatasets?.unmatchedJmsRows || rows.filter((r) => r.matchStatus === 'unmatched_jms'),
        allRows: allDatasets?.allRows || rows,
        initialFilter: activeFilter,
        initialView: viewMode,
        highlightedWorkOrders: Array.from(highlightedWos),
        bridgeFileName: bridgeFileState?.fileName,
        bridgeSheetName: bridgeFileState?.selectedSheet,
        jmsFileName: jmsFileState?.fileName,
        jmsSheetName: jmsFileState?.selectedSheet,
        rawBridgeData: bridgeFileState?.data,
        rawJmsData: jmsFileState?.data,
      },
      `${filterLabel} (${viewTag})`,
      filename,
      activeColumns
    );
  };

  return (
    <div id="results-section" className="bg-white border-2 border-neutral-300 rounded-xl shadow-xs overflow-hidden">
      {/* View Switcher Sub-header */}
      <div className="px-4 py-2.5 bg-neutral-100/90 border-b border-neutral-300 flex flex-col md:flex-row md:items-center justify-between gap-2.5">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-xs font-bold text-neutral-800">Report View:</span>
          <div className="inline-flex rounded-md shadow-2xs p-0.5 bg-neutral-200 border border-neutral-300 flex-wrap">
            <button
              type="button"
              onClick={() => {
                setViewMode('standard');
                setHiddenColumns(new Set());
                setCurrentPage(1);
              }}
              className={`px-3 py-1 text-xs font-semibold rounded transition-all ${
                viewMode === 'standard'
                  ? 'bg-white text-emerald-950 font-bold shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              Standard Output Headings (16 Cols)
            </button>
            <button
              type="button"
              onClick={() => {
                setViewMode('feedback');
                setHiddenColumns(new Set());
                setCurrentPage(1);
              }}
              className={`px-3 py-1 text-xs font-semibold rounded transition-all ${
                viewMode === 'feedback'
                  ? 'bg-white text-teal-950 font-bold shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              XBRL Feedback Audit View
            </button>
            <button
              type="button"
              onClick={() => {
                setViewMode('qc');
                setHiddenColumns(new Set());
                setCurrentPage(1);
              }}
              className={`px-3 py-1 text-xs font-semibold rounded transition-all ${
                viewMode === 'qc'
                  ? 'bg-white text-blue-950 font-bold shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              QC &amp; Secondary Details (24 Cols)
            </button>
            <button
              type="button"
              onClick={() => {
                setViewMode('all');
                setHiddenColumns(new Set());
                setCurrentPage(1);
              }}
              className={`px-3 py-1 text-xs font-semibold rounded transition-all ${
                viewMode === 'all'
                  ? 'bg-white text-indigo-950 font-bold shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              All Columns View
            </button>
          </div>
        </div>

        <div className="flex items-center gap-2 text-xs text-neutral-600 flex-wrap">
          <span>
            Active Columns: <strong className="text-neutral-900 font-mono">{activeColumns.length}</strong>
          </span>
          <span>·</span>
          <span>
            {viewMode === 'standard' && <strong className="text-emerald-800">16 Configured Headings (incl. Feedback Comments)</strong>}
            {viewMode === 'feedback' && <strong className="text-teal-800">Feedback Quality &amp; Comments Mapped</strong>}
            {viewMode === 'qc' && <span>Matched on <strong className="text-sky-800 font-mono">X Track ID</strong></span>}
            {viewMode === 'all' && <strong className="text-indigo-800">Complete Master Dataset</strong>}
          </span>
          <button
            type="button"
            onClick={() => setShowInputOptions(!showInputOptions)}
            className={`inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold rounded-md transition-colors ml-2 shadow-2xs ${
              showInputOptions
                ? 'bg-amber-200 text-amber-950 border border-amber-400'
                : 'bg-white text-neutral-700 hover:bg-neutral-50 border border-neutral-300'
            }`}
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-600" />
            <span>{showInputOptions ? 'Hide Input Options' : 'Show Input Options'}</span>
            {highlightedWos.size > 0 && (
              <span className="px-1.5 py-0.2 text-[10px] font-bold rounded-full bg-amber-400 text-amber-950 font-mono">
                {highlightedWos.size}
              </span>
            )}
          </button>
        </div>
      </div>

      {/* Dedicated Input Options Panel (Work Orders & File Inputs) */}
      {showInputOptions && (
        <div className="p-3.5 bg-amber-50/70 border-b border-amber-200 space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-amber-950 uppercase tracking-wider flex items-center gap-1.5">
                <Sparkles className="w-4 h-4 text-amber-600" />
                <span>Interactive Table Input Options</span>
              </span>
              <span className="text-[10px] text-amber-900 bg-amber-200/80 border border-amber-300 px-2 py-0.5 rounded-full font-semibold">
                Direct Work Order &amp; File Inputs
              </span>
            </div>

            {onLoadSample && (
              <button
                type="button"
                onClick={onLoadSample}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-medium text-neutral-700 bg-white border border-neutral-300 hover:bg-neutral-100 rounded-md transition-colors shadow-2xs self-start sm:self-auto"
                title="Reload the canonical sample data"
              >
                <RefreshCw className="w-3 h-3 text-neutral-500" />
                <span>Reload Sample Data</span>
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 pt-1">
            {/* 1. Work Order IDs Input Option */}
            <div className="lg:col-span-7 bg-white p-3 rounded-lg border border-amber-200/90 shadow-2xs space-y-2.5">
              <div className="flex items-center justify-between flex-wrap gap-1">
                <label className="text-[11px] font-bold text-amber-950 uppercase tracking-wider flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-amber-500" />
                  <span>Work Order Input Option (Highlight in Yellow)</span>
                </label>
                <button
                  type="button"
                  onClick={() => setIsHighlightModalOpen(true)}
                  className="text-[11px] font-semibold text-amber-900 hover:text-amber-950 underline underline-offset-2"
                >
                  Open Full Input Modal...
                </button>
              </div>

              {/* Batch Load Buttons */}
              <div className="flex items-center gap-1.5 flex-wrap">
                <span className="text-[11px] font-semibold text-neutral-600">Load Screenshot IDs:</span>
                <button
                  type="button"
                  onClick={() => handleLoadBatch(SCREENSHOT_CORE_5_WORK_ORDERS)}
                  className="px-2.5 py-1 text-[11px] font-semibold rounded bg-amber-100 text-amber-900 hover:bg-amber-200 border border-amber-300 transition-colors shadow-2xs"
                  title="Load the 5 core work orders from screenshot (WO-26-003172, WO-26-003125, etc.)"
                >
                  + Core 5 IDs
                </button>
                <button
                  type="button"
                  onClick={() => handleLoadBatch(SCREENSHOT_TABLE1_WORK_ORDERS)}
                  className="px-2.5 py-1 text-[11px] font-semibold rounded bg-amber-100 text-amber-900 hover:bg-amber-200 border border-amber-300 transition-colors shadow-2xs"
                  title="Load the 20 Table 1 work orders from Bridge screenshot"
                >
                  + Table 1 (20 IDs)
                </button>
                <button
                  type="button"
                  onClick={() => handleLoadBatch(SCREENSHOT_EXAMPLE_WORK_ORDERS)}
                  className="px-2.5 py-1 text-[11px] font-semibold rounded bg-amber-100 text-amber-900 hover:bg-amber-200 border border-amber-300 transition-colors shadow-2xs"
                  title="Load all 15 example screenshot work orders"
                >
                  + All Screenshot (15 IDs)
                </button>
              </div>

              {/* Inline Direct Input Bar */}
              <form onSubmit={handleAddInlineWo} className="flex items-center gap-1.5">
                <input
                  type="text"
                  value={inlineWoInput}
                  onChange={(e) => setInlineWoInput(e.target.value)}
                  placeholder="Type or paste Work Order ID (e.g. WO-26-003172 or comma-separated)..."
                  className="flex-1 text-xs px-2.5 py-1.5 bg-neutral-50 border border-neutral-300 rounded-md focus:bg-white focus:outline-hidden focus:ring-1 focus:ring-amber-500 focus:border-amber-500 font-mono"
                />
                <button
                  type="submit"
                  className="px-3 py-1.5 text-xs font-bold text-amber-950 bg-amber-400 hover:bg-amber-500 rounded-md transition-colors shadow-2xs whitespace-nowrap"
                >
                  + Highlight ID
                </button>
                {highlightedWos.size > 0 && (
                  <button
                    type="button"
                    onClick={handleClearAllHighlights}
                    className="px-2.5 py-1.5 text-xs font-medium text-red-700 bg-red-50 hover:bg-red-100 border border-red-200 rounded-md transition-colors whitespace-nowrap"
                    title="Clear all yellow highlighted work orders"
                  >
                    Clear
                  </button>
                )}
              </form>

              {/* Status and Show Highlighted Only Toggle */}
              {highlightedWos.size > 0 && (
                <div className="flex items-center justify-between text-[11px] pt-1.5 border-t border-neutral-100 flex-wrap gap-2">
                  <span className="text-amber-900 font-medium">
                    <strong>{highlightedWos.size}</strong> work orders loaded ({highlightedCountInTable} in current view highlighted in yellow)
                  </span>
                  <label className="flex items-center gap-1.5 font-semibold text-amber-950 cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={onlyHighlighted}
                      onChange={(e) => setOnlyHighlighted(e.target.checked)}
                      className="rounded text-amber-600 focus:ring-amber-500 h-3.5 w-3.5 cursor-pointer"
                    />
                    <span>Show Highlighted Only</span>
                  </label>
                </div>
              )}
            </div>

            {/* 2. File Input Options */}
            <div className="lg:col-span-5 bg-white p-3 rounded-lg border border-neutral-200 shadow-2xs space-y-2">
              <label className="text-[11px] font-bold text-neutral-800 uppercase tracking-wider block">
                File Input Options (Bridge &amp; JMS)
              </label>

              {/* Bridge File Input */}
              <div className="flex items-center justify-between text-xs gap-2 p-2 bg-neutral-50 rounded border border-neutral-200">
                <div className="min-w-0">
                  <span className="font-semibold text-neutral-900 block truncate">
                    1. Bridge: {bridgeFileState?.fileName || 'Bridge work order - All'}
                  </span>
                  <span className="text-[10px] text-teal-800 font-mono font-medium">
                    Sheet: {bridgeFileState?.selectedSheet || 'Table 1'} ({bridgeFileState?.data.length || rows.length} rows)
                  </span>
                </div>
                <div className="shrink-0 flex items-center gap-1">
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    ref={bridgeFileInputRef}
                    onChange={(e) => {
                      if (e.target.files && e.target.files[0] && onBridgeFileSelect) {
                        onBridgeFileSelect(e.target.files[0]);
                      }
                    }}
                    className="hidden"
                  />
                  <button
                    type="button"
                    onClick={() => bridgeFileInputRef.current?.click()}
                    className="px-2.5 py-1 text-[11px] font-semibold text-neutral-800 bg-white border border-neutral-300 hover:bg-neutral-100 rounded transition-colors shadow-2xs whitespace-nowrap"
                  >
                    Change
                  </button>
                </div>
              </div>

              {/* JMS File Input */}
              <div className="flex items-center justify-between text-xs gap-2 p-2 bg-neutral-50 rounded border border-neutral-200">
                <div className="min-w-0">
                  <span className="font-semibold text-neutral-900 block truncate">
                    2. JMS: {jmsFileState?.fileName || 'JMS Workbook'}
                  </span>
                  <span className="text-[10px] text-blue-800 font-mono font-medium">
                    Sheet: {jmsFileState?.selectedSheet || 'Active / Grabber'} ({jmsFileState?.data.length || 0} rows)
                  </span>
                </div>
                <div className="shrink-0 flex items-center gap-1">
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    ref={jmsFileInputRef}
                    onChange={(e) => {
                      if (e.target.files && e.target.files[0] && onJmsFileSelect) {
                        onJmsFileSelect(e.target.files[0]);
                      }
                    }}
                    className="hidden"
                  />
                  <button
                    type="button"
                    onClick={() => jmsFileInputRef.current?.click()}
                    className="px-2.5 py-1 text-[11px] font-semibold text-neutral-800 bg-white border border-neutral-300 hover:bg-neutral-100 rounded transition-colors shadow-2xs whitespace-nowrap"
                  >
                    Change
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Table Toolbar */}
      <div className="p-3.5 border-b border-neutral-300 bg-neutral-50/80 space-y-3">
        {/* Row 1: Filter Tabs and Export Actions */}
        <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-3">
          {/* Left: Tab selection */}
          <div className="flex items-center gap-1 p-1 bg-neutral-200/80 rounded-lg shrink-0 border border-neutral-300/80 flex-wrap">
            <button
              type="button"
              onClick={() => {
                onSelectFilter('matched');
                setCurrentPage(1);
              }}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                activeFilter === 'matched'
                  ? 'bg-white text-neutral-900 shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              Matched Records ({allDatasets?.matchedRows.length ?? rows.filter((r) => r.matchStatus === 'matched').length})
            </button>
            <button
              type="button"
              onClick={() => {
                onSelectFilter('unmatched_bridge');
                setCurrentPage(1);
              }}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                activeFilter === 'unmatched_bridge'
                  ? 'bg-white text-neutral-900 shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              Bridge Unmatched ({allDatasets?.unmatchedBridgeRows.length ?? rows.filter((r) => r.matchStatus === 'unmatched_bridge').length})
            </button>
            <button
              type="button"
              onClick={() => {
                onSelectFilter('matched_and_bridge_unmatched');
                setCurrentPage(1);
              }}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                activeFilter === 'matched_and_bridge_unmatched'
                  ? 'bg-teal-700 text-white shadow-xs font-bold'
                  : 'text-teal-900 bg-teal-50/80 hover:bg-teal-100 border border-teal-300 font-semibold'
              }`}
              title="Combined view of Matched and Bridge Unmatched work orders (100% of Base Bridge Table 1 records)"
            >
              Matched + Bridge Unmatched ({
                (allDatasets?.matchedRows.length ?? rows.filter((r) => r.matchStatus === 'matched').length) +
                (allDatasets?.unmatchedBridgeRows.length ?? rows.filter((r) => r.matchStatus === 'unmatched_bridge').length)
              })
            </button>
            <button
              type="button"
              onClick={() => {
                onSelectFilter('unmatched_jms');
                setCurrentPage(1);
              }}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                activeFilter === 'unmatched_jms'
                  ? 'bg-white text-neutral-900 shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              JMS Unmatched ({allDatasets?.unmatchedJmsRows.length ?? rows.filter((r) => r.matchStatus === 'unmatched_jms').length})
            </button>
            <button
              type="button"
              onClick={() => {
                onSelectFilter('all');
                setCurrentPage(1);
              }}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                activeFilter === 'all'
                  ? 'bg-white text-neutral-900 shadow-xs'
                  : 'text-neutral-700 hover:text-neutral-900'
              }`}
            >
              All ({allDatasets?.allRows.length ?? rows.length})
            </button>
          </div>

          {/* Right: Actions */}
          <div className="flex items-center gap-1.5 flex-wrap">
            {/* Fit to screen toggle */}
            <button
              type="button"
              onClick={() => setCompactMode(!compactMode)}
              className={`inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-md border transition-colors whitespace-nowrap ${
                compactMode
                  ? 'bg-neutral-800 text-white border-neutral-800'
                  : 'bg-white text-neutral-700 border-neutral-300 hover:bg-neutral-50'
              }`}
              title="Toggle Compact Mode to fit columns on screen"
            >
              {compactMode ? <Minimize2 className="w-3.5 h-3.5" /> : <Maximize2 className="w-3.5 h-3.5" />}
              <span>{compactMode ? 'Normal View' : 'Fit to Screen'}</span>
            </button>

            {/* Column Picker Dropdown */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setShowColumnsMenu(!showColumnsMenu)}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-neutral-700 bg-white border border-neutral-300 hover:bg-neutral-50 rounded-md transition-colors whitespace-nowrap"
                title="Select columns to show or hide"
              >
                <SlidersHorizontal className="w-3.5 h-3.5 text-neutral-500" />
                <span>Columns ({activeColumns.length})</span>
              </button>

              {showColumnsMenu && (
                <div className="absolute right-0 top-full mt-1.5 w-60 bg-white border border-neutral-300 rounded-lg shadow-xl p-2 z-50 max-h-72 overflow-y-auto text-xs">
                  <div className="flex items-center justify-between pb-1.5 mb-1.5 border-b border-neutral-200">
                    <span className="font-bold text-neutral-800 text-[11px] uppercase tracking-wider">
                      Toggle Columns
                    </span>
                    <button
                      type="button"
                      onClick={() => setHiddenColumns(new Set())}
                      className="text-[10px] text-blue-600 hover:underline"
                    >
                      Reset All
                    </button>
                  </div>
                  <div className="space-y-1">
                    {baseColumns.map((col) => {
                      const isChecked = !hiddenColumns.has(col.key);
                      return (
                        <label
                          key={col.key}
                          className="flex items-center gap-2 px-1.5 py-1 rounded hover:bg-neutral-100 cursor-pointer text-neutral-800"
                        >
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => toggleColumn(col.key)}
                            className="rounded text-blue-600 focus:ring-blue-500 h-3.5 w-3.5"
                          />
                          <span className="truncate">{col.header}</span>
                        </label>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            {/* Copy TSV */}
            <button
              type="button"
              onClick={handleCopyClipboard}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-neutral-700 bg-white border border-neutral-300 hover:bg-neutral-50 rounded-md transition-colors whitespace-nowrap"
              title="Copy table to clipboard as TSV to paste directly into Excel"
            >
              {copied ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                  <span className="text-emerald-700">Copied!</span>
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5 text-neutral-500" />
                  <span>Copy TSV</span>
                </>
              )}
            </button>

            {/* CSV */}
            <button
              type="button"
              onClick={handleExportCSV}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-neutral-700 bg-white border border-neutral-300 hover:bg-neutral-50 rounded-md transition-colors whitespace-nowrap"
              title="Export as CSV"
            >
              <Download className="w-3.5 h-3.5 text-neutral-500" />
              <span>CSV</span>
            </button>

            {/* Standalone HTML Export */}
            <button
              type="button"
              onClick={handleExportHTML}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-indigo-900 bg-indigo-50 border border-indigo-300 hover:bg-indigo-100 rounded-md transition-colors whitespace-nowrap shadow-2xs"
              title="Download standalone HTML report that can be opened offline in any browser with full tabs & views"
            >
              <Globe className="w-3.5 h-3.5 text-indigo-600" />
              <span>Download Standalone HTML</span>
            </button>

            {/* Dedicated Matched + Bridge Unmatched Excel Download Button */}
            <button
              type="button"
              onClick={handleExportMatchedAndBridgeUnmatchedExcel}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-teal-950 bg-teal-100 hover:bg-teal-200 border border-teal-300 rounded-md transition-colors whitespace-nowrap shadow-xs"
              title="Download Excel (.xlsx) combining Matched and Bridge Unmatched work orders (all Base Bridge Table 1 rows) with yellow highlights"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-teal-700" />
              <span>Download Matched + Bridge Unmatched (.xlsx)</span>
              <span className="text-[11px] font-mono text-teal-800 font-semibold">
                ({(allDatasets?.matchedRows.length ?? rows.filter((r) => r.matchStatus === 'matched').length) +
                  (allDatasets?.unmatchedBridgeRows.length ?? rows.filter((r) => r.matchStatus === 'unmatched_bridge').length)})
              </span>
            </button>

            {/* Primary Excel (.xlsx) Download */}
            <button
              type="button"
              onClick={handleExportExcel}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 rounded-md transition-colors whitespace-nowrap shadow-xs"
              title="Download current view as .xlsx file using SheetJS"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-200" />
              <span>Download Excel</span>
              <span className="text-[11px] font-mono text-emerald-100 font-normal">
                ({filteredAndSortedRows.length})
              </span>
            </button>
          </div>
        </div>

        {/* Row 2: Search Bar & Dedicated Work Order Highlighting Controls */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 pt-2.5 border-t border-neutral-200">
          {/* Search Bar */}
          <div className="relative min-w-[240px] max-w-sm flex-1">
            <Search className="w-3.5 h-3.5 text-neutral-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setCurrentPage(1);
              }}
              placeholder="Search work order, X Track ID, reviewer..."
              className="w-full text-xs pl-8 pr-7 py-1.5 bg-white border border-neutral-300 rounded-md focus:outline-hidden focus:ring-1 focus:ring-blue-500 focus:border-blue-500 transition-colors"
            />
            {searchQuery && (
              <button
                type="button"
                onClick={() => {
                  setSearchQuery('');
                  setCurrentPage(1);
                }}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-neutral-400 hover:text-neutral-600"
              >
                <X className="w-3 h-3" />
              </button>
            )}
          </div>

          {/* Work Order Highlighting Controls */}
          <div className="flex items-center gap-2 flex-wrap">
            {/* Main Highlight Input Button */}
            <button
              type="button"
              onClick={() => setIsHighlightModalOpen(true)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-md border transition-all whitespace-nowrap shadow-2xs ${
                highlightedWos.size > 0
                  ? 'bg-amber-100 text-amber-950 border-amber-400 hover:bg-amber-200 ring-2 ring-amber-300/40'
                  : 'bg-white text-neutral-800 border-neutral-300 hover:bg-amber-50 hover:border-amber-400'
              }`}
              title="Click to input or paste Work Order IDs to highlight in yellow"
            >
              <Sparkles className={`w-3.5 h-3.5 ${highlightedWos.size > 0 ? 'text-amber-700' : 'text-amber-500'}`} />
              <span>Highlight Work Orders</span>
              {highlightedWos.size > 0 ? (
                <span className="px-1.5 py-0.2 text-[10px] font-bold rounded-full bg-amber-300 text-amber-950">
                  {highlightedWos.size}
                </span>
              ) : (
                <span className="text-[10px] font-normal text-neutral-500">(Input IDs)</span>
              )}
            </button>

            {/* Quick Multiple IDs Sample Button when empty */}
            {highlightedWos.size === 0 && (
              <button
                type="button"
                onClick={() => {
                  setHighlightedWos(new Set(SCREENSHOT_EXAMPLE_WORK_ORDERS.map(normalizeWorkOrderKey)));
                }}
                className="hidden sm:inline-flex items-center gap-1.5 text-[11px] font-semibold text-amber-900 bg-amber-50 hover:bg-amber-100 border border-amber-300 px-2.5 py-1.5 rounded-md transition-colors shadow-2xs"
                title="Quickly load all 15 multiple screenshot Work Order IDs from the user screenshots"
              >
                <Sparkles className="w-3 h-3 text-amber-700" />
                <span>+ Load Screenshot IDs (15 IDs)</span>
              </button>
            )}

            {/* Active Highlight Info & Focus Filter */}
            {highlightedWos.size > 0 && (
              <div className="flex items-center gap-2 px-2.5 py-1 bg-amber-50 border border-amber-300 rounded-md text-xs shadow-2xs">
                <span className="font-bold text-amber-900">
                  ⭐ {highlightedCountInTable} of {highlightedWos.size} found
                </span>
                <label className="flex items-center gap-1.5 pl-2 border-l border-amber-300 cursor-pointer text-amber-950 select-none font-semibold">
                  <input
                    type="checkbox"
                    checked={onlyHighlighted}
                    onChange={(e) => {
                      setOnlyHighlighted(e.target.checked);
                      setCurrentPage(1);
                    }}
                    className="rounded text-amber-600 focus:ring-amber-500 h-3.5 w-3.5"
                  />
                  <span>Show Highlighted Only</span>
                </label>
                <button
                  type="button"
                  onClick={() => {
                    setHighlightedWos(new Set());
                    setOnlyHighlighted(false);
                  }}
                  className="text-amber-700 hover:text-amber-950 p-0.5 rounded hover:bg-amber-200 transition-colors ml-1"
                  title="Clear all highlighted Work Orders"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Table Container with colorful headings, sticky freeze columns, and sorting */}
      <div className="overflow-x-auto max-h-[600px] relative border-t border-neutral-200">
        <table className={`w-full text-left border-collapse ${compactMode ? 'text-[11px]' : 'text-xs'}`}>
          <thead className="sticky top-0 z-20 shadow-2xs">
            <tr>
              {/* Sticky # Column */}
              <th className="py-2.5 px-3 font-semibold text-slate-700 whitespace-nowrap border-r border-neutral-300 border-b-2 border-slate-300 bg-slate-100 w-10 text-center sticky left-0 z-30">
                #
              </th>

              {activeColumns.map((col) => {
                const style = COLUMN_STYLES[col.header] || {
                  th: 'bg-neutral-100 text-neutral-800 border-b-2 border-neutral-300',
                };
                const isStickyWo = col.key === 'workOrderNo';
                const isSorted = sortKey === col.key;

                return (
                  <th
                    key={col.header}
                    onClick={() => handleSort(col.key)}
                    className={`py-2.5 px-3 font-semibold whitespace-nowrap border-r border-neutral-300/80 last:border-r-0 cursor-pointer select-none transition-colors hover:brightness-95 ${
                      style.th
                    } ${isStickyWo ? 'sticky left-10 z-30 shadow-[2px_0_4px_-1px_rgba(0,0,0,0.08)]' : ''}`}
                  >
                    <div className="flex items-center gap-1.5">
                      <span>{col.header}</span>
                      <span className="text-neutral-500">
                        {isSorted ? (
                          sortAsc ? <ArrowUp className="w-3 h-3 text-neutral-900" /> : <ArrowDown className="w-3 h-3 text-neutral-900" />
                        ) : (
                          <ArrowUpDown className="w-2.5 h-2.5 opacity-30 group-hover:opacity-100" />
                        )}
                      </span>
                    </div>
                  </th>
                );
              })}

              <th className="py-2.5 px-3 font-semibold text-neutral-700 whitespace-nowrap border-b-2 border-neutral-300 bg-neutral-100 w-16 text-center">
                Details
              </th>
            </tr>
          </thead>
          <tbody className="divide-y divide-neutral-200 text-neutral-800">
            {paginatedRows.length === 0 ? (
              <tr>
                <td
                  colSpan={activeColumns.length + 2}
                  className="py-12 text-center text-neutral-500 bg-white"
                >
                  <p className="text-sm font-medium text-neutral-700">No records found</p>
                  <p className="text-xs text-neutral-400 mt-1">
                    {searchQuery ? 'Try clearing your search query' : 'Upload Bridge Work Orders Excel and JMS Workbook to view matched data'}
                  </p>
                </td>
              </tr>
            ) : (
              paginatedRows.map((row, idx) => {
                const rowIndex = (currentPage - 1) * pageSize + idx + 1;
                const isMatched = row.matchStatus === 'matched';
                const isBridgeOnly = row.matchStatus === 'unmatched_bridge';
                const isHighlighted = highlightedWos.has(normalizeWorkOrderKey(row.workOrderNo));

                return (
                  <tr
                    key={row.id}
                    className={`transition-colors ${
                      isHighlighted
                        ? 'bg-amber-100 hover:bg-amber-200/90 text-amber-950 border-l-4 border-l-amber-500 font-medium'
                        : !isMatched
                        ? isBridgeOnly
                          ? 'bg-amber-50/40 hover:bg-amber-100/40'
                          : 'bg-blue-50/40 hover:bg-blue-100/40'
                        : 'even:bg-neutral-50/60 hover:bg-blue-50/40'
                    }`}
                  >
                    {/* Sticky Row Index */}
                    <td
                      className={`py-2.5 px-3 font-mono tabular-nums text-center border-r border-neutral-200 sticky left-0 z-10 ${
                        isHighlighted
                          ? 'bg-amber-200/90 text-amber-950 font-bold border-r-amber-300'
                          : 'bg-neutral-50 text-neutral-400'
                      }`}
                    >
                      {rowIndex}
                    </td>

                    {activeColumns.map((col) => {
                      const val = (row as any)[col.key];
                      const style = COLUMN_STYLES[col.header];
                      const isNum = typeof val === 'number';
                      const isStickyWo = col.key === 'workOrderNo';

                      return (
                        <td
                          key={col.key}
                          className={`py-2.5 px-3 border-r border-neutral-200 whitespace-nowrap ${
                            style?.cell || ''
                          } ${
                            isStickyWo
                              ? isHighlighted
                                ? 'sticky left-10 z-10 bg-amber-100 font-bold text-amber-950 border-r-amber-300 shadow-[2px_0_4px_-1px_rgba(217,119,6,0.15)]'
                                : 'sticky left-10 z-10 bg-white shadow-[2px_0_4px_-1px_rgba(0,0,0,0.08)]'
                              : isHighlighted
                              ? 'bg-amber-50/60'
                              : ''
                          } ${
                            isNum && val > 0
                              ? 'text-red-700 font-bold bg-red-50/40'
                              : isNum
                              ? 'text-neutral-400'
                              : ''
                          }`}
                        >
                          {col.key === 'workOrderNo' ? (
                            <div className="flex items-center gap-1.5">
                              <span className={isHighlighted ? 'text-amber-950 font-bold' : ''}>{val}</span>
                              {isHighlighted && (
                                <span className="inline-flex items-center gap-0.5 text-[10px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-amber-300 text-amber-950 border border-amber-400 shadow-2xs">
                                  Highlighted
                                </span>
                              )}
                            </div>
                          ) : col.key === 'xTrackId' ? (
                            <div className="flex items-center gap-1.5">
                              <span>{val || <span className="text-neutral-400 italic font-normal">No match</span>}</span>
                              {row.secondaryReportType && row.secondaryReportType !== 'None' && (
                                <span className="text-[10px] font-sans text-emerald-800 bg-emerald-50 border border-emerald-200 px-1.5 py-0.2 rounded font-medium" title={`Enriched from ${row.secondaryReportType} report`}>
                                  {row.secondaryReportType}
                                </span>
                              )}
                            </div>
                          ) : col.key === 'jmsCycleNoFbQuality' && val ? (
                            <span className="inline-block bg-teal-50 border border-teal-200 text-teal-900 font-semibold px-2 py-0.5 rounded">
                              {cleanFeedbackQuality(val)}
                            </span>
                          ) : col.key === 'samplingQuality' && val ? (
                            <span className="inline-block bg-indigo-50 border border-indigo-200 text-indigo-900 font-semibold px-2 py-0.5 rounded">
                              {val}
                            </span>
                          ) : col.key === 'formType' && val ? (
                            <span className="inline-block bg-amber-50 border border-amber-200/80 px-2 py-0.5 rounded">
                              {val}
                            </span>
                          ) : col.key === 'xbrlFeedbackQuality' && val ? (
                            <span className="inline-block bg-teal-50 border border-teal-200 text-teal-900 font-semibold px-2 py-0.5 rounded">
                              {cleanFeedbackQuality(val)}
                            </span>
                          ) : col.key === 'xbrlFeedbackStatus' && val ? (
                            <span className="inline-block bg-emerald-50 border border-emerald-200 text-emerald-900 font-semibold px-2 py-0.5 rounded">
                              {val}
                            </span>
                          ) : col.key === 'xbrlFeedbackComments' && val ? (
                            <span className="italic text-neutral-800 max-w-xs block truncate" title={String(val)}>
                              {val}
                            </span>
                          ) : (
                            val !== undefined && val !== null && val !== '' ? val : '—'
                          )}
                        </td>
                      );
                    })}

                    {/* Inspect Action */}
                    <td className="py-2.5 px-3 text-center whitespace-nowrap">
                      <button
                        type="button"
                        onClick={() => setInspectingRow(row)}
                        className="text-neutral-500 hover:text-neutral-900 p-1.5 rounded hover:bg-neutral-200/80 transition-colors"
                        title="Inspect original source fields and rules"
                      >
                        <Eye className="w-4 h-4" />
                      </button>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Table Pagination & Footer */}
      <div className="p-3 border-t border-neutral-300 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-neutral-600 bg-neutral-50">
        <div className="flex items-center gap-2">
          <span>Showing</span>
          <span className="font-semibold font-mono tabular-nums text-neutral-900">
            {filteredAndSortedRows.length > 0 ? (currentPage - 1) * pageSize + 1 : 0}
          </span>
          <span>to</span>
          <span className="font-semibold font-mono tabular-nums text-neutral-900">
            {Math.min(currentPage * pageSize, filteredAndSortedRows.length)}
          </span>
          <span>of</span>
          <span className="font-semibold font-mono tabular-nums text-neutral-900">
            {filteredAndSortedRows.length}
          </span>
          <span>records</span>

          <div className="ml-4 flex items-center gap-1.5">
            <span>Per page:</span>
            <select
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setCurrentPage(1);
              }}
              className="bg-white border border-neutral-300 rounded px-1.5 py-0.5 font-mono text-xs text-neutral-700"
            >
              <option value={10}>10</option>
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={99999}>All</option>
            </select>
          </div>
        </div>

        {/* Page Nav */}
        <div className="flex items-center gap-1">
          <button
            type="button"
            disabled={currentPage <= 1}
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            className="p-1 rounded border border-neutral-300 bg-white text-neutral-700 hover:bg-neutral-100 disabled:opacity-40 disabled:pointer-events-none"
            title="Previous page"
          >
            <ChevronLeft className="w-3.5 h-3.5" />
          </button>
          <span className="px-2 font-mono">
            {currentPage} / {totalPages}
          </span>
          <button
            type="button"
            disabled={currentPage >= totalPages}
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            className="p-1 rounded border border-neutral-300 bg-white text-neutral-700 hover:bg-neutral-100 disabled:opacity-40 disabled:pointer-events-none"
            title="Next page"
          >
            <ChevronRight className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Row Inspector Modal */}
      {inspectingRow && (
        <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl border border-neutral-300 max-w-3xl w-full max-h-[88vh] overflow-hidden flex flex-col shadow-xl">
            <div className="p-4 border-b border-neutral-200 flex items-center justify-between bg-neutral-50">
              <div>
                <h3 className="text-sm font-bold text-neutral-900">
                  Record Audit: Work Order {inspectingRow.workOrderNo} · X Track ID: {inspectingRow.xTrackId || 'None'}
                </h3>
                <p className="text-xs text-neutral-500">
                  Status: {inspectingRow.matchStatus} · Secondary Source: {inspectingRow.secondaryReportType || 'None'}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setInspectingRow(null)}
                className="p-1 text-neutral-400 hover:text-neutral-700 rounded transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-5 text-xs">
              {/* XBRL Feedback & Sampling Quality Section */}
              <div className="p-3.5 bg-teal-50/50 rounded-lg border border-teal-200">
                <h4 className="font-bold text-teal-950 mb-2.5 flex items-center justify-between">
                  <span>XBRL Feedback &amp; Quality Audit</span>
                  <span className="text-[10px] text-teal-800 bg-white border border-teal-200 px-2 py-0.5 rounded font-mono font-medium">
                    Input Columns Mapping
                  </span>
                </h4>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 mb-2.5">
                  <div className="bg-white p-2.5 rounded border border-teal-100">
                    <span className="text-neutral-500 block text-[10px] uppercase font-bold tracking-wider">JMS Cycle No / FB Quality:</span>
                    <span className="font-bold text-teal-900 text-xs mt-0.5 block">
                      {cleanFeedbackQuality(inspectingRow.jmsCycleNoFbQuality) || <span className="text-neutral-400 font-normal italic">None</span>}
                    </span>
                    <span className="text-[10px] text-teal-600">Mapped from &ldquo;XBRL Feedback Quality&rdquo;</span>
                  </div>
                  <div className="bg-white p-2.5 rounded border border-indigo-100">
                    <span className="text-neutral-500 block text-[10px] uppercase font-bold tracking-wider">Sampling/Quality:</span>
                    <span className="font-bold text-indigo-950 text-xs mt-0.5 block">
                      {inspectingRow.samplingQuality || <span className="text-neutral-400 font-normal italic">None (blank)</span>}
                    </span>
                    <span className="text-[10px] text-indigo-600">Filled as &ldquo;STP Feedback&rdquo; when FB Quality present</span>
                  </div>
                  <div className="bg-white p-2.5 rounded border border-teal-100">
                    <span className="text-neutral-500 block text-[10px] uppercase font-bold tracking-wider">Feedback Status &amp; Validity:</span>
                    <span className="font-medium text-neutral-800 text-xs mt-0.5 block">
                      {inspectingRow.xbrlFeedbackStatus || 'Completed'} · {inspectingRow.xbrlFeedbackValidity || 'Valid'}
                    </span>
                    <span className="text-[10px] text-neutral-500">Reviewed: {inspectingRow.xbrlFeedbackReviewed || 'Yes'}</span>
                  </div>
                </div>
                {inspectingRow.xbrlFeedbackComments && (
                  <div className="bg-white p-2.5 rounded border border-teal-100 text-[11px]">
                    <span className="font-semibold text-neutral-700 block mb-0.5">XBRL Feedback Comments:</span>
                    <p className="text-neutral-800 italic">&ldquo;{inspectingRow.xbrlFeedbackComments}&rdquo;</p>
                  </div>
                )}
              </div>

              {/* Standard 15 Output Fields */}
              <div>
                <h4 className="font-bold text-neutral-900 mb-2">Canonical 15 Output Fields</h4>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-neutral-50 p-3 rounded-lg border border-neutral-200">
                  <div>
                    <span className="text-neutral-500 block">Team:</span>
                    <span className="font-medium text-neutral-900">{inspectingRow.team}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block">Work Order No:</span>
                    <span className="font-mono font-bold text-blue-900">{inspectingRow.workOrderNo}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block">X Track ID:</span>
                    <span className="font-mono font-semibold text-sky-900">{inspectingRow.xTrackId || 'None'}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block">Form Type:</span>
                    <span className="font-mono font-bold text-amber-900">{inspectingRow.formType || 'None'}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block">Milestone:</span>
                    <span className="font-medium text-neutral-900">{inspectingRow.milestone || 'None'}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block">Delivered Date:</span>
                    <span className="font-mono font-medium text-neutral-900">{inspectingRow.deliveredDate || 'None'}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block">Month:</span>
                    <span className="font-medium text-rose-900">{inspectingRow.month || 'None'}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block">JMS Team:</span>
                    <span className="font-medium text-neutral-900">{inspectingRow.jmsTeam || 'Bridge'}</span>
                  </div>
                </div>
              </div>

              {/* Secondary QC Error Summary */}
              <div>
                <h4 className="font-bold text-neutral-900 mb-2">Secondary QC Error Summary (Matched on X Track ID)</h4>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-neutral-50 p-3 rounded-lg border border-neutral-200 font-mono">
                  <div>
                    <span className="text-neutral-500 block font-sans">Validation Errors:</span>
                    <span className="font-bold text-red-600">{inspectingRow.validationErrors || 0}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block font-sans">Other Errors:</span>
                    <span className="font-bold text-amber-600">{inspectingRow.otherErrors || 0}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block font-sans">Label Errors:</span>
                    <span className="font-bold text-purple-600">{inspectingRow.labelErrors || 0}</span>
                  </div>
                  <div>
                    <span className="text-neutral-500 block font-sans">Total Errors:</span>
                    <span className="font-bold text-neutral-900">{inspectingRow.totalErrors || 0}</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="p-3.5 border-t border-neutral-200 bg-neutral-50 flex items-center justify-between text-xs text-neutral-500">
              <span>Match Key: <strong className="font-mono text-neutral-800">{inspectingRow.matchKey}</strong></span>
              <button
                type="button"
                onClick={() => setInspectingRow(null)}
                className="px-3.5 py-1.5 bg-neutral-800 text-white font-medium rounded hover:bg-neutral-900 transition-colors"
              >
                Close Audit
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Work Order Highlighting Modal */}
      <WorkOrderHighlightModal
        isOpen={isHighlightModalOpen}
        onClose={() => setIsHighlightModalOpen(false)}
        highlightedWos={highlightedWos}
        onApply={(newWos) => setHighlightedWos(newWos)}
        availableWorkOrdersInTable={availableWorkOrdersInTable}
        onlyHighlighted={onlyHighlighted}
        onToggleOnlyHighlighted={setOnlyHighlighted}
      />
    </div>
  );
};
