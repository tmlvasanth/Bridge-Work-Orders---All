import XLSX from 'xlsx-js-style';
import { MatchedOutputRow } from '../types';
import { normalizeWorkOrderKey, cleanFeedbackQuality } from './reconciliation';

export interface ColumnDefinition {
  header: string;
  key: string;
  width: number;
}

// 1. Standard Output Columns View - Primary Output Specification with XBRL Feedback Comments as last column
export const STANDARD_15_COLUMNS: ColumnDefinition[] = [
  { header: 'Team', key: 'team', width: 12 },
  { header: 'JMS Job Number/Work Order No', key: 'workOrderNo', width: 28 },
  { header: 'JMS Cycle No / FB Quality', key: 'jmsCycleNoFbQuality', width: 24 },
  { header: 'XOR Ver', key: 'xorVer', width: 12 },
  { header: 'X Track ID', key: 'xTrackId', width: 20 },
  { header: 'Job Name', key: 'jobName', width: 36 },
  { header: 'Form Type', key: 'formType', width: 18 },
  { header: 'Filing Period', key: 'filingPeriod', width: 16 },
  { header: 'Milestone', key: 'milestone', width: 22 },
  { header: 'JMS Team', key: 'jmsTeam', width: 14 },
  { header: 'Sampling/Quality', key: 'samplingQuality', width: 18 },
  { header: 'Consultant Name', key: 'consultantName', width: 20 },
  { header: '2nd Reviewer Name', key: 'secondReviewerName', width: 20 },
  { header: 'Delivered Date', key: 'deliveredDate', width: 16 },
  { header: 'Month', key: 'month', width: 14 },
  { header: 'XBRL Feedback Comments', key: 'xbrlFeedbackComments', width: 40 },
];

export const STANDARD_16_COLUMNS = STANDARD_15_COLUMNS;

// 2. XBRL Feedback Audit Columns View
export const XBRL_FEEDBACK_COLUMNS: ColumnDefinition[] = [
  { header: 'Team', key: 'team', width: 12 },
  { header: 'JMS Job Number/Work Order No', key: 'workOrderNo', width: 28 },
  { header: 'JMS Cycle No / FB Quality', key: 'jmsCycleNoFbQuality', width: 24 },
  { header: 'Sampling/Quality', key: 'samplingQuality', width: 18 },
  { header: 'XBRL Feedback Quality', key: 'xbrlFeedbackQuality', width: 22 },
  { header: 'XBRL Feedback Status', key: 'xbrlFeedbackStatus', width: 20 },
  { header: 'XBRL Feedback Reviewed', key: 'xbrlFeedbackReviewed', width: 22 },
  { header: 'XBRL Feedback Validity', key: 'xbrlFeedbackValidity', width: 22 },
  { header: 'XBRL Feedback Comments', key: 'xbrlFeedbackComments', width: 36 },
  { header: 'X Track ID', key: 'xTrackId', width: 20 },
  { header: 'Job Name', key: 'jobName', width: 36 },
  { header: 'Form Type', key: 'formType', width: 18 },
  { header: 'Milestone', key: 'milestone', width: 22 },
  { header: 'Delivered Date', key: 'deliveredDate', width: 16 },
  { header: 'Month', key: 'month', width: 14 },
];

// 3. QC & Error Details View (Optional Extended View)
export const QC_IMAGE_COLUMNS: ColumnDefinition[] = [
  { header: 'Team', key: 'team', width: 12 },
  { header: 'JMS Job Number/Work Order No', key: 'workOrderNo', width: 28 },
  { header: 'X Track ID', key: 'xTrackId', width: 20 },
  { header: 'Job Name', key: 'jobName', width: 36 },
  { header: 'Month', key: 'month', width: 14 },
  { header: 'Production Consultant Name', key: 'productionConsultantName', width: 24 },
  { header: 'Internal Reviewer name', key: 'internalReviewerName', width: 24 },
  { header: 'External Reviewer Name', key: 'externalReviewerName', width: 24 },
  { header: 'DEI Errors', key: 'deiErrors', width: 14 },
  { header: 'Element Errors', key: 'elementErrors', width: 14 },
  { header: 'Value Errors', key: 'valueErrors', width: 14 },
  { header: 'Untagged Errors', key: 'untaggedErrors', width: 16 },
  { header: 'L1-L3 Errors', key: 'l1L3Errors', width: 14 },
  { header: 'Reverse Sign Errors', key: 'reverseSignErrors', width: 18 },
  { header: 'Alternate Values Errors', key: 'alternateValuesErrors', width: 20 },
  { header: 'Date Context Errors', key: 'dateContextErrors', width: 18 },
  { header: 'Member Context Errors', key: 'memberContextErrors', width: 20 },
  { header: 'Scale/Precision Errors', key: 'scalePrecisionErrors', width: 20 },
  { header: 'Taxonomy/Structure/Invalid Spi', key: 'taxonomyStructureInvalidSpi', width: 26 },
  { header: 'Incorrect Mapping | Attribute', key: 'incorrectMappingAttribute', width: 26 },
  { header: 'Judgmental Errors', key: 'judgmentalErrors', width: 18 },
  { header: 'Validation Errors', key: 'validationErrors', width: 18 },
  { header: 'Other Errors', key: 'otherErrors', width: 16 },
  { header: 'Label Errors', key: 'labelErrors', width: 16 },
];

// 4. All Columns View (Complete Comprehensive View)
export const ALL_COLUMNS: ColumnDefinition[] = [
  ...STANDARD_15_COLUMNS,
  { header: 'XBRL Feedback Quality', key: 'xbrlFeedbackQuality', width: 22 },
  { header: 'XBRL Feedback Status', key: 'xbrlFeedbackStatus', width: 20 },
  { header: 'XBRL Feedback Reviewed', key: 'xbrlFeedbackReviewed', width: 22 },
  { header: 'XBRL Feedback Validity', key: 'xbrlFeedbackValidity', width: 22 },
  { header: 'Production Consultant Name', key: 'productionConsultantName', width: 24 },
  { header: 'Internal Reviewer name', key: 'internalReviewerName', width: 24 },
  { header: 'External Reviewer Name', key: 'externalReviewerName', width: 24 },
  { header: 'Validation Errors', key: 'validationErrors', width: 18 },
  { header: 'Other Errors', key: 'otherErrors', width: 16 },
  { header: 'Label Errors', key: 'labelErrors', width: 16 },
  { header: 'Total Errors', key: 'totalErrors', width: 16 },
];

// Default columns set to the Primary Standard Headings View
export const OUTPUT_COLUMNS = STANDARD_15_COLUMNS;

/**
 * Converts matched rows to array of objects with exact output headers.
 */
export function formatOutputRowsForExport(
  rows: MatchedOutputRow[],
  cols: ColumnDefinition[] = STANDARD_15_COLUMNS,
  highlightedWorkOrders?: Set<string>
): Record<string, any>[] {
  const hasHighlighted = highlightedWorkOrders && highlightedWorkOrders.size > 0;

  return rows.map((row) => {
    const item: Record<string, any> = {};

    if (hasHighlighted) {
      const norm = String(row.workOrderNo || '').trim().toUpperCase();
      item['Highlighted'] = highlightedWorkOrders.has(norm) ? 'YES (Highlighted)' : '';
    }

    cols.forEach((col) => {
      let val = (row as any)[col.key];
      if (col.key === 'jmsCycleNoFbQuality' || col.key === 'xbrlFeedbackQuality') {
        val = cleanFeedbackQuality(val);
      }
      item[col.header] = val !== undefined && val !== null ? val : '';
    });
    return item;
  });
}

/**
 * Generates and downloads an Excel file (.xlsx) with exact column headers,
 * professional formatting, and visual yellow highlighting for highlighted Work Orders.
 */
export function exportToExcel(
  rows: MatchedOutputRow[],
  filename = 'Bridge_Work_Orders_Reconciled.xlsx',
  cols: ColumnDefinition[] = STANDARD_15_COLUMNS,
  highlightedWorkOrders?: Set<string>
): void {
  const hasHighlighted = Boolean(highlightedWorkOrders && highlightedWorkOrders.size > 0);

  // Normalize all highlighted Work Orders for fast case-insensitive lookup
  const normalizedHighlightedSet = new Set(
    Array.from(highlightedWorkOrders || []).map(normalizeWorkOrderKey)
  );

  const exportData = formatOutputRowsForExport(rows, cols, highlightedWorkOrders);

  // Create workbook and worksheet
  const worksheet = XLSX.utils.json_to_sheet(exportData);

  // Build header keys list
  const headerKeys = hasHighlighted ? ['Highlighted', ...cols.map((c) => c.header)] : cols.map((c) => c.header);
  const totalCols = headerKeys.length;
  const totalRows = rows.length;

  // Set column widths
  const colWidths = [];
  if (hasHighlighted) {
    colWidths.push({ wch: 18 });
  }
  cols.forEach((col) => {
    colWidths.push({ wch: Math.max(col.header.length + 4, col.width) });
  });
  worksheet['!cols'] = colWidths;

  // Style Header Row (Excel row 1, index 0)
  for (let c = 0; c < totalCols; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
    if (!worksheet[cellAddr]) continue;

    const isHighlightCol = hasHighlighted && c === 0;
    worksheet[cellAddr].s = {
      fill: {
        patternType: 'solid',
        fgColor: { rgb: isHighlightCol ? 'F59E0B' : '065F46' }, // Amber or Dark Emerald Green
      },
      font: {
        name: 'Calibri',
        sz: 11,
        bold: true,
        color: { rgb: 'FFFFFF' },
      },
      border: {
        top: { style: 'thin', color: { rgb: '047857' } },
        bottom: { style: 'medium', color: { rgb: '047857' } },
        left: { style: 'thin', color: { rgb: '047857' } },
        right: { style: 'thin', color: { rgb: '047857' } },
      },
      alignment: {
        vertical: 'center',
        horizontal: 'center',
        wrapText: true,
      },
    };
  }

  // Style Data Rows (Excel rows 2 to totalRows + 1)
  for (let r = 0; r < totalRows; r++) {
    const rowData = rows[r];
    const woNorm = normalizeWorkOrderKey(rowData.workOrderNo);
    const jmsJobNoVal = (rowData as any).jmsJobNo;
    const jmsNorm = jmsJobNoVal ? normalizeWorkOrderKey(jmsJobNoVal) : '';
    const isHighlighted = normalizedHighlightedSet.has(woNorm) || (Boolean(jmsNorm) && normalizedHighlightedSet.has(jmsNorm));

    const excelRow = r + 1; // 0 is header, data starts at row 1

    for (let c = 0; c < totalCols; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r: excelRow, c });
      if (!worksheet[cellAddr]) {
        // Ensure cell exists so formatting spans across the entire row
        worksheet[cellAddr] = { t: 's', v: '' };
      }

      if (isHighlighted) {
        // VISIBLE VIBRANT YELLOW HIGHLIGHT IN EXCEL (.xlsx)
        worksheet[cellAddr].s = {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: 'FFFF00' }, // Universal standard Excel yellow highlight
          },
          font: {
            name: 'Calibri',
            sz: 11,
            bold: c <= 2, // Bold for 'Highlighted' and 'Work Order' columns
            color: { rgb: '000000' },
          },
          border: {
            top: { style: 'thin', color: { rgb: 'F59E0B' } },
            bottom: { style: 'thin', color: { rgb: 'F59E0B' } },
            left: { style: 'thin', color: { rgb: 'FDE047' } },
            right: { style: 'thin', color: { rgb: 'FDE047' } },
          },
          alignment: {
            vertical: 'center',
          },
        };
      } else {
        // Standard data row styling with alternating subtle tint
        const isZebra = r % 2 === 1;
        worksheet[cellAddr].s = {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: isZebra ? 'F9FAFB' : 'FFFFFF' },
          },
          font: {
            name: 'Calibri',
            sz: 11,
            color: { rgb: '111827' },
          },
          border: {
            top: { style: 'thin', color: { rgb: 'E5E7EB' } },
            bottom: { style: 'thin', color: { rgb: 'E5E7EB' } },
            left: { style: 'thin', color: { rgb: 'E5E7EB' } },
            right: { style: 'thin', color: { rgb: 'E5E7EB' } },
          },
          alignment: {
            vertical: 'center',
          },
        };
      }
    }
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Reconciled Output');

  // Trigger browser download with cell background colors preserved
  XLSX.writeFile(workbook, filename);
}

/**
 * Generates and downloads a CSV file.
 */
export function exportToCSV(
  rows: MatchedOutputRow[],
  filename = 'Bridge_Work_Orders_Reconciled.csv',
  cols: ColumnDefinition[] = STANDARD_15_COLUMNS,
  highlightedWorkOrders?: Set<string>
): void {
  const exportData = formatOutputRowsForExport(rows, cols, highlightedWorkOrders);
  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const csvContent = XLSX.utils.sheet_to_csv(worksheet);

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
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

/**
 * Copies the output rows to clipboard as TSV format (can be pasted directly into open Excel sheet).
 */
export async function copyTableToClipboard(
  rows: MatchedOutputRow[],
  cols: ColumnDefinition[] = STANDARD_15_COLUMNS,
  highlightedWorkOrders?: Set<string>
): Promise<boolean> {
  const hasHighlighted = highlightedWorkOrders && highlightedWorkOrders.size > 0;
  const headerList = hasHighlighted ? ['Highlighted', ...cols.map((c) => c.header)] : cols.map((c) => c.header);
  const headers = headerList.join('\t');

  const rowStrings = rows.map((row) => {
    const values: string[] = [];
    if (hasHighlighted) {
      const norm = String(row.workOrderNo || '').trim().toUpperCase();
      values.push(highlightedWorkOrders.has(norm) ? 'YES' : '');
    }
    cols.forEach((col) => {
      let val = (row as any)[col.key];
      if (col.key === 'jmsCycleNoFbQuality' || col.key === 'xbrlFeedbackQuality') {
        val = cleanFeedbackQuality(val);
      }
      values.push(val !== undefined && val !== null ? String(val) : '');
    });
    return values.join('\t');
  });

  const fullTsv = [headers, ...rowStrings].join('\n');

  try {
    await navigator.clipboard.writeText(fullTsv);
    return true;
  } catch (err) {
    console.error('Failed to copy to clipboard', err);
    return false;
  }
}
