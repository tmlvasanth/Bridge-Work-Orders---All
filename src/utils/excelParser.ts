import * as XLSX from 'xlsx';
import { ParsedSheet } from '../types';

/**
 * Format a date value or Excel serial date into MM/DD/YYYY strictly.
 */
export function formatDateToMMDDYYYY(val: any): { formatted: string; month: string } {
  if (val === null || val === undefined || val === '') {
    return { formatted: '', month: '' };
  }

  let date: Date | null = null;

  // Case 1: Already a Date object
  if (val instanceof Date && !isNaN(val.getTime())) {
    date = val;
  }
  // Case 2: Numeric Excel serial date (e.g., 45579)
  else if (typeof val === 'number') {
    // Excel base date is 1899-12-30 due to the 1900 leap year bug
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const msPerDay = 86400000;
    const dateMs = excelEpoch.getTime() + Math.round(val * msPerDay);
    date = new Date(dateMs);
  }
  // Case 3: String date parsing
  else if (typeof val === 'string') {
    const trimmed = val.trim();
    if (!trimmed) return { formatted: '', month: '' };

    // Check if numeric string serial
    if (/^\d{5}(\.\d+)?$/.test(trimmed)) {
      const numVal = parseFloat(trimmed);
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const msPerDay = 86400000;
      date = new Date(excelEpoch.getTime() + Math.round(numVal * msPerDay));
    }
    // Check for standard formats like YYYY-MM-DD or MM/DD/YYYY
    else if (/^\d{4}-\d{1,2}-\d{1,2}/.test(trimmed)) {
      const parts = trimmed.split(/[-T\s]/);
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const day = parseInt(parts[2], 10);
      date = new Date(year, month, day);
    } else if (/^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(trimmed)) {
      const parts = trimmed.split('/');
      const month = parseInt(parts[0], 10) - 1;
      const day = parseInt(parts[1], 10);
      let year = parseInt(parts[2], 10);
      if (year < 100) year += 2000;
      date = new Date(year, month, day);
    } else {
      const parsed = Date.parse(trimmed);
      if (!isNaN(parsed)) {
        date = new Date(parsed);
      }
    }
  }

  if (!date || isNaN(date.getTime())) {
    // If can't parse as date, return original string as fallback
    const fallbackStr = String(val).trim();
    return { formatted: fallbackStr, month: '' };
  }

  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const yyyy = date.getFullYear();

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  const month = monthNames[date.getMonth()] || '';

  return {
    formatted: `${mm}/${dd}/${yyyy}`,
    month: month,
  };
}

/**
 * Parses an Excel or CSV file from File object into structured sheets.
 */
export async function parseExcelFile(file: File): Promise<{
  sheetNames: string[];
  sheets: Record<string, ParsedSheet>;
}> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellDates: true,
    cellNF: false,
    cellText: false,
  });

  const sheets: Record<string, ParsedSheet> = {};

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    // Convert sheet to array of arrays to find real header row
    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      blankrows: false,
      raw: false,
      dateNF: 'yyyy-mm-dd',
    });

    if (rawData.length === 0) {
      sheets[sheetName] = {
        sheetName,
        headers: [],
        data: [],
        rowCount: 0,
      };
      return;
    }

    // Detect header row (first row with multiple non-empty string cells)
    let headerRowIndex = 0;
    for (let r = 0; r < Math.min(rawData.length, 5); r++) {
      const row = rawData[r];
      const nonEmptyCells = row.filter((cell) => cell !== null && cell !== undefined && String(cell).trim() !== '');
      if (nonEmptyCells.length >= 2) {
        headerRowIndex = r;
        break;
      }
    }

    const rawHeaders = rawData[headerRowIndex] || [];
    const headers: string[] = [];
    const headerCounts: Record<string, number> = {};

    rawHeaders.forEach((h: any, idx: number) => {
      let str = h ? String(h).trim() : `Column_${idx + 1}`;
      if (!str) str = `Column_${idx + 1}`;
      if (headerCounts[str]) {
        headerCounts[str]++;
        headers.push(`${str}_${headerCounts[str]}`);
      } else {
        headerCounts[str] = 1;
        headers.push(str);
      }
    });

    const data: Record<string, any>[] = [];
    for (let r = headerRowIndex + 1; r < rawData.length; r++) {
      const row = rawData[r];
      if (!row || row.length === 0) continue;

      const rowObj: Record<string, any> = {};
      let hasData = false;

      headers.forEach((header, colIdx) => {
        const val = row[colIdx];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          hasData = true;
          rowObj[header] = val;
        } else {
          rowObj[header] = '';
        }
      });

      if (hasData) {
        data.push(rowObj);
      }
    }

    sheets[sheetName] = {
      sheetName,
      headers,
      data,
      rowCount: data.length,
    };
  });

  return {
    sheetNames: workbook.SheetNames,
    sheets,
  };
}
