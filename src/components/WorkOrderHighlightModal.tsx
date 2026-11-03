import React, { useState, useRef } from 'react';
import {
  X,
  Sparkles,
  Upload,
  Trash2,
  Check,
  AlertCircle,
  Plus,
  Layers,
  ListFilter,
} from 'lucide-react';
import { normalizeWorkOrderKey } from '../utils/reconciliation';
import {
  SCREENSHOT_EXAMPLE_WORK_ORDERS,
  SCREENSHOT_CORE_5_WORK_ORDERS,
  SCREENSHOT_TABLE1_WORK_ORDERS,
  SCREENSHOT_BATCH_PRESETS,
} from '../data/sampleData';
import * as XLSX from 'xlsx';

interface WorkOrderHighlightModalProps {
  isOpen: boolean;
  onClose: () => void;
  highlightedWos: Set<string>;
  onApply: (wos: Set<string>) => void;
  availableWorkOrdersInTable: string[];
  onlyHighlighted: boolean;
  onToggleOnlyHighlighted: (val: boolean) => void;
}

export const WorkOrderHighlightModal: React.FC<WorkOrderHighlightModalProps> = ({
  isOpen,
  onClose,
  highlightedWos,
  onApply,
  availableWorkOrdersInTable,
  onlyHighlighted,
  onToggleOnlyHighlighted,
}) => {
  const [textInput, setTextInput] = useState<string>(() =>
    Array.from(highlightedWos).join('\n')
  );
  const [appendMode, setAppendMode] = useState<boolean>(false);
  const [fileError, setFileError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  if (!isOpen) return null;

  // Parse raw text into normalized and cleaned list of unique IDs
  const parsedList = textInput
    .split(/[\r\n,;\t]+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);

  const uniqueSet = new Set(parsedList.map(normalizeWorkOrderKey));

  // Count matches in currently loaded table
  const availableSet = new Set(availableWorkOrdersInTable.map(normalizeWorkOrderKey));
  const matchedInTableCount = Array.from(uniqueSet).filter((wo) => availableSet.has(wo)).length;

  const handleApply = () => {
    onApply(uniqueSet);
    onClose();
  };

  const handleClear = () => {
    setTextInput('');
    onApply(new Set());
  };

  /**
   * Loads a batch of multiple Work Order IDs.
   * If appendMode is checked, appends and deduplicates with existing IDs.
   */
  const handleLoadPreset = (wos: string[]) => {
    setFileError(null);
    if (appendMode && textInput.trim().length > 0) {
      const existing = textInput
        .split(/[\r\n,;\t]+/)
        .map((s) => s.trim())
        .filter(Boolean);
      const combined = Array.from(new Set([...existing, ...wos]));
      setTextInput(combined.join('\n'));
    } else {
      setTextInput(wos.join('\n'));
    }
  };

  /**
   * Toggles a single individual ID in the current textarea
   */
  const handleToggleSingleId = (id: string) => {
    const norm = normalizeWorkOrderKey(id);
    const existing = textInput
      .split(/[\r\n,;\t]+/)
      .map((s) => s.trim())
      .filter(Boolean);

    if (existing.some((item) => normalizeWorkOrderKey(item) === norm)) {
      // Remove it
      const filtered = existing.filter((item) => normalizeWorkOrderKey(item) !== norm);
      setTextInput(filtered.join('\n'));
    } else {
      // Add it
      setTextInput([...existing, id].join('\n'));
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFileError(null);
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    const isExcel =
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv');

    if (isExcel) {
      reader.onload = (evt) => {
        try {
          const buffer = evt.target?.result;
          const workbook = XLSX.read(buffer, { type: 'binary' });
          const firstSheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[firstSheetName];
          const jsonRows = XLSX.utils.sheet_to_json<any>(sheet, { header: 1 });

          const extractedIds: string[] = [];
          for (const row of jsonRows) {
            if (Array.isArray(row)) {
              for (const cell of row) {
                if (cell !== undefined && cell !== null) {
                  const str = String(cell).trim();
                  // Check if cell resembles a Work Order ID or number
                  if (str && !str.toLowerCase().includes('work order') && !str.toLowerCase().includes('total')) {
                    extractedIds.push(str);
                  }
                }
              }
            }
          }

          if (extractedIds.length === 0) {
            setFileError('No Work Order IDs could be found in the uploaded file.');
            return;
          }

          setTextInput((prev) => {
            const existing = prev ? prev.trim() + '\n' : '';
            return existing + extractedIds.join('\n');
          });
        } catch (err: any) {
          setFileError(`Error parsing spreadsheet: ${err.message || 'Invalid format'}`);
        }
      };
      reader.readAsBinaryString(file);
    } else {
      // Plain text file (.txt)
      reader.onload = (evt) => {
        const text = evt.target?.result as string;
        if (text) {
          setTextInput((prev) => {
            const existing = prev ? prev.trim() + '\n' : '';
            return existing + text;
          });
        }
      };
      reader.readAsText(file);
    }

    // Reset input so same file can be reloaded if needed
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="fixed inset-0 z-50 overflow-y-auto bg-neutral-900/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className="bg-white rounded-xl shadow-2xl max-w-xl w-full border border-neutral-200 overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        {/* Header */}
        <div className="bg-gradient-to-r from-amber-500 to-yellow-500 px-6 py-4 flex items-center justify-between text-white">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-white/20 rounded-lg">
              <Sparkles className="w-5 h-5 text-white" />
            </div>
            <div>
              <h3 className="font-bold text-base leading-tight">Highlight Work Orders in Yellow</h3>
              <p className="text-amber-100 text-xs mt-0.5">
                Load multiple Work Order IDs to visually highlight matching rows across all views
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="text-white/80 hover:text-white p-1 rounded-md hover:bg-white/10 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Content */}
        <div className="p-5 space-y-4 max-h-[75vh] overflow-y-auto">
          {/* Multiple Batch Options under "Load Screenshot IDs" */}
          <div className="bg-amber-50/90 border border-amber-300 rounded-xl p-3.5 space-y-2.5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="flex items-center gap-1.5">
                <Layers className="w-4 h-4 text-amber-800 shrink-0" />
                <span className="text-xs font-bold text-amber-950">
                  Load Screenshot IDs:
                </span>
                <span className="text-[11px] text-amber-800">
                  (Choose batch to load multiple IDs)
                </span>
              </div>
              <label className="flex items-center gap-1.5 text-[11px] font-semibold text-amber-900 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={appendMode}
                  onChange={(e) => setAppendMode(e.target.checked)}
                  className="rounded text-amber-600 focus:ring-amber-500 h-3.5 w-3.5"
                />
                <span>Append to existing list</span>
              </label>
            </div>

            {/* Batch Action Buttons */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
              <button
                type="button"
                onClick={() => handleLoadPreset(SCREENSHOT_EXAMPLE_WORK_ORDERS)}
                className="flex flex-col items-start p-2.5 bg-amber-200 hover:bg-amber-300 active:bg-amber-400 border border-amber-400 rounded-lg text-left transition-colors shadow-2xs group"
                title="Loads 15 Work Order IDs compiled from both screenshots"
              >
                <div className="flex items-center gap-1.5 w-full justify-between">
                  <span className="text-xs font-bold text-amber-950 flex items-center gap-1">
                    <Sparkles className="w-3.5 h-3.5 text-amber-800" />
                    All Screenshot IDs
                  </span>
                  <span className="px-1.5 py-0.2 text-[10px] font-bold rounded-full bg-amber-400/80 text-amber-950 font-mono">
                    15 IDs
                  </span>
                </div>
                <span className="text-[10px] text-amber-800 mt-1 leading-tight">
                  Target list from both screenshots
                </span>
              </button>

              <button
                type="button"
                onClick={() => handleLoadPreset(SCREENSHOT_CORE_5_WORK_ORDERS)}
                className="flex flex-col items-start p-2.5 bg-white hover:bg-amber-100/70 border border-amber-300 rounded-lg text-left transition-colors shadow-2xs group"
                title="Loads the 5 Work Order IDs shown in the crop screenshot: WO-26-003172, WO-26-003125, WO-26-003276, WO-26-003627, WO-26-003535"
              >
                <div className="flex items-center gap-1.5 w-full justify-between">
                  <span className="text-xs font-bold text-neutral-800">
                    Core 5 Crop IDs
                  </span>
                  <span className="px-1.5 py-0.2 text-[10px] font-bold rounded-full bg-neutral-200 text-neutral-800 font-mono">
                    5 IDs
                  </span>
                </div>
                <span className="text-[10px] text-neutral-600 mt-1 leading-tight">
                  WO-26-003172, WO-26-003125...
                </span>
              </button>

              <button
                type="button"
                onClick={() => handleLoadPreset(SCREENSHOT_TABLE1_WORK_ORDERS)}
                className="flex flex-col items-start p-2.5 bg-white hover:bg-amber-100/70 border border-amber-300 rounded-lg text-left transition-colors shadow-2xs group"
                title="Loads 20 Work Orders from Excel Table 1 screenshot (image_98a4351e.png)"
              >
                <div className="flex items-center gap-1.5 w-full justify-between">
                  <span className="text-xs font-bold text-neutral-800">
                    Table 1 Batch
                  </span>
                  <span className="px-1.5 py-0.2 text-[10px] font-bold rounded-full bg-neutral-200 text-neutral-800 font-mono">
                    20 IDs
                  </span>
                </div>
                <span className="text-[10px] text-neutral-600 mt-1 leading-tight">
                  From Table 1 Excel screenshot
                </span>
              </button>
            </div>

            {/* Quick click-to-toggle chips for individual screenshot IDs */}
            <div className="pt-2 border-t border-amber-200/70">
              <span className="text-[10.5px] font-bold text-amber-900 block mb-1.5">
                Quick click to add / remove individual IDs:
              </span>
              <div className="flex flex-wrap gap-1 max-h-24 overflow-y-auto pr-1">
                {SCREENSHOT_EXAMPLE_WORK_ORDERS.map((id) => {
                  const isPresent = uniqueSet.has(normalizeWorkOrderKey(id));
                  return (
                    <button
                      key={id}
                      type="button"
                      onClick={() => handleToggleSingleId(id)}
                      className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10.5px] font-mono transition-colors ${
                        isPresent
                          ? 'bg-amber-300 text-amber-950 font-bold border border-amber-400 shadow-2xs'
                          : 'bg-white text-neutral-700 border border-neutral-300 hover:bg-amber-100 hover:border-amber-400'
                      }`}
                      title={isPresent ? `Click to remove ${id}` : `Click to add ${id}`}
                    >
                      {isPresent ? <Check className="w-3 h-3 text-amber-900" /> : <Plus className="w-3 h-3 text-neutral-400" />}
                      <span>{id}</span>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Text Area Input */}
          <div>
            <div className="flex items-center justify-between mb-1.5">
              <label htmlFor="wo-input" className="text-xs font-bold text-neutral-800 flex items-center gap-1.5">
                <ListFilter className="w-3.5 h-3.5 text-neutral-600" />
                <span>Work Order IDs List (one per line or separated by commas / tabs):</span>
              </label>
              <span className="text-[11px] text-amber-900 font-bold bg-amber-100 px-2 py-0.5 rounded border border-amber-300">
                {uniqueSet.size} unique ID{uniqueSet.size === 1 ? '' : 's'}
              </span>
            </div>
            <textarea
              id="wo-input"
              rows={6}
              value={textInput}
              onChange={(e) => {
                setTextInput(e.target.value);
                setFileError(null);
              }}
              placeholder={`Paste multiple Work Order IDs here, e.g.:\nWO-26-003172\nWO-26-003125\nWO-26-003276\nWO-26-003627\nWO-26-003535\nWO-26-003361\nWO-26-003537\nWO-26-003558`}
              className="w-full text-xs font-mono p-3 bg-neutral-50 border border-neutral-300 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-amber-500 focus:border-amber-500 transition-colors"
            />
          </div>

          {/* File Upload Option */}
          <div className="flex items-center justify-between gap-3 text-xs border-t border-neutral-200 pt-3">
            <span className="text-neutral-600 font-medium">Or import work orders from file:</span>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".txt,.csv,.xlsx,.xls"
              className="hidden"
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-neutral-700 bg-white border border-neutral-300 hover:bg-neutral-50 rounded-md transition-colors"
            >
              <Upload className="w-3.5 h-3.5 text-neutral-500" />
              <span>Choose .xlsx / .csv / .txt</span>
            </button>
          </div>

          {fileError && (
            <div className="flex items-center gap-2 p-2.5 bg-red-50 border border-red-200 rounded-md text-red-700 text-xs">
              <AlertCircle className="w-4 h-4 shrink-0 text-red-600" />
              <span>{fileError}</span>
            </div>
          )}

          {/* Live Matching Statistics & Filter Option */}
          <div className="bg-neutral-50 border border-neutral-200 rounded-lg p-3 space-y-2">
            <div className="flex items-center justify-between text-xs">
              <span className="text-neutral-600 font-medium">Matching status in current table:</span>
              <span className="font-bold text-neutral-900">
                {matchedInTableCount} of {uniqueSet.size} Work Orders found
              </span>
            </div>

            <label className="flex items-center gap-2 pt-1 border-t border-neutral-200/80 cursor-pointer">
              <input
                type="checkbox"
                checked={onlyHighlighted}
                onChange={(e) => onToggleOnlyHighlighted(e.target.checked)}
                className="rounded text-amber-600 focus:ring-amber-500 h-3.5 w-3.5"
              />
              <span className="text-xs font-medium text-neutral-800">
                Show only highlighted Work Orders in table view (focus filter)
              </span>
            </label>
          </div>
        </div>

        {/* Footer */}
        <div className="bg-neutral-50 px-6 py-3 border-t border-neutral-200 flex items-center justify-between">
          <button
            type="button"
            onClick={handleClear}
            className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-red-700 hover:bg-red-50 rounded-md transition-colors"
          >
            <Trash2 className="w-3.5 h-3.5" />
            <span>Clear Highlighting</span>
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3 py-1.5 text-xs font-medium text-neutral-700 hover:bg-neutral-200 rounded-md transition-colors"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={handleApply}
              className="inline-flex items-center gap-1.5 px-4 py-1.5 text-xs font-bold text-amber-950 bg-amber-400 hover:bg-amber-500 rounded-md transition-colors shadow-xs"
            >
              <Check className="w-3.5 h-3.5" />
              <span>Apply Yellow Highlight ({uniqueSet.size})</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
