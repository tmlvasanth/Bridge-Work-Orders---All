import React, { useRef } from 'react';
import { UploadCloud, CheckCircle2, AlertCircle, FileText, ChevronDown } from 'lucide-react';
import { UploadedFileState } from '../types';

interface FileUploadCardProps {
  title: string;
  subtitle: string;
  expectedColumns: string[];
  fileState: UploadedFileState;
  onFileSelect: (file: File) => void;
  onSheetChange: (sheetName: string) => void;
  onRemove: () => void;
  accentColor?: string;
}

export const FileUploadCard: React.FC<FileUploadCardProps> = ({
  title,
  subtitle,
  expectedColumns,
  fileState,
  onFileSelect,
  onSheetChange,
  onRemove,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileSelect(e.dataTransfer.files[0]);
    }
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onFileSelect(e.target.files[0]);
    }
  };

  const hasData = fileState.data.length > 0;

  return (
    <div className="bg-white border-2 border-neutral-300 rounded-xl p-5 flex flex-col justify-between shadow-xs">
      <div>
        <div className="flex items-start justify-between mb-3">
          <div>
            <h2 className="text-sm font-bold text-neutral-900">{title}</h2>
            <p className="text-xs text-neutral-500 mt-0.5">{subtitle}</p>
          </div>
          {hasData && (
            <div className="flex flex-col sm:flex-row items-end sm:items-center gap-1.5">
              <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 px-2 py-0.5 rounded">
                <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                <span>{fileState.data.length} rows</span>
              </span>
              {fileState.selectedSheet && (
                <span className="text-[10px] text-teal-800 bg-teal-50 border border-teal-300 px-1.5 py-0.5 rounded font-mono font-medium">
                  Sheet: {fileState.selectedSheet}
                </span>
              )}
            </div>
          )}
        </div>

        {/* Dropzone or Loaded State */}
        {!fileState.fileName ? (
          <div
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className="border-2 border-dashed border-neutral-300 hover:border-blue-500 rounded-lg p-6 text-center cursor-pointer transition-colors bg-neutral-50/70 hover:bg-blue-50/20 flex flex-col items-center justify-center min-h-[160px]"
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileInput}
              className="hidden"
            />
            <div className="w-10 h-10 rounded-full bg-white border border-neutral-200 flex items-center justify-center mb-3 shadow-xs">
              <UploadCloud className="w-5 h-5 text-neutral-600" />
            </div>
            <p className="text-xs font-medium text-neutral-800">
              Drop Excel or CSV here, or <span className="text-neutral-900 underline underline-offset-2">browse</span>
            </p>
            <p className="text-[11px] text-neutral-400 mt-1">
              Supports .xlsx, .xls, .csv files
            </p>
          </div>
        ) : (
          <div className="border border-neutral-200 rounded-lg p-4 bg-neutral-50/60">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2.5 min-w-0">
                <div className="w-8 h-8 rounded bg-white border border-neutral-200 flex items-center justify-center shrink-0">
                  <FileText className="w-4 h-4 text-neutral-700" />
                </div>
                <div className="min-w-0">
                  <p className="text-xs font-medium text-neutral-900 truncate">
                    {fileState.fileName}
                  </p>
                  <p className="text-[11px] text-neutral-500 font-mono tabular-nums">
                    {fileState.data.length} records · {fileState.headers.length} columns
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={onRemove}
                className="text-xs text-neutral-500 hover:text-red-600 px-2 py-1 rounded transition-colors"
              >
                Change
              </button>
            </div>

            {/* Sheet Selector if multiple sheets exist */}
            {fileState.sheetNames.length > 1 && (
              <div className="mt-3 pt-3 border-t border-neutral-200/80">
                <label className="block text-[11px] font-medium text-neutral-600 mb-1">
                  Active Sheet:
                </label>
                <div className="relative">
                  <select
                    value={fileState.selectedSheet}
                    onChange={(e) => onSheetChange(e.target.value)}
                    className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 pr-8 appearance-none focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                  >
                    {fileState.sheetNames.map((sheet) => (
                      <option key={sheet} value={sheet}>
                        {sheet} ({fileState.sheets[sheet]?.rowCount || 0} rows)
                      </option>
                    ))}
                  </select>
                  <ChevronDown className="w-3.5 h-3.5 text-neutral-500 absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                </div>
              </div>
            )}
          </div>
        )}

        {/* Error message if any */}
        {fileState.error && (
          <div className="mt-2 text-xs text-red-600 flex items-center gap-1.5">
            <AlertCircle className="w-3.5 h-3.5 shrink-0" />
            <span>{fileState.error}</span>
          </div>
        )}
      </div>

      {/* Expected key fields hint */}
      <div className="mt-4 pt-3 border-t border-neutral-100 text-[11px] text-neutral-500">
        <span className="font-medium text-neutral-700">Expected columns: </span>
        <span>{expectedColumns.join(', ')}</span>
      </div>
    </div>
  );
};
