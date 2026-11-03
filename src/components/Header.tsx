import React from 'react';
import { FileSpreadsheet, RotateCcw, Sparkles } from 'lucide-react';

interface HeaderProps {
  onLoadSample: () => void;
  onReset: () => void;
  hasData: boolean;
}

export const Header: React.FC<HeaderProps> = ({ onLoadSample, onReset, hasData }) => {
  return (
    <header className="w-full bg-white border-b border-neutral-200 sticky top-0 z-30">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
        {/* Zone 1: Single text element wordmark */}
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-neutral-900 text-white flex items-center justify-center font-semibold text-sm">
            <FileSpreadsheet className="w-4 h-4" />
          </div>
          <div>
            <h1 className="text-base font-bold text-neutral-900 tracking-tight">
              Bridge Work Orders - All
            </h1>
            <p className="text-xs text-neutral-500 hidden sm:block">
              Work Order Reconciliation &amp; Formatting Tool
            </p>
          </div>
        </div>

        {/* Zone 2: Navigation Links */}
        <nav className="hidden md:flex items-center gap-6 text-xs font-medium text-neutral-600">
          <a href="#upload-section" className="hover:text-neutral-900 transition-colors">
            Upload Files
          </a>
          <a href="#mappings-section" className="hover:text-neutral-900 transition-colors">
            Column Rules
          </a>
          <a href="#results-section" className="hover:text-neutral-900 transition-colors">
            Reconciled Output
          </a>
        </nav>

        {/* Zone 3: Primary Actions */}
        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={onLoadSample}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-neutral-700 bg-neutral-100 hover:bg-neutral-200 rounded-md transition-colors whitespace-nowrap"
            title="Load authentic sample Bridge and JMS data to preview reconciliation"
          >
            <Sparkles className="w-3.5 h-3.5 text-neutral-600" />
            <span>Load Demo Data</span>
          </button>

          {hasData && (
            <button
              type="button"
              onClick={onReset}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-neutral-500 hover:text-neutral-900 hover:bg-neutral-50 rounded-md transition-colors"
              title="Clear all loaded files"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Reset</span>
            </button>
          )}
        </div>
      </div>
    </header>
  );
};
