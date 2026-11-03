import React, { useState } from 'react';
import { SlidersHorizontal, ChevronDown, ChevronUp, Info, CheckCircle2, RotateCcw, BookmarkCheck, Sparkles } from 'lucide-react';
import { ColumnMappingConfig, ParsedSheet } from '../types';
import { findBestColumnMatch } from '../utils/reconciliation';

interface ColumnMappingDrawerProps {
  bridgeHeaders: string[];
  jmsHeaders: string[];
  jmsSheetNames: string[];
  jmsSheets: Record<string, ParsedSheet>;
  config: ColumnMappingConfig;
  onChangeConfig: (newConfig: ColumnMappingConfig) => void;
  onResetToAutoDetect?: () => void;
}

export const ColumnMappingDrawer: React.FC<ColumnMappingDrawerProps> = ({
  bridgeHeaders,
  jmsHeaders,
  jmsSheetNames,
  jmsSheets,
  config,
  onChangeConfig,
  onResetToAutoDetect,
}) => {
  const [isOpen, setIsOpen] = useState(false);

  const handleUpdate = (key: keyof ColumnMappingConfig, val: string) => {
    // If updating production sheet, also auto-pick work order and track ID columns for that sheet
    if (key === 'jmsProductionSheet') {
      const prodSheet = jmsSheets[val];
      const prodHeaders = prodSheet ? prodSheet.headers : [];

      const newWoCol =
        findBestColumnMatch(
          prodHeaders,
          ['Work Order No', 'Work Order Number', 'Work Order', 'Job Number', 'WO No', 'WO#'],
          ['title', 'track', 'status', 'shift']
        ) || prodHeaders[0] || '';

      const newTrackCol =
        findBestColumnMatch(
          prodHeaders,
          ['X Track ID', 'X-Track ID', 'X Track', 'Track ID', 'Tracking ID'],
          ['order', 'work', 'status']
        ) || '';

      onChangeConfig({
        ...config,
        jmsProductionSheet: val,
        jmsProductionWorkOrderCol: newWoCol,
        jmsProductionTrackIdCol: newTrackCol,
      });
      return;
    }

    onChangeConfig({
      ...config,
      [key]: val,
    });
  };

  const selectedProdSheet = config.jmsProductionSheet ? jmsSheets[config.jmsProductionSheet] : null;
  const prodHeaders = selectedProdSheet ? selectedProdSheet.headers : [];

  return (
    <div id="mappings-section" className="bg-white border border-neutral-200 rounded-xl mb-6 shadow-xs overflow-hidden">
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full px-5 py-3.5 flex items-center justify-between text-left hover:bg-neutral-50/70 transition-colors"
      >
        <div className="flex items-center gap-2.5">
          <div className="w-7 h-7 rounded bg-neutral-100 flex items-center justify-center text-neutral-700">
            <SlidersHorizontal className="w-3.5 h-3.5" />
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h2 className="text-xs font-semibold text-neutral-900">
                Column Mapping &amp; Rule Configurations
              </h2>
              <span className="text-[10px] text-neutral-600 bg-neutral-100 px-2 py-0.5 rounded font-medium inline-flex items-center gap-1">
                <BookmarkCheck className="w-3 h-3 text-neutral-700" />
                <span>Auto-chosen on upload</span>
              </span>
              {config.jmsProductionSheet && (
                <span className="text-[10px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded font-medium inline-flex items-center gap-1">
                  <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                  <span>Production Day to Day linked</span>
                </span>
              )}
            </div>
            <p className="text-[11px] text-neutral-500">
              Your preferred options are automatically selected on upload and retained in browser storage
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 text-xs font-medium text-neutral-600">
          <span>{isOpen ? 'Collapse' : 'Configure Mappings'}</span>
          {isOpen ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
        </div>
      </button>

      {isOpen && (
        <div className="px-5 pb-5 pt-2 border-t border-neutral-100">
          {/* Automatic Profile Info Banner */}
          <div className="mb-4 p-3 bg-emerald-50/80 border border-emerald-200 rounded-lg flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 text-xs">
            <div className="flex items-center gap-2">
              <Sparkles className="w-4 h-4 text-emerald-600 shrink-0" />
              <div className="text-[11px] text-emerald-950 leading-snug">
                <span className="font-semibold text-emerald-900">Automatic Profile Active: </span>
                Work Order · Work Order Title · Due Date · Activity · Job Type · XBRL Feedback (Quality ➜ FB Quality) · Sampling/Quality (STP Feedback) · X Track ID
              </div>
            </div>
            {onResetToAutoDetect && (
              <button
                type="button"
                onClick={onResetToAutoDetect}
                className="px-2.5 py-1 text-[11px] font-semibold text-emerald-800 bg-white border border-emerald-300 hover:bg-emerald-100/60 rounded shadow-2xs whitespace-nowrap self-end sm:self-auto"
                title="Re-apply preferred automatic choices"
              >
                Auto-Select Preferred Options
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            {/* Bridge Core Columns */}
            <div className="bg-neutral-50/60 p-4 rounded-lg border border-neutral-200 flex flex-col justify-between">
              <div>
                <h3 className="text-xs font-semibold text-neutral-900 mb-3 flex items-center justify-between">
                  <span className="flex items-center gap-1.5 font-bold">
                    <span className="px-1.5 py-0.5 rounded bg-blue-100 text-blue-800 text-[10px] font-mono">1</span>
                    <span>Bridge Core Columns</span>
                  </span>
                </h3>

                <div className="space-y-3">
                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>Work Order (Match Key) *</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[Abc]</span>
                    </label>
                    <select
                      value={config.bridgeWorkOrderCol}
                      onChange={(e) => handleUpdate('bridgeWorkOrderCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>Work Order Title</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[Abc]</span>
                    </label>
                    <select
                      value={config.bridgeTitleCol}
                      onChange={(e) => handleUpdate('bridgeTitleCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                    <p className="text-[10px] text-neutral-500 mt-0.5">
                      Strips &ldquo;| Bridge XBRL&rdquo; automatically
                    </p>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>Due Date</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[Date]</span>
                    </label>
                    <select
                      value={config.bridgeDueDateCol}
                      onChange={(e) => handleUpdate('bridgeDueDateCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                    <p className="text-[10px] text-neutral-500 mt-0.5">
                      Formatted as MM/DD/YYYY + Month
                    </p>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>Activity</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[⊞]</span>
                    </label>
                    <select
                      value={config.bridgeActivityCol}
                      onChange={(e) => handleUpdate('bridgeActivityCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                    <p className="text-[10px] text-neutral-500 mt-0.5">
                      Applies &ldquo;XBRL Update - ...&rdquo; rules
                    </p>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>Job Type</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[⊞]</span>
                    </label>
                    <select
                      value={config.bridgeJobTypeCol}
                      onChange={(e) => handleUpdate('bridgeJobTypeCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            </div>

            {/* XBRL Feedback Input Columns (from screenshot) */}
            <div className="bg-neutral-50/60 p-4 rounded-lg border-2 border-teal-200/80 bg-teal-50/20 flex flex-col justify-between">
              <div>
                <h3 className="text-xs font-semibold text-teal-950 mb-3 flex items-center justify-between">
                  <span className="flex items-center gap-1.5 font-bold">
                    <span className="px-1.5 py-0.5 rounded bg-teal-100 text-teal-800 text-[10px] font-mono">2</span>
                    <span>XBRL Feedback Columns</span>
                  </span>
                  <span className="text-[10px] text-teal-800 bg-teal-100/70 border border-teal-300 px-1.5 py-0.2 rounded font-semibold">
                    New Mapping
                  </span>
                </h3>

                <div className="space-y-2.5">
                  <div className="p-2 bg-teal-100/40 border border-teal-300 rounded-md">
                    <label className="block text-[11px] font-bold text-teal-950 mb-1 flex items-center justify-between">
                      <span>XBRL Feedback Quality *</span>
                      <span className="text-[9px] bg-teal-700 text-white px-1.5 py-0.2 rounded font-mono font-medium">
                        Maps to FB Quality
                      </span>
                    </label>
                    <select
                      value={config.bridgeXbrlFeedbackQualityCol}
                      onChange={(e) => handleUpdate('bridgeXbrlFeedbackQualityCol', e.target.value)}
                      className="w-full text-xs bg-white border border-teal-400 rounded px-2 py-1.5 font-medium text-teal-900 focus:outline-hidden focus:ring-1 focus:ring-teal-500"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                    <p className="text-[10px] text-teal-800 mt-1 font-medium">
                      ➜ Output column: &ldquo;JMS Cycle No / FB Quality&rdquo;
                    </p>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>XBRL Feedback Comments</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[Abc]</span>
                    </label>
                    <select
                      value={config.bridgeXbrlFeedbackCommentsCol}
                      onChange={(e) => handleUpdate('bridgeXbrlFeedbackCommentsCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>XBRL Feedback Status</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[-]</span>
                    </label>
                    <select
                      value={config.bridgeXbrlFeedbackStatusCol}
                      onChange={(e) => handleUpdate('bridgeXbrlFeedbackStatusCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>XBRL Feedback Reviewed</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[-]</span>
                    </label>
                    <select
                      value={config.bridgeXbrlFeedbackReviewedCol}
                      onChange={(e) => handleUpdate('bridgeXbrlFeedbackReviewedCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-[11px] font-medium text-neutral-700 mb-1 flex items-center justify-between">
                      <span>XBRL Feedback Validity</span>
                      <span className="text-[10px] text-neutral-400 font-mono">[-]</span>
                    </label>
                    <select
                      value={config.bridgeXbrlFeedbackValidityCol}
                      onChange={(e) => handleUpdate('bridgeXbrlFeedbackValidityCol', e.target.value)}
                      className="w-full text-xs bg-white border border-neutral-300 rounded px-2 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                    >
                      <option value="">-- Select Column --</option>
                      {bridgeHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="pt-2 border-t border-teal-200/80">
                    <label className="block text-[11px] font-bold text-neutral-800 mb-1 flex items-center justify-between">
                      <span>Sampling/Quality Prefill</span>
                      <span className="text-[9px] bg-indigo-100 text-indigo-800 px-1.5 py-0.2 rounded font-mono font-medium">
                        Output Col #11
                      </span>
                    </label>
                    <input
                      type="text"
                      value={config.samplingQualityPrefill || 'STP Feedback'}
                      onChange={(e) => handleUpdate('samplingQualityPrefill', e.target.value)}
                      placeholder="STP Feedback"
                      className="w-full text-xs bg-white border border-indigo-300 rounded px-2 py-1.5 font-semibold text-indigo-950 focus:outline-hidden focus:ring-1 focus:ring-indigo-500"
                    />
                    <p className="text-[10px] text-indigo-700 mt-1 font-medium">
                      Fills &ldquo;STP Feedback&rdquo; wherever &ldquo;JMS Cycle No / FB Quality&rdquo; is populated
                    </p>
                  </div>
                </div>
              </div>
            </div>

            {/* Primary JMS / SaaS Columns */}
            <div className="bg-neutral-50/60 p-4 rounded-lg border border-neutral-200">
              <h3 className="text-xs font-semibold text-neutral-900 mb-3 flex items-center gap-1.5">
                <span>JMS Main Sheet Columns</span>
              </h3>

              <div className="space-y-3">
                <div>
                  <label className="block text-[11px] font-medium text-neutral-700 mb-1">
                    Work Order / Job Number (Match Key) *
                  </label>
                  <select
                    value={config.jmsWorkOrderCol}
                    onChange={(e) => handleUpdate('jmsWorkOrderCol', e.target.value)}
                    className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                  >
                    <option value="">-- Select Column --</option>
                    {jmsHeaders.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-medium text-neutral-700 mb-1">
                    X Track ID (Primary)
                  </label>
                  <select
                    value={config.jmsTrackIdCol}
                    onChange={(e) => handleUpdate('jmsTrackIdCol', e.target.value)}
                    className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                  >
                    <option value="">-- Select Column --</option>
                    {jmsHeaders.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-medium text-neutral-700 mb-1">
                    Grabber Day to Day (Priority 1 for Form Type)
                  </label>
                  <select
                    value={config.jmsGrabberCol}
                    onChange={(e) => handleUpdate('jmsGrabberCol', e.target.value)}
                    className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                  >
                    <option value="">-- Select Column --</option>
                    {jmsHeaders.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-medium text-neutral-700 mb-1">
                    Job Type (Fallback for Form Type)
                  </label>
                  <select
                    value={config.jmsJobTypeCol}
                    onChange={(e) => handleUpdate('jmsJobTypeCol', e.target.value)}
                    className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                  >
                    <option value="">-- Select Column --</option>
                    {jmsHeaders.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            {/* Production Day to Day Secondary Lookup */}
            <div className="bg-neutral-50/60 p-4 rounded-lg border border-neutral-200">
              <h3 className="text-xs font-semibold text-neutral-900 mb-3 flex items-center justify-between">
                <span>Production Day to Day (Lookup)</span>
                {config.jmsProductionSheet && (
                  <span className="text-[10px] text-emerald-700 font-normal">Active</span>
                )}
              </h3>

              <div className="space-y-3">
                <div>
                  <label className="block text-[11px] font-medium text-neutral-700 mb-1">
                    Production Sheet Name
                  </label>
                  <select
                    value={config.jmsProductionSheet}
                    onChange={(e) => handleUpdate('jmsProductionSheet', e.target.value)}
                    className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                  >
                    <option value="">-- None (No secondary sheet) --</option>
                    {jmsSheetNames.map((s) => (
                      <option key={s} value={s}>
                        {s} ({jmsSheets[s]?.rowCount || 0} rows)
                      </option>
                    ))}
                  </select>
                  <p className="text-[10px] text-neutral-500 mt-0.5">
                    Checks this sheet for X Track ID by work order number
                  </p>
                </div>

                {config.jmsProductionSheet && (
                  <>
                    <div>
                      <label className="block text-[11px] font-medium text-neutral-700 mb-1">
                        Work Order Column (in Production Sheet)
                      </label>
                      <select
                        value={config.jmsProductionWorkOrderCol}
                        onChange={(e) =>
                          handleUpdate('jmsProductionWorkOrderCol', e.target.value)
                        }
                        className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                      >
                        <option value="">-- Select Column --</option>
                        {prodHeaders.map((h) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[11px] font-medium text-neutral-700 mb-1">
                        X Track ID Column (in Production Sheet)
                      </label>
                      <select
                        value={config.jmsProductionTrackIdCol}
                        onChange={(e) =>
                          handleUpdate('jmsProductionTrackIdCol', e.target.value)
                        }
                        className="w-full text-xs bg-white border border-neutral-300 rounded px-2.5 py-1.5 focus:outline-hidden focus:ring-1 focus:ring-neutral-400"
                      >
                        <option value="">-- Select Column --</option>
                        {prodHeaders.map((h) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>

          {/* Footer with Storage Notice and Reset to Auto-Detect button */}
          <div className="mt-4 pt-3 border-t border-neutral-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
            <div className="flex items-start gap-2 text-neutral-600">
              <Info className="w-4 h-4 text-neutral-500 shrink-0 mt-0.5" />
              <div className="text-[11px] leading-relaxed">
                <span className="font-semibold text-neutral-800">Persistent Configuration: </span>
                All selections are saved automatically to your browser and will be retained when you upload new files or revisit the app.
              </div>
            </div>

            {onResetToAutoDetect && (
              <button
                type="button"
                onClick={onResetToAutoDetect}
                className="inline-flex items-center gap-1.5 px-3 py-1 text-xs font-medium text-neutral-700 bg-white border border-neutral-300 hover:bg-neutral-50 rounded-md transition-colors whitespace-nowrap self-end sm:self-auto shrink-0 shadow-2xs"
                title="Clear retained preferences and re-run automatic column detection"
              >
                <RotateCcw className="w-3.5 h-3.5 text-neutral-500" />
                <span>Reset to Auto-Detect</span>
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
