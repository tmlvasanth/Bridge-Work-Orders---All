import React from 'react';
import { ReconciliationSummary, ReconciliationFilterType } from '../types';
import { CheckCircle2, Layers, CheckSquare } from 'lucide-react';

interface SummaryStatsProps {
  summary: ReconciliationSummary;
  activeFilter: ReconciliationFilterType;
  onSelectFilter: (filter: ReconciliationFilterType) => void;
}

export const SummaryStats: React.FC<SummaryStatsProps> = ({
  summary,
  activeFilter,
  onSelectFilter,
}) => {
  const totalBridgeRecords = summary.matchedCount + summary.unmatchedBridgeCount;

  return (
    <div className="space-y-3 mb-6">
      {/* Primary KPI filter buttons */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3">
        {/* Matched Count */}
        <button
          type="button"
          onClick={() => onSelectFilter('matched')}
          className={`p-3.5 text-left rounded-xl border-2 transition-all shadow-xs ${
            activeFilter === 'matched'
              ? 'bg-neutral-900 text-white border-neutral-900 shadow-sm ring-2 ring-emerald-500/20'
              : 'bg-white text-neutral-900 border-neutral-200 hover:border-emerald-500/50 hover:shadow-sm'
          }`}
        >
          <div className="flex items-center justify-between mb-1">
            <span
              className={`block text-[11px] font-bold uppercase tracking-wider ${
                activeFilter === 'matched' ? 'text-emerald-300' : 'text-emerald-700'
              }`}
            >
              Matched Work Orders
            </span>
            <span
              className={`w-2 h-2 rounded-full ${
                activeFilter === 'matched' ? 'bg-emerald-400' : 'bg-emerald-500'
              }`}
            />
          </div>
          <div className="mt-1 flex items-baseline gap-2">
            <span className="text-2xl font-bold font-mono tabular-nums">
              {summary.matchedCount}
            </span>
            <span
              className={`text-xs font-mono tabular-nums ${
                activeFilter === 'matched' ? 'text-neutral-300' : 'text-neutral-500'
              }`}
            >
              ({summary.matchRatePercent}%)
            </span>
          </div>
          <p
            className={`text-[11px] mt-1.5 ${
              activeFilter === 'matched' ? 'text-neutral-300' : 'text-neutral-500'
            }`}
          >
            Ready for Excel / Standalone HTML export
          </p>
        </button>

        {/* Unmatched in Bridge */}
        <button
          type="button"
          onClick={() => onSelectFilter('unmatched_bridge')}
          className={`p-3.5 text-left rounded-xl border-2 transition-all shadow-xs ${
            activeFilter === 'unmatched_bridge'
              ? 'bg-neutral-900 text-white border-neutral-900 shadow-sm ring-2 ring-amber-500/20'
              : 'bg-white text-neutral-900 border-neutral-200 hover:border-amber-500/50 hover:shadow-sm'
          }`}
        >
          <div className="flex items-center justify-between mb-1">
            <span
              className={`block text-[11px] font-bold uppercase tracking-wider ${
                activeFilter === 'unmatched_bridge' ? 'text-amber-300' : 'text-amber-700'
              }`}
            >
              Bridge Unmatched
            </span>
            <span
              className={`w-2 h-2 rounded-full ${
                activeFilter === 'unmatched_bridge' ? 'bg-amber-400' : 'bg-amber-500'
              }`}
            />
          </div>
          <div className="mt-1 flex items-baseline gap-2">
            <span className="text-2xl font-bold font-mono tabular-nums">
              {summary.unmatchedBridgeCount}
            </span>
            <span
              className={`text-xs font-mono tabular-nums ${
                activeFilter === 'unmatched_bridge' ? 'text-neutral-300' : 'text-neutral-500'
              }`}
            >
              of {summary.totalBridgeRows}
            </span>
          </div>
          <p
            className={`text-[11px] mt-1.5 ${
              activeFilter === 'unmatched_bridge' ? 'text-neutral-300' : 'text-neutral-500'
            }`}
          >
            Present in Base Excel, missing in JMS
          </p>
        </button>

        {/* Matched + Bridge Unmatched (Together) */}
        <button
          type="button"
          onClick={() => onSelectFilter('matched_and_bridge_unmatched')}
          className={`p-3.5 text-left rounded-xl border-2 transition-all shadow-xs relative overflow-hidden ${
            activeFilter === 'matched_and_bridge_unmatched'
              ? 'bg-neutral-900 text-white border-teal-500 shadow-sm ring-2 ring-teal-400/30'
              : 'bg-white text-neutral-900 border-teal-300 hover:border-teal-500 hover:shadow-sm'
          }`}
          title="Select and download Matched work orders and Bridge unmatched work orders together"
        >
          <div className="flex items-center justify-between mb-1">
            <span
              className={`block text-[11px] font-bold uppercase tracking-wider ${
                activeFilter === 'matched_and_bridge_unmatched' ? 'text-teal-300' : 'text-teal-800'
              }`}
            >
              Matched + Bridge Unmatched
            </span>
            <span
              className={`w-2 h-2 rounded-full ${
                activeFilter === 'matched_and_bridge_unmatched' ? 'bg-teal-400' : 'bg-teal-500'
              }`}
            />
          </div>
          <div className="mt-1 flex items-baseline gap-2">
            <span className="text-2xl font-bold font-mono tabular-nums text-teal-600 dark:text-teal-400">
              {totalBridgeRecords}
            </span>
            <span
              className={`text-xs font-mono tabular-nums ${
                activeFilter === 'matched_and_bridge_unmatched' ? 'text-neutral-300' : 'text-neutral-500'
              }`}
            >
              (Bridge Total)
            </span>
          </div>
          <p
            className={`text-[11px] mt-1.5 ${
              activeFilter === 'matched_and_bridge_unmatched' ? 'text-teal-200' : 'text-teal-700'
            } font-medium`}
          >
            All Base Bridge Records together
          </p>
        </button>

        {/* Unmatched in JMS */}
        <button
          type="button"
          onClick={() => onSelectFilter('unmatched_jms')}
          className={`p-3.5 text-left rounded-xl border-2 transition-all shadow-xs ${
            activeFilter === 'unmatched_jms'
              ? 'bg-neutral-900 text-white border-neutral-900 shadow-sm ring-2 ring-blue-500/20'
              : 'bg-white text-neutral-900 border-neutral-200 hover:border-blue-500/50 hover:shadow-sm'
          }`}
        >
          <div className="flex items-center justify-between mb-1">
            <span
              className={`block text-[11px] font-bold uppercase tracking-wider ${
                activeFilter === 'unmatched_jms' ? 'text-blue-300' : 'text-blue-700'
              }`}
            >
              JMS Unmatched
            </span>
            <span
              className={`w-2 h-2 rounded-full ${
                activeFilter === 'unmatched_jms' ? 'bg-blue-400' : 'bg-blue-500'
              }`}
            />
          </div>
          <div className="mt-1 flex items-baseline gap-2">
            <span className="text-2xl font-bold font-mono tabular-nums">
              {summary.unmatchedJmsCount}
            </span>
            <span
              className={`text-xs font-mono tabular-nums ${
                activeFilter === 'unmatched_jms' ? 'text-neutral-300' : 'text-neutral-500'
              }`}
            >
              of {summary.totalJmsRows}
            </span>
          </div>
          <p
            className={`text-[11px] mt-1.5 ${
              activeFilter === 'unmatched_jms' ? 'text-neutral-300' : 'text-neutral-500'
            }`}
          >
            Present in JMS, missing in Bridge
          </p>
        </button>

        {/* Total Combined */}
        <button
          type="button"
          onClick={() => onSelectFilter('all')}
          className={`p-4 text-left rounded-xl border-2 transition-all shadow-xs ${
            activeFilter === 'all'
              ? 'bg-neutral-900 text-white border-neutral-900 shadow-sm ring-2 ring-purple-500/20'
              : 'bg-white text-neutral-900 border-neutral-200 hover:border-purple-500/50 hover:shadow-sm'
          }`}
        >
          <div className="flex items-center justify-between mb-1">
            <span
              className={`block text-[11px] font-bold uppercase tracking-wider ${
                activeFilter === 'all' ? 'text-neutral-300' : 'text-neutral-600'
              }`}
            >
              All Records
            </span>
            <span
              className={`w-2 h-2 rounded-full ${
                activeFilter === 'all' ? 'bg-purple-400' : 'bg-neutral-400'
              }`}
            />
          </div>
          <div className="mt-1 flex items-baseline gap-2">
            <span className="text-2xl font-bold font-mono tabular-nums">
              {summary.matchedCount + summary.unmatchedBridgeCount + summary.unmatchedJmsCount}
            </span>
            <span
              className={`text-xs font-mono tabular-nums ${
                activeFilter === 'all' ? 'text-neutral-300' : 'text-neutral-500'
              }`}
            >
              total
            </span>
          </div>
          <p
            className={`text-[11px] mt-1.5 ${
              activeFilter === 'all' ? 'text-neutral-300' : 'text-neutral-500'
            }`}
          >
            Complete reconciliation pool
          </p>
        </button>
      </div>

      {/* Production Day to Day Cross-Reference Status Strip */}
      {summary.productionCrossMatchedCount > 0 && (
        <div className="bg-emerald-50 border border-emerald-300 text-emerald-950 rounded-xl px-4 py-2.5 flex items-center justify-between gap-3 text-xs shadow-2xs">
          <div className="flex items-center gap-2">
            <div className="w-5 h-5 rounded bg-emerald-600 text-white flex items-center justify-center font-bold shrink-0">
              <CheckCircle2 className="w-3.5 h-3.5" />
            </div>
            <span className="font-semibold">
              Production Day to Day Cross-Reference Active:
            </span>
            <span className="text-emerald-800">
              Successfully populated <strong>{summary.productionCrossMatchedCount}</strong> missing X Track ID{summary.productionCrossMatchedCount > 1 ? 's' : ''} by matching Work Order Numbers in the Production Day to Day sheet.
            </span>
          </div>
          <span className="text-[11px] font-mono text-emerald-700 bg-white border border-emerald-200 px-2 py-0.5 rounded font-semibold shrink-0">
            JMS SaaS Cross-Lookup
          </span>
        </div>
      )}
    </div>
  );
};
