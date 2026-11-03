import React, { useState } from 'react';
import { ChevronDown, ChevronUp, CheckCircle, ArrowRight, Table, Layers, Sparkles } from 'lucide-react';

export const RuleCard: React.FC = () => {
  const [isOpen, setIsOpen] = useState(false);

  const transformationRules = [
    {
      heading: '0. Sheet Selection (Bridge)',
      logic: 'Exclusively reads sheet "Table 1" / "Table1" from "Bridge work order - All" Excel workbook, ignoring default "Sheet1"',
      type: 'Auto Sheet Filter',
      output: 'Table1',
    },
    {
      heading: '1. Team',
      logic: 'Static "Bridge" or Base Excel Cell B1 value',
      type: 'Fixed / Cell B1',
      output: 'Bridge',
    },
    {
      heading: '2. JMS Job Number / Work Order No',
      logic: 'Primary matching key normalized across Base Excel and JMS SaaS workbook',
      type: 'Match Key',
      output: 'WO-XXXXX',
    },
    {
      heading: '3. JMS Cycle No / FB Quality',
      logic: 'Mapped directly from the input column "XBRL Feedback Quality" from Bridge Excel',
      type: 'Direct Column Map',
      output: 'Pass / Cycle 1 / Cycle 2',
    },
    {
      heading: '4. X Track ID',
      logic: 'First checked in JMS active sheet. If missing or blank, looks up the respective Work Order Number in the "Production Day to Day" sheet in JMS SaaS workbook to pull the X Track ID.',
      type: 'Cross-Sheet Lookup',
      output: 'TRK-XXXX',
    },
    {
      heading: '5. Job Name',
      logic: 'Work Order Title with "| Bridge XBRL" or "- Bridge XBRL" stripped cleanly',
      type: 'Text Cleanup',
      output: 'Client Name Q3 Financials',
    },
    {
      heading: '6. Form Type',
      logic: 'Matched from "Grabber Day to Day"; if blank or not available, takes value from "Job Type"',
      type: 'Fallback Priority',
      output: '10-Q, 10-K, 8-K',
    },
    {
      heading: '7. Milestone Transformation',
      logic: 'Transforms "XBRL Update - Final to Client" → "Final to Client", "XBRL Update - Non Standard" → "Non Standard", "XBRL Update - Standard" → "Standard"',
      type: 'Transformation',
      output: 'Standard / Final to Client',
    },
    {
      heading: '8. Sampling/Quality',
      logic: 'Filled with "STP Feedback" whenever the "JMS Cycle No / FB Quality" column is filled in the output row (otherwise blank)',
      type: 'Conditional Rule',
      output: 'STP Feedback (or blank)',
    },
    {
      heading: '9. Delivered Date',
      logic: 'Due Date from Base Excel formatted strictly as MM/DD/YYYY',
      type: 'Date Formatting',
      output: 'MM/DD/YYYY',
    },
    {
      heading: '10. Month',
      logic: 'Full month name derived directly from Delivered Date (e.g. October, November)',
      type: 'Date Derived',
      output: 'October, November',
    },
    {
      heading: '11. XBRL Feedback Comments (Last Column)',
      logic: 'Directly mapped from "XBRL Feedback Comments" from Bridge work order - all for each respective job',
      type: 'Direct Column Map',
      output: 'Reviewer feedback / audit comments',
    },
  ];

  const inputColumnsScreenshot = [
    { name: 'Work Order', type: 'Abc', desc: 'Normalized key for matching against JMS Work Order No' },
    { name: 'Work Order Title', type: 'Abc', desc: 'Cleansed into Job Name (removes | Bridge XBRL)' },
    { name: 'Due Date', type: 'Date', desc: 'Formatted into Delivered Date (MM/DD/YYYY) and Month' },
    { name: 'Activity', type: 'Table', desc: 'Transformed into Milestone (Final to Client, Standard, etc.)' },
    { name: 'Job Type', type: 'Table', desc: 'Fallback for Form Type if Grabber Day to Day is not present' },
    { name: 'XBRL Feedback Comments', type: 'Abc/def', desc: 'Input feedback details from reviewer/client mapped to last column in output', highlight: true },
    { name: 'XBRL Feedback Quality', type: 'Toggle', desc: 'Mapped to output column "JMS Cycle No / FB Quality"', highlight: true },
    { name: 'XBRL Feedback Status', type: 'Toggle', desc: 'Feedback resolution status (Completed, In Review)' },
    { name: 'XBRL Feedback Reviewed', type: 'Toggle', desc: 'Reviewer audit confirmation (Yes/No)' },
    { name: 'XBRL Feedback Validity', type: 'Toggle', desc: 'Feedback validity assessment (Valid/Under Review)' },
  ];

  const standard15Headings = [
    'Team',
    'JMS Job Number/Work Order No',
    'JMS Cycle No / FB Quality',
    'XOR Ver',
    'X Track ID',
    'Job Name',
    'Form Type',
    'Filing Period',
    'Milestone',
    'JMS Team',
    'Sampling/Quality',
    'Consultant Name',
    '2nd Reviewer Name',
    'Delivered Date',
    'Month',
    'XBRL Feedback Comments',
  ];

  return (
    <div className="bg-white border-2 border-neutral-300 rounded-xl mb-6 shadow-xs overflow-hidden">
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full px-5 py-3.5 flex items-center justify-between text-left hover:bg-neutral-50/70 transition-colors"
      >
        <div className="flex items-center gap-2.5">
          <div className="w-7 h-7 rounded bg-blue-100 text-blue-800 flex items-center justify-center font-bold">
            <Layers className="w-3.5 h-3.5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-xs font-bold text-neutral-900">
                Bridge Work Orders - Transformation &amp; Cross-Referencing Rules
              </h2>
              <span className="text-[10px] text-emerald-800 bg-emerald-50 border border-emerald-300 px-2 py-0.5 rounded font-bold">
                16 Standard Headings (incl. Feedback Comments)
              </span>
            </div>
            <p className="text-[11px] text-neutral-500 mt-0.5">
              Automated rules for X Track ID (including Production Day to Day cross-lookup), Job Name cleanup, Milestones, and MM/DD/YYYY dates
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 text-xs font-medium text-neutral-600">
          <span>{isOpen ? 'Hide Rules' : 'View Rule Details'}</span>
          {isOpen ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
        </div>
      </button>

      {isOpen && (
        <div className="px-5 pb-5 pt-3 border-t border-neutral-200 space-y-4">
          <div>
            <h3 className="text-xs font-bold text-neutral-800 uppercase tracking-wider mb-2 flex items-center gap-1.5">
              <Table className="w-3.5 h-3.5 text-blue-600" />
              <span>Transformation &amp; Mapping Logic</span>
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-2.5 text-xs">
              {transformationRules.map((rule) => (
                <div
                  key={rule.heading}
                  className="p-3 bg-neutral-50 rounded-lg border border-neutral-200 flex flex-col justify-between"
                >
                  <div>
                    <div className="flex items-center justify-between gap-1 mb-1">
                      <span className="font-bold text-neutral-900 truncate">
                        {rule.heading}
                      </span>
                      <span className="text-[10px] text-blue-700 bg-blue-50 px-1.5 py-0.5 rounded font-mono font-medium">
                        {rule.type}
                      </span>
                    </div>
                    <p className="text-neutral-600 text-[11px] leading-relaxed">
                      {rule.logic}
                    </p>
                  </div>
                  <div className="mt-2 pt-1.5 border-t border-neutral-200/60 text-[10px] text-neutral-500 font-mono">
                    Example: {rule.output}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Input Columns from Screenshot */}
          <div>
            <h3 className="text-xs font-bold text-neutral-800 uppercase tracking-wider mb-2 flex items-center justify-between">
              <span className="flex items-center gap-1.5">
                <Sparkles className="w-3.5 h-3.5 text-teal-600" />
                <span>Recognized Input Columns (from Screenshot)</span>
              </span>
              <span className="text-[10px] text-teal-800 bg-teal-50 border border-teal-200 px-2 py-0.5 rounded font-medium">
                10 Configured Columns
              </span>
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-2 text-xs">
              {inputColumnsScreenshot.map((col) => (
                <div
                  key={col.name}
                  className={`p-2.5 rounded-lg border text-[11px] ${
                    col.highlight
                      ? 'bg-teal-50/70 border-teal-300 ring-1 ring-teal-200'
                      : 'bg-neutral-50 border-neutral-200'
                  }`}
                >
                  <div className="flex items-center justify-between gap-1 mb-1">
                    <span className="font-semibold text-neutral-900 truncate">
                      {col.name}
                    </span>
                    <span className="text-[9px] font-mono px-1 py-0.2 rounded bg-neutral-200/80 text-neutral-700">
                      {col.type}
                    </span>
                  </div>
                  <p className="text-[10px] text-neutral-500 leading-tight">
                    {col.desc}
                  </p>
                </div>
              ))}
            </div>
          </div>

          <div>
            <h3 className="text-xs font-bold text-neutral-800 uppercase tracking-wider mb-2">
              Canonical 15 Output Columns Order
            </h3>
            <div className="flex flex-wrap gap-1.5">
              {standard15Headings.map((heading, i) => (
                <span
                  key={heading}
                  className={`inline-flex items-center gap-1 text-[11px] border px-2.5 py-1 rounded font-mono ${
                    heading === 'JMS Cycle No / FB Quality'
                      ? 'bg-teal-50 text-teal-900 border-teal-300 font-bold'
                      : heading === 'Sampling/Quality'
                      ? 'bg-indigo-50 text-indigo-900 border-indigo-300 font-bold'
                      : 'bg-slate-100 text-slate-800 border-slate-300'
                  }`}
                >
                  <span className="text-neutral-400 font-bold">{i + 1}.</span>
                  <span className="font-semibold">{heading}</span>
                  {heading === 'JMS Cycle No / FB Quality' && (
                    <span className="text-[9px] bg-teal-200 text-teal-800 px-1 rounded ml-1 font-sans">
                      Mapped
                    </span>
                  )}
                  {heading === 'Sampling/Quality' && (
                    <span className="text-[9px] bg-indigo-200 text-indigo-800 px-1 rounded ml-1 font-sans">
                      Prefilled
                    </span>
                  )}
                </span>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
