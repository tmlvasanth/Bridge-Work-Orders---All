const fs = require('fs');
const path = require('path');

const xlsxPath = path.join(__dirname, '../node_modules/xlsx/dist/xlsx.full.min.js');
let xlsxScript = '';
if (fs.existsSync(xlsxPath)) {
  xlsxScript = fs.readFileSync(xlsxPath, 'utf8');
}

// Generate the standalone HTML content
function generateHtml() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Bridge Work Orders - Standalone Application (Online & Offline)</title>
  <meta name="description" content="Standalone Offline & Online Bridge Work Orders Reconciler with automated transformation rules and Excel export.">
  <style>
    :root {
      --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      --font-mono: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
      --color-primary: #0f172a;
      --color-primary-hover: #1e293b;
      --color-accent: #0284c7;
      --color-bg: #f8fafc;
      --color-surface: #ffffff;
      --color-border: #cbd5e1;
      --color-border-subtle: #e2e8f0;
      --color-text: #0f172a;
      --color-text-muted: #64748b;
    }

    *, *::before, *::after {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: var(--font-sans);
      background-color: var(--color-bg);
      color: var(--color-text);
      line-height: 1.5;
      font-size: 13px;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }

    header {
      background: var(--color-surface);
      border-bottom: 1px solid var(--color-border);
      position: sticky;
      top: 0;
      z-index: 40;
      box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }

    .header-inner {
      max-width: 1440px;
      margin: 0 auto;
      padding: 0 20px;
      height: 60px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }

    .logo-group {
      display: flex;
      align-items: center;
      gap: 12px;
    }

    .logo-icon {
      width: 34px;
      height: 34px;
      border-radius: 8px;
      background: #0f172a;
      color: #fff;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 700;
      font-size: 16px;
    }

    .logo-text h1 {
      font-size: 15px;
      font-weight: 800;
      letter-spacing: -0.01em;
      color: #0f172a;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .badge-standalone {
      font-size: 10px;
      font-weight: 700;
      background: #ecfdf5;
      color: #065f46;
      border: 1px solid #a7f3d0;
      padding: 2px 7px;
      border-radius: 4px;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }

    .status-badge {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      font-size: 11px;
      padding: 3px 8px;
      border-radius: 20px;
      background: #f1f5f9;
      border: 1px solid #cbd5e1;
      color: #334155;
    }

    .status-dot {
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background: #10b981;
    }

    .header-actions {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    main {
      max-width: 1440px;
      width: 100%;
      margin: 0 auto;
      padding: 24px 20px 48px;
      flex: 1;
    }

    .banner {
      background: var(--color-surface);
      border: 2px solid var(--color-border);
      border-radius: 12px;
      padding: 18px 22px;
      margin-bottom: 20px;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }

    .banner-title {
      font-size: 15px;
      font-weight: 800;
      color: #0f172a;
      margin-bottom: 4px;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .banner-desc {
      font-size: 12px;
      color: #475569;
      max-width: 900px;
      line-height: 1.5;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      padding: 7px 14px;
      font-size: 12px;
      font-weight: 600;
      border-radius: 6px;
      cursor: pointer;
      border: 1px solid var(--color-border);
      background: #ffffff;
      color: #334155;
      text-decoration: none;
      transition: all 0.15s ease;
      white-space: nowrap;
    }

    .btn:hover {
      background: #f1f5f9;
      border-color: #94a3b8;
    }

    .btn-primary {
      background: #0f172a;
      color: #ffffff;
      border-color: #0f172a;
    }

    .btn-primary:hover {
      background: #1e293b;
    }

    .btn-emerald {
      background: #047857;
      color: #ffffff;
      border-color: #047857;
      font-weight: 700;
    }

    .btn-emerald:hover {
      background: #065f46;
    }

    .btn-indigo {
      background: #eef2ff;
      color: #3730a3;
      border-color: #c7d2fe;
      font-weight: 600;
    }

    .btn-indigo:hover {
      background: #e0e7ff;
    }

    .btn-sm {
      padding: 4px 10px;
      font-size: 11px;
    }

    .grid-2 {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
      gap: 20px;
      margin-bottom: 20px;
    }

    .card {
      background: var(--color-surface);
      border: 2px solid var(--color-border);
      border-radius: 12px;
      padding: 18px 20px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }

    .card-header {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      margin-bottom: 14px;
    }

    .card-title {
      font-size: 14px;
      font-weight: 700;
      color: #0f172a;
    }

    .card-subtitle {
      font-size: 11px;
      color: #64748b;
      margin-top: 2px;
    }

    .dropzone {
      border: 2px dashed var(--color-border);
      border-radius: 8px;
      padding: 24px 16px;
      text-align: center;
      background: #f8fafc;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      min-height: 140px;
    }

    .dropzone:hover, .dropzone.dragover {
      border-color: #2563eb;
      background: #eff6ff;
    }

    .dropzone-icon {
      width: 40px;
      height: 40px;
      border-radius: 50%;
      background: #ffffff;
      border: 1px solid var(--color-border-subtle);
      display: flex;
      align-items: center;
      justify-content: center;
      margin-bottom: 8px;
      font-size: 18px;
    }

    .loaded-file-bar {
      background: #f8fafc;
      border: 1px solid var(--color-border);
      border-radius: 8px;
      padding: 12px 14px;
    }

    .file-info-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }

    .sheet-select-row {
      margin-top: 10px;
      padding-top: 10px;
      border-top: 1px solid #e2e8f0;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    select, input[type="text"] {
      width: 100%;
      padding: 6px 10px;
      font-size: 12px;
      font-family: inherit;
      border: 1px solid var(--color-border);
      border-radius: 6px;
      background: #ffffff;
      color: #0f172a;
      outline: none;
    }

    select:focus, input[type="text"]:focus {
      border-color: #2563eb;
      box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.12);
    }

    .expected-cols {
      margin-top: 12px;
      padding-top: 10px;
      border-top: 1px solid #f1f5f9;
      font-size: 11px;
      color: #64748b;
      line-height: 1.4;
    }

    .expected-cols strong {
      color: #334155;
    }

    /* Accordions & Drawers */
    .drawer {
      background: var(--color-surface);
      border: 2px solid var(--color-border);
      border-radius: 12px;
      margin-bottom: 20px;
      overflow: hidden;
      box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }

    .drawer-trigger {
      width: 100%;
      padding: 14px 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: none;
      border: none;
      cursor: pointer;
      font-family: inherit;
      text-align: left;
    }

    .drawer-trigger:hover {
      background: #f8fafc;
    }

    .drawer-content {
      padding: 16px 20px 20px;
      border-top: 1px solid var(--color-border);
      background: #ffffff;
    }

    .drawer-content.hidden {
      display: none;
    }

    .mapping-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 16px;
      margin-top: 12px;
    }

    .mapping-card {
      background: #f8fafc;
      border: 1px solid var(--color-border);
      border-radius: 8px;
      padding: 14px;
    }

    .mapping-card-accent {
      border: 2px solid #5eead4;
      background: #f0fdfa;
    }

    .mapping-card h4 {
      font-size: 12px;
      font-weight: 700;
      color: #0f172a;
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .field-group {
      margin-bottom: 10px;
    }

    .field-group:last-child {
      margin-bottom: 0;
    }

    .field-group label {
      display: flex;
      justify-content: space-between;
      font-size: 11px;
      font-weight: 600;
      color: #334155;
      margin-bottom: 4px;
    }

    .field-hint {
      font-size: 10px;
      color: #64748b;
      margin-top: 2px;
    }

    /* Summary KPI Bar */
    .summary-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
      margin-bottom: 20px;
    }

    .summary-card {
      background: var(--color-surface);
      border: 2px solid var(--color-border);
      border-radius: 10px;
      padding: 14px 16px;
      cursor: pointer;
      transition: all 0.15s ease;
    }

    .summary-card:hover {
      border-color: #94a3b8;
    }

    .summary-card.active {
      border-color: #0f172a;
      background: #f8fafc;
      box-shadow: 0 0 0 1px #0f172a;
    }

    .summary-label {
      font-size: 11px;
      font-weight: 600;
      color: #64748b;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }

    .summary-value {
      font-size: 24px;
      font-weight: 800;
      color: #0f172a;
      margin-top: 4px;
      font-family: var(--font-mono);
    }

    .summary-subtext {
      font-size: 11px;
      color: #64748b;
      margin-top: 2px;
    }

    /* Table Component */
    .table-section {
      background: var(--color-surface);
      border: 2px solid var(--color-border);
      border-radius: 12px;
      overflow: hidden;
      box-shadow: 0 1px 4px rgba(0,0,0,0.03);
    }

    .table-toolbar {
      padding: 12px 18px;
      border-bottom: 2px solid var(--color-border);
      background: #f8fafc;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }

    .tab-pills {
      display: flex;
      gap: 4px;
      background: #e2e8f0;
      padding: 3px;
      border-radius: 8px;
    }

    .tab-btn {
      padding: 5px 12px;
      font-size: 11px;
      font-weight: 600;
      border-radius: 6px;
      border: none;
      background: transparent;
      color: #475569;
      cursor: pointer;
      transition: all 0.15s;
    }

    .tab-btn.active {
      background: #ffffff;
      color: #0f172a;
      box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }

    .table-controls {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }

    .search-box {
      width: 240px;
    }

    .table-wrapper {
      overflow-x: auto;
      max-height: 600px;
      position: relative;
    }

    table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 12px;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 20;
      padding: 9px 12px;
      text-align: left;
      font-weight: 700;
      white-space: nowrap;
      border-right: 1px solid var(--color-border);
      border-bottom: 2px solid var(--color-border);
      background: #f1f5f9;
      color: #1e293b;
    }

    td {
      padding: 8px 12px;
      white-space: nowrap;
      border-right: 1px solid #f1f5f9;
      border-bottom: 1px solid #e2e8f0;
      color: #334155;
    }

    tr:nth-child(even) td {
      background-color: #fafbfc;
    }

    tr:hover td {
      background-color: #f0f7ff !important;
    }

    /* Col headers custom themes */
    th.th-wo { background: #eff6ff; color: #1e3a8a; border-bottom-color: #93c5fd; }
    th.th-fb { background: #f0fdfa; color: #0f766e; border-bottom-color: #99f6e4; }
    th.th-trk { background: #f0f9ff; color: #0369a1; border-bottom-color: #7dd3fc; }
    th.th-job { background: #f0fdf4; color: #166534; border-bottom-color: #86efac; }
    th.th-form { background: #fffbeb; color: #92400e; border-bottom-color: #fcd34d; }
    th.th-milestone { background: #faf5ff; color: #6b21a8; border-bottom-color: #d8b4fe; }
    th.th-sampling { background: #eef2ff; color: #3730a3; border-bottom-color: #c7d2fe; }
    th.th-date { background: #fff1f2; color: #9f1239; border-bottom-color: #fda4af; }

    .tag-badge {
      display: inline-block;
      padding: 2px 7px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 600;
    }

    .tag-fb {
      background: #ccfbf1;
      color: #0f766e;
      border: 1px solid #99f6e4;
    }

    .tag-stp {
      background: #e0e7ff;
      color: #3730a3;
      border: 1px solid #c7d2fe;
    }

    .tag-form {
      background: #fef3c7;
      color: #92400e;
      border: 1px solid #fde68a;
      font-family: var(--font-mono);
    }

    .table-footer {
      padding: 10px 18px;
      border-top: 1px solid var(--color-border);
      background: #f8fafc;
      display: flex;
      align-items: center;
      justify-content: space-between;
      font-size: 12px;
      color: #64748b;
    }

    .pagination {
      display: flex;
      align-items: center;
      gap: 6px;
    }

    /* Modal */
    .modal-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(0, 0, 0, 0.45);
      backdrop-filter: blur(2px);
      z-index: 100;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 16px;
    }

    .modal-backdrop.hidden {
      display: none;
    }

    .modal-box {
      background: #ffffff;
      border-radius: 12px;
      border: 2px solid var(--color-border);
      max-width: 760px;
      width: 100%;
      max-height: 90vh;
      display: flex;
      flex-direction: column;
      box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1);
      overflow: hidden;
    }

    .modal-header {
      padding: 14px 20px;
      border-bottom: 1px solid var(--color-border);
      background: #f8fafc;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .modal-body {
      padding: 20px;
      overflow-y: auto;
      font-size: 12px;
    }

    .audit-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }

    .audit-card {
      background: #f8fafc;
      border: 1px solid var(--color-border);
      border-radius: 6px;
      padding: 10px;
    }

    .audit-card span.label {
      font-size: 10px;
      font-weight: 700;
      text-transform: uppercase;
      color: #64748b;
      display: block;
    }

    .audit-card span.val {
      font-size: 13px;
      font-weight: 700;
      color: #0f172a;
      margin-top: 2px;
      display: block;
    }

    pre.code-block {
      background: #0f172a;
      color: #f8fafc;
      padding: 12px;
      border-radius: 6px;
      font-family: var(--font-mono);
      font-size: 11px;
      overflow-x: auto;
      max-height: 220px;
    }

    footer {
      background: #ffffff;
      border-top: 1px solid var(--color-border);
      padding: 16px 20px;
      text-align: center;
      font-size: 12px;
      color: #64748b;
    }

    @media print {
      header, .banner, .grid-2, .drawer, .summary-grid, .table-toolbar, .table-footer, footer {
        display: none !important;
      }
      .table-section { border: none !important; box-shadow: none !important; }
      .table-wrapper { max-height: none !important; overflow: visible !important; }
      body { background: #fff !important; }
    }
  </style>
</head>
<body>

  <!-- Header -->
  <header>
    <div class="header-inner">
      <div class="logo-group">
        <div class="logo-icon">&#128202;</div>
        <div class="logo-text">
          <h1>
            Bridge Work Orders - All
            <span class="badge-standalone">Standalone Offline &amp; Online App</span>
          </h1>
        </div>
      </div>

      <div class="header-actions">
        <div id="connection-status" class="status-badge">
          <span class="status-dot"></span>
          <span id="connection-text">Online &amp; Offline Ready</span>
        </div>
        <button type="button" class="btn btn-sm" onclick="loadSampleData()">
          &#10024; Load Sample Data
        </button>
        <button type="button" class="btn btn-sm" onclick="resetAll()">
          &#8634; Reset
        </button>
      </div>
    </div>
  </header>

  <main>
    <!-- Top Hero Banner -->
    <div class="banner">
      <div>
        <div class="banner-title">
          <span>Bridge Work Orders &amp; JMS/SaaS Reconciler</span>
          <span class="badge-standalone">15 Standard Headings</span>
        </div>
        <p class="banner-desc">
          Single-file portable application that runs anywhere completely offline. Matches Work Orders, strips &ldquo;| Bridge XBRL&rdquo;,
          transforms Milestones, formats Delivered Dates to MM/DD/YYYY, pulls X Track ID (including Production Day to Day cross-lookup),
          <strong>maps &ldquo;XBRL Feedback Quality&rdquo; to &ldquo;JMS Cycle No / FB Quality&rdquo;</strong>, and
          <strong>conditionally fills &ldquo;Sampling/Quality&rdquo; with &ldquo;STP Feedback&rdquo;</strong>.
        </p>
      </div>

      <div style="display: flex; gap: 8px; flex-wrap: wrap;">
        <button type="button" class="btn btn-indigo" onclick="exportExcel()">
          &#128196; Export Excel (.xlsx)
        </button>
        <button type="button" class="btn" onclick="copyTSV()">
          &#128203; Copy TSV
        </button>
      </div>
    </div>

    <!-- Dropzones: Bridge & JMS files -->
    <div class="grid-2">
      <!-- Bridge Excel Upload -->
      <div class="card">
        <div class="card-header">
          <div>
            <h2 class="card-title">1. Bridge Work Orders Base Excel</h2>
            <p class="card-subtitle">Contains Work Order, Title, Due Date, Activity, Job Type, &amp; XBRL Feedback columns</p>
          </div>
          <span id="bridge-status-badge" class="badge-standalone" style="display: none;">Loaded</span>
        </div>

        <div id="bridge-dropzone" class="dropzone" onclick="document.getElementById('bridge-file-input').click()">
          <input type="file" id="bridge-file-input" accept=".xlsx,.xls,.csv" style="display: none;" onchange="handleBridgeFile(this.files[0])">
          <div class="dropzone-icon">&#128196;</div>
          <p style="font-weight: 600; color: #1e293b;">Drop Bridge Excel here, or <span style="text-decoration: underline; color: #2563eb;">browse</span></p>
          <p style="font-size: 11px; color: #64748b; margin-top: 4px;">Supports .xlsx, .xls, .csv files</p>
        </div>

        <div id="bridge-file-bar" class="loaded-file-bar" style="display: none;">
          <div class="file-info-row">
            <div>
              <strong id="bridge-file-name" style="font-size: 12px; color: #0f172a;"></strong>
              <div id="bridge-file-meta" style="font-size: 11px; color: #64748b;"></div>
            </div>
            <button type="button" class="btn btn-sm" onclick="document.getElementById('bridge-file-input').click()">Change</button>
          </div>
          <div id="bridge-sheet-select-row" class="sheet-select-row" style="display: none;">
            <label style="font-size: 11px; font-weight: 600; color: #475569;">Active Sheet:</label>
            <select id="bridge-sheet-select" onchange="onBridgeSheetChange(this.value)"></select>
          </div>
        </div>

        <div class="expected-cols">
          <strong>Recognized screenshot columns:</strong> Work Order, Work Order Title, Due Date, Activity, Job Type, XBRL Feedback Comments, XBRL Feedback Quality, XBRL Feedback Status, XBRL Feedback Reviewed, XBRL Feedback Validity
        </div>
      </div>

      <!-- JMS / SaaS Workbook Upload -->
      <div class="card">
        <div class="card-header">
          <div>
            <h2 class="card-title">2. JMS / SaaS Active Workbook</h2>
            <p class="card-subtitle">Contains Active Tracker, X Track ID, Grabber Day to Day, &amp; Production Day to Day sheet</p>
          </div>
          <span id="jms-status-badge" class="badge-standalone" style="display: none;">Loaded</span>
        </div>

        <div id="jms-dropzone" class="dropzone" onclick="document.getElementById('jms-file-input').click()">
          <input type="file" id="jms-file-input" accept=".xlsx,.xls,.csv" style="display: none;" onchange="handleJmsFile(this.files[0])">
          <div class="dropzone-icon">&#128202;</div>
          <p style="font-weight: 600; color: #1e293b;">Drop JMS SaaS Workbook here, or <span style="text-decoration: underline; color: #2563eb;">browse</span></p>
          <p style="font-size: 11px; color: #64748b; margin-top: 4px;">Supports .xlsx, .xls, .csv workbooks</p>
        </div>

        <div id="jms-file-bar" class="loaded-file-bar" style="display: none;">
          <div class="file-info-row">
            <div>
              <strong id="jms-file-name" style="font-size: 12px; color: #0f172a;"></strong>
              <div id="jms-file-meta" style="font-size: 11px; color: #64748b;"></div>
            </div>
            <button type="button" class="btn btn-sm" onclick="document.getElementById('jms-file-input').click()">Change</button>
          </div>
          <div id="jms-sheet-select-row" class="sheet-select-row" style="display: none;">
            <label style="font-size: 11px; font-weight: 600; color: #475569;">Active Sheet:</label>
            <select id="jms-sheet-select" onchange="onJmsSheetChange(this.value)"></select>
          </div>
        </div>

        <div class="expected-cols">
          <strong>Cross-lookup sheets:</strong> JMS Active Jobs, Production Day to Day (for missing X Track IDs)
        </div>
      </div>
    </div>

    <!-- Column Mapping & Rule Configuration Drawer -->
    <div class="drawer">
      <button type="button" class="drawer-trigger" onclick="toggleDrawer('mapping-content', 'mapping-caret')">
        <div style="display: flex; align-items: center; gap: 8px;">
          <span style="font-weight: 700; font-size: 13px; color: #0f172a;">&#9881; Column Mapping &amp; Transformation Rules</span>
          <span class="badge-standalone">Auto-Configured</span>
        </div>
        <div style="display: flex; align-items: center; gap: 6px; font-size: 12px; color: #64748b;">
          <span>Configure Mappings</span>
          <span id="mapping-caret">&#9660;</span>
        </div>
      </button>

      <div id="mapping-content" class="drawer-content hidden">
        <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 10px 14px; margin-bottom: 14px; font-size: 11px; color: #166534; display: flex; justify-content: space-between; align-items: center;">
          <span><strong>Automatic Detection Active:</strong> Preferred column names from user specification are matched automatically upon upload and retained in browser local storage.</span>
          <button type="button" class="btn btn-sm" onclick="resetToAutoDetect()">Auto-Detect Preferred</button>
        </div>

        <div class="mapping-grid">
          <!-- Bridge Core -->
          <div class="mapping-card">
            <h4>1. Bridge Core Columns</h4>
            <div class="field-group">
              <label>Work Order (Match Key) *</label>
              <select id="map-bridge-wo" onchange="updateConfig('bridgeWorkOrderCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>Work Order Title</label>
              <select id="map-bridge-title" onchange="updateConfig('bridgeTitleCol', this.value)"></select>
              <div class="field-hint">Removes &ldquo;| Bridge XBRL&rdquo;</div>
            </div>
            <div class="field-group">
              <label>Due Date</label>
              <select id="map-bridge-date" onchange="updateConfig('bridgeDueDateCol', this.value)"></select>
              <div class="field-hint">Formatted as MM/DD/YYYY + Month</div>
            </div>
            <div class="field-group">
              <label>Activity (Milestone)</label>
              <select id="map-bridge-activity" onchange="updateConfig('bridgeActivityCol', this.value)"></select>
              <div class="field-hint">Applies &ldquo;XBRL Update - ...&rdquo; rules</div>
            </div>
            <div class="field-group">
              <label>Job Type</label>
              <select id="map-bridge-jobtype" onchange="updateConfig('bridgeJobTypeCol', this.value)"></select>
            </div>
          </div>

          <!-- XBRL Feedback Columns -->
          <div class="mapping-card mapping-card-accent">
            <h4>
              <span>2. XBRL Feedback Columns</span>
              <span class="badge-standalone" style="background: #ccfbf1; color: #0f766e; border-color: #99f6e4;">Mapped</span>
            </h4>
            <div class="field-group" style="background: #ffffff; padding: 8px; border-radius: 6px; border: 1px solid #99f6e4;">
              <label style="color: #0f766e; font-weight: 700;">
                <span>XBRL Feedback Quality *</span>
                <span style="font-size: 9px; background: #0f766e; color: #fff; padding: 1px 4px; border-radius: 3px;">Maps to FB Quality</span>
              </label>
              <select id="map-bridge-fbquality" onchange="updateConfig('bridgeXbrlFeedbackQualityCol', this.value)"></select>
              <div class="field-hint" style="color: #0f766e; font-weight: 600;">&#10142; Mapped to output: &ldquo;JMS Cycle No / FB Quality&rdquo;</div>
            </div>
            <div class="field-group">
              <label>XBRL Feedback Comments</label>
              <select id="map-bridge-fbcomments" onchange="updateConfig('bridgeXbrlFeedbackCommentsCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>XBRL Feedback Status</label>
              <select id="map-bridge-fbstatus" onchange="updateConfig('bridgeXbrlFeedbackStatusCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>XBRL Feedback Reviewed</label>
              <select id="map-bridge-fbreviewed" onchange="updateConfig('bridgeXbrlFeedbackReviewedCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>XBRL Feedback Validity</label>
              <select id="map-bridge-fbvalidity" onchange="updateConfig('bridgeXbrlFeedbackValidityCol', this.value)"></select>
            </div>
            <div class="field-group" style="padding-top: 8px; border-top: 1px dashed #99f6e4;">
              <label style="color: #3730a3;">Sampling/Quality Prefill Value</label>
              <input type="text" id="map-sampling-prefill" value="STP Feedback" onchange="updateConfig('samplingQualityPrefill', this.value)">
              <div class="field-hint" style="color: #3730a3;">Fills whenever &ldquo;JMS Cycle No / FB Quality&rdquo; is present</div>
            </div>
          </div>

          <!-- JMS Tracker Columns -->
          <div class="mapping-card">
            <h4>3. JMS Main Sheet</h4>
            <div class="field-group">
              <label>Work Order / Job No *</label>
              <select id="map-jms-wo" onchange="updateConfig('jmsWorkOrderCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>X Track ID (Primary)</label>
              <select id="map-jms-track" onchange="updateConfig('jmsTrackIdCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>Grabber Day to Day (Priority 1)</label>
              <select id="map-jms-grabber" onchange="updateConfig('jmsGrabberCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>Job Type (Priority 2 Fallback)</label>
              <select id="map-jms-jobtype" onchange="updateConfig('jmsJobTypeCol', this.value)"></select>
            </div>
          </div>

          <!-- Production Day to Day Lookup -->
          <div class="mapping-card">
            <h4>4. Production Day to Day (Lookup)</h4>
            <div class="field-group">
              <label>Production Sheet Name</label>
              <select id="map-prod-sheet" onchange="onProdSheetChange(this.value)"></select>
              <div class="field-hint">Checks this sheet for X Track ID</div>
            </div>
            <div class="field-group">
              <label>Work Order Column</label>
              <select id="map-prod-wo" onchange="updateConfig('jmsProductionWorkOrderCol', this.value)"></select>
            </div>
            <div class="field-group">
              <label>X Track ID Column</label>
              <select id="map-prod-track" onchange="updateConfig('jmsProductionTrackIdCol', this.value)"></select>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Summary KPI Cards -->
    <div class="summary-grid">
      <div class="summary-card active" id="card-filter-matched" onclick="setFilter('matched')">
        <div class="summary-label">Matched Records</div>
        <div class="summary-value" id="kpi-matched">0</div>
        <div class="summary-subtext" id="kpi-rate">0% match rate</div>
      </div>
      <div class="summary-card" id="card-filter-unmatched-bridge" onclick="setFilter('unmatched_bridge')">
        <div class="summary-label">Bridge Unmatched</div>
        <div class="summary-value" id="kpi-unmatched-bridge">0</div>
        <div class="summary-subtext">Missing in JMS</div>
      </div>
      <div class="summary-card" id="card-filter-unmatched-jms" onclick="setFilter('unmatched_jms')">
        <div class="summary-label">JMS Unmatched</div>
        <div class="summary-value" id="kpi-unmatched-jms">0</div>
        <div class="summary-subtext">Missing in Bridge</div>
      </div>
      <div class="summary-card" id="card-filter-all" onclick="setFilter('all')">
        <div class="summary-label">Total Records</div>
        <div class="summary-value" id="kpi-total">0</div>
        <div class="summary-subtext" id="kpi-prod-cross">0 cross-lookup IDs</div>
      </div>
    </div>

    <!-- Results Table Section -->
    <div class="table-section">
      <div class="table-toolbar">
        <div class="tab-pills">
          <button type="button" class="tab-btn active" id="tab-matched" onclick="setFilter('matched')">Matched Records (<span id="count-tab-matched">0</span>)</button>
          <button type="button" class="tab-btn" id="tab-unmatched-bridge" onclick="setFilter('unmatched_bridge')">Bridge Unmatched (<span id="count-tab-bridge">0</span>)</button>
          <button type="button" class="tab-btn" id="tab-unmatched-jms" onclick="setFilter('unmatched_jms')">JMS Unmatched (<span id="count-tab-jms">0</span>)</button>
          <button type="button" class="tab-btn" id="tab-all" onclick="setFilter('all')">All (<span id="count-tab-all">0</span>)</button>
        </div>

        <div class="table-controls">
          <input type="text" id="table-search" class="search-box" placeholder="Search work order, X Track ID, title..." oninput="onSearchChange(this.value)">
          <button type="button" class="btn btn-sm" onclick="copyTSV()">&#128203; Copy TSV</button>
          <button type="button" class="btn btn-sm" onclick="exportCSV()">&#128190; CSV</button>
          <button type="button" class="btn btn-sm btn-emerald" onclick="exportExcel()">&#128196; Download Excel</button>
        </div>
      </div>

      <div class="table-wrapper">
        <table id="results-table">
          <thead>
            <tr>
              <th style="width: 44px; text-align: center;">#</th>
              <th>Team</th>
              <th class="th-wo">JMS Job Number/Work Order No</th>
              <th class="th-fb">JMS Cycle No / FB Quality</th>
              <th>XOR Ver</th>
              <th class="th-trk">X Track ID</th>
              <th class="th-job">Job Name</th>
              <th class="th-form">Form Type</th>
              <th>Filing Period</th>
              <th class="th-milestone">Milestone</th>
              <th>JMS Team</th>
              <th class="th-sampling">Sampling/Quality</th>
              <th>Consultant Name</th>
              <th>2nd Reviewer Name</th>
              <th class="th-date">Delivered Date</th>
              <th class="th-date">Month</th>
              <th style="width: 60px; text-align: center;">Audit</th>
            </tr>
          </thead>
          <tbody id="table-body">
            <tr>
              <td colspan="17" style="text-align: center; padding: 48px; color: #94a3b8;">
                No files loaded yet. Click <strong>Load Sample Data</strong> above or drop your Bridge &amp; JMS Excel files.
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="table-footer">
        <div>
          Showing <span id="pag-start">0</span> to <span id="pag-end">0</span> of <span id="pag-total">0</span> records
          <span style="margin-left: 14px;">Per page:</span>
          <select id="pag-size" style="width: auto; padding: 2px 6px; font-size: 11px;" onchange="onPageSizeChange(Number(this.value))">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>

        <div class="pagination">
          <button type="button" class="btn btn-sm" id="btn-prev-page" onclick="changePage(-1)">&#9664; Prev</button>
          <span id="page-indicator" style="font-family: var(--font-mono); font-size: 11px; padding: 0 6px;">1 / 1</span>
          <button type="button" class="btn btn-sm" id="btn-next-page" onclick="changePage(1)">Next &#9654;</button>
        </div>
      </div>
    </div>
  </main>

  <!-- Row Audit Modal -->
  <div id="audit-modal" class="modal-backdrop hidden" onclick="if(event.target === this) closeModal()">
    <div class="modal-box">
      <div class="modal-header">
        <div>
          <strong id="modal-wo-title" style="font-size: 13px; color: #0f172a;"></strong>
          <div id="modal-wo-sub" style="font-size: 11px; color: #64748b;"></div>
        </div>
        <button type="button" class="btn btn-sm" onclick="closeModal()">&#10005;</button>
      </div>
      <div class="modal-body" id="modal-body-content"></div>
    </div>
  </div>

  <footer>
    Bridge Work Orders Reconciler &amp; QMT Standard Output &middot; Single-file Standalone Application (Works 100% Offline &amp; Online)
  </footer>

  <!-- Embed SheetJS library -->
  ` + (xlsxScript ? '<script>' + xlsxScript + '</script>' : '<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>') + `

  <!-- Standalone App Logic -->
  <script>
    // Live Online/Offline listener
    function updateOnlineStatus() {
      const isOnline = navigator.onLine;
      const statusText = document.getElementById('connection-text');
      const dot = document.querySelector('.status-dot');
      if (isOnline) {
        statusText.textContent = 'Online & Offline Ready';
        if (dot) dot.style.background = '#10b981';
      } else {
        statusText.textContent = 'Offline Mode Active';
        if (dot) dot.style.background = '#f59e0b';
      }
    }
    window.addEventListener('online', updateOnlineStatus);
    window.addEventListener('offline', updateOnlineStatus);
    updateOnlineStatus();

    // Default configuration
    const DEFAULT_CONFIG = {
      bridgeWorkOrderCol: 'Work Order',
      bridgeTitleCol: 'Work Order Title',
      bridgeActivityCol: 'Activity',
      bridgeDueDateCol: 'Due Date',
      bridgeJobTypeCol: 'Job Type',
      bridgeGrabberCol: 'Grabber Day to Day',
      bridgeTeamStaticValue: 'Bridge',
      bridgeTrackIdCol: 'X Track ID',

      bridgeXbrlFeedbackCommentsCol: 'XBRL Feedback Comments',
      bridgeXbrlFeedbackQualityCol: 'XBRL Feedback Quality',
      bridgeXbrlFeedbackStatusCol: 'XBRL Feedback Status',
      bridgeXbrlFeedbackReviewedCol: 'XBRL Feedback Reviewed',
      bridgeXbrlFeedbackValidityCol: 'XBRL Feedback Validity',

      samplingQualityPrefill: 'STP Feedback',

      jmsWorkOrderCol: 'Work Order No',
      jmsTrackIdCol: 'X Track ID',
      jmsGrabberCol: 'Grabber Day to Day',
      jmsJobTypeCol: 'Job Type',
      jmsFormTypeCol: 'Form Type',

      jmsProductionSheet: 'Production Day to Day',
      jmsProductionWorkOrderCol: 'Work Order No',
      jmsProductionTrackIdCol: 'X Track ID',
    };

    let config = Object.assign({}, DEFAULT_CONFIG);

    // Saved localStorage config
    try {
      const saved = localStorage.getItem('bridge_reconcile_standalone_cfg_v2');
      if (saved) config = Object.assign({}, config, JSON.parse(saved));
    } catch (e) {}

    function saveConfig() {
      try {
        localStorage.setItem('bridge_reconcile_standalone_cfg_v2', JSON.stringify(config));
      } catch (e) {}
    }

    // Application State
    let bridgeFile = { fileName: '', sheetNames: [], selectedSheet: '', sheets: {}, headers: [], data: [] };
    let jmsFile = { fileName: '', sheetNames: [], selectedSheet: '', sheets: {}, headers: [], data: [] };
    let reconciledData = { matchedRows: [], unmatchedBridgeRows: [], unmatchedJmsRows: [], allRows: [] };
    let activeFilter = 'matched';
    let searchQuery = '';
    let currentPage = 1;
    let pageSize = 25;

    // Authentic Demo Dataset
    const DEMO_BRIDGE_DATA = [
      { 'Work Order': 'WO-98421', 'Work Order Number': 'WO-98421', 'Work Order Title': 'Alphabet Inc. Q3 Financial Statement | Bridge XBRL', 'Due Date': '2026-10-15', 'Activity': 'XBRL Update - Final to Client', 'Job Type': '10-Q', 'XBRL Feedback Comments': 'Footnote tagging updated as per client review notes', 'XBRL Feedback Quality': 'Pass', 'XBRL Feedback Status': 'Completed', 'XBRL Feedback Reviewed': 'Yes', 'XBRL Feedback Validity': 'Valid', 'Grabber Day to Day': '10-Q', 'X Track ID': 'TRK-GL-5541' },
      { 'Work Order': 'WO-98422', 'Work Order Number': 'WO-98422', 'Work Order Title': 'Microsoft Corp Annual Report 10-K | Bridge XBRL', 'Due Date': '2026-10-22', 'Activity': 'XBRL Update - Non Standard', 'Job Type': '10-K', 'XBRL Feedback Comments': 'Scale precision reviewed and adjusted for segment data', 'XBRL Feedback Quality': 'Cycle 2', 'XBRL Feedback Status': 'In Review', 'XBRL Feedback Reviewed': 'Yes', 'XBRL Feedback Validity': 'Valid', 'Grabber Day to Day': '', 'X Track ID': 'TRK-MS-7712' },
      { 'Work Order': 'WO-98423', 'Work Order Number': 'WO-98423', 'Work Order Title': 'Apple Inc. Current Report 8-K | Bridge XBRL', 'Due Date': '2026-11-05', 'Activity': 'XBRL Update - Standard', 'Job Type': '8-K', 'XBRL Feedback Comments': 'Member context alignment completed without error', 'XBRL Feedback Quality': 'Cycle 1', 'XBRL Feedback Status': 'Completed', 'XBRL Feedback Reviewed': 'Yes', 'XBRL Feedback Validity': 'Valid', 'Grabber Day to Day': '8-K', 'X Track ID': 'TRK-AP-3320' },
      { 'Work Order': 'WO-98424', 'Work Order Number': 'WO-98424', 'Work Order Title': 'NVIDIA Corp Registration Statement S-3 | Bridge XBRL', 'Due Date': '2026-10-28', 'Activity': 'XBRL Update - Final to Client', 'Job Type': 'S-3', 'XBRL Feedback Comments': 'Cover page taxonomy references checked and verified', 'XBRL Feedback Quality': 'Pass', 'XBRL Feedback Status': 'Completed', 'XBRL Feedback Reviewed': 'Yes', 'XBRL Feedback Validity': 'Valid', 'Grabber Day to Day': 'S-3', 'X Track ID': 'TRK-NV-9014' },
      { 'Work Order': 'WO-98425', 'Work Order Number': 'WO-98425', 'Work Order Title': 'Amazon.com Inc Q3 10-Q | Bridge XBRL', 'Due Date': '2026-11-12', 'Activity': 'XBRL Update - Standard', 'Job Type': '10-Q', 'XBRL Feedback Comments': 'Initial filing - no quality score assigned', 'XBRL Feedback Quality': '', 'XBRL Feedback Status': 'Completed', 'XBRL Feedback Reviewed': 'Yes', 'XBRL Feedback Validity': 'Valid', 'Grabber Day to Day': '', 'X Track ID': 'TRK-PROD-AMZ-9921' },
      { 'Work Order': 'WO-98426', 'Work Order Number': 'WO-98426', 'Work Order Title': 'Tesla Inc Interim Filing 10-Q | Bridge XBRL', 'Due Date': '2026-10-30', 'Activity': 'XBRL Update - Non Standard', 'Job Type': '10-Q', 'XBRL Feedback Comments': 'Balance sheet cash and cash equivalents line confirmed', 'XBRL Feedback Quality': 'Cycle 2', 'XBRL Feedback Status': 'Approved', 'XBRL Feedback Reviewed': 'Yes', 'XBRL Feedback Validity': 'Valid', 'Grabber Day to Day': '10-Q', 'X Track ID': 'TRK-PROD-TSL-8834' },
      { 'Work Order': 'WO-98427', 'Work Order Number': 'WO-98427', 'Work Order Title': 'Meta Platforms Form 8-K Filing | Bridge XBRL', 'Due Date': '2026-11-18', 'Activity': 'XBRL Update - Final to Client', 'Job Type': '8-K', 'XBRL Feedback Comments': 'Item 5.02 executive officer departure notes tagged', 'XBRL Feedback Quality': 'Pass', 'XBRL Feedback Status': 'Completed', 'XBRL Feedback Reviewed': 'Yes', 'XBRL Feedback Validity': 'Valid', 'Grabber Day to Day': '8-K', 'X Track ID': 'TRK-FB-4109' },
      { 'Work Order': 'WO-98499', 'Work Order Number': 'WO-98499', 'Work Order Title': 'Berkshire Hathaway Q3 Filing | Bridge XBRL', 'Due Date': '2026-10-19', 'Activity': 'XBRL Update - Standard', 'Job Type': '10-Q', 'XBRL Feedback Comments': 'Pending reviewer sign-off on investment schedule', 'XBRL Feedback Quality': 'Review Required', 'XBRL Feedback Status': 'Pending', 'XBRL Feedback Reviewed': 'No', 'XBRL Feedback Validity': 'Under Review', 'Grabber Day to Day': '10-Q', 'X Track ID': 'TRK-BRK-9900' }
    ];

    const DEMO_JMS_DATA = [
      { 'Work Order No': 'WO-98421', 'X Track ID': 'TRK-GL-5541', 'Grabber Day to Day': '10-Q', 'Job Type': '10-Q' },
      { 'Work Order No': 'WO-98422', 'X Track ID': 'TRK-MS-7712', 'Grabber Day to Day': '', 'Job Type': '10-K' },
      { 'Work Order No': 'WO-98423', 'X Track ID': 'TRK-AP-3320', 'Grabber Day to Day': '8-K', 'Job Type': '8-K' },
      { 'Work Order No': 'WO-98424', 'X Track ID': 'TRK-NV-9014', 'Grabber Day to Day': 'S-3', 'Job Type': 'S-3' },
      { 'Work Order No': 'WO-98425', 'X Track ID': '', 'Grabber Day to Day': '', 'Job Type': '10-Q' },
      { 'Work Order No': 'WO-98426', 'X Track ID': '', 'Grabber Day to Day': '10-Q', 'Job Type': '10-Q' },
      { 'Work Order No': 'WO-98427', 'X Track ID': 'TRK-FB-4109', 'Grabber Day to Day': '8-K', 'Job Type': '8-K' },
      { 'Work Order No': 'WO-99100', 'X Track ID': 'TRK-IBM-1102', 'Grabber Day to Day': '10-K', 'Job Type': '10-K' }
    ];

    const DEMO_PROD_DATA = [
      { 'Work Order No': 'WO-98425', 'X Track ID': 'TRK-PROD-AMZ-9921', 'Shift': 'Day Shift' },
      { 'Work Order No': 'WO-98426', 'X Track ID': 'TRK-PROD-TSL-8834', 'Shift': 'Night Shift' },
      { 'Work Order No': 'WO-98421', 'X Track ID': 'TRK-GL-5541', 'Shift': 'Day Shift' }
    ];

    // Core String / Normalization Helpers
    function normalizeKey(val) {
      if (!val) return '';
      return String(val).replace(/[\\u200B-\\u200D\\uFEFF]/g, '').trim().toUpperCase();
    }

    function cleanJobName(raw) {
      if (!raw) return '';
      let title = String(raw).trim();
      title = title.replace(/\\s*\\|\\s*Bridge\\s*XBRL(\\b|$)/gi, '');
      title = title.replace(/\\s*-\\s*Bridge\\s*XBRL(\\b|$)/gi, '');
      title = title.replace(/\\s*Bridge\\s*XBRL$/gi, '');
      return title.replace(/[\\s|\\-]+$/, '').trim();
    }

    function transformMilestone(raw) {
      if (!raw) return '';
      const act = String(raw).trim();
      if (/^XBRL\\s*Update\\s*-\\s*Final\\s*to\\s*Client$/i.test(act)) return 'Final to Client';
      if (/^XBRL\\s*Update\\s*-\\s*Non\\s*Standard$/i.test(act)) return 'Non Standard';
      if (/^XBRL\\s*Update\\s*-\\s*Standard$/i.test(act)) return 'Standard';
      const m = act.match(/^XBRL\\s*Update\\s*[-–—:]\\s*(.+)$/i);
      if (m && m[1]) return m[1].trim();
      return act;
    }

    function formatDate(raw) {
      if (!raw) return { formatted: '', month: '' };
      const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      let d = null;
      if (raw instanceof Date && !isNaN(raw.getTime())) d = raw;
      else if (typeof raw === 'number' && raw > 25569) {
        d = new Date(Math.round((raw - 25569) * 86400 * 1000));
      } else {
        const s = String(raw).trim();
        const parsed = Date.parse(s);
        if (!isNaN(parsed)) d = new Date(parsed);
      }
      if (!d || isNaN(d.getTime())) return { formatted: String(raw).trim(), month: '' };

      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(d.getUTCDate()).padStart(2, '0');
      const yyyy = d.getUTCFullYear();
      return {
        formatted: mm + '/' + dd + '/' + yyyy,
        month: MONTHS[d.getUTCMonth()] || ''
      };
    }

    function resolveFormType(grabber, jobType, fallback) {
      if (grabber && String(grabber).trim()) return String(grabber).trim();
      if (jobType && String(jobType).trim()) return String(jobType).trim();
      if (fallback && String(fallback).trim()) return String(fallback).trim();
      return '';
    }

    function findMatch(headers, targets, excludes) {
      if (!headers || !headers.length) return '';
      const norm = headers.map(h => ({ orig: h, clean: h.toLowerCase().replace(/[^a-z0-9]/g, ' ').replace(/\\s+/g, ' ').trim() }));
      for (const t of targets) {
        const cleanT = t.toLowerCase().replace(/[^a-z0-9]/g, ' ').replace(/\\s+/g, ' ').trim();
        const exact = norm.find(item => item.clean === cleanT);
        if (exact) return exact.orig;
      }
      for (const t of targets) {
        const cleanT = t.toLowerCase().replace(/[^a-z0-9]/g, ' ').replace(/\\s+/g, ' ').trim();
        for (const item of norm) {
          if (excludes && excludes.some(e => item.clean.includes(e.toLowerCase()))) continue;
          if (item.clean.includes(cleanT) || cleanT.includes(item.clean) || item.orig.toLowerCase().includes(t.toLowerCase())) {
            return item.orig;
          }
        }
      }
      return '';
    }

    function autoDetect() {
      const bh = bridgeFile.headers;
      const jh = jmsFile.headers;

      config.bridgeWorkOrderCol = findMatch(bh, ['Work Order', 'Work Order Number', 'Work Order No', 'WO#'], ['title','date','status','type']) || bh[0] || '';
      config.bridgeTitleCol = findMatch(bh, ['Work Order Title', 'Job Name', 'Title', 'Client Name'], ['no','number','id','date']) || '';
      config.bridgeDueDateCol = findMatch(bh, ['Due Date', 'Delivered Date', 'Target Date'], ['created','update']) || '';
      config.bridgeActivityCol = findMatch(bh, ['Activity', 'Milestone', 'Task'], ['date','number','id']) || '';
      config.bridgeJobTypeCol = findMatch(bh, ['Job Type', 'Form Type', 'Type'], ['number','title']) || '';
      config.bridgeGrabberCol = findMatch(bh, ['Grabber Day to Day', 'Grabber', 'Form Type']) || '';

      config.bridgeXbrlFeedbackQualityCol = findMatch(bh, ['XBRL Feedback Quality', 'Feedback Quality', 'XBRL Quality', 'FB Quality', 'Cycle No / FB Quality'], ['status','comment','review','valid']) || '';
      config.bridgeXbrlFeedbackCommentsCol = findMatch(bh, ['XBRL Feedback Comments', 'Feedback Comments', 'XBRL Comments', 'Comments'], ['quality','status','review','valid']) || '';
      config.bridgeXbrlFeedbackStatusCol = findMatch(bh, ['XBRL Feedback Status', 'Feedback Status', 'XBRL Status'], ['quality','comment','review']) || '';
      config.bridgeXbrlFeedbackReviewedCol = findMatch(bh, ['XBRL Feedback Reviewed', 'Feedback Reviewed', 'XBRL Reviewed', 'Reviewed'], ['quality','comment','status']) || '';
      config.bridgeXbrlFeedbackValidityCol = findMatch(bh, ['XBRL Feedback Validity', 'Feedback Validity', 'XBRL Validity', 'Validity'], ['quality','comment','status']) || '';

      config.jmsWorkOrderCol = findMatch(jh, ['Work Order No', 'Work Order Number', 'Work Order', 'Job Number'], ['title','date','status']) || jh[0] || '';
      config.jmsTrackIdCol = findMatch(jh, ['X Track ID', 'X-Track ID', 'X Track', 'Track ID'], ['work','order','title']) || '';
      config.jmsGrabberCol = findMatch(jh, ['Grabber Day to Day', 'Grabber', 'Form Type']) || '';
      config.jmsJobTypeCol = findMatch(jh, ['Job Type', 'Type']) || '';

      const jSheets = Object.keys(jmsFile.sheets);
      const prodSheet = jSheets.find(s => /production.*day.*to.*day/i.test(s)) || jSheets.find(s => /production/i.test(s)) || '';
      if (prodSheet && jmsFile.sheets[prodSheet]) {
        config.jmsProductionSheet = prodSheet;
        const ph = jmsFile.sheets[prodSheet].headers;
        config.jmsProductionWorkOrderCol = findMatch(ph, ['Work Order No', 'Work Order Number', 'Work Order', 'WO#'], ['title','track']) || ph[0] || '';
        config.jmsProductionTrackIdCol = findMatch(ph, ['X Track ID', 'X-Track ID', 'X Track'], ['order','work']) || '';
      }

      saveConfig();
      renderMappingControls();
    }

    function runReconciliation() {
      const bData = bridgeFile.data;
      const jData = jmsFile.data;

      if (!bData.length && !jData.length) {
        reconciledData = { matchedRows: [], unmatchedBridgeRows: [], unmatchedJmsRows: [], allRows: [] };
        renderKPIs();
        renderTable();
        return;
      }

      const jmsMap = new Map();
      const matchedJmsKeys = new Set();
      jData.forEach((row, idx) => {
        const key = normalizeKey(row[config.jmsWorkOrderCol]);
        if (key && !jmsMap.has(key)) jmsMap.set(key, { row, idx });
      });

      const prodMap = new Map();
      if (config.jmsProductionSheet && jmsFile.sheets[config.jmsProductionSheet]) {
        const prodData = jmsFile.sheets[config.jmsProductionSheet].data || [];
        prodData.forEach(row => {
          const key = normalizeKey(row[config.jmsProductionWorkOrderCol]);
          if (key && !prodMap.has(key)) prodMap.set(key, row);
        });
      }

      const matchedRows = [];
      const unmatchedBridgeRows = [];

      bData.forEach((bRow, idx) => {
        const rawWo = bRow[config.bridgeWorkOrderCol];
        const key = normalizeKey(rawWo);
        if (!key) return;

        const jMatch = jmsMap.get(key);
        const pMatch = prodMap.get(key);

        const rawTitle = bRow[config.bridgeTitleCol];
        const jobName = cleanJobName(rawTitle);

        const rawAct = bRow[config.bridgeActivityCol];
        const milestone = transformMilestone(rawAct);

        const rawDate = bRow[config.bridgeDueDateCol];
        const resDate = formatDate(rawDate);
        const deliveredDate = resDate.formatted;
        const month = resDate.month;

        let grabberVal = jMatch ? (jMatch.row[config.jmsGrabberCol] || bRow[config.bridgeGrabberCol] || '') : (bRow[config.bridgeGrabberCol] || '');
        let jobTypeVal = jMatch ? (jMatch.row[config.jmsJobTypeCol] || bRow[config.bridgeJobTypeCol] || '') : (bRow[config.bridgeJobTypeCol] || '');
        let formTypeVal = jMatch ? (jMatch.row[config.jmsFormTypeCol] || '') : '';
        const formType = resolveFormType(grabberVal, jobTypeVal, formTypeVal);

        let xTrackId = '';
        let xTrackSource = 'none';
        if (jMatch && config.jmsTrackIdCol && jMatch.row[config.jmsTrackIdCol]) {
          xTrackId = String(jMatch.row[config.jmsTrackIdCol]).trim();
          xTrackSource = 'main_sheet';
        }
        if (!xTrackId && pMatch && config.jmsProductionTrackIdCol && pMatch[config.jmsProductionTrackIdCol]) {
          xTrackId = String(pMatch[config.jmsProductionTrackIdCol]).trim();
          xTrackSource = 'production_sheet';
        }

        const fbQuality = (config.bridgeXbrlFeedbackQualityCol && bRow[config.bridgeXbrlFeedbackQualityCol] !== undefined && bRow[config.bridgeXbrlFeedbackQualityCol] !== null)
          ? String(bRow[config.bridgeXbrlFeedbackQualityCol]).trim() : '';
        const fbComments = (config.bridgeXbrlFeedbackCommentsCol && bRow[config.bridgeXbrlFeedbackCommentsCol]) ? String(bRow[config.bridgeXbrlFeedbackCommentsCol]).trim() : '';
        const fbStatus = (config.bridgeXbrlFeedbackStatusCol && bRow[config.bridgeXbrlFeedbackStatusCol]) ? String(bRow[config.bridgeXbrlFeedbackStatusCol]).trim() : '';
        const fbReviewed = (config.bridgeXbrlFeedbackReviewedCol && bRow[config.bridgeXbrlFeedbackReviewedCol]) ? String(bRow[config.bridgeXbrlFeedbackReviewedCol]).trim() : '';
        const fbValidity = (config.bridgeXbrlFeedbackValidityCol && bRow[config.bridgeXbrlFeedbackValidityCol]) ? String(bRow[config.bridgeXbrlFeedbackValidityCol]).trim() : '';

        // Conditional rule: Whenever "JMS Cycle No / FB Quality" is filled, "Sampling/Quality" is filled as "STP Feedback"
        const hasFbQuality = Boolean(fbQuality && fbQuality.length > 0);
        const samplingQuality = hasFbQuality ? (config.samplingQualityPrefill || 'STP Feedback') : '';

        const isMatch = Boolean(jMatch || pMatch);

        const outRow = {
          id: (isMatch ? 'match-' : 'bridge-') + idx + '-' + key,
          team: 'Bridge',
          workOrderNo: String(rawWo).trim(),
          jmsCycleNoFbQuality: fbQuality,
          xorVer: '',
          xTrackId: xTrackId,
          jobName: jobName,
          formType: formType,
          filingPeriod: '',
          milestone: milestone,
          jmsTeam: 'Bridge',
          samplingQuality: samplingQuality,
          consultantName: '',
          secondReviewerName: '',
          deliveredDate: deliveredDate,
          month: month,
          xbrlFeedbackComments: fbComments,
          xbrlFeedbackQuality: fbQuality,
          xbrlFeedbackStatus: fbStatus,
          xbrlFeedbackReviewed: fbReviewed,
          xbrlFeedbackValidity: fbValidity,
          xTrackIdSource: xTrackSource,
          rawBridge: bRow,
          rawJms: jMatch ? jMatch.row : null,
          rawProd: pMatch,
          matchStatus: isMatch ? 'matched' : 'unmatched_bridge'
        };

        if (isMatch) {
          matchedJmsKeys.add(key);
          matchedRows.push(outRow);
        } else {
          unmatchedBridgeRows.push(outRow);
        }
      });

      const unmatchedJmsRows = [];
      jData.forEach((jRow, idx) => {
        const rawWo = jRow[config.jmsWorkOrderCol];
        const key = normalizeKey(rawWo);
        if (key && !matchedJmsKeys.has(key)) {
          let xTrackId = jRow[config.jmsTrackIdCol] ? String(jRow[config.jmsTrackIdCol]).trim() : '';
          let xTrackSource = xTrackId ? 'main_sheet' : 'none';
          const pMatch = prodMap.get(key);
          if (!xTrackId && pMatch && config.jmsProductionTrackIdCol && pMatch[config.jmsProductionTrackIdCol]) {
            xTrackId = String(pMatch[config.jmsProductionTrackIdCol]).trim();
            xTrackSource = 'production_sheet';
          }
          const formType = resolveFormType(jRow[config.jmsGrabberCol], jRow[config.jmsJobTypeCol], jRow[config.jmsFormTypeCol]);

          unmatchedJmsRows.push({
            id: 'jms-' + idx + '-' + key,
            team: 'Bridge',
            workOrderNo: String(rawWo).trim(),
            jmsCycleNoFbQuality: '',
            xorVer: '',
            xTrackId: xTrackId,
            jobName: '',
            formType: formType,
            filingPeriod: '',
            milestone: '',
            jmsTeam: 'Bridge',
            samplingQuality: '',
            consultantName: '',
            secondReviewerName: '',
            deliveredDate: '',
            month: '',
            xTrackIdSource: xTrackSource,
            rawJms: jRow,
            rawProd: pMatch,
            matchStatus: 'unmatched_jms'
          });
        }
      });

      reconciledData = {
        matchedRows: matchedRows,
        unmatchedBridgeRows: unmatchedBridgeRows,
        unmatchedJmsRows: unmatchedJmsRows,
        allRows: [].concat(matchedRows, unmatchedBridgeRows, unmatchedJmsRows)
      };

      currentPage = 1;
      renderKPIs();
      renderTable();
    }

    // UI Renderers
    function renderKPIs() {
      const mCount = reconciledData.matchedRows.length;
      const bCount = reconciledData.unmatchedBridgeRows.length;
      const jCount = reconciledData.unmatchedJmsRows.length;
      const tCount = reconciledData.allRows.length;
      const totalBridge = mCount + bCount;
      const rate = totalBridge > 0 ? Math.round((mCount / totalBridge) * 100) : 0;
      const crossCount = reconciledData.matchedRows.filter(r => r.xTrackIdSource === 'production_sheet').length;

      document.getElementById('kpi-matched').textContent = mCount;
      document.getElementById('kpi-rate').textContent = rate + '% match rate';
      document.getElementById('kpi-unmatched-bridge').textContent = bCount;
      document.getElementById('kpi-unmatched-jms').textContent = jCount;
      document.getElementById('kpi-total').textContent = tCount;
      document.getElementById('kpi-prod-cross').textContent = crossCount + ' cross-lookup IDs';

      document.getElementById('count-tab-matched').textContent = mCount;
      document.getElementById('count-tab-bridge').textContent = bCount;
      document.getElementById('count-tab-jms').textContent = jCount;
      document.getElementById('count-tab-all').textContent = tCount;
    }

    function getFilteredRows() {
      let list = [];
      if (activeFilter === 'matched') list = reconciledData.matchedRows;
      else if (activeFilter === 'unmatched_bridge') list = reconciledData.unmatchedBridgeRows;
      else if (activeFilter === 'unmatched_jms') list = reconciledData.unmatchedJmsRows;
      else list = reconciledData.allRows;

      if (!searchQuery.trim()) return list;
      const q = searchQuery.toLowerCase().trim();
      return list.filter(r => (
        r.workOrderNo.toLowerCase().includes(q) ||
        r.jobName.toLowerCase().includes(q) ||
        r.xTrackId.toLowerCase().includes(q) ||
        r.formType.toLowerCase().includes(q) ||
        r.milestone.toLowerCase().includes(q) ||
        r.jmsCycleNoFbQuality.toLowerCase().includes(q) ||
        r.deliveredDate.toLowerCase().includes(q) ||
        r.month.toLowerCase().includes(q)
      ));
    }

    function renderTable() {
      const rows = getFilteredRows();
      const total = rows.length;
      const totalPages = Math.max(1, Math.ceil(total / pageSize));
      if (currentPage > totalPages) currentPage = totalPages;

      const start = total === 0 ? 0 : (currentPage - 1) * pageSize;
      const end = Math.min(start + pageSize, total);
      const pageRows = rows.slice(start, end);

      document.getElementById('pag-start').textContent = total === 0 ? 0 : start + 1;
      document.getElementById('pag-end').textContent = end;
      document.getElementById('pag-total').textContent = total;
      document.getElementById('page-indicator').textContent = currentPage + ' / ' + totalPages;
      document.getElementById('btn-prev-page').disabled = currentPage <= 1;
      document.getElementById('btn-next-page').disabled = currentPage >= totalPages;

      const tbody = document.getElementById('table-body');
      if (!pageRows.length) {
        tbody.innerHTML = '<tr><td colspan="17" style="text-align: center; padding: 48px; color: #94a3b8;">No matching records found.</td></tr>';
        return;
      }

      let html = '';
      pageRows.forEach((r, idx) => {
        const rowNum = start + idx + 1;
        const fbQualityBadge = r.jmsCycleNoFbQuality
          ? '<span class="tag-badge tag-fb">' + escapeHtml(r.jmsCycleNoFbQuality) + '</span>'
          : '<span style="color: #94a3b8;">—</span>';

        const samplingBadge = r.samplingQuality
          ? '<span class="tag-badge tag-stp">' + escapeHtml(r.samplingQuality) + '</span>'
          : '<span style="color: #94a3b8;">—</span>';

        const formTypeBadge = r.formType
          ? '<span class="tag-badge tag-form">' + escapeHtml(r.formType) + '</span>'
          : '<span style="color: #94a3b8;">—</span>';

        const trackBadge = r.xTrackId
          ? '<span style="font-family: var(--font-mono); font-weight: 700; color: #0284c7;">' + escapeHtml(r.xTrackId) + (r.xTrackIdSource === 'production_sheet' ? ' <span style="font-size: 9px; background: #ecfdf5; color: #065f46; border: 1px solid #a7f3d0; padding: 1px 4px; border-radius: 3px;">Prod Sheet</span>' : '') + '</span>'
          : '<span style="color: #94a3b8; font-style: italic;">No match</span>';

        html += '<tr>' +
          '<td style="text-align: center; color: #94a3b8; font-family: var(--font-mono);">' + rowNum + '</td>' +
          '<td style="font-weight: 600;">' + escapeHtml(r.team) + '</td>' +
          '<td style="font-family: var(--font-mono); font-weight: 700; color: #1e3a8a;">' + escapeHtml(r.workOrderNo) + '</td>' +
          '<td>' + fbQualityBadge + '</td>' +
          '<td style="color: #94a3b8;">—</td>' +
          '<td>' + trackBadge + '</td>' +
          '<td style="font-weight: 600; color: #0f172a;">' + escapeHtml(r.jobName) + '</td>' +
          '<td>' + formTypeBadge + '</td>' +
          '<td style="color: #94a3b8;">—</td>' +
          '<td style="font-weight: 600; color: #6b21a8;">' + escapeHtml(r.milestone) + '</td>' +
          '<td>' + escapeHtml(r.jmsTeam) + '</td>' +
          '<td>' + samplingBadge + '</td>' +
          '<td style="color: #94a3b8;">—</td>' +
          '<td style="color: #94a3b8;">—</td>' +
          '<td style="font-family: var(--font-mono); font-weight: 600; color: #9f1239;">' + escapeHtml(r.deliveredDate) + '</td>' +
          '<td style="font-weight: 600; color: #9f1239;">' + escapeHtml(r.month) + '</td>' +
          '<td style="text-align: center;"><button type="button" class="btn btn-sm" onclick="openAuditModal(' + (start + idx) + ')">&#128065;</button></td>' +
        '</tr>';
      });

      tbody.innerHTML = html;
    }

    function populateSelect(selectId, options, selectedVal) {
      const sel = document.getElementById(selectId);
      if (!sel) return;
      sel.innerHTML = '<option value="">-- Select Column --</option>' +
        options.map(o => '<option value="' + escapeHtml(o) + '" ' + (o === selectedVal ? 'selected' : '') + '>' + escapeHtml(o) + '</option>').join('');
    }

    function renderMappingControls() {
      const bh = bridgeFile.headers;
      const jh = jmsFile.headers;

      populateSelect('map-bridge-wo', bh, config.bridgeWorkOrderCol);
      populateSelect('map-bridge-title', bh, config.bridgeTitleCol);
      populateSelect('map-bridge-date', bh, config.bridgeDueDateCol);
      populateSelect('map-bridge-activity', bh, config.bridgeActivityCol);
      populateSelect('map-bridge-jobtype', bh, config.bridgeJobTypeCol);

      populateSelect('map-bridge-fbquality', bh, config.bridgeXbrlFeedbackQualityCol);
      populateSelect('map-bridge-fbcomments', bh, config.bridgeXbrlFeedbackCommentsCol);
      populateSelect('map-bridge-fbstatus', bh, config.bridgeXbrlFeedbackStatusCol);
      populateSelect('map-bridge-fbreviewed', bh, config.bridgeXbrlFeedbackReviewedCol);
      populateSelect('map-bridge-fbvalidity', bh, config.bridgeXbrlFeedbackValidityCol);

      const prefillInp = document.getElementById('map-sampling-prefill');
      if (prefillInp) prefillInp.value = config.samplingQualityPrefill || 'STP Feedback';

      populateSelect('map-jms-wo', jh, config.jmsWorkOrderCol);
      populateSelect('map-jms-track', jh, config.jmsTrackIdCol);
      populateSelect('map-jms-grabber', jh, config.jmsGrabberCol);
      populateSelect('map-jms-jobtype', jh, config.jmsJobTypeCol);

      const jSheets = Object.keys(jmsFile.sheets);
      const prodSheetSel = document.getElementById('map-prod-sheet');
      if (prodSheetSel) {
        prodSheetSel.innerHTML = '<option value="">-- None (No lookup) --</option>' +
          jSheets.map(s => '<option value="' + escapeHtml(s) + '" ' + (s === config.jmsProductionSheet ? 'selected' : '') + '>' + escapeHtml(s) + '</option>').join('');
      }

      const ph = config.jmsProductionSheet && jmsFile.sheets[config.jmsProductionSheet] ? jmsFile.sheets[config.jmsProductionSheet].headers : [];
      populateSelect('map-prod-wo', ph, config.jmsProductionWorkOrderCol);
      populateSelect('map-prod-track', ph, config.jmsProductionTrackIdCol);
    }

    function updateConfig(key, val) {
      config[key] = val;
      saveConfig();
      runReconciliation();
    }

    function onProdSheetChange(sheetName) {
      config.jmsProductionSheet = sheetName;
      if (sheetName && jmsFile.sheets[sheetName]) {
        const ph = jmsFile.sheets[sheetName].headers;
        config.jmsProductionWorkOrderCol = findMatch(ph, ['Work Order No', 'Work Order Number', 'WO#']) || ph[0] || '';
        config.jmsProductionTrackIdCol = findMatch(ph, ['X Track ID', 'X-Track ID', 'X Track']) || '';
      }
      saveConfig();
      renderMappingControls();
      runReconciliation();
    }

    function setFilter(f) {
      activeFilter = f;
      document.querySelectorAll('.summary-card').forEach(c => c.classList.remove('active'));
      const activeCard = document.getElementById('card-filter-' + f.replace('_', '-'));
      if (activeCard) activeCard.classList.add('active');

      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      const activeTab = document.getElementById(f === 'unmatched_bridge' ? 'tab-unmatched-bridge' : f === 'unmatched_jms' ? 'tab-unmatched-jms' : 'tab-' + f);
      if (activeTab) activeTab.classList.add('active');

      currentPage = 1;
      renderTable();
    }

    function onSearchChange(val) {
      searchQuery = val;
      currentPage = 1;
      renderTable();
    }

    function onPageSizeChange(val) {
      pageSize = val;
      currentPage = 1;
      renderTable();
    }

    function changePage(delta) {
      const rows = getFilteredRows();
      const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
      const next = currentPage + delta;
      if (next >= 1 && next <= totalPages) {
        currentPage = next;
        renderTable();
      }
    }

    function toggleDrawer(contentId, caretId) {
      const c = document.getElementById(contentId);
      const caret = document.getElementById(caretId);
      if (c.classList.contains('hidden')) {
        c.classList.remove('hidden');
        if (caret) caret.innerHTML = '&#9650;';
      } else {
        c.classList.add('hidden');
        if (caret) caret.innerHTML = '&#9660;';
      }
    }

    function openAuditModal(rowIndex) {
      const rows = getFilteredRows();
      const row = rows[rowIndex];
      if (!row) return;

      document.getElementById('modal-wo-title').textContent = 'Record Audit: Work Order ' + row.workOrderNo;
      document.getElementById('modal-wo-sub').textContent = 'Match Status: ' + (row.matchStatus === 'matched' ? 'Matched with JMS' : row.matchStatus === 'unmatched_bridge' ? 'Bridge Only' : 'JMS Only') + ' · X Track Source: ' + row.xTrackIdSource;

      const body = document.getElementById('modal-body-content');
      body.innerHTML =
        '<div style="background: #f0fdfa; border: 1px solid #99f6e4; border-radius: 8px; padding: 12px; margin-bottom: 14px;">' +
          '<h4 style="font-size: 11px; font-weight: 700; color: #0f766e; text-transform: uppercase; margin-bottom: 8px;">XBRL Feedback &amp; Quality Audit</h4>' +
          '<div class="audit-grid">' +
            '<div class="audit-card">' +
              '<span class="label">JMS Cycle No / FB Quality:</span>' +
              '<span class="val" style="color: #0f766e;">' + escapeHtml(row.jmsCycleNoFbQuality || '—') + '</span>' +
              '<span style="font-size: 10px; color: #0f766e;">Mapped from "XBRL Feedback Quality"</span>' +
            '</div>' +
            '<div class="audit-card">' +
              '<span class="label">Sampling/Quality:</span>' +
              '<span class="val" style="color: #3730a3;">' + escapeHtml(row.samplingQuality || '—') + '</span>' +
              '<span style="font-size: 10px; color: #3730a3;">Filled when FB Quality present</span>' +
            '</div>' +
            '<div class="audit-card">' +
              '<span class="label">Feedback Status &amp; Validity:</span>' +
              '<span class="val">' + escapeHtml(row.xbrlFeedbackStatus || 'Completed') + ' &middot; ' + escapeHtml(row.xbrlFeedbackValidity || 'Valid') + '</span>' +
              '<span style="font-size: 10px; color: #64748b;">Reviewed: ' + escapeHtml(row.xbrlFeedbackReviewed || 'Yes') + '</span>' +
            '</div>' +
          '</div>' +
          (row.xbrlFeedbackComments ? '<div style="background: #fff; padding: 8px 10px; border-radius: 4px; border: 1px solid #ccfbf1; font-size: 11px;"><strong>Feedback Comments:</strong> &ldquo;' + escapeHtml(row.xbrlFeedbackComments) + '&rdquo;</div>' : '') +
        '</div>' +

        '<h4 style="font-size: 12px; font-weight: 700; color: #0f172a; margin-bottom: 8px;">Standard 15 Output Values</h4>' +
        '<div class="audit-grid">' +
          '<div class="audit-card"><span class="label">Team:</span><span class="val">' + escapeHtml(row.team) + '</span></div>' +
          '<div class="audit-card"><span class="label">Work Order No:</span><span class="val">' + escapeHtml(row.workOrderNo) + '</span></div>' +
          '<div class="audit-card"><span class="label">X Track ID:</span><span class="val" style="color: #0284c7;">' + escapeHtml(row.xTrackId || 'None') + '</span></div>' +
          '<div class="audit-card"><span class="label">Job Name:</span><span class="val">' + escapeHtml(row.jobName) + '</span></div>' +
          '<div class="audit-card"><span class="label">Form Type:</span><span class="val">' + escapeHtml(row.formType) + '</span></div>' +
          '<div class="audit-card"><span class="label">Milestone:</span><span class="val">' + escapeHtml(row.milestone) + '</span></div>' +
          '<div class="audit-card"><span class="label">Delivered Date:</span><span class="val">' + escapeHtml(row.deliveredDate) + '</span></div>' +
          '<div class="audit-card"><span class="label">Month:</span><span class="val">' + escapeHtml(row.month) + '</span></div>' +
        '</div>' +

        (row.rawBridge ? '<h4 style="font-size: 12px; font-weight: 700; margin-top: 14px; margin-bottom: 6px;">Raw Bridge Record:</h4><pre class="code-block">' + escapeHtml(JSON.stringify(row.rawBridge, null, 2)) + '</pre>' : '') +
        (row.rawJms ? '<h4 style="font-size: 12px; font-weight: 700; margin-top: 14px; margin-bottom: 6px;">Raw JMS Record:</h4><pre class="code-block">' + escapeHtml(JSON.stringify(row.rawJms, null, 2)) + '</pre>' : '');

      document.getElementById('audit-modal').classList.remove('hidden');
    }

    function closeModal() {
      document.getElementById('audit-modal').classList.add('hidden');
    }

    // Excel Parser
    function parseWorkbookFile(file, callback) {
      if (!window.XLSX) {
        alert('SheetJS Excel engine is not loaded. Please ensure the file is opened properly.');
        return;
      }
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });
          const sheets = {};
          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
            const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
            sheets[sheetName] = { sheetName, headers, data: rows, rowCount: rows.length };
          });
          callback(null, { fileName: file.name, sheetNames: workbook.SheetNames, sheets });
        } catch (err) {
          callback(err);
        }
      };
      reader.onerror = callback;
      reader.readAsArrayBuffer(file);
    }

    function handleBridgeFile(file) {
      if (!file) return;
      parseWorkbookFile(file, function(err, result) {
        if (err) { alert('Failed to read file: ' + err.message); return; }
        const preferred = result.sheetNames.find(s => /bridge|work.*order|order|active/i.test(s)) || result.sheetNames[0] || '';
        const sheetData = result.sheets[preferred] || { headers: [], data: [] };

        bridgeFile = {
          fileName: result.fileName,
          sheetNames: result.sheetNames,
          selectedSheet: preferred,
          sheets: result.sheets,
          headers: sheetData.headers,
          data: sheetData.data
        };

        document.getElementById('bridge-dropzone').style.display = 'none';
        document.getElementById('bridge-file-bar').style.display = 'block';
        document.getElementById('bridge-status-badge').style.display = 'inline-block';
        document.getElementById('bridge-file-name').textContent = result.fileName;
        document.getElementById('bridge-file-meta').innerHTML = sheetData.data.length + ' records &middot; ' + sheetData.headers.length + ' columns';

        if (result.sheetNames.length > 1) {
          const sel = document.getElementById('bridge-sheet-select');
          sel.innerHTML = result.sheetNames.map(s => '<option value="' + escapeHtml(s) + '" ' + (s === preferred ? 'selected' : '') + '>' + escapeHtml(s) + '</option>').join('');
          document.getElementById('bridge-sheet-select-row').style.display = 'flex';
        }

        autoDetect();
        runReconciliation();
      });
    }

    function onBridgeSheetChange(sheetName) {
      bridgeFile.selectedSheet = sheetName;
      const sheetData = bridgeFile.sheets[sheetName] || { headers: [], data: [] };
      bridgeFile.headers = sheetData.headers;
      bridgeFile.data = sheetData.data;
      document.getElementById('bridge-file-meta').innerHTML = sheetData.data.length + ' records &middot; ' + sheetData.headers.length + ' columns';
      autoDetect();
      runReconciliation();
    }

    function handleJmsFile(file) {
      if (!file) return;
      parseWorkbookFile(file, function(err, result) {
        if (err) { alert('Failed to read workbook: ' + err.message); return; }
        const prodSheetName = result.sheetNames.find(s => /production.*day.*to.*day/i.test(s)) || result.sheetNames.find(s => /production/i.test(s)) || '';
        const nonProd = result.sheetNames.filter(s => s !== prodSheetName);
        let chosen = nonProd[0] || result.sheetNames[0] || '';
        for (const s of (nonProd.length ? nonProd : result.sheetNames)) {
          if (/grabber|jms.*active|active.*job|jms|saas|active/i.test(s)) {
            chosen = s;
            break;
          }
        }
        const sheetData = result.sheets[chosen] || { headers: [], data: [] };

        jmsFile = {
          fileName: result.fileName,
          sheetNames: result.sheetNames,
          selectedSheet: chosen,
          sheets: result.sheets,
          headers: sheetData.headers,
          data: sheetData.data
        };

        document.getElementById('jms-dropzone').style.display = 'none';
        document.getElementById('jms-file-bar').style.display = 'block';
        document.getElementById('jms-status-badge').style.display = 'inline-block';
        document.getElementById('jms-file-name').textContent = result.fileName;
        document.getElementById('jms-file-meta').innerHTML = sheetData.data.length + ' records &middot; ' + sheetData.headers.length + ' columns';

        if (result.sheetNames.length > 1) {
          const sel = document.getElementById('jms-sheet-select');
          sel.innerHTML = result.sheetNames.map(s => '<option value="' + escapeHtml(s) + '" ' + (s === chosen ? 'selected' : '') + '>' + escapeHtml(s) + '</option>').join('');
          document.getElementById('jms-sheet-select-row').style.display = 'flex';
        }

        autoDetect();
        runReconciliation();
      });
    }

    function onJmsSheetChange(sheetName) {
      jmsFile.selectedSheet = sheetName;
      const sheetData = jmsFile.sheets[sheetName] || { headers: [], data: [] };
      jmsFile.headers = sheetData.headers;
      jmsFile.data = sheetData.data;
      document.getElementById('jms-file-meta').innerHTML = sheetData.data.length + ' records &middot; ' + sheetData.headers.length + ' columns';
      autoDetect();
      runReconciliation();
    }

    function resetToAutoDetect() {
      autoDetect();
      runReconciliation();
    }

    function loadSampleData() {
      const bh = Object.keys(DEMO_BRIDGE_DATA[0]);
      const jh = Object.keys(DEMO_JMS_DATA[0]);
      const ph = Object.keys(DEMO_PROD_DATA[0]);

      bridgeFile = {
        fileName: 'Sample_Bridge_Work_Orders.xlsx',
        sheetNames: ['Bridge Work Orders'],
        selectedSheet: 'Bridge Work Orders',
        sheets: {
          'Bridge Work Orders': { sheetName: 'Bridge Work Orders', headers: bh, data: DEMO_BRIDGE_DATA, rowCount: DEMO_BRIDGE_DATA.length }
        },
        headers: bh,
        data: DEMO_BRIDGE_DATA
      };

      document.getElementById('bridge-dropzone').style.display = 'none';
      document.getElementById('bridge-file-bar').style.display = 'block';
      document.getElementById('bridge-status-badge').style.display = 'inline-block';
      document.getElementById('bridge-file-name').textContent = bridgeFile.fileName;
      document.getElementById('bridge-file-meta').innerHTML = DEMO_BRIDGE_DATA.length + ' records &middot; ' + bh.length + ' columns';

      jmsFile = {
        fileName: 'Sample_JMS_SaaS_Workbook.xlsx',
        sheetNames: ['JMS Active Jobs', 'Production Day to Day'],
        selectedSheet: 'JMS Active Jobs',
        sheets: {
          'JMS Active Jobs': { sheetName: 'JMS Active Jobs', headers: jh, data: DEMO_JMS_DATA, rowCount: DEMO_JMS_DATA.length },
          'Production Day to Day': { sheetName: 'Production Day to Day', headers: ph, data: DEMO_PROD_DATA, rowCount: DEMO_PROD_DATA.length }
        },
        headers: jh,
        data: DEMO_JMS_DATA
      };

      document.getElementById('jms-dropzone').style.display = 'none';
      document.getElementById('jms-file-bar').style.display = 'block';
      document.getElementById('jms-status-badge').style.display = 'inline-block';
      document.getElementById('jms-file-name').textContent = jmsFile.fileName;
      document.getElementById('jms-file-meta').innerHTML = DEMO_JMS_DATA.length + ' records &middot; ' + jh.length + ' columns';

      autoDetect();
      runReconciliation();
    }

    function resetAll() {
      bridgeFile = { fileName: '', sheetNames: [], selectedSheet: '', sheets: {}, headers: [], data: [] };
      jmsFile = { fileName: '', sheetNames: [], selectedSheet: '', sheets: {}, headers: [], data: [] };
      reconciledData = { matchedRows: [], unmatchedBridgeRows: [], unmatchedJmsRows: [], allRows: [] };

      document.getElementById('bridge-dropzone').style.display = 'flex';
      document.getElementById('bridge-file-bar').style.display = 'none';
      document.getElementById('bridge-status-badge').style.display = 'none';
      document.getElementById('bridge-file-input').value = '';

      document.getElementById('jms-dropzone').style.display = 'flex';
      document.getElementById('jms-file-bar').style.display = 'none';
      document.getElementById('jms-status-badge').style.display = 'none';
      document.getElementById('jms-file-input').value = '';

      renderKPIs();
      renderTable();
    }

    // Drag and drop event listeners
    function setupDragDrop(zoneId, inputId, handler) {
      const el = document.getElementById(zoneId);
      if (!el) return;
      ['dragenter', 'dragover'].forEach(name => {
        el.addEventListener(name, (e) => { e.preventDefault(); e.stopPropagation(); el.classList.add('dragover'); });
      });
      ['dragleave', 'drop'].forEach(name => {
        el.addEventListener(name, (e) => { e.preventDefault(); e.stopPropagation(); el.classList.remove('dragover'); });
      });
      el.addEventListener('drop', (e) => {
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          handler(e.dataTransfer.files[0]);
        }
      });
    }

    setupDragDrop('bridge-dropzone', 'bridge-file-input', handleBridgeFile);
    setupDragDrop('jms-dropzone', 'jms-file-input', handleJmsFile);

    // Export Utilities
    const EXPORT_COLUMNS = [
      { header: 'Team', key: 'team' },
      { header: 'JMS Job Number/Work Order No', key: 'workOrderNo' },
      { header: 'JMS Cycle No / FB Quality', key: 'jmsCycleNoFbQuality' },
      { header: 'XOR Ver', key: 'xorVer' },
      { header: 'X Track ID', key: 'xTrackId' },
      { header: 'Job Name', key: 'jobName' },
      { header: 'Form Type', key: 'formType' },
      { header: 'Filing Period', key: 'filingPeriod' },
      { header: 'Milestone', key: 'milestone' },
      { header: 'JMS Team', key: 'jmsTeam' },
      { header: 'Sampling/Quality', key: 'samplingQuality' },
      { header: 'Consultant Name', key: 'consultantName' },
      { header: '2nd Reviewer Name', key: 'secondReviewerName' },
      { header: 'Delivered Date', key: 'deliveredDate' },
      { header: 'Month', key: 'month' }
    ];

    function exportExcel() {
      const rows = getFilteredRows();
      if (!rows.length) { alert('No rows to export'); return; }
      if (!window.XLSX) { alert('XLSX library not loaded'); return; }

      const exportData = rows.map(r => {
        const item = {};
        EXPORT_COLUMNS.forEach(c => { item[c.header] = r[c.key] !== undefined ? r[c.key] : ''; });
        return item;
      });

      const ws = XLSX.utils.json_to_sheet(exportData);
      ws['!cols'] = EXPORT_COLUMNS.map(c => ({ wch: Math.max(c.header.length + 4, 16) }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Reconciled Output');

      const filterTag = activeFilter === 'matched' ? 'Matched' : activeFilter === 'unmatched_bridge' ? 'Bridge_Unmatched' : activeFilter === 'unmatched_jms' ? 'JMS_Unmatched' : 'All';
      const dateStr = new Date().toISOString().slice(0, 10);
      XLSX.writeFile(wb, 'Bridge_Work_Orders_Reconciled_' + filterTag + '_' + dateStr + '.xlsx');
    }

    function exportCSV() {
      const rows = getFilteredRows();
      if (!rows.length) { alert('No rows to export'); return; }
      if (!window.XLSX) { alert('XLSX library not loaded'); return; }

      const exportData = rows.map(r => {
        const item = {};
        EXPORT_COLUMNS.forEach(c => { item[c.header] = r[c.key] !== undefined ? r[c.key] : ''; });
        return item;
      });
      const ws = XLSX.utils.json_to_sheet(exportData);
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'Bridge_Work_Orders_' + activeFilter + '_' + (new Date().toISOString().slice(0,10)) + '.csv';
      link.click();
    }

    function copyTSV() {
      const rows = getFilteredRows();
      if (!rows.length) { alert('No rows to copy'); return; }
      const headers = EXPORT_COLUMNS.map(c => c.header).join('\\t');
      const lines = rows.map(r => EXPORT_COLUMNS.map(c => r[c.key] !== undefined ? r[c.key] : '').join('\\t'));
      const tsv = [headers, ...lines].join('\\n');
      navigator.clipboard.writeText(tsv).then(() => {
        alert('Copied ' + rows.length + ' rows to clipboard as TSV! You can now paste directly into Excel.');
      }).catch(err => {
        alert('Copy failed: ' + err);
      });
    }

    function escapeHtml(s) {
      if (s === null || s === undefined) return '';
      return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
    }

    // Auto-load demo data on launch for instant preview
    loadSampleData();
  </script>
</body>
</html>`;
}

const htmlContent = generateHtml();
const outputPath = path.join(__dirname, '../public/Bridge_Work_Orders_Standalone_App.html');
fs.writeFileSync(outputPath, htmlContent, 'utf8');
console.log('Successfully generated Standalone App (.html) at: ' + outputPath + ' (' + Math.round(fs.statSync(outputPath).size / 1024) + ' KB)');
