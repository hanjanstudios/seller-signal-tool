import React, { useState, useMemo, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';

const FIELD_DEFS = [
  { key: 'ownerName', label: 'Owner name', required: false, hints: ['owner 1 name', 'owner1 name', 'owner name', 'owner full name', 'name'] },
  { key: 'propertyAddress', label: 'Property (situs) address', required: false, hints: ['situs address', 'property address', 'site address', 'address'] },
  { key: 'mailingAddress', label: 'Owner mailing address', required: false, hints: ['mailing address', 'owner mailing address', 'mail address'] },
  { key: 'ownerOccupied', label: 'Owner occupied (Y/N)', required: false, hints: ['owner occupied', 'owner occ', 'owneroccupied'] },
  { key: 'vacant', label: 'Vacant (Y/N)', required: true, hints: ['vacant'] },
  { key: 'taxDelinquent', label: 'Tax delinquent', required: true, hints: ['delinquent', 'tax delinq'] },
  { key: 'bedrooms', label: 'Bedrooms', required: true, hints: ['bedrooms', 'beds', 'bed cnt', 'bed'] },
  { key: 'bathrooms', label: 'Bathrooms', required: true, hints: ['bathrooms', 'baths', 'bath cnt', 'bath'] },
  { key: 'saleDate', label: 'Sale date', required: true, hints: ['sale date', 'last sale date', 'recording date', 'deed date', 'transfer date'] },
];

function guessColumn(headers, hints) {
  const lower = headers.map((h) => String(h).toLowerCase().trim());
  for (const hint of hints) {
    const idx = lower.findIndex((h) => h === hint);
    if (idx !== -1) return headers[idx];
  }
  for (const hint of hints) {
    const idx = lower.findIndex((h) => h.includes(hint));
    if (idx !== -1) return headers[idx];
  }
  return '';
}

function parseBool(val) {
  if (val === null || val === undefined || val === '') return null;
  const s = String(val).trim().toLowerCase();
  if (['y', 'yes', 'true', 't'].includes(s)) return true;
  if (['n', 'no', 'false', 'f'].includes(s)) return false;
  const num = parseFloat(s);
  if (!isNaN(num)) return num > 0;
  return null;
}

function normalizeAddress(addr) {
  return String(addr || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

function parseYearsOwned(row, map) {
  if (map.saleDate && row[map.saleDate] !== '') {
    const raw = row[map.saleDate];
    let d = new Date(raw);
    if (isNaN(d.getTime())) {
      const serial = parseFloat(raw);
      if (!isNaN(serial) && serial > 20000 && serial < 60000) {
        d = new Date(Math.round((serial - 25569) * 86400 * 1000));
      }
    }
    if (!isNaN(d.getTime())) {
      const years = (Date.now() - d.getTime()) / (1000 * 60 * 60 * 24 * 365.25);
      if (years >= 0 && years < 150) return years;
    }
  }
  return null;
}

function computeRow(row, map, idx) {
  const beds = map.bedrooms && row[map.bedrooms] !== '' ? parseFloat(row[map.bedrooms]) : null;
  const baths = map.bathrooms && row[map.bathrooms] !== '' ? parseFloat(row[map.bathrooms]) : null;
  const yearsOwned = parseYearsOwned(row, map);
  const taxDelinquent = map.taxDelinquent ? !!parseBool(row[map.taxDelinquent]) : false;
  const vacant = map.vacant ? !!parseBool(row[map.vacant]) : false;

  let ownerOccupied = map.ownerOccupied ? parseBool(row[map.ownerOccupied]) : null;
  if (ownerOccupied === null && map.mailingAddress && map.propertyAddress) {
    const a = normalizeAddress(row[map.propertyAddress]);
    const b = normalizeAddress(row[map.mailingAddress]);
    if (a && b) ownerOccupied = a === b;
  }

  let score = 0;
  const tags = [];

  if (taxDelinquent) {
    score += 30;
    tags.push('Tax delinquent');
  }
  if (vacant) {
    score += 25;
    tags.push('Vacant');
  }

  const isUpsize =
    beds !== null && baths !== null && beds <= 2 && baths <= 1 &&
    yearsOwned !== null && yearsOwned >= 1 && yearsOwned <= 6;
  if (isUpsize) {
    score += 25;
    tags.push('Upsize candidate');
  }

  let isDownsize = false;
  if (yearsOwned !== null && yearsOwned >= 10) {
    isDownsize = true;
    score += yearsOwned >= 15 ? 20 : 15;
    tags.push('Downsize candidate');
  }

  if (ownerOccupied === false) {
    score += 15;
    tags.push('Non-owner occupied');
  }

  score = Math.min(100, Math.round(score));
  let category = 'Low';
  if (score >= 60) category = 'High';
  else if (score >= 30) category = 'Medium';

  return {
    id: idx,
    ownerName: map.ownerName ? String(row[map.ownerName] || '').trim() : '',
    propertyAddress: map.propertyAddress ? String(row[map.propertyAddress] || '').trim() : '',
    beds,
    baths,
    yearsOwned,
    taxDelinquent,
    vacant,
    ownerOccupied,
    isUpsize,
    isDownsize,
    score,
    category,
    tags,
  };
}

function downloadCSV(rows) {
  const headers = ['Owner name', 'Property address', 'Beds', 'Baths', 'Years owned', 'Owner occupied', 'Vacant', 'Tax delinquent', 'Score', 'Category', 'Signals'];
  const lines = [headers.join(',')];
  rows.forEach((r) => {
    const vals = [
      r.ownerName,
      r.propertyAddress,
      r.beds ?? '',
      r.baths ?? '',
      r.yearsOwned !== null ? r.yearsOwned.toFixed(1) : '',
      r.ownerOccupied === null ? '' : r.ownerOccupied ? 'Y' : 'N',
      r.vacant ? 'Y' : 'N',
      r.taxDelinquent ? 'Y' : 'N',
      r.score,
      r.category,
      r.tags.join('; '),
    ].map((v) => `"${String(v).replace(/"/g, '""')}"`);
    lines.push(vals.join(','));
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'seller-signal-results.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export default function App() {
  const [stage, setStage] = useState('upload');
  const [fileName, setFileName] = useState('');
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [columnMap, setColumnMap] = useState({});
  const [error, setError] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [showMethod, setShowMethod] = useState(false);
  const fileInputRef = useRef(null);

  const [search, setSearch] = useState('');
  const [filters, setFilters] = useState({
    taxDelinquent: false,
    vacant: false,
    ownerOcc: 'any',
    upsize: false,
    downsize: false,
  });
  const [sortConfig, setSortConfig] = useState({ key: 'score', dir: 'desc' });

  const readFile = useCallback((file) => {
    if (!file) return;
    setError('');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
        if (!json.length) {
          setError('That file loaded, but no rows were found on the first sheet.');
          return;
        }
        const hdrs = Object.keys(json[0]);
        const guessedMap = {};
        FIELD_DEFS.forEach((f) => {
          guessedMap[f.key] = guessColumn(hdrs, f.hints);
        });
        setHeaders(hdrs);
        setRawRows(json);
        setColumnMap(guessedMap);
        setFileName(file.name);
        setStage('mapping');
      } catch (err) {
        setError('Could not read this file. Make sure it is a .csv, .xls, or .xlsx export from your title company.');
      }
    };
    reader.onerror = () => setError('Something went wrong reading that file. Please try again.');
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    readFile(file);
  }, [readFile]);

  const onFileInput = useCallback((e) => {
    const file = e.target.files && e.target.files[0];
    readFile(file);
  }, [readFile]);

  const processedRows = useMemo(() => {
    if (stage !== 'results' || !rawRows.length) return [];
    return rawRows.map((row, idx) => computeRow(row, columnMap, idx));
  }, [stage, rawRows, columnMap]);

  const summary = useMemo(() => {
    const total = processedRows.length;
    const high = processedRows.filter((r) => r.category === 'High').length;
    const taxDelinquent = processedRows.filter((r) => r.taxDelinquent).length;
    const vacant = processedRows.filter((r) => r.vacant).length;
    const upsize = processedRows.filter((r) => r.isUpsize).length;
    const downsize = processedRows.filter((r) => r.isDownsize).length;
    return { total, high, taxDelinquent, vacant, upsize, downsize };
  }, [processedRows]);

  const visibleRows = useMemo(() => {
    let rows = processedRows;
    if (filters.taxDelinquent) rows = rows.filter((r) => r.taxDelinquent);
    if (filters.vacant) rows = rows.filter((r) => r.vacant);
    if (filters.upsize) rows = rows.filter((r) => r.isUpsize);
    if (filters.downsize) rows = rows.filter((r) => r.isDownsize);
    if (filters.ownerOcc === 'occupied') rows = rows.filter((r) => r.ownerOccupied === true);
    if (filters.ownerOcc === 'nonoccupied') rows = rows.filter((r) => r.ownerOccupied === false);
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      rows = rows.filter(
        (r) => r.ownerName.toLowerCase().includes(q) || r.propertyAddress.toLowerCase().includes(q)
      );
    }
    const { key, dir } = sortConfig;
    const sorted = [...rows].sort((a, b) => {
      let av = a[key];
      let bv = b[key];
      if (av === null || av === undefined) av = -Infinity;
      if (bv === null || bv === undefined) bv = -Infinity;
      if (typeof av === 'string') av = av.toLowerCase();
      if (typeof bv === 'string') bv = bv.toLowerCase();
      if (av < bv) return dir === 'asc' ? -1 : 1;
      if (av > bv) return dir === 'asc' ? 1 : -1;
      return 0;
    });
    return sorted;
  }, [processedRows, filters, search, sortConfig]);

  const toggleSort = (key) => {
    setSortConfig((prev) => {
      if (prev.key === key) return { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' };
      return { key, dir: key === 'score' || key === 'yearsOwned' ? 'desc' : 'asc' };
    });
  };

  const reset = () => {
    setStage('upload');
    setFileName('');
    setHeaders([]);
    setRawRows([]);
    setColumnMap({});
    setError('');
    setSearch('');
    setFilters({ taxDelinquent: false, vacant: false, ownerOcc: 'any', upsize: false, downsize: false });
    setSortConfig({ key: 'score', dir: 'desc' });
  };

  const missingCritical = FIELD_DEFS.filter((f) => f.required && !columnMap[f.key]);
  const noOwnershipLength = !columnMap.saleDate;

  const sortArrow = (key) => {
    if (sortConfig.key !== key) return '';
    return sortConfig.dir === 'asc' ? ' \u2191' : ' \u2193';
  };

  return (
    <div className="fss-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Lora:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap');

        .fss-root {
          --paper: #FAF7F1;
          --paper-line: #DED6C4;
          --ink: #232A2E;
          --ink-soft: #5B6469;
          --navy: #223345;
          --stamp: #A8412F;
          --stamp-soft: #F3E3DF;
          --brass: #96742E;
          --brass-soft: #F1E8D3;
          --sage: #57705C;
          --sage-soft: #E3EBE1;
          font-family: 'IBM Plex Sans', sans-serif;
          color: var(--ink);
          background: var(--paper);
          min-height: 100vh;
          padding: 40px 24px 80px;
          box-sizing: border-box;
        }
        .fss-root * { box-sizing: border-box; }
        .fss-shell { max-width: 980px; margin: 0 auto; }

        .fss-head { border-bottom: 3px double var(--paper-line); padding-bottom: 22px; margin-bottom: 32px; }
        .fss-eyebrow { font-size: 13px; color: var(--brass); letter-spacing: 0.02em; margin: 0 0 8px; font-weight: 600; }
        .fss-title { font-family: 'Lora', serif; font-size: 32px; font-weight: 600; margin: 0 0 10px; color: var(--navy); line-height: 1.2; }
        .fss-sub { font-size: 15px; color: var(--ink-soft); margin: 0; max-width: 60ch; line-height: 1.55; }

        /* Upload stage */
        .fss-dropzone {
          border: 1.5px dashed var(--paper-line);
          background: #fff;
          border-radius: 4px;
          padding: 56px 24px;
          text-align: center;
          cursor: pointer;
          transition: border-color 0.15s ease, background 0.15s ease;
        }
        .fss-dropzone.dragging { border-color: var(--brass); background: var(--brass-soft); }
        .fss-dropzone:focus-visible { outline: 2px solid var(--navy); outline-offset: 3px; }
        .fss-drop-icon { font-family: 'Lora', serif; font-size: 15px; color: var(--navy); margin-bottom: 6px; }
        .fss-drop-main { font-size: 16px; font-weight: 600; color: var(--ink); margin-bottom: 6px; }
        .fss-drop-sub { font-size: 13.5px; color: var(--ink-soft); }
        .fss-drop-formats { margin-top: 18px; font-size: 12.5px; color: var(--ink-soft); }

        .fss-error {
          margin-top: 16px;
          background: var(--stamp-soft);
          color: var(--stamp);
          border: 1px solid rgba(168, 65, 47, 0.3);
          padding: 12px 14px;
          border-radius: 4px;
          font-size: 13.5px;
        }

        /* Mapping stage */
        .fss-card { background: #fff; border: 1px solid var(--paper-line); border-radius: 4px; padding: 28px; }
        .fss-card-title { font-family: 'Lora', serif; font-size: 20px; font-weight: 600; color: var(--navy); margin: 0 0 4px; }
        .fss-card-note { font-size: 13.5px; color: var(--ink-soft); margin: 0 0 24px; line-height: 1.5; }
        .fss-map-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 24px; }
        @media (max-width: 640px) { .fss-map-grid { grid-template-columns: 1fr; } }
        .fss-map-row { display: flex; flex-direction: column; gap: 6px; padding-bottom: 12px; border-bottom: 1px solid #F0ECE2; }
        .fss-map-label { font-size: 13px; font-weight: 600; color: var(--ink); }
        .fss-map-label .req { color: var(--stamp); margin-left: 3px; }
        .fss-select {
          font-family: 'IBM Plex Sans', sans-serif;
          font-size: 13.5px;
          padding: 8px 10px;
          border: 1px solid var(--paper-line);
          border-radius: 3px;
          background: #fff;
          color: var(--ink);
        }
        .fss-select:focus-visible { outline: 2px solid var(--navy); outline-offset: 1px; }

        .fss-warn {
          margin-top: 20px;
          font-size: 13px;
          color: var(--brass);
          background: var(--brass-soft);
          padding: 10px 14px;
          border-radius: 4px;
          line-height: 1.5;
        }

        .fss-actions { margin-top: 26px; display: flex; align-items: center; gap: 16px; }
        .fss-btn {
          font-family: 'IBM Plex Sans', sans-serif;
          font-size: 14px;
          font-weight: 600;
          padding: 11px 22px;
          border-radius: 3px;
          border: none;
          cursor: pointer;
          transition: opacity 0.15s ease;
        }
        .fss-btn:hover { opacity: 0.88; }
        .fss-btn:disabled { opacity: 0.45; cursor: not-allowed; }
        .fss-btn-primary { background: var(--navy); color: #fff; }
        .fss-btn-ghost { background: transparent; color: var(--ink-soft); text-decoration: underline; padding: 11px 4px; }

        /* Results stage */
        .fss-toolbar-top { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 22px; flex-wrap: wrap; gap: 10px; }
        .fss-file-info { font-size: 13px; color: var(--ink-soft); }
        .fss-file-info strong { color: var(--ink); }

        .fss-stats { display: grid; grid-template-columns: repeat(6, 1fr); border: 1px solid var(--paper-line); border-radius: 4px; overflow: hidden; background: #fff; margin-bottom: 26px; }
        @media (max-width: 800px) { .fss-stats { grid-template-columns: repeat(3, 1fr); } }
        .fss-stat { padding: 16px 12px; text-align: center; border-right: 1px solid var(--paper-line); }
        .fss-stat:last-child { border-right: none; }
        .fss-stat-num { font-family: 'Lora', serif; font-size: 26px; font-weight: 600; color: var(--navy); }
        .fss-stat-label { font-size: 11.5px; color: var(--ink-soft); margin-top: 4px; }

        .fss-filterbar { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; margin-bottom: 18px; }
        .fss-search {
          flex: 1 1 220px;
          font-size: 13.5px;
          padding: 9px 12px;
          border: 1px solid var(--paper-line);
          border-radius: 3px;
          background: #fff;
        }
        .fss-chip {
          font-size: 13px;
          padding: 8px 14px;
          border-radius: 20px;
          border: 1px solid var(--paper-line);
          background: #fff;
          color: var(--ink-soft);
          cursor: pointer;
          white-space: nowrap;
          transition: all 0.12s ease;
        }
        .fss-chip.active { background: var(--navy); border-color: var(--navy); color: #fff; }
        .fss-owner-select {
          font-size: 13px;
          padding: 8px 10px;
          border: 1px solid var(--paper-line);
          border-radius: 3px;
          background: #fff;
          color: var(--ink);
        }

        .fss-table-wrap { border: 1px solid var(--paper-line); border-radius: 4px; overflow: auto; max-height: 560px; background: #fff; }
        table.fss-table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
        .fss-table thead th {
          position: sticky; top: 0; background: var(--paper);
          text-align: left; padding: 12px 14px; font-weight: 600; color: var(--navy);
          border-bottom: 2px solid var(--paper-line); cursor: pointer; user-select: none; white-space: nowrap;
        }
        .fss-table tbody td { padding: 11px 14px; border-bottom: 1px solid #F0ECE2; vertical-align: top; }
        .fss-table tbody tr:hover { background: #FBF9F4; }
        .fss-name { font-weight: 600; color: var(--ink); }
        .fss-addr { color: var(--ink-soft); font-size: 12.5px; margin-top: 2px; }

        .fss-tags { display: flex; flex-wrap: wrap; gap: 5px; }
        .fss-tag { font-size: 11px; padding: 3px 8px; border-radius: 3px; white-space: nowrap; }
        .fss-tag-delinquent { background: var(--stamp-soft); color: var(--stamp); }
        .fss-tag-vacant { background: var(--brass-soft); color: var(--brass); }
        .fss-tag-upsize { background: var(--sage-soft); color: var(--sage); }
        .fss-tag-downsize { background: #E4E9EF; color: var(--navy); }
        .fss-tag-nonocc { background: #EFEFEF; color: var(--ink-soft); }

        .fss-score-wrap { display: flex; align-items: center; gap: 8px; }
        .fss-score-num { font-family: 'Lora', serif; font-weight: 600; font-size: 15px; width: 26px; text-align: right; }
        .fss-score-bar { width: 44px; height: 6px; border-radius: 3px; background: #EDEAE1; overflow: hidden; }
        .fss-score-fill { height: 100%; }
        .fss-score-High .fss-score-num, .fss-score-High .fss-score-fill { color: var(--stamp); background: var(--stamp); }
        .fss-score-Medium .fss-score-num, .fss-score-Medium .fss-score-fill { color: var(--brass); background: var(--brass); }
        .fss-score-Low .fss-score-num, .fss-score-Low .fss-score-fill { color: var(--ink-soft); background: #B7BEC1; }

        .fss-empty { padding: 48px 20px; text-align: center; color: var(--ink-soft); font-size: 14px; }

        .fss-footer-bar { display: flex; justify-content: space-between; align-items: center; margin-top: 20px; flex-wrap: wrap; gap: 12px; }
        .fss-count { font-size: 13px; color: var(--ink-soft); }
        .fss-export-btn { background: var(--brass); color: #fff; }

        .fss-method-toggle { font-size: 13px; color: var(--navy); text-decoration: underline; cursor: pointer; background: none; border: none; padding: 0; margin-top: 4px; }
        .fss-method-panel { margin-top: 14px; background: var(--sage-soft); border-radius: 4px; padding: 18px 20px; font-size: 13px; color: var(--ink); line-height: 1.7; }
        .fss-method-panel ul { margin: 6px 0 0; padding-left: 18px; }
      `}</style>

      <div className="fss-shell">
        <div className="fss-head">
          <p className="fss-eyebrow">Seller signal</p>
          <h1 className="fss-title">Find your next likely sellers</h1>
          <p className="fss-sub">
            Upload a farm export from Title Toolbox or your title rep, and this tool ranks every
            home by how likely the owner is to sell next — based on tax delinquency, vacancy,
            and ownership-stage patterns.
          </p>
        </div>

        {stage === 'upload' && (
          <>
            <div
              className={`fss-dropzone${isDragging ? ' dragging' : ''}`}
              role="button"
              tabIndex={0}
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') fileInputRef.current.click(); }}
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={onDrop}
            >
              <div className="fss-drop-icon">— farm file —</div>
              <div className="fss-drop-main">Drop your farm file here, or click to browse</div>
              <div className="fss-drop-sub">One file, usually 200–500 homes</div>
              <div className="fss-drop-formats">Accepts .csv, .xls, .xlsx</div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".csv,.xls,.xlsx"
                onChange={onFileInput}
                style={{ display: 'none' }}
              />
            </div>
            {error && <div className="fss-error">{error}</div>}
          </>
        )}

        {stage === 'mapping' && (
          <div className="fss-card">
            <h2 className="fss-card-title">Check your columns</h2>
            <p className="fss-card-note">
              We matched these automatically from <strong>{fileName}</strong> ({rawRows.length} rows).
              Double-check the ones marked required, especially bedrooms, bathrooms, vacancy, and
              tax delinquency — the ranking depends on them.
            </p>
            <div className="fss-map-grid">
              {FIELD_DEFS.map((f) => (
                <div className="fss-map-row" key={f.key}>
                  <label className="fss-map-label" htmlFor={`map-${f.key}`}>
                    {f.label}
                    {f.required && <span className="req">required</span>}
                  </label>
                  <select
                    id={`map-${f.key}`}
                    className="fss-select"
                    value={columnMap[f.key] || ''}
                    onChange={(e) => setColumnMap((prev) => ({ ...prev, [f.key]: e.target.value }))}
                  >
                    <option value="">— not in this file —</option>
                    {headers.map((h) => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>

            {noOwnershipLength && (
              <div className="fss-warn">
                No "sale date" column is mapped, so upsize and downsize signals won't be
                calculated for this farm — only vacancy and tax delinquency will count toward
                the score.
              </div>
            )}

            <div className="fss-actions">
              <button
                className="fss-btn fss-btn-primary"
                disabled={missingCritical.length === FIELD_DEFS.filter((f) => f.required).length}
                onClick={() => setStage('results')}
              >
                Analyze this farm
              </button>
              <button className="fss-btn fss-btn-ghost" onClick={reset}>Upload a different file</button>
            </div>
          </div>
        )}

        {stage === 'results' && (
          <>
            <div className="fss-toolbar-top">
              <div className="fss-file-info">
                <strong>{fileName}</strong> — {summary.total} homes analyzed
              </div>
              <button className="fss-btn fss-btn-ghost" onClick={reset}>Upload a different file</button>
            </div>

            <div className="fss-stats">
              <div className="fss-stat"><div className="fss-stat-num">{summary.total}</div><div className="fss-stat-label">Total homes</div></div>
              <div className="fss-stat"><div className="fss-stat-num">{summary.high}</div><div className="fss-stat-label">High priority</div></div>
              <div className="fss-stat"><div className="fss-stat-num">{summary.taxDelinquent}</div><div className="fss-stat-label">Tax delinquent</div></div>
              <div className="fss-stat"><div className="fss-stat-num">{summary.vacant}</div><div className="fss-stat-label">Vacant</div></div>
              <div className="fss-stat"><div className="fss-stat-num">{summary.upsize}</div><div className="fss-stat-label">Upsize candidates</div></div>
              <div className="fss-stat"><div className="fss-stat-num">{summary.downsize}</div><div className="fss-stat-label">Downsize candidates</div></div>
            </div>

            <div className="fss-filterbar">
              <input
                className="fss-search"
                placeholder="Search by owner name or address…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              <button
                className={`fss-chip${filters.taxDelinquent ? ' active' : ''}`}
                onClick={() => setFilters((p) => ({ ...p, taxDelinquent: !p.taxDelinquent }))}
              >Tax delinquent</button>
              <button
                className={`fss-chip${filters.vacant ? ' active' : ''}`}
                onClick={() => setFilters((p) => ({ ...p, vacant: !p.vacant }))}
              >Vacant</button>
              <button
                className={`fss-chip${filters.upsize ? ' active' : ''}`}
                onClick={() => setFilters((p) => ({ ...p, upsize: !p.upsize }))}
              >Upsize candidates</button>
              <button
                className={`fss-chip${filters.downsize ? ' active' : ''}`}
                onClick={() => setFilters((p) => ({ ...p, downsize: !p.downsize }))}
              >Downsize candidates</button>
              <select
                className="fss-owner-select"
                value={filters.ownerOcc}
                onChange={(e) => setFilters((p) => ({ ...p, ownerOcc: e.target.value }))}
              >
                <option value="any">Owner occupied: any</option>
                <option value="occupied">Owner occupied only</option>
                <option value="nonoccupied">Non-owner occupied only</option>
              </select>
            </div>

            <div className="fss-table-wrap">
              {visibleRows.length === 0 ? (
                <div className="fss-empty">No homes match these filters. Try clearing one.</div>
              ) : (
                <table className="fss-table">
                  <thead>
                    <tr>
                      <th onClick={() => toggleSort('ownerName')}>Owner{sortArrow('ownerName')}</th>
                      <th onClick={() => toggleSort('beds')}>Beds / Baths{sortArrow('beds')}</th>
                      <th onClick={() => toggleSort('yearsOwned')}>Years owned{sortArrow('yearsOwned')}</th>
                      <th>Signals</th>
                      <th onClick={() => toggleSort('score')}>Score{sortArrow('score')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {visibleRows.map((r) => (
                      <tr key={r.id}>
                        <td>
                          <div className="fss-name">{r.ownerName || '—'}</div>
                          {r.propertyAddress && <div className="fss-addr">{r.propertyAddress}</div>}
                        </td>
                        <td>{r.beds ?? '—'} / {r.baths ?? '—'}</td>
                        <td>{r.yearsOwned !== null ? r.yearsOwned.toFixed(1) : '—'}</td>
                        <td>
                          <div className="fss-tags">
                            {r.taxDelinquent && <span className="fss-tag fss-tag-delinquent">Tax delinquent</span>}
                            {r.vacant && <span className="fss-tag fss-tag-vacant">Vacant</span>}
                            {r.isUpsize && <span className="fss-tag fss-tag-upsize">Upsize</span>}
                            {r.isDownsize && <span className="fss-tag fss-tag-downsize">Downsize</span>}
                            {r.ownerOccupied === false && <span className="fss-tag fss-tag-nonocc">Non-owner occ.</span>}
                            {r.tags.length === 0 && '—'}
                          </div>
                        </td>
                        <td>
                          <div className={`fss-score-wrap fss-score-${r.category}`}>
                            <span className="fss-score-num">{r.score}</span>
                            <span className="fss-score-bar"><span className="fss-score-fill" style={{ width: `${r.score}%` }} /></span>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>

            <div className="fss-footer-bar">
              <span className="fss-count">Showing {visibleRows.length} of {summary.total} homes</span>
              <button className="fss-btn fss-export-btn" onClick={() => downloadCSV(visibleRows)}>
                Export this list as CSV
              </button>
            </div>

            <button className="fss-method-toggle" onClick={() => setShowMethod((v) => !v)}>
              {showMethod ? 'Hide' : 'How is this score calculated?'}
            </button>
            {showMethod && (
              <div className="fss-method-panel">
                Each home starts at 0 and picks up points for the signals that tend to come before a sale:
                <ul>
                  <li>Tax delinquent: +30</li>
                  <li>Vacant: +25</li>
                  <li>Upsize pattern — 1–2 bed, 1 bath, owned 1–6 years: +25</li>
                  <li>Downsize pattern — owned 10–15 years: +15, owned 15+ years: +20</li>
                  <li>Non-owner occupied (absentee owner): +15</li>
                </ul>
                Scores of 60+ are High priority, 30–59 are Medium, and under 30 are Low. This is a
                prioritization tool, not a guarantee — always confirm before reaching out.
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
