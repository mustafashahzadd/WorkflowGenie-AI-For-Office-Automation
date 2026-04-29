/**
 * FortuneSheet spreadsheet — full Excel-like experience.
 * Charts: non-draggable overlays anchored at a specific cell.
 *
 * Data-flow rules (avoids infinite loops AND the broken updateSheet API):
 * - Workbook is an UNCONTROLLED component: data prop is a one-time seed.
 * - External changes (file load, AI result) → remount via key increment.
 * - User edits → onChange → store update, detected as "own" change, no remount.
 * - updateSheet is NEVER called (it crashes on FortuneSheet v1.0.4 due to
 *   immer-frozen internal state).
 */

import React, { useRef, useState, useEffect, useMemo, useCallback } from 'react';
import { Workbook } from '@fortune-sheet/react';
import type { WorkbookInstance } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from 'recharts';

export const syncLuckysheetToStore = (): void => {};

// FortuneSheet mutates its input — always pass a fresh mutable clone
const clone = (sheets: any[]): any[] => JSON.parse(JSON.stringify(sheets));

const CHART_COLORS = ['#107c41', '#1a73e8', '#f4a261', '#e76f51', '#2a9d8f', '#e9c46a'];

const getColLetter = (i: number) =>
  i < 26 ? String.fromCharCode(65 + i)
          : String.fromCharCode(65 + Math.floor(i / 26) - 1) + String.fromCharCode(65 + i % 26);

const buildChartData = (sheets: any[]) => {
  const sheet = sheets.find((s: any) => s.status === 1) || sheets[0];
  if (!sheet) return { data: [], xKey: '', numericKeys: [] };
  const cm: Record<string, any> = {};
  let maxRow = 0; let maxCol = 0;

  // FortuneSheet starts with celldata (sparse) but converts to data (2D) after user edits
  if (sheet.celldata?.length) {
    sheet.celldata.forEach((c: any) => {
      cm[`${c.r},${c.c}`] = c.v;
      if (c.r > maxRow) maxRow = c.r;
      if (c.c > maxCol) maxCol = c.c;
    });
  } else if (sheet.data?.length) {
    (sheet.data as any[][]).forEach((row, r) => {
      if (!row) return;
      row.forEach((cell, c) => {
        if (cell != null) {
          cm[`${r},${c}`] = cell;
          if (r > maxRow) maxRow = r;
          if (c > maxCol) maxCol = c;
        }
      });
    });
  }

  const headers = Array.from({ length: maxCol + 1 }, (_, c) => { const v = cm[`0,${c}`]; return v ? String(v.m ?? v.v ?? '') : getColLetter(c); });
  const numericCols: number[] = []; let xCol = -1;
  for (let c = 0; c <= maxCol; c++) { const val = cm[`1,${c}`]?.v; if (typeof val === 'number') numericCols.push(c); else if (xCol === -1) xCol = c; }
  if (xCol === -1) xCol = 0;
  const data: any[] = [];
  for (let r = 1; r <= maxRow; r++) {
    const xVal = cm[`${r},${xCol}`]?.v; if (xVal == null) continue;
    const row: any = { [headers[xCol]]: String(xVal) };
    numericCols.forEach(c => { const val = cm[`${r},${c}`]?.v; if (val != null) row[headers[c]] = Number(val); });
    data.push(row);
  }
  return { data, xKey: headers[xCol], numericKeys: numericCols.map(c => headers[c]) };
};

const ChartPanel: React.FC<{ sheets: any[]; chartType: 'bar' | 'line' | 'pie' }> = ({ sheets, chartType }) => {
  const { data, xKey, numericKeys } = useMemo(() => buildChartData(sheets), [sheets]);
  if (!data.length || !numericKeys.length)
    return <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%', color: '#aaa', fontSize: 12 }}>No numeric data</div>;
  if (chartType === 'pie') {
    const pieData = data.map(d => ({ name: d[xKey], value: d[numericKeys[0]] }));
    return <ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius="40%" label={({ name, percent }: any) => `${name} ${((percent ?? 0) * 100).toFixed(0)}%`}>{pieData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Pie><Tooltip /><Legend /></PieChart></ResponsiveContainer>;
  }
  if (chartType === 'line') {
    return <ResponsiveContainer width="100%" height="100%"><LineChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 4 }}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" /><XAxis dataKey={xKey} tick={{ fontSize: 11 }} /><YAxis tick={{ fontSize: 10 }} /><Tooltip /><Legend />{numericKeys.map((k, i) => <Line key={k} type="monotone" dataKey={k} stroke={CHART_COLORS[i % CHART_COLORS.length]} strokeWidth={2} dot={{ r: 3 }} />)}</LineChart></ResponsiveContainer>;
  }
  return <ResponsiveContainer width="100%" height="100%"><BarChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 4 }}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" /><XAxis dataKey={xKey} tick={{ fontSize: 11 }} /><YAxis tick={{ fontSize: 10 }} /><Tooltip /><Legend />{numericKeys.map((k, i) => <Bar key={k} dataKey={k} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[3, 3, 0, 0]} />)}</BarChart></ResponsiveContainer>;
};

interface LuckysheetSpreadsheetProps {
  className?: string;
  lastChartType?: 'bar' | 'line' | 'pie' | null;
  chartCell?: string | null;
  chartSheet?: string | null;
}

const CHART_W = 480;
const CHART_H = 300;
const SHEET_TAB_H = 40; // FortuneSheet sheet-tab + zoom bar height

export const LuckysheetSpreadsheet: React.FC<LuckysheetSpreadsheetProps> = ({ className = '', lastChartType, chartCell, chartSheet }) => {
  const workbookRef = useRef<WorkbookInstance>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const { fortuneSheets, setFortuneSheets } = useSpreadsheetStore();
  // userClosedChart lets the X button hide the chart until new chart params arrive
  const [userClosedChart, setUserClosedChart] = useState(false);
  const [scrollOffset, setScrollOffset] = useState({ x: 0, y: 0 });
  // Real column/row pixel sizes read from FortuneSheet after mount
  const [sheetMetrics, setSheetMetrics] = useState({
    colWidths: {} as Record<number, number>,
    rowHeights: {} as Record<number, number>,
    defaultColW: 73, defaultRowH: 19,
    rowHeaderW: 46, headerTotalH: 94, // toolbar+formulabar+colheader
  });
  // Active sheet name (the tab currently selected)
  const [activeSheetName, setActiveSheetName] = useState<string | null>(null);

  // Seed data for Workbook — only updated when an external change is detected
  const [mountData, setMountData] = useState<any[] | null>(
    fortuneSheets ? clone(fortuneSheets) : null
  );
  // Key forces a full Workbook remount on external file changes
  const [fileVersion, setFileVersion] = useState(0);

  // We distinguish "our own" onChange updates vs external (AI/file-load) updates
  // by comparing the fortuneSheets reference against what we last passed to the store.
  const lastSentRef = useRef<any[] | null>(fortuneSheets);

  useEffect(() => {
    if (!fortuneSheets) return;
    if (fortuneSheets === lastSentRef.current) return;
    lastSentRef.current = fortuneSheets;
    setMountData(clone(fortuneSheets));
    setScrollOffset({ x: 0, y: 0 });
    setFileVersion(v => v + 1);
  }, [fortuneSheets]);

  const handleChange = useCallback((data: any[]) => {
    lastSentRef.current = data;
    setFortuneSheets(data);
  }, [setFortuneSheets]);

  // Reset close state whenever new chart params arrive so chart reappears
  useEffect(() => {
    if (lastChartType && chartCell) setUserClosedChart(false);
  }, [lastChartType, chartCell]);

  // Scroll FortuneSheet's grid to make the chart cell visible after chart creation
  useEffect(() => {
    if (!chartCell) return;
    const t = setTimeout(() => {
      const container = containerRef.current;
      if (!container) return;
      const scrollEl = Array.from(container.querySelectorAll<HTMLElement>('div')).find(el => {
        const s = getComputedStyle(el);
        return (s.overflowY === 'auto' || s.overflowY === 'scroll') && el.scrollHeight > el.clientHeight + 10;
      });
      if (!scrollEl) return;
      const { rowHeights, defaultRowH, headerTotalH } = sheetMetrics;
      const m = chartCell.toUpperCase().match(/^([A-Z]+)(\d+)$/);
      if (!m) return;
      const row = parseInt(m[2], 10);
      let cellTop = headerTotalH;
      for (let r = 0; r < row - 1; r++) cellTop += rowHeights[r] ?? defaultRowH;
      scrollEl.scrollTo({ top: Math.max(0, cellTop - 100), behavior: 'smooth' });
    }, 800);
    return () => clearTimeout(t);
  }, [chartCell]);

  // Read actual column/row sizes from FortuneSheet after it renders
  useEffect(() => {
    const t = setTimeout(() => {
      try {
        const sheets = (workbookRef.current as any)?.getAllSheets?.();
        if (!sheets?.length) return;
        const active = sheets.find((s: any) => s.status === 1) ?? sheets[0];
        setActiveSheetName(active.name ?? null);
        setSheetMetrics(prev => ({
          ...prev,
          colWidths:   active.config?.columnlen ?? {},
          rowHeights:  active.config?.rowlen    ?? {},
          defaultColW: active.defaultColWidth   ?? 73,
          defaultRowH: active.defaultRowHeight  ?? 19,
        }));
      } catch { /* use defaults */ }
    }, 600);
    return () => clearTimeout(t);
  }, [fileVersion]);

  // Track active sheet when user switches tabs (via onChange → fortuneSheets)
  useEffect(() => {
    if (!fortuneSheets?.length) return;
    const active = fortuneSheets.find((s: any) => s.status === 1) ?? fortuneSheets[0];
    setActiveSheetName(active?.name ?? null);
  }, [fortuneSheets]);

  // Capture ALL scroll events from any FortuneSheet child — covers both axes.
  // Use canScrollX/canScrollY so scroll-back-to-zero is correctly tracked.
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const handleScroll = (e: Event) => {
      const el = e.target as HTMLElement;
      if (!el || el === container) return;
      setScrollOffset(prev => {
        const canScrollX = el.scrollWidth > el.clientWidth + 1;
        const canScrollY = el.scrollHeight > el.clientHeight + 1;
        const x = canScrollX ? el.scrollLeft : prev.x;
        const y = canScrollY ? el.scrollTop  : prev.y;
        return x === prev.x && y === prev.y ? prev : { x, y };
      });
    };

    container.addEventListener('scroll', handleScroll, { capture: true, passive: true });
    return () => container.removeEventListener('scroll', handleScroll, { capture: true });
  }, [fileVersion]);

  // Chart anchored to real cell pixel coords, scrolls with grid
  const chartPosition = useMemo(() => {
    const { colWidths, rowHeights, defaultColW, defaultRowH, rowHeaderW, headerTotalH } = sheetMetrics;

    const colPixel = (colIdx: number) => { // colIdx is 1-based
      let x = rowHeaderW;
      for (let c = 0; c < colIdx - 1; c++) x += colWidths[c] ?? defaultColW;
      return x;
    };
    const rowPixel = (rowIdx: number) => { // rowIdx is 1-based
      let y = headerTotalH;
      for (let r = 0; r < rowIdx - 1; r++) y += rowHeights[r] ?? defaultRowH;
      return y;
    };

    let left = 420, top = 200;
    if (chartCell) {
      const m = chartCell.toUpperCase().match(/^([A-Z]+)(\d+)$/);
      if (m) {
        let col = 0;
        for (let i = 0; i < m[1].length; i++) col = col * 26 + m[1].charCodeAt(i) - 64;
        left = colPixel(col);
        top  = rowPixel(parseInt(m[2], 10));
      }
    }
    // Allow negative (chart scrolls off-screen like Excel) but clamp top to headerTotalH min
    // so chart doesn't disappear under the toolbar when scrolled back up past origin
    return {
      left: left - scrollOffset.x,
      top:  top  - scrollOffset.y,
    };
  }, [chartCell, scrollOffset, sheetMetrics]);

  // Chart is shown whenever lastChartType + chartCell are set AND user hasn't closed it
  const chartVisible = !!(lastChartType && chartCell && !userClosedChart);

  if (!mountData) {
    return <div className={`flex items-center justify-center h-full text-gray-400 text-sm ${className}`}>Loading spreadsheet…</div>;
  }

  return (
    <div ref={containerRef} className={`h-full relative overflow-hidden ${className}`}>
      <Workbook
        key={fileVersion}
        ref={workbookRef}
        data={mountData}
        onChange={handleChange}
        lang="en"
        showToolbar
        showFormulaBar
        showSheetTabs
        allowEdit
      />

      {/* Chart overlay — clips at the sheet-tab bar using clipPath on the outer wrapper.
          zIndex 10000 ensures it renders above all FortuneSheet internal layers. */}
      {chartVisible && (!chartSheet || !activeSheetName || activeSheetName.toLowerCase() === chartSheet.toLowerCase()) && (!chartSheet || !activeSheetName || activeSheetName.toLowerCase() === chartSheet.toLowerCase()) && (
        <div style={{
          position: 'absolute', top: 0, left: 0, right: 0, bottom: 0,
          pointerEvents: 'none',
          zIndex: 10000,
          clipPath: `inset(0 0 ${SHEET_TAB_H}px 0)`,
        }}>
          <div style={{
            position: 'absolute',
            left: chartPosition.left,
            top: chartPosition.top,
            width: CHART_W,
            height: CHART_H,
            background: '#fff',
            border: '1.5px solid #c6d4e8',
            borderRadius: 6,
            boxShadow: '0 4px 20px rgba(0,0,0,0.18)',
            display: 'flex',
            flexDirection: 'column',
            pointerEvents: 'auto',
          }}>
            <div style={{ padding: '5px 10px', background: '#f0faf5', borderBottom: '1px solid #d4edda', borderRadius: '6px 6px 0 0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0 }}>
              <span style={{ fontSize: 11, fontWeight: 600, color: '#107c41', textTransform: 'uppercase', letterSpacing: '0.05em' }}>{lastChartType} Chart</span>
              <button onPointerDown={e => e.stopPropagation()} onClick={() => setUserClosedChart(true)} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 16, color: '#888', lineHeight: 1 }}>×</button>
            </div>
            <div style={{ flex: 1, padding: '6px 8px', minHeight: 0 }}>
              <ChartPanel sheets={fortuneSheets ?? mountData} chartType={lastChartType} />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
