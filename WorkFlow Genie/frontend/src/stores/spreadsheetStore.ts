/**
 * Zustand store for managing spreadsheet data and application state
 * This store handles workbooks, worksheets, cells, and AI operations
 */

import { create } from 'zustand';
import { devtools } from 'zustand/middleware';
import { v4 as uuidv4 } from 'uuid';
import type { 
  Workbook, 
  Worksheet, 
  CellValue, 
  CellPosition, 
  CellRange, 
  AIOperation, 
  AIPrompt,
  UIState,
  CellFormat,
  FindOptions,
  FilterCriteria,
  SortSettings,
  DataValidation,
  CellComment,
  Chart,
  ConditionalFormat,
  HistoryEntry
} from '../types/index';

interface SpreadsheetState {
  // FortuneSheet data (primary spreadsheet state)
  fortuneSheets: any[] | null;
  setFortuneSheets: (sheets: any[]) => void;
  loadFortuneSheets: (sheets: any[]) => void;

  // Workbook data
  workbooks: Workbook[];
  currentWorkbook: Workbook | null;
  currentWorksheet: Worksheet | null;
  
  // UI state
  ui: UIState;
  
  // History
  history: HistoryEntry[];
  historyIndex: number;
  
  // AI operations
  aiPrompts: AIPrompt[];
  currentOperations: AIOperation[];
  
  // Actions for workbook management
  createWorkbook: (name: string) => void;
  loadWorkbook: (workbook: Workbook) => void;
  updateWorkbookName: (name: string) => void;
  setActiveWorksheet: (worksheetId: string) => void;
  
  // Actions for worksheet management
  createWorksheet: (name: string) => void;
  renameWorksheet: (worksheetId: string, name: string) => void;
  deleteWorksheet: (worksheetId: string) => void;
  duplicateWorksheet: (worksheetId: string) => void;
  
  // Actions for cell management
  setCellValue: (position: CellPosition, value: CellValue) => void;
  getCellValue: (position: CellPosition) => CellValue | null;
  setCellRange: (range: CellRange, values: CellValue[][]) => void;
  setSelectedRange: (range: CellRange | null) => void;
  syncWorksheetFromLuckysheet: (cells: Record<string, CellValue>) => void;
  
  // Cell formatting
  setCellFormat: (position: CellPosition, format: Partial<CellFormat>) => void;
  setRangeFormat: (range: CellRange, format: Partial<CellFormat>) => void;
  
  // Clipboard operations
  copy: (range?: CellRange) => void;
  cut: (range?: CellRange) => void;
  paste: (position: CellPosition, pasteSpecial?: 'values' | 'formats' | 'formulas') => void;
  
  // Row & Column operations
  insertRow: (index: number, count?: number) => void;
  deleteRow: (index: number, count?: number) => void;
  insertColumn: (index: number, count?: number) => void;
  deleteColumn: (index: number, count?: number) => void;
  setColumnWidth: (index: number, width: number) => void;
  setRowHeight: (index: number, height: number) => void;
  hideColumn: (index: number) => void;
  unhideColumn: (index: number) => void;
  hideRow: (index: number) => void;
  unhideRow: (index: number) => void;
  
  // Merge cells
  mergeCells: (range: CellRange) => void;
  unmergeCells: (range: CellRange) => void;
  
  // Find & Replace
  find: (searchText: string, options?: FindOptions) => CellPosition[];
  replace: (searchText: string, replaceText: string, options?: FindOptions) => number;
  
  // Sorting & Filtering
  sortRange: (range: CellRange, sortSettings: SortSettings[]) => void;
  setFilter: (range: CellRange, filters: Record<number, FilterCriteria>) => void;
  clearFilter: () => void;
  
  // Freeze panes
  freezePanes: (rows: number, columns: number) => void;
  unfreezePanes: () => void;
  
  // Named ranges
  addNamedRange: (name: string, range: CellRange) => void;
  deleteNamedRange: (name: string) => void;
  
  // Data validation
  setDataValidation: (range: CellRange, validation: DataValidation) => void;
  removeDataValidation: (range: CellRange) => void;
  
  // Comments
  addComment: (position: CellPosition, comment: CellComment) => void;
  deleteComment: (position: CellPosition) => void;
  
  // Charts
  addChart: (chart: Chart) => void;
  updateChart: (chartId: string, updates: Partial<Chart>) => void;
  deleteChart: (chartId: string) => void;
  
  // Conditional formatting
  addConditionalFormat: (range: CellRange, format: ConditionalFormat) => void;
  removeConditionalFormat: (range: CellRange) => void;
  
  // Undo/Redo
  undo: () => void;
  redo: () => void;
  addToHistory: (entry: Omit<HistoryEntry, 'id' | 'timestamp'>) => void;
  
  // Actions for AI operations
  addAIPrompt: (message: string) => Promise<void>;
  updateOperation: (operationId: string, updates: Partial<AIOperation>) => void;
  completeOperation: (operationId: string, data?: any) => void;
  failOperation: (operationId: string, error: string) => void;
  
  // UI actions
  toggleSidebar: () => void;
  setActiveTab: (tab: UIState['activeTab']) => void;
  setProcessingState: (isProcessing: boolean) => void;
  setZoom: (zoom: number) => void;
}

/**
 * Creates a default empty worksheet
 */
const createEmptyWorksheet = (name: string): Worksheet => ({
  id: uuidv4(),
  name,
  cells: {},
  rowCount: 1000,
  colCount: 26,
  createdAt: new Date(),
  updatedAt: new Date(),
});

/**
 * Creates a default workbook with one worksheet
 */
const createDefaultWorkbook = (name: string): Workbook => {
  const worksheet = createEmptyWorksheet('Sheet1');
  return {
    id: uuidv4(),
    name,
    worksheets: [worksheet],
    activeWorksheetId: worksheet.id,
    createdAt: new Date(),
    updatedAt: new Date(),
  };
};

/**
 * Converts cell position to string key for storage
 */
const positionToKey = (position: CellPosition): string => 
  `${position.row},${position.col}`;

/**
 * Main Zustand store for the application
 */
export const useSpreadsheetStore = create<SpreadsheetState>()(
  devtools(
    (set, get) => ({
      // Initial state
      fortuneSheets: null,
      workbooks: [],
      currentWorkbook: null,
      currentWorksheet: null,
      
      history: [],
      historyIndex: -1,
      
      ui: {
        selectedCells: null,
        isAIProcessing: false,
        sidebarOpen: false,
        activeTab: 'prompt',
        theme: 'light',
        showFormulaBar: true,
        zoom: 100,
      },
      
      aiPrompts: [],
      currentOperations: [],
      
      // FortuneSheet actions
      setFortuneSheets: (sheets: any[]) => {
        set({ fortuneSheets: sheets });
      },

      loadFortuneSheets: (sheets: any[]) => {
        set({ fortuneSheets: sheets });
      },

      // Workbook actions
      createWorkbook: (name: string) => {
        const newWorkbook = createDefaultWorkbook(name);
        set((state) => ({
          workbooks: [...state.workbooks, newWorkbook],
          currentWorkbook: newWorkbook,
          currentWorksheet: newWorkbook.worksheets[0],
        }));
      },
      
      loadWorkbook: (workbook: Workbook) => {
        set({
          currentWorkbook: workbook,
          currentWorksheet: workbook.worksheets.find(
            (ws) => ws.id === workbook.activeWorksheetId
          ) || workbook.worksheets[0],
        });
      },

      updateWorkbookName: (name: string) => {
        set((state) => {
          if (!state.currentWorkbook) return state;
          return {
            currentWorkbook: {
              ...state.currentWorkbook,
              name,
              updatedAt: new Date(),
            },
          };
        });
      },
      
      setActiveWorksheet: (worksheetId: string) => {
        const { currentWorkbook } = get();
        if (!currentWorkbook) return;
        
        const worksheet = currentWorkbook.worksheets.find(
          (ws) => ws.id === worksheetId
        );
        if (worksheet) {
          set((state) => ({
            currentWorksheet: worksheet,
            currentWorkbook: {
              ...state.currentWorkbook!,
              activeWorksheetId: worksheetId,
            },
          }));
        }
      },
      
      // Worksheet actions
      createWorksheet: (name: string) => {
        const { currentWorkbook } = get();
        if (!currentWorkbook) return;
        
        const newWorksheet = createEmptyWorksheet(name);
        set((state) => ({
          currentWorkbook: {
            ...state.currentWorkbook!,
            worksheets: [...state.currentWorkbook!.worksheets, newWorksheet],
            activeWorksheetId: newWorksheet.id,
          },
          currentWorksheet: newWorksheet,
        }));
      },
      
      renameWorksheet: (worksheetId: string, name: string) => {
        set((state) => {
          if (!state.currentWorkbook) return state;
          
          const updatedWorksheets = state.currentWorkbook.worksheets.map(
            (ws) => (ws.id === worksheetId ? { ...ws, name, updatedAt: new Date() } : ws)
          );
          
          return {
            currentWorkbook: {
              ...state.currentWorkbook,
              worksheets: updatedWorksheets,
            },
            currentWorksheet: state.currentWorksheet?.id === worksheetId
              ? { ...state.currentWorksheet, name, updatedAt: new Date() }
              : state.currentWorksheet,
          };
        });
      },
      
      deleteWorksheet: (worksheetId: string) => {
        set((state) => {
          if (!state.currentWorkbook || state.currentWorkbook.worksheets.length <= 1) {
            return state; // Don't delete the last worksheet
          }
          
          const remainingWorksheets = state.currentWorkbook.worksheets.filter(
            (ws) => ws.id !== worksheetId
          );
          
          const newActiveId = state.currentWorkbook.activeWorksheetId === worksheetId
            ? remainingWorksheets[0].id
            : state.currentWorkbook.activeWorksheetId;
          
          return {
            currentWorkbook: {
              ...state.currentWorkbook,
              worksheets: remainingWorksheets,
              activeWorksheetId: newActiveId,
            },
            currentWorksheet: remainingWorksheets.find(ws => ws.id === newActiveId) || remainingWorksheets[0],
          };
        });
      },
      
      duplicateWorksheet: (worksheetId: string) => {
        const { currentWorkbook } = get();
        if (!currentWorkbook) return;
        
        const worksheet = currentWorkbook.worksheets.find(ws => ws.id === worksheetId);
        if (!worksheet) return;
        
        const newWorksheet: Worksheet = {
          ...worksheet,
          id: uuidv4(),
          name: `${worksheet.name} (Copy)`,
          cells: { ...worksheet.cells },
          createdAt: new Date(),
          updatedAt: new Date(),
        };
        
        set((state) => ({
          currentWorkbook: {
            ...state.currentWorkbook!,
            worksheets: [...state.currentWorkbook!.worksheets, newWorksheet],
          },
        }));
      },
      
      // Cell actions
      setCellValue: (position: CellPosition, value: CellValue) => {
        set((state) => {
          if (!state.currentWorksheet) return state;
          
          const key = positionToKey(position);
          const updatedCells = { ...state.currentWorksheet.cells };
          
          if (value.value === null || value.value === '') {
            delete updatedCells[key];
          } else {
            updatedCells[key] = value;
          }
          
          const updatedWorksheet = {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          };
          
          return {
            currentWorksheet: updatedWorksheet,
            currentWorkbook: state.currentWorkbook ? {
              ...state.currentWorkbook,
              worksheets: state.currentWorkbook.worksheets.map(
                (ws) => (ws.id === updatedWorksheet.id ? updatedWorksheet : ws)
              ),
            } : null,
          };
        });
      },
      
      getCellValue: (position: CellPosition) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return null;
        
        const key = positionToKey(position);
        return currentWorksheet.cells[key] || null;
      },
      
      setCellRange: (range: CellRange, values: CellValue[][]) => {
        set((state) => {
          if (!state.currentWorksheet) return state;
          
          const updatedCells = { ...state.currentWorksheet.cells };
          
          for (let row = range.startRow; row <= range.endRow; row++) {
            for (let col = range.startCol; col <= range.endCol; col++) {
              const valueRow = row - range.startRow;
              const valueCol = col - range.startCol;
              
              if (values[valueRow] && values[valueRow][valueCol]) {
                const key = positionToKey({ row, col });
                updatedCells[key] = values[valueRow][valueCol];
              }
            }
          }
          
          const updatedWorksheet = {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          };
          
          return {
            currentWorksheet: updatedWorksheet,
            currentWorkbook: state.currentWorkbook ? {
              ...state.currentWorkbook,
              worksheets: state.currentWorkbook.worksheets.map(
                (ws) => (ws.id === updatedWorksheet.id ? updatedWorksheet : ws)
              ),
            } : null,
          };
        });
      },
      
      setSelectedRange: (range: CellRange | null) => {
        set((state) => ({
          ui: { ...state.ui, selectedCells: range },
        }));
      },

      syncWorksheetFromLuckysheet: (cells: Record<string, CellValue>) => {
        set((state) => {
          if (!state.currentWorksheet) return state;

          const updatedWorksheet = {
            ...state.currentWorksheet,
            cells,
            updatedAt: new Date(),
          };

          return {
            currentWorksheet: updatedWorksheet,
            currentWorkbook: state.currentWorkbook ? {
              ...state.currentWorkbook,
              worksheets: state.currentWorkbook.worksheets.map(
                (ws) => (ws.id === updatedWorksheet.id ? updatedWorksheet : ws)
              ),
            } : null,
          };
        });
      },
      
      // Cell formatting
      setCellFormat: (position: CellPosition, format: Partial<CellFormat>) => {
        set((state) => {
          if (!state.currentWorksheet) return state;
          
          const key = positionToKey(position);
          const cell = state.currentWorksheet.cells[key] || { value: null };
          
          const updatedCells = {
            ...state.currentWorksheet.cells,
            [key]: {
              ...cell,
              format: { ...cell.format, ...format },
            },
          };
          
          const updatedWorksheet = {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          };
          
          return {
            currentWorksheet: updatedWorksheet,
            currentWorkbook: state.currentWorkbook ? {
              ...state.currentWorkbook,
              worksheets: state.currentWorkbook.worksheets.map(
                (ws) => (ws.id === updatedWorksheet.id ? updatedWorksheet : ws)
              ),
            } : null,
          };
        });
      },
      
      setRangeFormat: (range: CellRange, format: Partial<CellFormat>) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            get().setCellFormat({ row, col }, format);
          }
        }
      },
      
      // Clipboard operations
      copy: (range?: CellRange) => {
        const { currentWorksheet, ui } = get();
        if (!currentWorksheet) return;
        
        const copyRange = range || ui.selectedCells;
        if (!copyRange) return;
        
        const cells: Record<string, CellValue> = {};
        for (let row = copyRange.startRow; row <= copyRange.endRow; row++) {
          for (let col = copyRange.startCol; col <= copyRange.endCol; col++) {
            const key = positionToKey({ row, col });
            if (currentWorksheet.cells[key]) {
              cells[key] = { ...currentWorksheet.cells[key] };
            }
          }
        }
        
        set((state) => ({
          ui: {
            ...state.ui,
            clipboard: {
              type: 'copy',
              range: copyRange,
              cells,
              timestamp: new Date(),
            },
          },
        }));
      },
      
      cut: (range?: CellRange) => {
        const { currentWorksheet, ui } = get();
        if (!currentWorksheet) return;
        
        const cutRange = range || ui.selectedCells;
        if (!cutRange) return;
        
        const cells: Record<string, CellValue> = {};
        const updatedCells = { ...currentWorksheet.cells };
        
        for (let row = cutRange.startRow; row <= cutRange.endRow; row++) {
          for (let col = cutRange.startCol; col <= cutRange.endCol; col++) {
            const key = positionToKey({ row, col });
            if (currentWorksheet.cells[key]) {
              cells[key] = { ...currentWorksheet.cells[key] };
              delete updatedCells[key];
            }
          }
        }
        
        set((state) => ({
          ui: {
            ...state.ui,
            clipboard: {
              type: 'cut',
              range: cutRange,
              cells,
              timestamp: new Date(),
            },
          },
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      paste: (position: CellPosition, pasteSpecial?: 'values' | 'formats' | 'formulas') => {
        const { currentWorksheet, ui } = get();
        if (!currentWorksheet || !ui.clipboard) return;
        
        const { cells, range } = ui.clipboard;
        const updatedCells = { ...currentWorksheet.cells };
        
        const rowOffset = position.row - range.startRow;
        const colOffset = position.col - range.startCol;
        
        Object.entries(cells).forEach(([key, cell]) => {
          const [row, col] = key.split(',').map(Number);
          const newRow = row + rowOffset;
          const newCol = col + colOffset;
          const newKey = positionToKey({ row: newRow, col: newCol });
          
          if (pasteSpecial === 'values') {
            updatedCells[newKey] = { value: cell.value };
          } else if (pasteSpecial === 'formats') {
            const existing = updatedCells[newKey] || { value: null };
            updatedCells[newKey] = { ...existing, format: cell.format };
          } else if (pasteSpecial === 'formulas') {
            updatedCells[newKey] = { value: cell.value, formula: cell.formula };
          } else {
            updatedCells[newKey] = { ...cell };
          }
        });
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Row & Column operations
      insertRow: (index: number, count: number = 1) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const updatedCells: Record<string, CellValue> = {};
        
        Object.entries(currentWorksheet.cells).forEach(([key, cell]) => {
          const [row, col] = key.split(',').map(Number);
          if (row >= index) {
            const newKey = positionToKey({ row: row + count, col });
            updatedCells[newKey] = cell;
          } else {
            updatedCells[key] = cell;
          }
        });
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            rowCount: state.currentWorksheet.rowCount + count,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      deleteRow: (index: number, count: number = 1) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const updatedCells: Record<string, CellValue> = {};
        
        Object.entries(currentWorksheet.cells).forEach(([key, cell]) => {
          const [row, col] = key.split(',').map(Number);
          if (row < index) {
            updatedCells[key] = cell;
          } else if (row >= index + count) {
            const newKey = positionToKey({ row: row - count, col });
            updatedCells[newKey] = cell;
          }
          // Cells in the deleted range are not included
        });
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            rowCount: Math.max(state.currentWorksheet.rowCount - count, 1),
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      insertColumn: (index: number, count: number = 1) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const updatedCells: Record<string, CellValue> = {};
        
        Object.entries(currentWorksheet.cells).forEach(([key, cell]) => {
          const [row, col] = key.split(',').map(Number);
          if (col >= index) {
            const newKey = positionToKey({ row, col: col + count });
            updatedCells[newKey] = cell;
          } else {
            updatedCells[key] = cell;
          }
        });
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            colCount: state.currentWorksheet.colCount + count,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      deleteColumn: (index: number, count: number = 1) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const updatedCells: Record<string, CellValue> = {};
        
        Object.entries(currentWorksheet.cells).forEach(([key, cell]) => {
          const [row, col] = key.split(',').map(Number);
          if (col < index) {
            updatedCells[key] = cell;
          } else if (col >= index + count) {
            const newKey = positionToKey({ row, col: col - count });
            updatedCells[newKey] = cell;
          }
        });
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            colCount: Math.max(state.currentWorksheet.colCount - count, 1),
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      setColumnWidth: (index: number, width: number) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            columnWidths: {
              ...state.currentWorksheet.columnWidths,
              [index]: width,
            },
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      setRowHeight: (index: number, height: number) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            rowHeights: {
              ...state.currentWorksheet.rowHeights,
              [index]: height,
            },
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      hideColumn: (index: number) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            hiddenColumns: new Set([...(state.currentWorksheet.hiddenColumns || []), index]),
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      unhideColumn: (index: number) => {
        set((state) => {
          if (!state.currentWorksheet) return state;
          const hidden = new Set(state.currentWorksheet.hiddenColumns || []);
          hidden.delete(index);
          return {
            currentWorksheet: {
              ...state.currentWorksheet,
              hiddenColumns: hidden,
              updatedAt: new Date(),
            },
          };
        });
      },
      
      hideRow: (index: number) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            hiddenRows: new Set([...(state.currentWorksheet.hiddenRows || []), index]),
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      unhideRow: (index: number) => {
        set((state) => {
          if (!state.currentWorksheet) return state;
          const hidden = new Set(state.currentWorksheet.hiddenRows || []);
          hidden.delete(index);
          return {
            currentWorksheet: {
              ...state.currentWorksheet,
              hiddenRows: hidden,
              updatedAt: new Date(),
            },
          };
        });
      },
      
      // Merge cells
      mergeCells: (range: CellRange) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const topLeft = { row: range.startRow, col: range.startCol };
        const topLeftKey = positionToKey(topLeft);
        const topLeftCell = currentWorksheet.cells[topLeftKey] || { value: null };
        
        const updatedCells = { ...currentWorksheet.cells };
        updatedCells[topLeftKey] = {
          ...topLeftCell,
          isMerged: true,
          mergeRange: range,
        };
        
        // Mark other cells in range as merged
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            if (row !== topLeft.row || col !== topLeft.col) {
              const key = positionToKey({ row, col });
              updatedCells[key] = {
                value: null,
                mergedWith: topLeft,
              };
            }
          }
        }
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      unmergeCells: (range: CellRange) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const updatedCells = { ...currentWorksheet.cells };
        
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            const key = positionToKey({ row, col });
            const cell = updatedCells[key];
            if (cell) {
              delete cell.isMerged;
              delete cell.mergeRange;
              delete cell.mergedWith;
            }
          }
        }
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Find & Replace
      find: (searchText: string, options?: FindOptions) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return [];
        
        const results: CellPosition[] = [];
        const search = options?.matchCase ? searchText : searchText.toLowerCase();
        
        Object.entries(currentWorksheet.cells).forEach(([key, cell]) => {
          const [row, col] = key.split(',').map(Number);
          let cellText = '';
          
          if (options?.searchFormulas && cell.formula) {
            cellText = cell.formula;
          } else {
            cellText = String(cell.value || '');
          }
          
          if (!options?.matchCase) {
            cellText = cellText.toLowerCase();
          }
          
          const matches = options?.matchEntireCell 
            ? cellText === search
            : cellText.includes(search);
          
          if (matches) {
            results.push({ row, col });
          }
        });
        
        return results;
      },
      
      replace: (searchText: string, replaceText: string, options?: FindOptions) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return 0;
        
        const positions = get().find(searchText, options);
        const updatedCells = { ...currentWorksheet.cells };
        
        positions.forEach(pos => {
          const key = positionToKey(pos);
          const cell = updatedCells[key];
          if (!cell) return;
          
          if (options?.searchFormulas && cell.formula) {
            cell.formula = cell.formula.replace(new RegExp(searchText, options.matchCase ? 'g' : 'gi'), replaceText);
          } else {
            const oldValue = String(cell.value || '');
            cell.value = oldValue.replace(new RegExp(searchText, options?.matchCase ? 'g' : 'gi'), replaceText);
          }
        });
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
        
        return positions.length;
      },
      
      // Sorting & Filtering
      sortRange: (range: CellRange, sortSettings: SortSettings[]) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        // Extract rows in range
        const rows: CellValue[][] = [];
        for (let row = range.startRow; row <= range.endRow; row++) {
          const rowData: CellValue[] = [];
          for (let col = range.startCol; col <= range.endCol; col++) {
            const key = positionToKey({ row, col });
            rowData.push(currentWorksheet.cells[key] || { value: null });
          }
          rows.push(rowData);
        }
        
        // Sort rows
        rows.sort((a, b) => {
          for (const setting of sortSettings) {
            const aVal = a[setting.column]?.value;
            const bVal = b[setting.column]?.value;
            
            let comparison = 0;
            if (aVal !== null && bVal !== null) {
              if (aVal < bVal) comparison = -1;
              else if (aVal > bVal) comparison = 1;
            } else if (aVal === null && bVal !== null) {
              comparison = -1;
            } else if (aVal !== null && bVal === null) {
              comparison = 1;
            }
            
            if (comparison !== 0) {
              return setting.ascending ? comparison : -comparison;
            }
          }
          return 0;
        });
        
        // Write back sorted data
        const updatedCells = { ...currentWorksheet.cells };
        rows.forEach((rowData, rowIdx) => {
          rowData.forEach((cell, colIdx) => {
            const key = positionToKey({ 
              row: range.startRow + rowIdx, 
              col: range.startCol + colIdx 
            });
            updatedCells[key] = cell;
          });
        });
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      setFilter: (range: CellRange, filters: Record<number, FilterCriteria>) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            filters: {
              range,
              filters,
              active: true,
            },
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      clearFilter: () => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            filters: undefined,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Freeze panes
      freezePanes: (rows: number, columns: number) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            frozenRows: rows,
            frozenColumns: columns,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      unfreezePanes: () => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            frozenRows: undefined,
            frozenColumns: undefined,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Named ranges
      addNamedRange: (name: string, range: CellRange) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            namedRanges: {
              ...state.currentWorksheet.namedRanges,
              [name]: range,
            },
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      deleteNamedRange: (name: string) => {
        set((state) => {
          if (!state.currentWorksheet?.namedRanges) return state;
          const { [name]: _, ...rest } = state.currentWorksheet.namedRanges;
          return {
            currentWorksheet: {
              ...state.currentWorksheet,
              namedRanges: rest,
              updatedAt: new Date(),
            },
          };
        });
      },
      
      // Data validation
      setDataValidation: (range: CellRange, validation: DataValidation) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const updatedCells = { ...currentWorksheet.cells };
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            const key = positionToKey({ row, col });
            const cell = updatedCells[key] || { value: null };
            updatedCells[key] = { ...cell, validation };
          }
        }
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      removeDataValidation: (range: CellRange) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const updatedCells = { ...currentWorksheet.cells };
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            const key = positionToKey({ row, col });
            const cell = updatedCells[key];
            if (cell) {
              delete cell.validation;
            }
          }
        }
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: updatedCells,
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Comments
      addComment: (position: CellPosition, comment: CellComment) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const key = positionToKey(position);
        const cell = currentWorksheet.cells[key] || { value: null };
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: {
              ...state.currentWorksheet.cells,
              [key]: { ...cell, comment },
            },
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      deleteComment: (position: CellPosition) => {
        const { currentWorksheet } = get();
        if (!currentWorksheet) return;
        
        const key = positionToKey(position);
        const cell = currentWorksheet.cells[key];
        if (!cell) return;
        
        delete cell.comment;
        
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            cells: {
              ...state.currentWorksheet.cells,
              [key]: cell,
            },
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Charts
      addChart: (chart: Chart) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            charts: [...(state.currentWorksheet.charts || []), chart],
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      updateChart: (chartId: string, updates: Partial<Chart>) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            charts: (state.currentWorksheet.charts || []).map(chart =>
              chart.id === chartId ? { ...chart, ...updates } : chart
            ),
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      deleteChart: (chartId: string) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            charts: (state.currentWorksheet.charts || []).filter(chart => chart.id !== chartId),
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Conditional formatting
      addConditionalFormat: (range: CellRange, format: ConditionalFormat) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            conditionalFormats: [...(state.currentWorksheet.conditionalFormats || []), { ...format, rules: format.rules.map(r => ({ ...r, range })) }],
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      removeConditionalFormat: (range: CellRange) => {
        set((state) => ({
          currentWorksheet: state.currentWorksheet ? {
            ...state.currentWorksheet,
            conditionalFormats: (state.currentWorksheet.conditionalFormats || []).filter(cf =>
              !cf.rules.some(r => JSON.stringify(r) === JSON.stringify({ ...r, range }))
            ),
            updatedAt: new Date(),
          } : null,
        }));
      },
      
      // Undo/Redo
      undo: () => {
        const { history, historyIndex } = get();
        if (historyIndex < 0) return;
        
        const entry = history[historyIndex];
        entry.undo();
        
        set({ historyIndex: historyIndex - 1 });
      },
      
      redo: () => {
        const { history, historyIndex } = get();
        if (historyIndex >= history.length - 1) return;
        
        const entry = history[historyIndex + 1];
        entry.redo();
        
        set({ historyIndex: historyIndex + 1 });
      },
      
      addToHistory: (entry: Omit<HistoryEntry, 'id' | 'timestamp'>) => {
        const { history, historyIndex } = get();
        
        const newEntry: HistoryEntry = {
          ...entry,
          id: uuidv4(),
          timestamp: new Date(),
        };
        
        // Remove any entries after current index
        const newHistory = history.slice(0, historyIndex + 1);
        newHistory.push(newEntry);
        
        // Limit history to 100 entries
        if (newHistory.length > 100) {
          newHistory.shift();
        }
        
        set({
          history: newHistory,
          historyIndex: newHistory.length - 1,
        });
      },
      
      // AI operation actions
      addAIPrompt: async (message: string) => {
        const promptId = uuidv4();
        const newPrompt: AIPrompt = {
          id: promptId,
          message,
          timestamp: new Date(),
          operations: [],
          status: 'processing',
        };
        
        set((state) => ({
          aiPrompts: [newPrompt, ...state.aiPrompts],
          ui: { ...state.ui, isAIProcessing: true },
        }));
        
        // TODO: Integrate with AI service
        // This is where you would call your backend AI service
        // For now, we'll simulate with a timeout
        setTimeout(() => {
          const mockOperations: AIOperation[] = [
            {
              id: uuidv4(),
              type: 'populate_data',
              description: `Processing: ${message}`,
              status: 'in-progress',
              timestamp: new Date(),
              worksheetId: get().currentWorksheet?.id,
            },
          ];
          
          set((state) => ({
            aiPrompts: state.aiPrompts.map((prompt) =>
              prompt.id === promptId
                ? { ...prompt, operations: mockOperations, status: 'completed' }
                : prompt
            ),
            currentOperations: [...state.currentOperations, ...mockOperations],
            ui: { ...state.ui, isAIProcessing: false },
          }));
        }, 2000);
      },
      
      updateOperation: (operationId: string, updates: Partial<AIOperation>) => {
        set((state) => ({
          currentOperations: state.currentOperations.map((op) =>
            op.id === operationId ? { ...op, ...updates } : op
          ),
        }));
      },
      
      completeOperation: (operationId: string, data?: any) => {
        set((state) => ({
          currentOperations: state.currentOperations.map((op) =>
            op.id === operationId
              ? { ...op, status: 'completed', data }
              : op
          ),
        }));
      },
      
      failOperation: (operationId: string, error: string) => {
        set((state) => ({
          currentOperations: state.currentOperations.map((op) =>
            op.id === operationId
              ? { ...op, status: 'failed', error }
              : op
          ),
        }));
      },
      
      // UI actions
      toggleSidebar: () => {
        set((state) => ({
          ui: { ...state.ui, sidebarOpen: !state.ui.sidebarOpen },
        }));
      },
      
      setActiveTab: (tab: UIState['activeTab']) => {
        set((state) => ({
          ui: { ...state.ui, activeTab: tab },
        }));
      },
      
      setProcessingState: (isProcessing: boolean) => {
        set((state) => ({
          ui: { ...state.ui, isAIProcessing: isProcessing },
        }));
      },
      
      setZoom: (zoom: number) => {
        set((state) => ({
          ui: { ...state.ui, zoom: Math.max(10, Math.min(400, zoom)) },
        }));
      },
    }),
    {
      name: 'excelerate-store',
    }
  )
);