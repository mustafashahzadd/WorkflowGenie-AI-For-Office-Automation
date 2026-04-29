/**
 * Core type definitions for the Excelerate application
 * These types define the structure of spreadsheet data, AI operations, and application state
 */

// Spreadsheet Cell Types
export interface CellValue {
  value: string | number | boolean | null;
  formula?: string;
  format?: CellFormat;
  comment?: CellComment;
  validation?: DataValidation;
  hyperlink?: string;
  mergedWith?: CellPosition; // Points to the top-left cell of a merged range
  isMerged?: boolean; // True if this is the top-left cell of a merge
  mergeRange?: CellRange; // The range this cell is merged across
}

export interface CellFormat {
  // Font properties
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  
  // Background
  backgroundColor?: string;
  pattern?: string;
  
  // Borders
  border?: BorderStyle;
  
  // Alignment
  horizontalAlign?: 'left' | 'center' | 'right' | 'justify' | 'fill' | 'centerContinuous' | 'distributed';
  verticalAlign?: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
  textRotation?: number;
  indent?: number;
  
  // Number formatting
  numberFormat?: string;
  
  // Conditional formatting
  conditionalFormat?: ConditionalFormat;
}

export interface BorderStyle {
  top?: string;
  bottom?: string;
  left?: string;
  right?: string;
  style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
  color?: string;
}

export interface ConditionalFormat {
  type: 'cellValue' | 'colorScale' | 'dataBar' | 'iconSet' | 'formula';
  rules: ConditionalRule[];
}

export interface ConditionalRule {
  condition: string; // e.g., '>100', 'between 0 and 100'
  format: Partial<CellFormat>;
  priority?: number;
}

export interface CellComment {
  text: string;
  author: string;
  timestamp: Date;
  resolved?: boolean;
}

export interface DataValidation {
  type: 'list' | 'number' | 'date' | 'text' | 'custom';
  criteria: any;
  allowBlank?: boolean;
  showDropdown?: boolean;
  errorMessage?: string;
}

export interface FindOptions {
  matchCase?: boolean;
  matchEntireCell?: boolean;
  searchFormulas?: boolean;
  direction?: 'forward' | 'backward';
}

export interface CellPosition {
  row: number;
  col: number;
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

// Worksheet Types
export interface Worksheet {
  id: string;
  name: string;
  cells: Record<string, CellValue>; // Key format: "row,col"
  rowCount: number;
  colCount: number;
  columnWidths?: Record<number, number>; // Custom column widths
  rowHeights?: Record<number, number>; // Custom row heights
  hiddenColumns?: Set<number>;
  hiddenRows?: Set<number>;
  frozenRows?: number; // Number of rows frozen at top
  frozenColumns?: number; // Number of columns frozen at left
  selectedRange?: CellRange;
  namedRanges?: Record<string, CellRange>;
  filters?: FilterSettings;
  sorts?: SortSettings[];
  conditionalFormats?: ConditionalFormat[];
  charts?: Chart[];
  pivotTables?: PivotTable[];
  printSettings?: PrintSettings;
  protection?: WorksheetProtection;
  createdAt: Date;
  updatedAt: Date;
}

export interface FilterSettings {
  range: CellRange;
  filters: Record<number, FilterCriteria>; // Column index to filter
  active: boolean;
}

export interface FilterCriteria {
  type: 'value' | 'condition' | 'color';
  values?: Set<string>;
  condition?: string;
  color?: string;
}

export interface SortSettings {
  column: number;
  ascending: boolean;
  caseSensitive?: boolean;
}

export interface Chart {
  id: string;
  type: 'bar' | 'line' | 'pie' | 'scatter' | 'area' | 'column';
  dataRange: CellRange;
  position: CellPosition;
  size: { width: number; height: number };
  title?: string;
  legend?: boolean;
  axes?: { xAxis?: string; yAxis?: string };
}

export interface PivotTable {
  id: string;
  sourceRange: CellRange;
  position: CellPosition;
  rows: string[];
  columns: string[];
  values: string[];
  filters: string[];
}

export interface PrintSettings {
  paperSize: string;
  orientation: 'portrait' | 'landscape';
  margins: { top: number; bottom: number; left: number; right: number };
  fitToPage?: boolean;
  scale?: number;
  headerFooter?: { header: string; footer: string };
  gridlines?: boolean;
  rowColumnHeadings?: boolean;
}

export interface WorksheetProtection {
  enabled: boolean;
  password?: string;
  allowSelectLockedCells?: boolean;
  allowSelectUnlockedCells?: boolean;
  allowFormatCells?: boolean;
  allowInsertRows?: boolean;
  allowDeleteRows?: boolean;
}

export interface Workbook {
  id: string;
  name: string;
  worksheets: Worksheet[];
  activeWorksheetId: string;
  createdAt: Date;
  updatedAt: Date;
}

// AI Agent Types
export interface AIOperation {
  id: string;
  type: AIOperationType;
  description: string;
  status: 'pending' | 'in-progress' | 'completed' | 'failed';
  timestamp: Date;
  worksheetId?: string;
  cellRange?: CellRange;
  data?: any;
  error?: string;
}

export type AIOperationType = 
  | 'create_worksheet'
  | 'populate_data'
  | 'format_cells'
  | 'create_formula'
  | 'import_json'
  | 'generate_chart'
  | 'sort_data'
  | 'filter_data'
  | 'calculate_summary';

export interface AIPrompt {
  id: string;
  message: string;
  timestamp: Date;
  response?: string;
  operations: AIOperation[];
  status: 'processing' | 'completed' | 'failed';
}

// Real-time Communication Types
export interface WebSocketMessage {
  type: 'operation_update' | 'operation_complete' | 'error' | 'heartbeat';
  payload: any;
  timestamp: Date;
}

export interface OperationUpdate {
  operationId: string;
  status: AIOperation['status'];
  progress?: number;
  data?: any;
  error?: string;
}

// UI State Types
export interface UIState {
  selectedCells: CellRange | null;
  isAIProcessing: boolean;
  sidebarOpen: boolean;
  activeTab: 'prompt' | 'operations' | 'history';
  theme: 'light' | 'dark';
  clipboard?: ClipboardData;
  showFormulaBar?: boolean;
  zoom?: number;
}

export interface ClipboardData {
  type: 'copy' | 'cut';
  range: CellRange;
  cells: Record<string, CellValue>;
  timestamp: Date;
}

export interface HistoryEntry {
  id: string;
  action: string;
  timestamp: Date;
  data: any;
  undo: () => void;
  redo: () => void;
}

// API Response Types
export interface APIResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  timestamp: Date;
}

export interface ProcessPromptResponse {
  promptId: string;
  operations: AIOperation[];
  estimatedTime?: number;
}

// Configuration Types
export interface AppConfig {
  apiBaseUrl: string;
  websocketUrl: string;
  maxWorksheets: number;
  maxCellsPerWorksheet: number;
  enableRealTimeUpdates: boolean;
  theme: 'light' | 'dark';
}