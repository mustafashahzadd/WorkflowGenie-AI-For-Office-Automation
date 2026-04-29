/**
 * Type definitions for LuckySheet library
 * Since LuckySheet may not have complete TypeScript definitions,
 * this file provides basic type support for our usage
 */

declare module 'luckysheet' {
  interface CellValue {
    v?: any;
    m?: string;
    ct?: {
      fa?: string;
      t?: string;
    };
    f?: string;
    bl?: number;
    it?: number;
    un?: number;
    cl?: number;
    fs?: number;
    fc?: string;
    bg?: string;
    ht?: number;
    vt?: number;
    tb?: number;
    ff?: string;
  }

  interface CellData {
    r: number;
    c: number;
    v: CellValue;
  }

  interface SheetConfig {
    columnlen?: Record<number, number>;
    rowlen?: Record<number, number>;
    merge?: Record<string, any>;
  }

  interface SheetData {
    name: string;
    color?: string;
    status?: number;
    order?: number;
    data?: any[][];
    celldata?: CellData[];
    config?: SheetConfig;
    frozen?: {
      type: string;
      range: {
        row_focus: number;
        column_focus: number;
      };
    };
  }

  interface Range {
    row: number[];
    column: number[];
  }

  interface LuckySheetOptions {
    container: string;
    title?: string;
    lang?: string;
    showtoolbar?: boolean;
    showinfobar?: boolean;
    showsheetbar?: boolean;
    showstatisticBar?: boolean;
    sheetBottomConfig?: boolean;
    allowEdit?: boolean;
    enableAddRow?: boolean;
    enableAddCol?: boolean;
    userInfo?: boolean;
    showGridLines?: boolean;
    data?: SheetData[];
    hook?: {
      cellUpdated?: (r: number, c: number, oldValue: any, newValue: any, isRefresh: boolean) => void;
      rangeSelect?: (sheet: any, range: Range[]) => void;
      updated?: (operate: string) => void;
      cellEditBefore?: (range: Range[]) => void;
      cellUpdateBefore?: (r: number, c: number, value: any, isRefresh: boolean) => boolean;
      sheetActivate?: (index: number, isPivotInitial: boolean, isNewSheet: boolean) => void;
      [key: string]: any;
    };
  }

  interface LuckySheet {
    create(options: LuckySheetOptions): void;
    destroy(): void;
    luckysheetfile: SheetData[];
    jfrefreshgrid(): void;
    getSheetData(sheetIndex?: number): any;
    getCellValue(row: number, col: number, options?: any): any;
    setCellValue(row: number, col: number, value: any, options?: any): void;
    getRange(): Range[];
    setRangeValue(range: Range, value: any): void;
    [key: string]: any;
  }

  const luckysheet: LuckySheet;
  export default luckysheet;
}
