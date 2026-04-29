/**
 * Excel file import utilities
 * Primary parser: SheetJS (xlsx) — fast, reliable, handles all xlsx variants
 * Export path: ExcelJS — rich formatting support
 */

import * as ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import type { Workbook, Worksheet, CellValue, CellFormat } from '../types/index';
import { v4 as uuidv4 } from 'uuid';

/**
 * Import Excel file using SheetJS — primary import path, never throws on valid xlsx
 */
const importWithSheetJS = async (file: File): Promise<Workbook> => {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true, cellDates: true });

  const worksheets: Worksheet[] = wb.SheetNames.map((sheetName) => {
    const ws = wb.Sheets[sheetName];
    const cells: Record<string, CellValue> = {};

    Object.entries(ws).forEach(([addr, cell]: [string, any]) => {
      if (addr.startsWith('!')) return; // skip metadata keys
      const ref = XLSX.utils.decode_cell(addr);
      const r = ref.r; // 0-based row
      const c = ref.c; // 0-based col
      const key = `${r},${c}`;

      let value: string | number | boolean | null = null;
      if (cell.v !== undefined && cell.v !== null) {
        if (cell.v instanceof Date) {
          value = cell.v.toISOString();
        } else {
          value = cell.v as string | number | boolean;
        }
      }

      const cellValue: CellValue = { value };
      if (cell.f) cellValue.formula = `=${cell.f}`;
      cells[key] = cellValue;
    });

    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { e: { r: 99, c: 25 } };

    return {
      id: uuidv4(),
      name: sheetName,
      cells,
      rowCount: Math.max(range.e.r + 1, 100),
      colCount: Math.max(range.e.c + 1, 26),
      columnWidths: {},
      rowHeights: {},
      createdAt: new Date(),
      updatedAt: new Date(),
    };
  });

  return {
    id: uuidv4(),
    name: file.name.replace(/\.(xlsx?|xlsm|xltx?|xlam)$/i, ''),
    worksheets,
    activeWorksheetId: worksheets[0]?.id || '',
    createdAt: new Date(),
    updatedAt: new Date(),
  };
};

/**
 * Parse Excel cell format and convert to our format
 */
const parseExcelFormat = (cell: ExcelJS.Cell): Partial<CellFormat> => {
  const format: Partial<CellFormat> = {};

  if (cell.style) {
    const style = cell.style;
    
    // Font
    if (style.font) {
      if (style.font.bold) format.bold = true;
      if (style.font.italic) format.italic = true;
      if (style.font.underline) format.underline = true;
      if (style.font.name) format.fontFamily = style.font.name;
      if (style.font.size) format.fontSize = style.font.size;
      if (style.font.color) {
        const color = style.font.color as any;
        format.fontColor = color.argb ? `#${color.argb.slice(2)}` : '#000000';
      }
    }

    // Fill
    if (style.fill && (style.fill as any).fgColor) {
      const fill = style.fill as any;
      if (fill.fgColor && fill.fgColor.argb) {
        format.backgroundColor = `#${fill.fgColor.argb.slice(2)}`;
      }
    }

    // Alignment
    if (style.alignment) {
      if (style.alignment.horizontal) {
        format.horizontalAlign = style.alignment.horizontal;
      }
      if (style.alignment.vertical) {
        format.verticalAlign = style.alignment.vertical;
      }
      if (style.alignment.wrapText) {
        format.wrapText = true;
      }
    }

    // Borders
    if (style.border) {
      const border: any = {};
      
      const parseBorder = (borderSide: any) => {
        if (!borderSide) return undefined;
        let style = 'solid';
        let width = '1px';
        let color = '#000000';
        
        if (borderSide.style) {
          switch (borderSide.style) {
            case 'thin': style = 'solid'; width = '1px'; break;
            case 'medium': style = 'solid'; width = '2px'; break;
            case 'thick': style = 'solid'; width = '3px'; break;
            case 'double': style = 'double'; width = '3px'; break;
            case 'dashed': style = 'dashed'; width = '1px'; break;
            case 'dotted': style = 'dotted'; width = '1px'; break;
            default: style = 'solid'; width = '1px'; break;
          }
        }
        
        if (borderSide.color && borderSide.color.argb) {
          color = `#${borderSide.color.argb.slice(2)}`;
        }
        
        return `${width} ${style} ${color}`;
      };
      
      if (style.border.top) border.top = parseBorder(style.border.top);
      if (style.border.bottom) border.bottom = parseBorder(style.border.bottom);
      if (style.border.left) border.left = parseBorder(style.border.left);
      if (style.border.right) border.right = parseBorder(style.border.right);
      
      if (Object.keys(border).length > 0) {
        format.border = border;
      }
    }

    // Number format
    if (style.numFmt) {
      format.numberFormat = style.numFmt;
    }
  }

  return format;
};

/**
 * Parse Excel cell value
 */
const parseExcelCellValue = (cell: ExcelJS.Cell): CellValue => {
  const cellValue: CellValue = {
    value: null,
  };

  if (!cell) return cellValue;

  // Get the actual value
  if (cell.value !== undefined && cell.value !== null) {
    if (typeof cell.value === 'string' || typeof cell.value === 'number' || typeof cell.value === 'boolean') {
      cellValue.value = cell.value;
    } else if (cell.value instanceof Date) {
      cellValue.value = cell.value.toISOString();
    } else {
      // Handle complex types like errors, rich text, etc.
      cellValue.value = String(cell.value);
    }
  }

  // Check if it's a formula
  if (cell.formula) {
    cellValue.formula = `=${cell.formula}`;
  }

  // Parse formatting
  const format = parseExcelFormat(cell);
  if (Object.keys(format).length > 0) {
    cellValue.format = format;
  }

  // Handle comments - ExcelJS has notes
  if (cell.note) {
    const noteText = typeof cell.note === 'string' ? cell.note : 
                     (cell.note.texts ? cell.note.texts.map(t => typeof t === 'string' ? t : t.text).join('') : '');
    cellValue.comment = {
      text: noteText,
      author: 'Unknown',
      timestamp: new Date(),
    };
  }

  return cellValue;
};

/**
 * Convert Excel worksheet to our Worksheet format
 */
const convertExcelWorksheet = (
  excelSheet: ExcelJS.Worksheet,
  name: string
): Worksheet => {
  const cells: Record<string, CellValue> = {};

  // Parse all cells
  excelSheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      const key = `${rowNumber - 1},${colNumber - 1}`;
      cells[key] = parseExcelCellValue(cell);
    });
  });

  // Parse column widths
  const columnWidths: Record<number, number> = {};
  excelSheet.columns.forEach((col, index) => {
    if (col && col.width) {
      columnWidths[index] = col.width * 7; // Approximate conversion
    }
  });

  // Parse row heights
  const rowHeights: Record<number, number> = {};
  excelSheet.eachRow((row, rowNumber) => {
    if (row.height) {
      rowHeights[rowNumber - 1] = row.height;
    }
  });

  // Parse merged cells - TODO: Implement merged cells parsing for ExcelJS
  const mergedCells: Array<{
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  }> = [];

  // Note: ExcelJS merged cells parsing is more complex and requires different approach
  // For now, merged cells from import will not be preserved

  return {
    id: uuidv4(),
    name,
    cells,
    rowCount: Math.max(excelSheet.rowCount, 1000),
    colCount: Math.max(excelSheet.columnCount, 26),
    columnWidths,
    rowHeights,
    createdAt: new Date(),
    updatedAt: new Date(),
  };
};

/**
 * Convert LuckySheet data to our Worksheet format
 */
const convertLuckySheetToWorksheet = (luckySheet: any): Worksheet => {
  const cells: Record<string, CellValue> = {};
  
  if (Array.isArray(luckySheet.celldata)) {
    luckySheet.celldata.forEach((cellData: any) => {
      const { r, c, v } = cellData;
      const key = `${r},${c}`;
      
      const cellValue: CellValue = {
        value: v.v,
      };

      if (v.f) {
        cellValue.formula = `=${v.f}`;
      }

      const format: any = {};
      // Basic formatting mapping
      if (v.bl) format.bold = true;
      if (v.it) format.italic = true;
      if (v.un) format.underline = true;
      if (v.cl) format.strikethrough = true;
      if (v.fs) format.fontSize = Number(v.fs);
      if (v.fc) format.fontColor = v.fc;
      if (v.bg) format.backgroundColor = v.bg;
      
      if (Object.keys(format).length > 0) {
        cellValue.format = format;
      }

      cells[key] = cellValue;
    });
  }

  const columnWidths: Record<number, number> = {};
  const rowHeights: Record<number, number> = {};

  if (luckySheet.config) {
    if (luckySheet.config.columnlen) {
       Object.entries(luckySheet.config.columnlen).forEach(([k, v]) => {
          columnWidths[Number(k)] = Number(v);
       });
    }
    if (luckySheet.config.rowlen) {
       Object.entries(luckySheet.config.rowlen).forEach(([k, v]) => {
          rowHeights[Number(k)] = Number(v);
       });
    }
  }

  return {
    id: uuidv4(),
    name: luckySheet.name,
    cells,
    rowCount: Math.max(luckySheet.data?.length || 100, 100),
    colCount: Math.max(luckySheet.data?.[0]?.length || 26, 26),
    columnWidths,
    rowHeights,
    createdAt: new Date(),
    updatedAt: new Date(),
  };
};

/**
 * Import using ExcelJS (Fallback)
 */
const importWithExcelJS = async (file: File): Promise<Workbook> => {
  const workbook = new ExcelJS.Workbook();

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        const data = e.target?.result as ArrayBuffer;
        if (!data) {
          reject(new Error('Failed to read file'));
          return;
        }

        // Load Excel file
        await workbook.xlsx.load(data);

        // Convert sheets to our format
        const worksheets: Worksheet[] = [];
        workbook.eachSheet((excelSheet, index) => {
          const worksheet = convertExcelWorksheet(excelSheet, excelSheet.name || `Sheet${index + 1}`);
          worksheets.push(worksheet);
        });

        // Create workbook
        const newWorkbook: Workbook = {
          id: uuidv4(),
          name: file.name.replace(/\.(xlsx?|xlsm|xltx?|xlam)$/i, ''),
          worksheets,
          activeWorksheetId: worksheets[0]?.id || '',
          createdAt: new Date(),
          updatedAt: new Date(),
        };

        resolve(newWorkbook);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => {
      reject(new Error('Failed to read file'));
    };

    reader.readAsArrayBuffer(file);
  });
};

/**
 * Import Excel file and convert to our Workbook format.
 * Uses SheetJS as the primary parser (fast, always works).
 * Falls back to ExcelJS if SheetJS fails.
 */
export const importExcelFile = async (file: File): Promise<Workbook> => {
  try {
    return await importWithSheetJS(file);
  } catch (sheetJsErr) {
    console.warn('SheetJS import failed, falling back to ExcelJS:', sheetJsErr);
    return importWithExcelJS(file);
  }
};

/**
 * Export workbook to Excel format
 */
export const exportToExcel = async (workbook: Workbook, filename?: string): Promise<void> => {
  const excelWorkbook = new ExcelJS.Workbook();

  workbook.worksheets.forEach((worksheet) => {
    // Create a new worksheet
    const excelSheet = excelWorkbook.addWorksheet(worksheet.name);

    // Convert cells
    Object.entries(worksheet.cells).forEach(([key, cellValue]) => {
      const [rowStr, colStr] = key.split(',');
      const row = parseInt(rowStr, 10) + 1; // ExcelJS uses 1-based
      const col = parseInt(colStr, 10) + 1;
      const cell = excelSheet.getCell(row, col);

      cell.value = cellValue.value ?? '';

      if (cellValue.formula) {
        cell.value = { formula: cellValue.formula.substring(1) }; // Remove leading '='
      }

      // Apply formatting if available
      if (cellValue.format) {
        const format = cellValue.format;
        if (format.bold || format.italic || format.underline || format.fontFamily || format.fontSize || format.fontColor) {
          cell.font = {
            bold: format.bold,
            italic: format.italic,
            underline: format.underline,
            name: format.fontFamily,
            size: format.fontSize,
            color: format.fontColor ? { argb: `FF${format.fontColor.slice(1)}` } : undefined,
          };
        }

        if (format.backgroundColor) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: `FF${format.backgroundColor.slice(1)}` },
          };
        }

        if (format.horizontalAlign || format.verticalAlign || format.wrapText) {
          cell.alignment = {
            horizontal: format.horizontalAlign,
            vertical: format.verticalAlign,
            wrapText: format.wrapText,
          };
        }

        if (format.border) {
          cell.border = {
            top: format.border.top ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
            bottom: format.border.bottom ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
            left: format.border.left ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
            right: format.border.right ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
          };
        }

        if (format.numberFormat) {
          cell.numFmt = format.numberFormat;
        }
      }
    });

    // Add column widths
    if (worksheet.columnWidths) {
      Object.entries(worksheet.columnWidths).forEach(([colIndex, width]) => {
        const col = parseInt(colIndex, 10) + 1;
        excelSheet.getColumn(col).width = width / 7; // Convert back
      });
    }

    // Add row heights
    if (worksheet.rowHeights) {
      Object.entries(worksheet.rowHeights).forEach(([rowIndex, height]) => {
        const row = parseInt(rowIndex, 10) + 1;
        excelSheet.getRow(row).height = height;
      });
    }

    // Handle merged cells
    Object.values(worksheet.cells).forEach((cellValue) => {
      if (cellValue.isMerged && cellValue.mergeRange) {
        const range = cellValue.mergeRange;
        excelSheet.mergeCells(range.startRow + 1, range.startCol + 1, range.endRow + 1, range.endCol + 1);
      }
    });
  });

  // Write file - Use browser download instead of writeFile
  const finalFilename = filename ? filename.replace(/\.xlsx?$/, '') + '.xlsx' : `${workbook.name}.xlsx`;
  
  // Generate buffer and create download
  const buffer = await excelWorkbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = finalFilename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
};

/**
 * Convert workbook to Blob for upload/save operations
 * IMPORTANT: This function creates a FRESH conversion every time it's called
 * No caching - always converts the current workbook state to a new Excel file
 */
export const workbookToBlob = async (workbook: Workbook): Promise<Blob> => {
  // Create a NEW ExcelJS Workbook instance (not cached)
  const excelWorkbook = new ExcelJS.Workbook();

  // Iterate through ALL worksheets and convert them FRESH
  workbook.worksheets.forEach((worksheet) => {
    // Create a new worksheet
    const excelSheet = excelWorkbook.addWorksheet(worksheet.name);

    // Convert ALL current cells (fresh data)
    Object.entries(worksheet.cells).forEach(([key, cellValue]) => {
      const [rowStr, colStr] = key.split(',');
      const row = parseInt(rowStr, 10) + 1; // ExcelJS uses 1-based
      const col = parseInt(colStr, 10) + 1;
      const cell = excelSheet.getCell(row, col);

      cell.value = cellValue.value ?? '';

      if (cellValue.formula) {
        cell.value = { formula: cellValue.formula.substring(1) }; // Remove leading '='
      }

      // Apply formatting if available
      if (cellValue.format) {
        const format = cellValue.format;
        if (format.bold || format.italic || format.underline || format.fontFamily || format.fontSize || format.fontColor) {
          cell.font = {
            bold: format.bold,
            italic: format.italic,
            underline: format.underline,
            name: format.fontFamily,
            size: format.fontSize,
            color: format.fontColor ? { argb: `FF${format.fontColor.slice(1)}` } : undefined,
          };
        }

        if (format.backgroundColor) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: `FF${format.backgroundColor.slice(1)}` },
          };
        }

        if (format.horizontalAlign || format.verticalAlign || format.wrapText) {
          cell.alignment = {
            horizontal: format.horizontalAlign,
            vertical: format.verticalAlign,
            wrapText: format.wrapText,
          };
        }

        if (format.border) {
          cell.border = {
            top: format.border.top ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
            bottom: format.border.bottom ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
            left: format.border.left ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
            right: format.border.right ? { style: 'thin', color: { argb: 'FF000000' } } : undefined,
          };
        }

        if (format.numberFormat) {
          cell.numFmt = format.numberFormat;
        }
      }
    });

    // Add column widths
    if (worksheet.columnWidths) {
      Object.entries(worksheet.columnWidths).forEach(([colIndex, width]) => {
        const col = parseInt(colIndex, 10) + 1;
        excelSheet.getColumn(col).width = width / 7; // Convert back
      });
    }

    // Add row heights
    if (worksheet.rowHeights) {
      Object.entries(worksheet.rowHeights).forEach(([rowIndex, height]) => {
        const row = parseInt(rowIndex, 10) + 1;
        excelSheet.getRow(row).height = height;
      });
    }

    // Handle merged cells
    Object.values(worksheet.cells).forEach((cellValue) => {
      if (cellValue.isMerged && cellValue.mergeRange) {
        const range = cellValue.mergeRange;
        excelSheet.mergeCells(range.startRow + 1, range.startCol + 1, range.endRow + 1, range.endCol + 1);
      }
    });
  });

  // Generate FRESH buffer from the converted workbook (not cached)
  const buffer = await excelWorkbook.xlsx.writeBuffer();
  // Create FRESH blob from buffer (not cached)
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
  return blob;
};

/**
 * Import Excel file and convert directly to FortuneSheet sheet array.
 * This is the primary import path for the FortuneSheet spreadsheet view.
 */
export const importExcelToFortuneSheets = async (file: File): Promise<any[]> => {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true, cellDates: true });
  return wb.SheetNames.map((sheetName, index) => {
    const ws = wb.Sheets[sheetName];
    const celldata: any[] = [];
    Object.entries(ws).forEach(([addr, cell]: [string, any]) => {
      if (addr.startsWith('!')) return;
      const ref = XLSX.utils.decode_cell(addr);
      const v: any = {};
      if (cell.v !== undefined && cell.v !== null) {
        v.v = cell.v instanceof Date ? cell.v.toISOString() : cell.v;
        v.m = cell.w || String(v.v);
      }
      if (cell.f) v.f = '=' + cell.f;
      if (Object.keys(v).length > 0) {
        celldata.push({ r: ref.r, c: ref.c, v });
      }
    });
    return { name: sheetName, celldata, status: index === 0 ? 1 : 0, order: index };
  });
};

/**
 * Convert FortuneSheet sheet array back to an Excel blob for saving.
 */
export const fortuneSheetsToBlob = async (sheets: any[]): Promise<Blob> => {
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const celldata: any[] = sheet.celldata || [];
    if (!celldata.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[]]), sheet.name || 'Sheet1');
      return;
    }
    const maxRow = Math.max(...celldata.map((c: any) => c.r)) + 1;
    const maxCol = Math.max(...celldata.map((c: any) => c.c)) + 1;
    const aoa: any[][] = Array.from({ length: maxRow }, () => Array(maxCol).fill(null));
    celldata.forEach((c: any) => {
      if (c.v) {
        aoa[c.r][c.c] = c.v.f ? { f: c.v.f.replace(/^=/, ''), v: c.v.v } : (c.v.v ?? null);
      }
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sheet.name || 'Sheet1');
  });
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};

export default {
  importExcelFile,
  exportToExcel,
  workbookToBlob,
  importExcelToFortuneSheets,
  fortuneSheetsToBlob,
};
