/**
 * Utility functions for spreadsheet operations, formatting, and data manipulation
 */

import type { CellPosition, CellRange, CellValue } from '../types/index';

/**
 * Convert column number to Excel-style letter (A, B, C, ..., AA, AB, etc.)
 */
export const numberToColumnLetter = (num: number): string => {
  let result = '';
  while (num >= 0) {
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26) - 1;
  }
  return result;
};

/**
 * Convert Excel-style column letter to number
 */
export const columnLetterToNumber = (letter: string): number => {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
};

/**
 * Convert cell position to Excel-style address (e.g., A1, B2)
 */
export const positionToAddress = (position: CellPosition): string => {
  return `${numberToColumnLetter(position.col)}${position.row + 1}`;
};

/**
 * Convert Excel-style address to cell position
 */
export const addressToPosition = (address: string): CellPosition => {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  
  const [, letters, numbers] = match;
  return {
    col: columnLetterToNumber(letters),
    row: parseInt(numbers) - 1,
  };
};

/**
 * Convert cell range to Excel-style range string (e.g., A1:C3)
 */
export const rangeToString = (range: CellRange): string => {
  const start = positionToAddress({ row: range.startRow, col: range.startCol });
  const end = positionToAddress({ row: range.endRow, col: range.endCol });
  return `${start}:${end}`;
};

/**
 * Convert Excel-style range string to cell range
 */
export const stringToRange = (rangeStr: string): CellRange => {
  const [startAddr, endAddr] = rangeStr.split(':');
  const start = addressToPosition(startAddr);
  const end = endAddr ? addressToPosition(endAddr) : start;
  
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
};

/**
 * Check if a position is within a range
 */
export const isPositionInRange = (position: CellPosition, range: CellRange): boolean => {
  return (
    position.row >= range.startRow &&
    position.row <= range.endRow &&
    position.col >= range.startCol &&
    position.col <= range.endCol
  );
};

/**
 * Generate a unique ID
 */
export const generateId = (): string => {
  return `${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
};

/**
 * Format cell value for display
 */
export const formatCellValue = (value: CellValue | null): string => {
  if (!value || value.value === null || value.value === undefined) {
    return '';
  }
  
  const rawValue = value.value;
  
  // Handle different data types
  if (typeof rawValue === 'number') {
    // Apply number formatting if specified
    if (value.format?.numberFormat) {
      return formatNumber(rawValue, value.format.numberFormat);
    }
    return rawValue.toString();
  }
  
  if (typeof rawValue === 'boolean') {
    return rawValue ? 'TRUE' : 'FALSE';
  }
  
  return rawValue.toString();
};

/**
 * Format number based on format string
 */
export const formatNumber = (num: number, format: string): string => {
  switch (format) {
    case 'currency':
      return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
      }).format(num);
    case 'percentage':
      return new Intl.NumberFormat('en-US', {
        style: 'percent',
        minimumFractionDigits: 2,
      }).format(num / 100);
    case 'date':
      return new Date(num).toLocaleDateString();
    case 'integer':
      return Math.round(num).toString();
    case 'decimal':
      return num.toFixed(2);
    default:
      return num.toString();
  }
};

/**
 * Parse value from string input
 */
export const parseValue = (input: string): CellValue => {
  const trimmed = input.trim();
  
  // Empty input
  if (!trimmed) {
    return { value: null };
  }
  
  // Formula (starts with =)
  if (trimmed.startsWith('=')) {
    return {
      value: trimmed, // For now, store formula as-is
      formula: trimmed,
    };
  }
  
  // Boolean values
  if (trimmed.toLowerCase() === 'true') {
    return { value: true };
  }
  if (trimmed.toLowerCase() === 'false') {
    return { value: false };
  }
  
  // Number values
  const numValue = Number(trimmed);
  if (!isNaN(numValue) && isFinite(numValue)) {
    return { value: numValue };
  }
  
  // Default to string
  return { value: trimmed };
};

/**
 * Deep clone an object
 */
export const deepClone = <T>(obj: T): T => {
  return JSON.parse(JSON.stringify(obj));
};

/**
 * Debounce function for performance optimization
 */
export const debounce = <T extends (...args: any[]) => any>(
  func: T,
  wait: number
): ((...args: Parameters<T>) => void) => {
  let timeout: number;
  
  return (...args: Parameters<T>) => {
    clearTimeout(timeout);
    timeout = setTimeout(() => func(...args), wait) as any;
  };
};

/**
 * Throttle function for performance optimization
 */
export const throttle = <T extends (...args: any[]) => any>(
  func: T,
  limit: number
): ((...args: Parameters<T>) => void) => {
  let inThrottle: boolean;
  
  return (...args: Parameters<T>) => {
    if (!inThrottle) {
      func(...args);
      inThrottle = true;
      setTimeout(() => (inThrottle = false), limit);
    }
  };
};

/**
 * Export data to CSV format
 */
export const exportToCSV = (data: Record<string, CellValue>, range: CellRange): string => {
  const rows: string[] = [];
  
  for (let row = range.startRow; row <= range.endRow; row++) {
    const rowData: string[] = [];
    
    for (let col = range.startCol; col <= range.endCol; col++) {
      const key = `${row},${col}`;
      const cell = data[key];
      const value = formatCellValue(cell);
      
      // Escape CSV values
      const escaped = value.includes(',') || value.includes('"') || value.includes('\n')
        ? `"${value.replace(/"/g, '""')}"`
        : value;
      
      rowData.push(escaped);
    }
    
    rows.push(rowData.join(','));
  }
  
  return rows.join('\n');
};

/**
 * Import data from CSV format
 */
export const importFromCSV = (csvText: string): CellValue[][] => {
  const rows = csvText.split('\n').filter(row => row.trim());
  const data: CellValue[][] = [];
  
  rows.forEach(row => {
    const cells = parseCSVRow(row);
    const rowData: CellValue[] = cells.map(cell => parseValue(cell));
    data.push(rowData);
  });
  
  return data;
};

/**
 * Parse a CSV row, handling quoted values
 */
const parseCSVRow = (row: string): string[] => {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < row.length; i++) {
    const char = row[i];
    
    if (char === '"') {
      if (inQuotes && row[i + 1] === '"') {
        current += '"';
        i++; // Skip next quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  
  result.push(current);
  return result;
};

/**
 * Generate sample data for testing
 */
export const generateSampleData = (rows: number, cols: number): Record<string, CellValue> => {
  const data: Record<string, CellValue> = {};
  const sampleProducts = ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Tablet', 'Phone'];
  const sampleRegions = ['North', 'South', 'East', 'West'];
  
  // Headers
  data['0,0'] = { value: 'Product', format: { bold: true } };
  data['0,1'] = { value: 'Sales', format: { bold: true } };
  data['0,2'] = { value: 'Region', format: { bold: true } };
  data['0,3'] = { value: 'Date', format: { bold: true } };
  
  // Sample data
  for (let row = 1; row < rows; row++) {
    data[`${row},0`] = { value: sampleProducts[row % sampleProducts.length] };
    data[`${row},1`] = { value: Math.floor(Math.random() * 5000) + 100 };
    data[`${row},2`] = { value: sampleRegions[row % sampleRegions.length] };
    data[`${row},3`] = { value: new Date(2024, Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1).toLocaleDateString() };
  }
  
  return data;
};