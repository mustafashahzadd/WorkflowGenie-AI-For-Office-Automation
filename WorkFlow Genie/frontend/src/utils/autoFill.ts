/**
 * Auto-fill and pattern detection utilities for Excel-like behavior
 * Detects patterns in selected cells and auto-fills accordingly
 */

import type { CellValue } from '../types/index';

export interface FillPattern {
  type: 'linear' | 'growth' | 'date' | 'series' | 'copy';
  step?: number;
  multiplier?: number;
}

/**
 * Detect pattern in a series of values
 */
export const detectPattern = (values: (string | number | null)[]): FillPattern => {
  const numbers = values.filter(v => typeof v === 'number') as number[];
  
  // Check for numeric linear progression
  if (numbers.length >= 2) {
    const differences = [];
    for (let i = 1; i < numbers.length; i++) {
      differences.push(numbers[i] - numbers[i - 1]);
    }
    
    const avgDiff = differences.reduce((sum, d) => sum + d, 0) / differences.length;
    const isLinear = differences.every(d => Math.abs(d - avgDiff) < 0.001);
    
    if (isLinear) {
      return { type: 'linear', step: avgDiff };
    }
    
    // Check for growth (geometric progression)
    const ratios = [];
    for (let i = 1; i < numbers.length; i++) {
      if (numbers[i - 1] !== 0) {
        ratios.push(numbers[i] / numbers[i - 1]);
      }
    }
    
    if (ratios.length > 0) {
      const avgRatio = ratios.reduce((sum, r) => sum + r, 0) / ratios.length;
      const isGrowth = ratios.every(r => Math.abs(r - avgRatio) < 0.001);
      
      if (isGrowth && avgRatio !== 1) {
        return { type: 'growth', multiplier: avgRatio };
      }
    }
  }
  
  // Check for date patterns
  if (values.length > 0 && typeof values[0] === 'string') {
    const firstValue = values[0] as string;
    if (isDateString(firstValue)) {
      return { type: 'date', step: 1 };
    }
  }
  
  // Check for series patterns (Mon, Tue, Wed... or Jan, Feb, Mar...)
  if (values.length > 0 && typeof values[0] === 'string') {
    const pattern = detectSeriesPattern(values as string[]);
    if (pattern) {
      return pattern;
    }
  }
  
  // Default to copy
  return { type: 'copy' };
};

/**
 * Generate auto-fill values based on pattern
 */
export const autoFill = (
  sourceValues: CellValue[],
  targetCount: number,
  _direction: 'down' | 'right' | 'up' | 'left'
): CellValue[] => {
  if (sourceValues.length === 0) return [];
  
  // Extract raw values
  const rawValues = sourceValues.map(cv => cv.value).filter(v => typeof v !== 'boolean') as (string | number | null)[];
  const pattern = detectPattern(rawValues);
  
  const result: CellValue[] = [];
  
  for (let i = 0; i < targetCount; i++) {
    const sourceIndex = i % sourceValues.length;
    const cycleCount = Math.floor(i / sourceValues.length);
    
    let newValue: string | number | boolean | null = null;
    
    switch (pattern.type) {
      case 'linear':
        if (typeof sourceValues[sourceIndex].value === 'number' && pattern.step !== undefined) {
          newValue = (sourceValues[sourceIndex].value as number) + (pattern.step * (cycleCount + 1));
        }
        break;
        
      case 'growth':
        if (typeof sourceValues[sourceIndex].value === 'number' && pattern.multiplier !== undefined) {
          newValue = (sourceValues[sourceIndex].value as number) * Math.pow(pattern.multiplier, cycleCount + 1);
        }
        break;
        
      case 'date':
        if (typeof sourceValues[sourceIndex].value === 'string') {
          const date = new Date(sourceValues[sourceIndex].value as string);
          if (!isNaN(date.getTime())) {
            date.setDate(date.getDate() + (cycleCount + 1));
            newValue = date.toISOString().split('T')[0];
          }
        }
        break;
        
      case 'series':
        newValue = getNextInSeries(sourceValues[sourceIndex].value as string, cycleCount + 1);
        break;
        
      case 'copy':
      default:
        newValue = sourceValues[sourceIndex].value;
        break;
    }
    
    result.push({
      value: newValue,
      format: sourceValues[sourceIndex].format,
    });
  }
  
  return result;
};

/**
 * Check if string is a date
 */
const isDateString = (str: string): boolean => {
  const date = new Date(str);
  return !isNaN(date.getTime());
};

/**
 * Detect series patterns (days, months, etc.)
 */
const detectSeriesPattern = (values: string[]): FillPattern | null => {
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const daysShort = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  
  const series = [days, daysShort, months, monthsShort];
  
  for (const s of series) {
    const indices = values.map(v => s.findIndex(item => item.toLowerCase() === v.toLowerCase()));
    if (indices.every(i => i !== -1)) {
      return { type: 'series', step: 1 };
    }
  }
  
  return null;
};

/**
 * Get next item in a series
 */
const getNextInSeries = (value: string, offset: number): string => {
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const daysShort = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  
  const series = [
    { array: days, name: 'days' },
    { array: daysShort, name: 'daysShort' },
    { array: months, name: 'months' },
    { array: monthsShort, name: 'monthsShort' },
  ];
  
  for (const s of series) {
    const index = s.array.findIndex(item => item.toLowerCase() === value.toLowerCase());
    if (index !== -1) {
      const newIndex = (index + offset) % s.array.length;
      return s.array[newIndex];
    }
  }
  
  return value;
};

/**
 * Flash Fill - detect pattern and fill based on examples
 * This is a simplified version of Excel's Flash Fill feature
 */
export const flashFill = (
  examples: Array<{ input: string; output: string }>,
  newInput: string
): string => {
  if (examples.length === 0) return newInput;
  
  // Detect common patterns
  const patterns = [];
  
  for (const example of examples) {
    const { input, output } = example;
    
    // Check for substring extraction
    if (output.length < input.length && input.includes(output)) {
      const startIndex = input.indexOf(output);
      patterns.push({
        type: 'substring',
        start: startIndex,
        length: output.length,
      });
    }
    
    // Check for case conversion
    if (output.toLowerCase() === input.toLowerCase()) {
      if (output === input.toUpperCase()) {
        patterns.push({ type: 'uppercase' });
      } else if (output === input.toLowerCase()) {
        patterns.push({ type: 'lowercase' });
      } else if (output === input.charAt(0).toUpperCase() + input.slice(1).toLowerCase()) {
        patterns.push({ type: 'capitalize' });
      }
    }
    
    // Check for concatenation with delimiter
    const words = output.split(/[\s,.-]+/);
    if (words.length > 1) {
      patterns.push({ type: 'split', delimiter: output.match(/[\s,.-]+/)?.[0] || ' ' });
    }
  }
  
  // Apply most common pattern
  if (patterns.length > 0) {
    const pattern = patterns[0];
    
    switch (pattern.type) {
      case 'substring':
        if (pattern.start !== undefined && pattern.length !== undefined) {
          return newInput.substring(pattern.start, pattern.start + pattern.length);
        }
        return newInput;
      case 'uppercase':
        return newInput.toUpperCase();
      case 'lowercase':
        return newInput.toLowerCase();
      case 'capitalize':
        return newInput.charAt(0).toUpperCase() + newInput.slice(1).toLowerCase();
      default:
        return newInput;
    }
  }
  
  return newInput;
};

/**
 * Smart paste - detect and apply formatting from copied data
 */
export const smartPaste = (
  data: string,
  _detectFormat: boolean = true
): { values: string[][]; format?: 'csv' | 'tsv' | 'html' | 'plain' } => {
  // Detect format
  let format: 'csv' | 'tsv' | 'html' | 'plain' = 'plain';
  
  if (data.includes('\t') && data.includes('\n')) {
    format = 'tsv';
  } else if (data.includes(',') && data.includes('\n')) {
    format = 'csv';
  } else if (data.includes('<table') || data.includes('<tr')) {
    format = 'html';
  }
  
  let values: string[][] = [];
  
  switch (format) {
    case 'csv':
      values = parseCSV(data);
      break;
    case 'tsv':
      values = parseTSV(data);
      break;
    case 'html':
      values = parseHTMLTable(data);
      break;
    default:
      values = [[data]];
      break;
  }
  
  return { values, format };
};

/**
 * Parse CSV data
 */
const parseCSV = (data: string): string[][] => {
  const rows = data.split('\n').filter(row => row.trim());
  return rows.map(row => {
    const cells: string[] = [];
    let currentCell = '';
    let inQuotes = false;
    
    for (let i = 0; i < row.length; i++) {
      const char = row[i];
      
      if (char === '"') {
        if (inQuotes && row[i + 1] === '"') {
          currentCell += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        cells.push(currentCell);
        currentCell = '';
      } else {
        currentCell += char;
      }
    }
    
    cells.push(currentCell);
    return cells;
  });
};

/**
 * Parse TSV (tab-separated) data
 */
const parseTSV = (data: string): string[][] => {
  return data.split('\n').map(row => row.split('\t'));
};

/**
 * Parse HTML table
 */
const parseHTMLTable = (data: string): string[][] => {
  const parser = new DOMParser();
  const doc = parser.parseFromString(data, 'text/html');
  const rows = doc.querySelectorAll('tr');
  
  return Array.from(rows).map(row => {
    const cells = row.querySelectorAll('td, th');
    return Array.from(cells).map(cell => cell.textContent || '');
  });
};

export default {
  detectPattern,
  autoFill,
  flashFill,
  smartPaste,
};
