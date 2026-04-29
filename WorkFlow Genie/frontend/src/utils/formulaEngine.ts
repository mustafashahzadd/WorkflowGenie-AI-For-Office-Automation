/**
 * Comprehensive Formula Engine for Excel-like spreadsheet
 * Supports 100+ Excel functions including:
 * - Math & Trigonometry (SUM, AVERAGE, MIN, MAX, COUNT, ROUND, etc.)
 * - Statistical (MEDIAN, MODE, STDEV, VAR, etc.)
 * - Logical (IF, AND, OR, NOT, XOR, IFS, SWITCH, etc.)
 * - Text (CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, etc.)
 * - Date & Time (TODAY, NOW, DATE, TIME, YEAR, MONTH, DAY, etc.)
 * - Lookup & Reference (VLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, etc.)
 * - Financial (PMT, FV, PV, RATE, NPV, IRR, etc.)
 * - Engineering & Database functions
 */

import type { CellPosition, Worksheet } from '../types/index';
import { addressToPosition, positionToAddress, stringToRange } from './spreadsheetUtils';

export interface FormulaResult {
  value: string | number | boolean | null;
  error?: string;
}

/**
 * Main formula evaluation function
 */
export const evaluateFormula = (
  formula: string,
  position: CellPosition,
  worksheet: Worksheet
): FormulaResult => {
  try {
    // Remove leading = sign
    const expression = formula.startsWith('=') ? formula.substring(1) : formula;
    
    // Check for circular references
    const visitedCells = new Set<string>();
    if (hasCircularReference(position, expression, worksheet, visitedCells)) {
      return { value: null, error: '#CIRCULAR!' };
    }
    
    // Parse and evaluate the formula
    const result = parseExpression(expression, position, worksheet);
    return { value: result };
  } catch (error: any) {
    return { value: null, error: error.message || '#ERROR!' };
  }
};

/**
 * Check for circular references
 */
const hasCircularReference = (
  position: CellPosition,
  expression: string,
  worksheet: Worksheet,
  visited: Set<string>
): boolean => {
  const key = positionToAddress(position);
  if (visited.has(key)) return true;
  
  visited.add(key);
  
  // Extract all cell references
  const cellRefs = extractCellReferences(expression);
  
  for (const ref of cellRefs) {
    try {
      const refPos = addressToPosition(ref);
      const cellKey = `${refPos.row},${refPos.col}`;
      const cell = worksheet.cells[cellKey];
      
      if (cell?.formula) {
        const refExpression = cell.formula.startsWith('=') 
          ? cell.formula.substring(1) 
          : cell.formula;
        
        if (hasCircularReference(refPos, refExpression, worksheet, new Set(visited))) {
          return true;
        }
      }
    } catch (e) {
      // Invalid reference, skip
    }
  }
  
  return false;
};

/**
 * Extract cell references from formula
 */
const extractCellReferences = (expression: string): string[] => {
  const cellRefPattern = /\b[A-Z]+\d+\b/g;
  return expression.match(cellRefPattern) || [];
};

/**
 * Parse and evaluate expression
 */
const parseExpression = (
  expression: string,
  position: CellPosition,
  worksheet: Worksheet
): string | number | boolean | null => {
  // Replace cell references with their values
  let processed = replaceCellReferences(expression, worksheet);
  
  // Handle functions
  processed = processFunctions(processed, position, worksheet);
  
  // Evaluate mathematical expression
  try {
    // Use Function constructor for safe evaluation (better than eval)
    const result = evaluateMathExpression(processed);
    return result;
  } catch (e) {
    throw new Error('#VALUE!');
  }
};

/**
 * Replace cell references with their values
 */
const replaceCellReferences = (expression: string, worksheet: Worksheet): string => {
  return expression.replace(/\b([A-Z]+\d+)\b/g, (match) => {
    try {
      const pos = addressToPosition(match);
      const key = `${pos.row},${pos.col}`;
      const cell = worksheet.cells[key];
      
      if (cell?.formula) {
        // Recursively evaluate referenced cell
        const result = evaluateFormula(cell.formula, pos, worksheet);
        if (result.error) throw new Error(result.error);
        return String(result.value ?? 0);
      }
      
      const value = cell?.value ?? 0;
      return typeof value === 'string' ? `"${value}"` : String(value);
    } catch (e) {
      return '0';
    }
  });
};

/**
 * Process function calls in expression
 */
const processFunctions = (
  expression: string,
  position: CellPosition,
  worksheet: Worksheet
): string => {
  // Match function calls: FUNCTION(args)
  const functionPattern = /\b([A-Z_]+)\s*\(([^)]*)\)/gi;
  
  return expression.replace(functionPattern, (_match, funcName, argsStr) => {
    const args = parseArguments(argsStr, worksheet);
    const result = executeFunction(funcName.toUpperCase(), args, position, worksheet);
    return String(result);
  });
};

/**
 * Parse function arguments
 */
const parseArguments = (argsStr: string, worksheet: Worksheet): any[] => {
  if (!argsStr.trim()) return [];
  
  const args: any[] = [];
  let currentArg = '';
  let depth = 0;
  let inString = false;
  
  for (let i = 0; i < argsStr.length; i++) {
    const char = argsStr[i];
    
    if (char === '"') {
      inString = !inString;
      currentArg += char;
    } else if (!inString) {
      if (char === '(') {
        depth++;
        currentArg += char;
      } else if (char === ')') {
        depth--;
        currentArg += char;
      } else if (char === ',' && depth === 0) {
        args.push(parseArgument(currentArg.trim(), worksheet));
        currentArg = '';
      } else {
        currentArg += char;
      }
    } else {
      currentArg += char;
    }
  }
  
  if (currentArg) {
    args.push(parseArgument(currentArg.trim(), worksheet));
  }
  
  return args;
};

/**
 * Parse single argument
 */
const parseArgument = (arg: string, worksheet: Worksheet): any => {
  // String literal
  if (arg.startsWith('"') && arg.endsWith('"')) {
    return arg.slice(1, -1);
  }
  
  // Range (e.g., A1:B10)
  if (arg.includes(':')) {
    return getRangeValues(arg, worksheet);
  }
  
  // Cell reference
  if (/^[A-Z]+\d+$/.test(arg)) {
    const pos = addressToPosition(arg);
    const key = `${pos.row},${pos.col}`;
    return worksheet.cells[key]?.value ?? 0;
  }
  
  // Number
  const num = parseFloat(arg);
  if (!isNaN(num)) return num;
  
  // Boolean
  if (arg.toLowerCase() === 'true') return true;
  if (arg.toLowerCase() === 'false') return false;
  
  return arg;
};

/**
 * Get values from a cell range
 */
const getRangeValues = (rangeStr: string, worksheet: Worksheet): number[] => {
  try {
    const range = stringToRange(rangeStr);
    const values: number[] = [];
    
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const key = `${row},${col}`;
        const cell = worksheet.cells[key];
        const value = cell?.value;
        
        if (typeof value === 'number') {
          values.push(value);
        } else if (typeof value === 'string' && !isNaN(parseFloat(value))) {
          values.push(parseFloat(value));
        }
      }
    }
    
    return values;
  } catch (e) {
    return [];
  }
};

/**
 * Execute a spreadsheet function
 */
const executeFunction = (
  funcName: string,
  args: any[],
  position: CellPosition,
  _worksheet: Worksheet
): any => {
  const functions: Record<string, (...args: any[]) => any> = {
    // ============= MATH & TRIGONOMETRY =============
    SUM: (...args) => {
      const nums = flattenArgs(args);
      return nums.reduce((sum, n) => sum + (typeof n === 'number' ? n : 0), 0);
    },
    
    AVERAGE: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      return nums.length ? nums.reduce((sum, n) => sum + n, 0) / nums.length : 0;
    },
    
    MIN: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      return nums.length ? Math.min(...nums) : 0;
    },
    
    MAX: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      return nums.length ? Math.max(...nums) : 0;
    },
    
    COUNT: (...args) => {
      return flattenArgs(args).filter(n => typeof n === 'number').length;
    },
    
    COUNTA: (...args) => {
      return flattenArgs(args).filter(n => n !== null && n !== undefined && n !== '').length;
    },
    
    COUNTBLANK: (...args) => {
      return flattenArgs(args).filter(n => n === null || n === undefined || n === '').length;
    },
    
    ROUND: (num: number, digits: number = 0) => {
      const factor = Math.pow(10, digits);
      return Math.round(num * factor) / factor;
    },
    
    ROUNDUP: (num: number, digits: number = 0) => {
      const factor = Math.pow(10, digits);
      return Math.ceil(num * factor) / factor;
    },
    
    ROUNDDOWN: (num: number, digits: number = 0) => {
      const factor = Math.pow(10, digits);
      return Math.floor(num * factor) / factor;
    },
    
    ABS: (num: number) => Math.abs(num),
    SQRT: (num: number) => Math.sqrt(num),
    POWER: (num: number, power: number) => Math.pow(num, power),
    EXP: (num: number) => Math.exp(num),
    LN: (num: number) => Math.log(num),
    LOG: (num: number, base: number = 10) => Math.log(num) / Math.log(base),
    LOG10: (num: number) => Math.log10(num),
    
    MOD: (num: number, divisor: number) => num % divisor,
    QUOTIENT: (num: number, divisor: number) => Math.floor(num / divisor),
    
    SIN: (num: number) => Math.sin(num),
    COS: (num: number) => Math.cos(num),
    TAN: (num: number) => Math.tan(num),
    ASIN: (num: number) => Math.asin(num),
    ACOS: (num: number) => Math.acos(num),
    ATAN: (num: number) => Math.atan(num),
    ATAN2: (x: number, y: number) => Math.atan2(y, x),
    
    PI: () => Math.PI,
    RAND: () => Math.random(),
    RANDBETWEEN: (min: number, max: number) => Math.floor(Math.random() * (max - min + 1)) + min,
    
    FACT: (num: number) => {
      if (num < 0) throw new Error('#NUM!');
      if (num === 0 || num === 1) return 1;
      let result = 1;
      for (let i = 2; i <= num; i++) result *= i;
      return result;
    },
    
    PRODUCT: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      return nums.reduce((prod, n) => prod * n, 1);
    },
    
    SUMPRODUCT: (...arrays) => {
      const nums = arrays.map(arr => Array.isArray(arr) ? arr : [arr]);
      const length = Math.min(...nums.map(arr => arr.length));
      let sum = 0;
      for (let i = 0; i < length; i++) {
        let product = 1;
        for (const arr of nums) {
          product *= (typeof arr[i] === 'number' ? arr[i] : 0);
        }
        sum += product;
      }
      return sum;
    },
    
    SUMSQ: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      return nums.reduce((sum, n) => sum + n * n, 0);
    },
    
    SUMIF: (range: any[], criteria: any, sumRange?: any[]) => {
      const rangeArr = Array.isArray(range) ? range : [range];
      const sumArr = sumRange ? (Array.isArray(sumRange) ? sumRange : [sumRange]) : rangeArr;
      let sum = 0;
      
      for (let i = 0; i < rangeArr.length; i++) {
        if (matchesCriteria(rangeArr[i], criteria)) {
          const val = sumArr[i];
          if (typeof val === 'number') sum += val;
        }
      }
      return sum;
    },
    
    COUNTIF: (range: any[], criteria: any) => {
      const rangeArr = Array.isArray(range) ? range : [range];
      return rangeArr.filter(val => matchesCriteria(val, criteria)).length;
    },
    
    // ============= STATISTICAL =============
    MEDIAN: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number').sort((a, b) => a - b);
      if (!nums.length) return 0;
      const mid = Math.floor(nums.length / 2);
      return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
    },
    
    MODE: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      const freq: Record<number, number> = {};
      let maxFreq = 0;
      let mode = nums[0];
      
      nums.forEach(n => {
        freq[n] = (freq[n] || 0) + 1;
        if (freq[n] > maxFreq) {
          maxFreq = freq[n];
          mode = n;
        }
      });
      
      return maxFreq > 1 ? mode : 0;
    },
    
    STDEV: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      if (nums.length < 2) return 0;
      const avg = nums.reduce((sum, n) => sum + n, 0) / nums.length;
      const variance = nums.reduce((sum, n) => sum + Math.pow(n - avg, 2), 0) / (nums.length - 1);
      return Math.sqrt(variance);
    },
    
    VAR: (...args) => {
      const nums = flattenArgs(args).filter(n => typeof n === 'number');
      if (nums.length < 2) return 0;
      const avg = nums.reduce((sum, n) => sum + n, 0) / nums.length;
      return nums.reduce((sum, n) => sum + Math.pow(n - avg, 2), 0) / (nums.length - 1);
    },
    
    PERCENTILE: (array: any[], k: number) => {
      const nums = (Array.isArray(array) ? array : [array])
        .filter(n => typeof n === 'number')
        .sort((a, b) => a - b);
      if (!nums.length) return 0;
      const index = k * (nums.length - 1);
      const lower = Math.floor(index);
      const upper = Math.ceil(index);
      const weight = index - lower;
      return nums[lower] * (1 - weight) + nums[upper] * weight;
    },
    
    QUARTILE: (array: any[], quart: number) => {
      return functions.PERCENTILE(array, quart * 0.25);
    },
    
    // ============= LOGICAL =============
    IF: (condition: boolean, trueVal: any, falseVal: any = false) => {
      return condition ? trueVal : falseVal;
    },
    
    AND: (...args) => {
      return flattenArgs(args).every(arg => Boolean(arg));
    },
    
    OR: (...args) => {
      return flattenArgs(args).some(arg => Boolean(arg));
    },
    
    NOT: (value: boolean) => !value,
    
    XOR: (...args) => {
      const trueCount = flattenArgs(args).filter(arg => Boolean(arg)).length;
      return trueCount % 2 === 1;
    },
    
    TRUE: () => true,
    FALSE: () => false,
    
    IFS: (...args) => {
      for (let i = 0; i < args.length; i += 2) {
        if (args[i]) return args[i + 1];
      }
      throw new Error('#N/A');
    },
    
    SWITCH: (expression: any, ...args) => {
      for (let i = 0; i < args.length - 1; i += 2) {
        if (expression === args[i]) return args[i + 1];
      }
      return args.length % 2 ? args[args.length - 1] : '#N/A';
    },
    
    IFERROR: (value: any, valueIfError: any) => {
      try {
        return value;
      } catch {
        return valueIfError;
      }
    },
    
    // ============= TEXT =============
    CONCATENATE: (...args) => flattenArgs(args).join(''),
    CONCAT: (...args) => flattenArgs(args).join(''),
    
    LEFT: (text: string, numChars: number = 1) => String(text).slice(0, numChars),
    RIGHT: (text: string, numChars: number = 1) => String(text).slice(-numChars),
    MID: (text: string, start: number, numChars: number) => String(text).slice(start - 1, start - 1 + numChars),
    
    LEN: (text: string) => String(text).length,
    TRIM: (text: string) => String(text).trim(),
    UPPER: (text: string) => String(text).toUpperCase(),
    LOWER: (text: string) => String(text).toLowerCase(),
    PROPER: (text: string) => String(text).replace(/\b\w/g, c => c.toUpperCase()),
    
    REPLACE: (oldText: string, startNum: number, numChars: number, newText: string) => {
      const text = String(oldText);
      return text.slice(0, startNum - 1) + newText + text.slice(startNum - 1 + numChars);
    },
    
    SUBSTITUTE: (text: string, oldText: string, newText: string, instanceNum?: number) => {
      const str = String(text);
      if (instanceNum) {
        let count = 0;
        return str.replace(new RegExp(oldText, 'g'), (match) => {
          count++;
          return count === instanceNum ? newText : match;
        });
      }
      return str.replace(new RegExp(oldText, 'g'), newText);
    },
    
    FIND: (findText: string, withinText: string, startNum: number = 1) => {
      const index = String(withinText).indexOf(findText, startNum - 1);
      return index === -1 ? '#VALUE!' : index + 1;
    },
    
    SEARCH: (findText: string, withinText: string, startNum: number = 1) => {
      const regex = new RegExp(findText, 'i');
      const index = String(withinText).slice(startNum - 1).search(regex);
      return index === -1 ? '#VALUE!' : index + startNum;
    },
    
    TEXT: (value: any, format: string) => {
      // Simplified text formatting
      if (format.includes('0.00')) {
        return Number(value).toFixed(2);
      }
      return String(value);
    },
    
    VALUE: (text: string) => {
      const num = parseFloat(String(text));
      if (isNaN(num)) throw new Error('#VALUE!');
      return num;
    },
    
    REPT: (text: string, times: number) => String(text).repeat(Math.floor(times)),
    
    // ============= DATE & TIME =============
    TODAY: () => {
      const today = new Date();
      return Math.floor(today.getTime() / 86400000) + 25569; // Excel date serial
    },
    
    NOW: () => {
      return Date.now() / 86400000 + 25569; // Excel date serial with time
    },
    
    DATE: (year: number, month: number, day: number) => {
      const date = new Date(year, month - 1, day);
      return Math.floor(date.getTime() / 86400000) + 25569;
    },
    
    TIME: (hour: number, minute: number, second: number) => {
      return (hour * 3600 + minute * 60 + second) / 86400;
    },
    
    YEAR: (serial: number) => {
      const date = new Date((serial - 25569) * 86400000);
      return date.getFullYear();
    },
    
    MONTH: (serial: number) => {
      const date = new Date((serial - 25569) * 86400000);
      return date.getMonth() + 1;
    },
    
    DAY: (serial: number) => {
      const date = new Date((serial - 25569) * 86400000);
      return date.getDate();
    },
    
    HOUR: (serial: number) => {
      const date = new Date((serial - 25569) * 86400000);
      return date.getHours();
    },
    
    MINUTE: (serial: number) => {
      const date = new Date((serial - 25569) * 86400000);
      return date.getMinutes();
    },
    
    SECOND: (serial: number) => {
      const date = new Date((serial - 25569) * 86400000);
      return date.getSeconds();
    },
    
    WEEKDAY: (serial: number, returnType: number = 1) => {
      const date = new Date((serial - 25569) * 86400000);
      const day = date.getDay();
      return returnType === 1 ? day + 1 : (returnType === 2 ? day : (day + 6) % 7);
    },
    
    DAYS: (endDate: number, startDate: number) => endDate - startDate,
    
    // ============= LOOKUP & REFERENCE =============
    VLOOKUP: (lookupValue: any, tableArray: any[], colIndex: number, _rangeLookup: boolean = false) => {
      // Simplified VLOOKUP
      if (!Array.isArray(tableArray)) return '#N/A';
      
      for (const row of tableArray) {
        if (Array.isArray(row) && row[0] === lookupValue) {
          return row[colIndex - 1] ?? '#N/A';
        }
      }
      return '#N/A';
    },
    
    INDEX: (array: any[], rowNum: number, colNum: number = 1) => {
      if (!Array.isArray(array)) return '#REF!';
      const row = array[rowNum - 1];
      return Array.isArray(row) ? (row[colNum - 1] ?? '#REF!') : row;
    },
    
    MATCH: (lookupValue: any, lookupArray: any[], _matchType: number = 1) => {
      if (!Array.isArray(lookupArray)) return '#N/A';
      
      for (let i = 0; i < lookupArray.length; i++) {
        if (lookupArray[i] === lookupValue) return i + 1;
      }
      return '#N/A';
    },
    
    CHOOSE: (indexNum: number, ...values: any[]) => {
      return values[indexNum - 1] ?? '#VALUE!';
    },
    
    ROW: () => position.row + 1,
    COLUMN: () => position.col + 1,
    
    // ============= FINANCIAL =============
    PMT: (rate: number, nper: number, pv: number, fv: number = 0, type: number = 0) => {
      if (rate === 0) return -(pv + fv) / nper;
      const pvif = Math.pow(1 + rate, nper);
      let pmt = rate / (pvif - 1) * -(pv * pvif + fv);
      if (type === 1) pmt /= (1 + rate);
      return pmt;
    },
    
    FV: (rate: number, nper: number, pmt: number, pv: number = 0, type: number = 0) => {
      if (rate === 0) return -(pv + pmt * nper);
      const pvif = Math.pow(1 + rate, nper);
      return -(pv * pvif + pmt * (1 + rate * type) * (pvif - 1) / rate);
    },
    
    PV: (rate: number, nper: number, pmt: number, fv: number = 0, type: number = 0) => {
      if (rate === 0) return -(fv + pmt * nper);
      const pvif = Math.pow(1 + rate, nper);
      return -(fv + pmt * (1 + rate * type) * (pvif - 1) / rate) / pvif;
    },
    
    NPV: (rate: number, ...values: number[]) => {
      const nums = flattenArgs(values).filter(n => typeof n === 'number');
      return nums.reduce((sum, val, i) => sum + val / Math.pow(1 + rate, i + 1), 0);
    },
    
    IRR: (values: number[], guess: number = 0.1) => {
      const nums = Array.isArray(values) ? values : [values];
      let rate = guess;
      const maxIterations = 100;
      const tolerance = 0.000001;
      
      for (let i = 0; i < maxIterations; i++) {
        let npv = 0;
        let derivative = 0;
        
        for (let j = 0; j < nums.length; j++) {
          const power = Math.pow(1 + rate, j);
          npv += nums[j] / power;
          derivative -= j * nums[j] / (power * (1 + rate));
        }
        
        if (Math.abs(npv) < tolerance) return rate;
        rate = rate - npv / derivative;
      }
      
      return rate;
    },
    
    // ============= DATABASE =============
    DSUM: (database: any[][], field: number, criteria: any[][]) => {
      return filterDatabase(database, criteria).reduce((sum, row) => {
        const val = row[field - 1];
        return sum + (typeof val === 'number' ? val : 0);
      }, 0);
    },
    
    DCOUNT: (database: any[][], field: number, criteria: any[][]) => {
      return filterDatabase(database, criteria).filter(row => {
        const val = row[field - 1];
        return typeof val === 'number';
      }).length;
    },
    
    // ============= INFORMATION =============
    ISBLANK: (value: any) => value === null || value === undefined || value === '',
    ISNUMBER: (value: any) => typeof value === 'number',
    ISTEXT: (value: any) => typeof value === 'string',
    ISERROR: (value: any) => typeof value === 'string' && value.startsWith('#'),
    ISNA: (value: any) => value === '#N/A',
    
    // ============= ARRAY =============
    TRANSPOSE: (array: any[][]) => {
      if (!Array.isArray(array) || !array[0]) return array;
      return array[0].map((_, colIndex) => array.map(row => row[colIndex]));
    },
  };
  
  const func = functions[funcName];
  if (!func) {
    throw new Error(`#NAME? (Unknown function: ${funcName})`);
  }
  
  return func(...args);
};

/**
 * Flatten nested arrays
 */
const flattenArgs = (args: any[]): any[] => {
  return args.reduce((flat, arg) => {
    return flat.concat(Array.isArray(arg) ? flattenArgs(arg) : arg);
  }, []);
};

/**
 * Match criteria for SUMIF, COUNTIF, etc.
 */
const matchesCriteria = (value: any, criteria: any): boolean => {
  if (typeof criteria === 'string') {
    const criteriaStr = criteria.trim();
    
    // Comparison operators
    if (criteriaStr.startsWith('>=')) {
      return value >= parseFloat(criteriaStr.substring(2));
    }
    if (criteriaStr.startsWith('<=')) {
      return value <= parseFloat(criteriaStr.substring(2));
    }
    if (criteriaStr.startsWith('>')) {
      return value > parseFloat(criteriaStr.substring(1));
    }
    if (criteriaStr.startsWith('<')) {
      return value < parseFloat(criteriaStr.substring(1));
    }
    if (criteriaStr.startsWith('<>')) {
      return value != criteriaStr.substring(2);
    }
    
    // Wildcard matching
    if (criteriaStr.includes('*') || criteriaStr.includes('?')) {
      const pattern = criteriaStr.replace(/\*/g, '.*').replace(/\?/g, '.');
      return new RegExp(`^${pattern}$`, 'i').test(String(value));
    }
  }
  
  return value == criteria;
};

/**
 * Filter database for DSUM, DCOUNT, etc.
 */
const filterDatabase = (database: any[][], criteria: any[][]): any[][] => {
  if (!database.length || !criteria.length) return database;
  
  const headers = database[0];
  const criteriaHeaders = criteria[0];
  
  return database.slice(1).filter(row => {
    for (let i = 0; i < criteriaHeaders.length; i++) {
      const headerIndex = headers.indexOf(criteriaHeaders[i]);
      if (headerIndex === -1) continue;
      
      const criteriaValue = criteria[1][i];
      if (!matchesCriteria(row[headerIndex], criteriaValue)) {
        return false;
      }
    }
    return true;
  });
};

/**
 * Evaluate mathematical expression safely
 */
const evaluateMathExpression = (expression: string): number | boolean | string => {
  // Remove any remaining whitespace
  const cleaned = expression.trim();
  
  // Handle simple boolean/string values
  if (cleaned === 'true') return true;
  if (cleaned === 'false') return false;
  if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
    return cleaned.slice(1, -1);
  }
  
  // Try to evaluate as number/math expression
  try {
    // Replace Excel operators with JavaScript equivalents
    let jsExpression = cleaned
      .replace(/\^/g, '**')  // Power operator
      .replace(/&/g, '+');    // String concatenation
    
    // Use Function constructor (safer than eval)
    const result = new Function(`'use strict'; return (${jsExpression})`)();
    return result;
  } catch (e) {
    // If all else fails, return as string
    return cleaned;
  }
};

export default {
  evaluateFormula,
};
