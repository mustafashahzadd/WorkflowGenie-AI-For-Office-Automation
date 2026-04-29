/**
 * Formula Bar component - displays and allows editing of cell formulas and values
 * Located between toolbar and spreadsheet grid, just like Excel
 */

import React, { useState, useEffect, useRef } from 'react';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import { positionToAddress } from '../../utils/spreadsheetUtils';
import type { CellPosition } from '../../types/index';

interface FormulaBarProps {
  selectedCell: CellPosition;
  className?: string;
}

export const FormulaBar: React.FC<FormulaBarProps> = ({ selectedCell, className = '' }) => {
  const { getCellValue, setCellValue } = useSpreadsheetStore();
  const [formulaValue, setFormulaValue] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const cell = getCellValue(selectedCell);
    // Show formula if it exists, otherwise show value
    if (cell?.formula) {
      setFormulaValue(cell.formula);
    } else {
      setFormulaValue(cell?.value?.toString() || '');
    }
  }, [selectedCell, getCellValue]);

  const handleSubmit = () => {
    setCellValue(selectedCell, {
      value: formulaValue,
      formula: formulaValue.startsWith('=') ? formulaValue : undefined,
    });
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      handleSubmit();
      inputRef.current?.blur();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      const cell = getCellValue(selectedCell);
      setFormulaValue(cell?.formula || cell?.value?.toString() || '');
      inputRef.current?.blur();
    }
  };

  const cellAddress = positionToAddress(selectedCell);

  return (
    <div className={`flex items-center gap-0 px-1 py-0 bg-white border-b border-gray-300 h-8 ${className}`}>
      {/* Cell address indicator */}
      <div className="flex items-center justify-center min-w-[40px] px-2 border-r border-gray-300 h-full bg-gray-50">
        <span className="text-sm font-medium text-gray-600">{cellAddress}</span>
      </div>

      {/* Function button (fx) */}
      <div className="flex items-center justify-center px-3 h-full border-r border-gray-300 bg-gray-50">
        <span className="text-gray-500 font-serif italic font-bold text-lg leading-none">fx</span>
      </div>

      {/* Formula input */}
      <input
        ref={inputRef}
        type="text"
        value={formulaValue}
        onChange={(e) => setFormulaValue(e.target.value)}
        onKeyDown={handleKeyDown}
        onBlur={handleSubmit}
        className="flex-1 px-2 h-full text-sm outline-none border-none focus:ring-0"
      />
    </div>
  );
};
