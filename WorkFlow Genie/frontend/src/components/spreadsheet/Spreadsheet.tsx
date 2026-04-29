/**
 * Core spreadsheet component that renders an Excel-like grid interface
 * Features include cell editing, selection, formulas, formatting, and real-time updates
 * 
 * NOTE: This is the legacy custom implementation. The application now uses
 * LuckysheetSpreadsheet.tsx which leverages the LuckySheet library for better
 * functionality and maintainability. This file is kept for reference and as a fallback.
 * 
 * See: src/components/spreadsheet/LuckysheetSpreadsheet.tsx
 * Documentation: LUCKYSHEET_INTEGRATION.md
 */

import React, { useState, useCallback, useRef, useEffect } from 'react';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import { evaluateFormula } from '../../utils/formulaEngine';
import { FormulaBar } from '../ui/FormulaBar';
import { ContextMenu } from '../ui/ContextMenu';
import type { CellPosition, CellValue } from '../../types/index';

interface SpreadsheetProps {
  className?: string;
}

/**
 * Convert column number to Excel-style letter (A, B, C, ..., AA, AB, etc.)
 */
const numberToColumnLetter = (num: number): string => {
  let result = '';
  while (num >= 0) {
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26) - 1;
  }
  return result;
};

/**
 * Individual cell component with editing capabilities
 */
interface CellProps {
  position: CellPosition;
  value: CellValue | null;
  isSelected: boolean;
  onSelect: (position: CellPosition, shiftKey: boolean) => void;
  onMouseDown: (position: CellPosition, event: React.MouseEvent) => void;
  onMouseEnter: (position: CellPosition, isPrimaryDown: boolean) => void;
  onEdit: (position: CellPosition, value: string) => void;
  onContextMenu: (e: React.MouseEvent, position: CellPosition) => void;
}

const Cell: React.FC<CellProps> = React.memo(({
  position,
  value,
  isSelected,
  onSelect,
  onMouseDown,
  onMouseEnter,
  onEdit,
  onContextMenu,
}) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editValue, setEditValue] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const { currentWorksheet } = useSpreadsheetStore();

  const handleClick = (e: React.MouseEvent) => {
    onSelect(position, e.shiftKey);
  };

  const handleMouseDown = (e: React.MouseEvent) => {
    onMouseDown(position, e);
  };

  const handleMouseEnter = (e: React.MouseEvent) => {
    onMouseEnter(position, e.buttons === 1);
  };

  const handleDoubleClick = () => {
    setIsEditing(true);
    // When editing, show the formula if it exists
    setEditValue(value?.formula || value?.value?.toString() || '');
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSubmit();
    } else if (e.key === 'Escape') {
      setIsEditing(false);
      setEditValue('');
    }
  };

  const handleSubmit = () => {
    onEdit(position, editValue);
    setIsEditing(false);
  };

  const handleBlur = () => {
    handleSubmit();
  };

  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [isEditing]);

  // Compute display value (evaluate formulas)
  const displayValue = React.useMemo(() => {
    if (!value) return '';
    
    if (value.formula && currentWorksheet) {
      const result = evaluateFormula(value.formula, position, currentWorksheet);
      if (result.error) return result.error;
      return result.value?.toString() || '';
    }
    
    return value.value?.toString() || '';
  }, [value, position, currentWorksheet]);

  // Apply cell formatting
  // Map Excel alignments to CSS
  const mapHorizontalAlign = (align?: string): 'left' | 'center' | 'right' | 'justify' => {
    switch (align) {
      case 'center':
      case 'centerContinuous':
        return 'center';
      case 'right':
        return 'right';
      case 'justify':
        return 'justify';
      case 'fill':
      case 'distributed':
      default:
        return 'left';
    }
  };

  const mapVerticalAlign = (align?: string): 'top' | 'middle' | 'bottom' => {
    switch (align) {
      case 'top':
        return 'top';
      case 'middle':
        return 'middle';
      case 'bottom':
        return 'bottom';
      case 'justify':
      case 'distributed':
      default:
        return 'middle';
    }
  };

  const cellStyle: React.CSSProperties = {
    fontWeight: value?.format?.bold ? 'bold' : 'normal',
    fontStyle: value?.format?.italic ? 'italic' : 'normal',
    textDecoration: [
      value?.format?.underline ? 'underline' : '',
      value?.format?.strikethrough ? 'line-through' : '',
    ].filter(Boolean).join(' ') || 'none',
    fontSize: value?.format?.fontSize ? `${value.format.fontSize}px` : undefined,
    color: value?.format?.fontColor || undefined,
    backgroundColor: value?.format?.backgroundColor || undefined,
    textAlign: mapHorizontalAlign(value?.format?.horizontalAlign),
    verticalAlign: mapVerticalAlign(value?.format?.verticalAlign),
    fontFamily: value?.format?.fontFamily || undefined,
    whiteSpace: value?.format?.wrapText ? 'pre-wrap' : 'nowrap',
    textIndent: value?.format?.indent ? `${value.format.indent}em` : undefined,
    borderTop: value?.format?.border?.top || undefined,
    borderBottom: value?.format?.border?.bottom || undefined,
    borderLeft: value?.format?.border?.left || undefined,
    borderRight: value?.format?.border?.right || undefined,
  };

  const textStyle: React.CSSProperties = {
    transform: value?.format?.textRotation ? `rotate(${value.format.textRotation}deg)` : undefined,
  };

  const cellClasses = [
    'excel-cell',
    isSelected && 'selected',
    isEditing && 'editing',
  ].filter(Boolean).join(' ');

  return (
    <div
      className={cellClasses}
      style={cellStyle}
      onClick={handleClick}
      onMouseDown={handleMouseDown}
      onMouseEnter={handleMouseEnter}
      onDoubleClick={handleDoubleClick}
      onContextMenu={(e) => onContextMenu(e, position)}
      role="gridcell"
      tabIndex={0}
    >
      {isEditing ? (
        <input
          ref={inputRef}
          type="text"
          value={editValue}
          onChange={(e) => setEditValue(e.target.value)}
          onKeyDown={handleKeyDown}
          onBlur={handleBlur}
          className="w-full h-full border-none outline-none bg-transparent p-0 m-0"
        />
      ) : (
        <span 
          className={`${value?.format?.wrapText ? 'w-full block' : 'truncate w-full block'}`} 
          style={textStyle}
          title={displayValue}
        >
          {displayValue}
        </span>
      )}
      {value?.comment && (
        <div className="absolute top-0 right-0 w-0 h-0 border-t-[6px] border-t-red-500 border-l-[6px] border-l-transparent"></div>
      )}
    </div>
  );
}, (prevProps, nextProps) => {
  // Custom comparison for memo optimization
  return (
    prevProps.position.row === nextProps.position.row &&
    prevProps.position.col === nextProps.position.col &&
    prevProps.isSelected === nextProps.isSelected &&
    prevProps.value === nextProps.value
  );
});

Cell.displayName = 'Cell';

/**
 * Main spreadsheet component
 */
export const Spreadsheet: React.FC<SpreadsheetProps> = ({ className = '' }) => {
  const {
    currentWorksheet,
    ui,
    setCellValue,
    getCellValue,
    setSelectedRange,
  } = useSpreadsheetStore();

  const [selectedCell, setSelectedCell] = useState<CellPosition>({ row: 0, col: 0 });
  const [selectionStart, setSelectionStart] = useState<CellPosition | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number; cell: CellPosition } | null>(null);
  const [visibleRange, setVisibleRange] = useState({
    startRow: 0,
    endRow: 50,
    startCol: 0,
    endCol: 26,
  });
  const [resizingColumn, setResizingColumn] = useState<{ col: number; startX: number; startWidth: number } | null>(null);
  const [resizingRow, setResizingRow] = useState<{ row: number; startY: number; startHeight: number } | null>(null);

  const containerRef = useRef<HTMLDivElement>(null);
  const scrollTimeoutRef = useRef<NodeJS.Timeout>();
  const isPointerDownRef = useRef(false);
  const maxRows = currentWorksheet?.rowCount || 10000;
  const maxCols = currentWorksheet?.colCount || 1000;

  const handleCellSelect = useCallback((position: CellPosition, shiftKey: boolean = false) => {
    if (shiftKey && selectionStart) {
      // Shift-click: select range from selectionStart to current position
      const range = {
        startRow: Math.min(selectionStart.row, position.row),
        startCol: Math.min(selectionStart.col, position.col),
        endRow: Math.max(selectionStart.row, position.row),
        endCol: Math.max(selectionStart.col, position.col),
      };
      setSelectedRange(range);
      setSelectedCell(position);
    } else {
      // Normal click: select single cell
      setSelectedCell(position);
      setSelectionStart(position);
      setSelectedRange({
        startRow: position.row,
        startCol: position.col,
        endRow: position.row,
        endCol: position.col,
      });
    }
  }, [setSelectedRange, selectionStart]);

  const handleCellMouseDown = useCallback((position: CellPosition, event: React.MouseEvent) => {
    if (event.button !== 0) return; // only start selection on primary button

    isPointerDownRef.current = true;

    if (!event.shiftKey) {
      setIsDragging(true);
      setSelectionStart(position);
    }
    handleCellSelect(position, event.shiftKey);
  }, [handleCellSelect]);

  const handleCellMouseEnter = useCallback((position: CellPosition, isPrimaryDown: boolean) => {
    // Guard against stray mousemove events when no button is pressed
    if (!isDragging || !selectionStart) return;
    if (!isPrimaryDown || !isPointerDownRef.current) {
      setIsDragging(false);
      return;
    }

    const range = {
      startRow: Math.min(selectionStart.row, position.row),
      startCol: Math.min(selectionStart.col, position.col),
      endRow: Math.max(selectionStart.row, position.row),
      endCol: Math.max(selectionStart.col, position.col),
    };
    setSelectedRange(range);
    setSelectedCell(position);
  }, [isDragging, selectionStart, setSelectedRange]);

  const handleMouseUp = useCallback(() => {
    isPointerDownRef.current = false;
    setIsDragging(false);
  }, []);

  // Reset dragging if window loses focus (prevents stuck selection state)
  useEffect(() => {
    const handleWindowBlur = () => {
      isPointerDownRef.current = false;
      setIsDragging(false);
    };

    window.addEventListener('blur', handleWindowBlur);
    return () => window.removeEventListener('blur', handleWindowBlur);
  }, []);

  // Setup mouse up listener for drag selection
  useEffect(() => {
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [handleMouseUp]);

  const handleContextMenu = useCallback((e: React.MouseEvent, position: CellPosition) => {
    e.preventDefault();
    setContextMenu({
      x: e.clientX,
      y: e.clientY,
      cell: position,
    });
  }, []);

  const handleCellEdit = useCallback((position: CellPosition, value: string) => {
    const cellValue: CellValue = {
      value: value === '' ? null : value,
      formula: value.startsWith('=') ? value : undefined,
    };
    setCellValue(position, cellValue);
  }, [setCellValue]);

  const isCellSelected = useCallback((position: CellPosition) => {
    if (!ui.selectedCells) {
      return selectedCell.row === position.row && selectedCell.col === position.col;
    }
    return (
      position.row >= ui.selectedCells.startRow &&
      position.row <= ui.selectedCells.endRow &&
      position.col >= ui.selectedCells.startCol &&
      position.col <= ui.selectedCells.endCol
    );
  }, [selectedCell, ui.selectedCells]);

  // Handle column resize
  const handleColumnResizeStart = useCallback((e: React.MouseEvent, col: number) => {
    e.preventDefault();
    e.stopPropagation();
    const currentWidth = currentWorksheet?.columnWidths?.[col] || 100;
    setResizingColumn({ col, startX: e.clientX, startWidth: currentWidth });
  }, [currentWorksheet]);

  const handleColumnResizeMove = useCallback((e: MouseEvent) => {
    if (!resizingColumn) return;
    const delta = e.clientX - resizingColumn.startX;
    const newWidth = Math.max(30, resizingColumn.startWidth + delta);
    const { setColumnWidth } = useSpreadsheetStore.getState();
    setColumnWidth(resizingColumn.col, newWidth);
  }, [resizingColumn]);

  const handleColumnResizeEnd = useCallback(() => {
    setResizingColumn(null);
  }, []);

  // Handle row resize
  const handleRowResizeStart = useCallback((e: React.MouseEvent, row: number) => {
    e.preventDefault();
    e.stopPropagation();
    const currentHeight = currentWorksheet?.rowHeights?.[row] || 24;
    setResizingRow({ row, startY: e.clientY, startHeight: currentHeight });
  }, [currentWorksheet]);

  const handleRowResizeMove = useCallback((e: MouseEvent) => {
    if (!resizingRow) return;
    const delta = e.clientY - resizingRow.startY;
    const newHeight = Math.max(20, resizingRow.startHeight + delta);
    const { setRowHeight } = useSpreadsheetStore.getState();
    setRowHeight(resizingRow.row, newHeight);
  }, [resizingRow]);

  const handleRowResizeEnd = useCallback(() => {
    setResizingRow(null);
  }, []);

  // Setup resize event listeners
  useEffect(() => {
    if (resizingColumn) {
      document.addEventListener('mousemove', handleColumnResizeMove);
      document.addEventListener('mouseup', handleColumnResizeEnd);
      return () => {
        document.removeEventListener('mousemove', handleColumnResizeMove);
        document.removeEventListener('mouseup', handleColumnResizeEnd);
      };
    }
  }, [resizingColumn, handleColumnResizeMove, handleColumnResizeEnd]);

  useEffect(() => {
    if (resizingRow) {
      document.addEventListener('mousemove', handleRowResizeMove);
      document.addEventListener('mouseup', handleRowResizeEnd);
      return () => {
        document.removeEventListener('mousemove', handleRowResizeMove);
        document.removeEventListener('mouseup', handleRowResizeEnd);
      };
    }
  }, [resizingRow, handleRowResizeMove, handleRowResizeEnd]);

  // Cleanup scroll timeout on unmount
  useEffect(() => {
    return () => {
      if (scrollTimeoutRef.current) {
        clearTimeout(scrollTimeoutRef.current);
      }
    };
  }, []);

  // Handle scrolling to load more cells (throttled for performance)
  const handleScroll = useCallback((e: React.UIEvent<HTMLDivElement>) => {
    // Clear existing timeout
    if (scrollTimeoutRef.current) {
      clearTimeout(scrollTimeoutRef.current);
    }

    // Throttle scroll updates to every 100ms
    scrollTimeoutRef.current = setTimeout(() => {
      const target = e.currentTarget;
      const scrollPercentX = (target.scrollLeft + target.clientWidth) / target.scrollWidth;
      const scrollPercentY = (target.scrollTop + target.clientHeight) / target.scrollHeight;

      setVisibleRange(prev => {
        let updated = { ...prev };
        
        // Load more columns if scrolled near right edge
        if (scrollPercentX > 0.8 && prev.endCol < maxCols - 1) {
          updated.endCol = Math.min(maxCols - 1, prev.endCol + 26);
        }
        
        // Load more rows if scrolled near bottom edge
        if (scrollPercentY > 0.8 && prev.endRow < maxRows - 1) {
          updated.endRow = Math.min(maxRows - 1, prev.endRow + 50);
        }
        
        return updated;
      });
    }, 100);
  }, [maxRows, maxCols]);

  // Handle keyboard navigation and shortcuts
  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    const { row, col } = selectedCell;
    let newPosition = { ...selectedCell };

    // Copy (Ctrl+C)
    if (e.ctrlKey && e.key === 'c') {
      e.preventDefault();
      const { copy, ui } = useSpreadsheetStore.getState();
      copy(ui.selectedCells || undefined);
      return;
    }

    // Cut (Ctrl+X)
    if (e.ctrlKey && e.key === 'x') {
      e.preventDefault();
      const { cut, ui } = useSpreadsheetStore.getState();
      cut(ui.selectedCells || undefined);
      return;
    }

    // Paste (Ctrl+V)
    if (e.ctrlKey && e.key === 'v') {
      e.preventDefault();
      const { paste, ui } = useSpreadsheetStore.getState();
      // Paste at the top-left corner of the selected range
      const pastePos = ui.selectedCells 
        ? { row: ui.selectedCells.startRow, col: ui.selectedCells.startCol }
        : selectedCell;
      paste(pastePos);
      return;
    }

    // Undo (Ctrl+Z)
    if (e.ctrlKey && e.key === 'z') {
      e.preventDefault();
      const { undo } = useSpreadsheetStore.getState();
      undo();
      return;
    }

    // Redo (Ctrl+Y)
    if (e.ctrlKey && e.key === 'y') {
      e.preventDefault();
      const { redo } = useSpreadsheetStore.getState();
      redo();
      return;
    }

    // Bold (Ctrl+B)
    if (e.ctrlKey && e.key === 'b') {
      e.preventDefault();
      const { setRangeFormat, ui } = useSpreadsheetStore.getState();
      if (ui.selectedCells) {
        setRangeFormat(ui.selectedCells, { bold: true });
      }
      return;
    }

    // Italic (Ctrl+I)
    if (e.ctrlKey && e.key === 'i') {
      e.preventDefault();
      const { setRangeFormat, ui } = useSpreadsheetStore.getState();
      if (ui.selectedCells) {
        setRangeFormat(ui.selectedCells, { italic: true });
      }
      return;
    }

    // Underline (Ctrl+U)
    if (e.ctrlKey && e.key === 'u') {
      e.preventDefault();
      const { setRangeFormat, ui } = useSpreadsheetStore.getState();
      if (ui.selectedCells) {
        setRangeFormat(ui.selectedCells, { underline: true });
      }
      return;
    }

    // Delete cell content
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      const { ui, setCellValue: setCellValueStore } = useSpreadsheetStore.getState();
      if (ui.selectedCells) {
        // Delete all cells in range
        for (let row = ui.selectedCells.startRow; row <= ui.selectedCells.endRow; row++) {
          for (let col = ui.selectedCells.startCol; col <= ui.selectedCells.endCol; col++) {
            setCellValueStore({ row, col }, { value: null });
          }
        }
      } else {
        setCellValue(selectedCell, { value: null });
      }
      return;
    }

    // Navigation
    switch (e.key) {
      case 'ArrowUp':
        newPosition.row = Math.max(0, row - 1);
        break;
      case 'ArrowDown':
        newPosition.row = Math.min(maxRows - 1, row + 1);
        break;
      case 'ArrowLeft':
        newPosition.col = Math.max(0, col - 1);
        break;
      case 'ArrowRight':
        newPosition.col = Math.min(maxCols - 1, col + 1);
        break;
      case 'Tab':
        e.preventDefault();
        newPosition.col = Math.min(maxCols - 1, col + 1);
        break;
      default:
        return;
    }

    if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Tab'].includes(e.key)) {
        e.preventDefault();
        handleCellSelect(newPosition, false);
        
        // Update visible range if necessary
        setVisibleRange(prev => {
          let updated = { ...prev };
          if (newPosition.row < prev.startRow) {
            updated.startRow = newPosition.row;
            updated.endRow = newPosition.row + 50;
          } else if (newPosition.row > prev.endRow) {
            updated.endRow = newPosition.row;
            updated.startRow = Math.max(0, newPosition.row - 50);
          }
          if (newPosition.col < prev.startCol) {
            updated.startCol = newPosition.col;
            updated.endCol = newPosition.col + 26;
          } else if (newPosition.col > prev.endCol) {
            updated.endCol = newPosition.col;
            updated.startCol = Math.max(0, newPosition.col - 26);
          }
          return updated;
        });
        
        // Scroll into view
        const element = document.getElementById(`cell-${newPosition.row}-${newPosition.col}`);
        if (element) {
            element.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }
    }
  }, [selectedCell, currentWorksheet, handleCellSelect, getCellValue, setCellValue, maxRows, maxCols]);

  // Render column headers
  const renderColumnHeaders = () => {
    const headers = [];
    for (let col = visibleRange.startCol; col <= visibleRange.endCol; col++) {
      const isSelected = selectedCell.col === col;
      headers.push(
        <div
          key={col}
          className={`excel-col-header ${isSelected ? 'bg-[#e1e1e1] text-[#107c41] font-bold border-b-2 border-b-[#107c41]' : ''} relative`}
          style={{ gridColumn: col + 2, gridRow: 1 }}
        >
          {numberToColumnLetter(col)}
          <div
            className="absolute top-0 right-0 w-1 h-full cursor-col-resize hover:bg-[#107c41] hover:w-0.5 z-10"
            onMouseDown={(e) => handleColumnResizeStart(e, col)}
            title="Drag to resize column"
          />
        </div>
      );
    }
    return headers;
  };

  // Render row headers
  const renderRowHeaders = () => {
    const headers = [];
    for (let row = visibleRange.startRow; row <= visibleRange.endRow; row++) {
      const isSelected = selectedCell.row === row;
      headers.push(
        <div
          key={row}
          className={`excel-row-header ${isSelected ? 'bg-[#e1e1e1] text-[#107c41] font-bold border-r-2 border-r-[#107c41]' : ''} relative`}
          style={{ gridColumn: 1, gridRow: row + 2 }}
        >
          {row + 1}
          <div
            className="absolute bottom-0 left-0 w-full h-1 cursor-row-resize hover:bg-[#107c41] hover:h-0.5 z-10"
            onMouseDown={(e) => handleRowResizeStart(e, row)}
            title="Drag to resize row"
          />
        </div>
      );
    }
    return headers;
  };

  // Render cells
  const renderCells = () => {
    const cells = [];
    for (let row = visibleRange.startRow; row <= visibleRange.endRow; row++) {
      for (let col = visibleRange.startCol; col <= visibleRange.endCol; col++) {
        const position = { row, col };
        const value = getCellValue(position);
        cells.push(
          <div
            key={`${row}-${col}`}
            id={`cell-${row}-${col}`}
            style={{ gridColumn: col + 2, gridRow: row + 2 }}
          >
            <Cell
              position={position}
              value={value}
              isSelected={isCellSelected(position)}
              onSelect={handleCellSelect}
              onMouseDown={handleCellMouseDown}
              onMouseEnter={handleCellMouseEnter}
              onEdit={handleCellEdit}
              onContextMenu={handleContextMenu}
            />
          </div>
        );
      }
    }
    return cells;
  };

  if (!currentWorksheet) {
    return (
      <div className={`flex items-center justify-center h-full ${className}`}>
        <div className="text-gray-500 text-lg">
          No worksheet selected. Create a new workbook to get started.
        </div>
      </div>
    );
  }

  return (
    <div 
      ref={containerRef}
      className={`relative h-full overflow-auto ${className}`}
      onKeyDown={handleKeyDown}
      onScroll={handleScroll}
      tabIndex={0}
    >
      {/* Formula Bar */}
      {ui.showFormulaBar && <FormulaBar selectedCell={selectedCell} />}
      
      {/* Grid container */}
      <div 
        className="excel-grid custom-scrollbar"
        style={{
          gridTemplateColumns: `50px ${Array.from({ length: visibleRange.endCol - visibleRange.startCol + 1 }, (_, i) => {
            const col = visibleRange.startCol + i;
            const width = currentWorksheet?.columnWidths?.[col] || 100;
            return `${width}px`;
          }).join(' ')}`,
          gridTemplateRows: `24px ${Array.from({ length: visibleRange.endRow - visibleRange.startRow + 1 }, (_, i) => {
            const row = visibleRange.startRow + i;
            const height = currentWorksheet?.rowHeights?.[row] || 24;
            return `${height}px`;
          }).join(' ')}`,
        }}
      >
        {/* Top-left corner */}
        <div className="excel-row-header" style={{ gridColumn: 1, gridRow: 1, zIndex: 30 }} />
        
        {/* Headers and cells */}
        {renderColumnHeaders()}
        {renderRowHeaders()}
        {renderCells()}
      </div>
      
      {/* Context Menu */}
      {contextMenu && (
        <ContextMenu
          position={{ x: contextMenu.x, y: contextMenu.y }}
          cellPosition={contextMenu.cell}
          selectedRange={ui.selectedCells || undefined}
          onClose={() => setContextMenu(null)}
        />
      )}
    </div>
  );
};
