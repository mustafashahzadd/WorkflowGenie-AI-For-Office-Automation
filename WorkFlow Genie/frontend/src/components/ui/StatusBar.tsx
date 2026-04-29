/**
 * Status bar component showing spreadsheet information and AI status
 * Displays current cell info, worksheet statistics, and operation status
 */

import React from 'react';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import { Activity, Minus, Plus, Monitor, LayoutGrid, BookOpen } from 'lucide-react';

interface StatusBarProps {
  className?: string;
}

/**
 * Main status bar component
 */
export const StatusBar: React.FC<StatusBarProps> = ({ className = '' }) => {
  const {
    currentOperations,
    ui,
  } = useSpreadsheetStore();

  const getActiveOperationsCount = () => {
    return currentOperations.filter(op => 
      op.status === 'pending' || op.status === 'in-progress'
    ).length;
  };

  const getSelectedRangeInfo = () => {
    if (!ui.selectedCells) return '';
    
    const rowCount = ui.selectedCells.endRow - ui.selectedCells.startRow + 1;
    const colCount = ui.selectedCells.endCol - ui.selectedCells.startCol + 1;
    const totalCells = rowCount * colCount;
    
    if (totalCells === 1) return '';
    
    return `${rowCount}R x ${colCount}C (${totalCells} cells selected)`;
  };

  return (
    <div className={`bg-[#107c41] text-white px-4 h-6 flex items-center justify-between text-xs select-none ${className}`}>
      {/* Left section - Status */}
      <div className="flex items-center space-x-4">
        <span className="font-medium">Ready</span>
        
        {getSelectedRangeInfo() && (
          <span className="text-white/90">{getSelectedRangeInfo()}</span>
        )}
        
        {ui.isAIProcessing && (
            <div className="flex items-center gap-1">
              <Activity className="w-3 h-3 animate-pulse" />
              <span>AI Processing...</span>
            </div>
        )}

        {getActiveOperationsCount() > 0 && (
            <span>{getActiveOperationsCount()} operations running</span>
        )}
      </div>

      {/* Right section - View Controls & Zoom */}
      <div className="flex items-center space-x-4">
        {/* View Modes */}
        <div className="flex items-center gap-2 mr-4">
            <button className="p-0.5 hover:bg-white/20 rounded" title="Normal View">
                <LayoutGrid className="w-3 h-3" />
            </button>
            <button className="p-0.5 hover:bg-white/20 rounded" title="Page Layout">
                <BookOpen className="w-3 h-3" />
            </button>
            <button className="p-0.5 hover:bg-white/20 rounded" title="Page Break Preview">
                <Monitor className="w-3 h-3" />
            </button>
        </div>

        {/* Zoom Control */}
        <div className="flex items-center gap-2">
            <button className="hover:bg-white/20 rounded p-0.5"><Minus className="w-3 h-3" /></button>
            <div className="w-24 h-1 bg-white/30 rounded-full relative cursor-pointer">
                <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-2 h-3 bg-white rounded-sm shadow-sm hover:bg-gray-100"></div>
            </div>
            <button className="hover:bg-white/20 rounded p-0.5"><Plus className="w-3 h-3" /></button>
            <span className="w-8 text-right">100%</span>
        </div>
      </div>
    </div>
  );
};
