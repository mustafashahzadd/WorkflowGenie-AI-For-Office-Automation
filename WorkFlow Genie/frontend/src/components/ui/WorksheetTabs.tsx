/**
 * Worksheet tabs component for navigating between different sheets
 * Provides Excel-like tab interface with add, rename, and delete functionality
 */

import React, { useState } from 'react';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import { Plus, ChevronLeft, ChevronRight } from 'lucide-react';

interface WorksheetTabsProps {
  className?: string;
}

/**
 * Individual worksheet tab component
 */
interface WorksheetTabProps {
  worksheetId: string;
  name: string;
  isActive: boolean;
  onSelect: () => void;
  onRename: (newName: string) => void;
  onDelete: () => void;
  canDelete: boolean;
}

const WorksheetTab: React.FC<WorksheetTabProps> = ({
  worksheetId,
  name,
  isActive,
  onSelect,
  onRename,
  onDelete,
  canDelete,
}) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editName, setEditName] = useState(name);

  const handleDoubleClick = () => {
    setIsEditing(true);
    setEditName(name);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSubmit();
    } else if (e.key === 'Escape') {
      setIsEditing(false);
      setEditName(name);
    }
  };

  const handleSubmit = () => {
    if (editName.trim() && editName !== name) {
      onRename(editName.trim());
    }
    setIsEditing(false);
  };

  const handleBlur = () => {
    handleSubmit();
  };

  return (
    <div
      className={`relative group flex items-center px-4 py-1 cursor-pointer transition-colors select-none ${
        isActive
          ? 'bg-white text-[#0F9D58] font-semibold border-b-2 border-[#0F9D58]'
          : 'bg-[#f1f3f4] hover:bg-[#e8eaed] text-gray-700 border-r border-gray-300'
      }`}
      onClick={onSelect}
      onDoubleClick={handleDoubleClick}
    >
      {isEditing ? (
        <input
          type="text"
          value={editName}
          onChange={(e) => setEditName(e.target.value)}
          onKeyDown={handleKeyDown}
          onBlur={handleBlur}
          className="bg-white border border-[#1a73e8] outline-none text-xs w-20 px-1"
          autoFocus
        />
      ) : (
        <span className="text-xs">{name}</span>
      )}
    </div>
  );
};

/**
 * Main worksheet tabs container
 */
export const WorksheetTabs: React.FC<WorksheetTabsProps> = ({ className = '' }) => {
  const {
    currentWorkbook,
    currentWorksheet,
    createWorksheet,
    deleteWorksheet,
    renameWorksheet,
    setActiveWorksheet,
  } = useSpreadsheetStore();

  if (!currentWorkbook) return null;

  const handleAddWorksheet = () => {
    const name = `Sheet${Object.keys(currentWorkbook.worksheets).length + 1}`;
    createWorksheet(name);
  };

  return (
    <div className={`flex items-center bg-[#f8f9fa] border-t border-gray-300 h-8 ${className}`}>
      {/* Navigation Controls */}
      <div className="flex items-center px-2 gap-1 text-gray-500">
        <button className="p-0.5 hover:bg-gray-200 rounded disabled:opacity-30">
            <ChevronLeft className="w-4 h-4" />
        </button>
        <button className="p-0.5 hover:bg-gray-200 rounded disabled:opacity-30">
            <ChevronRight className="w-4 h-4" />
        </button>
      </div>

      {/* Tabs Container */}
      <div className="flex items-end h-full overflow-x-auto no-scrollbar">
        {Object.values(currentWorkbook.worksheets).map((worksheet) => (
          <WorksheetTab
            key={worksheet.id}
            worksheetId={worksheet.id}
            name={worksheet.name}
            isActive={currentWorksheet?.id === worksheet.id}
            onSelect={() => setActiveWorksheet(worksheet.id)}
            onRename={(newName) => renameWorksheet(worksheet.id, newName)}
            onDelete={() => deleteWorksheet(worksheet.id)}
            canDelete={Object.keys(currentWorkbook.worksheets).length > 1}
          />
        ))}
      </div>

      {/* Add Button */}
      <button
        onClick={handleAddWorksheet}
        className="flex items-center justify-center w-8 h-full hover:bg-gray-200 text-gray-600 transition-colors"
        title="New Sheet"
      >
        <Plus className="w-4 h-4" />
      </button>
    </div>
  );
};
