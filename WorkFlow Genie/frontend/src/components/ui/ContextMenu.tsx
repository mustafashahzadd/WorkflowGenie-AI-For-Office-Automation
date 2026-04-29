/**
 * Context Menu component for right-click actions on cells
 * Provides Excel-like context menu with common operations
 */

import React, { useEffect, useRef } from 'react';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import type { CellPosition, CellRange } from '../../types/index';

interface ContextMenuProps {
  position: { x: number; y: number };
  cellPosition: CellPosition;
  selectedRange?: CellRange;
  onClose: () => void;
}

export const ContextMenu: React.FC<ContextMenuProps> = ({
  position,
  cellPosition,
  selectedRange,
  onClose,
}) => {
  const menuRef = useRef<HTMLDivElement>(null);
  const {
    copy,
    cut,
    paste,
    setCellValue,
    insertRow,
    deleteRow,
    insertColumn,
    deleteColumn,
    mergeCells,
    unmergeCells,
    setCellFormat,
    addComment,
  } = useSpreadsheetStore();

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        onClose();
      }
    };

    const handleEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose();
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    document.addEventListener('keydown', handleEscape);

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleEscape);
    };
  }, [onClose]);

  const handleCopy = () => {
    copy(selectedRange);
    onClose();
  };

  const handleCut = () => {
    cut(selectedRange);
    onClose();
  };

  const handlePaste = () => {
    paste(cellPosition);
    onClose();
  };

  const handleDelete = () => {
    if (selectedRange) {
      for (let row = selectedRange.startRow; row <= selectedRange.endRow; row++) {
        for (let col = selectedRange.startCol; col <= selectedRange.endCol; col++) {
          setCellValue({ row, col }, { value: null });
        }
      }
    } else {
      setCellValue(cellPosition, { value: null });
    }
    onClose();
  };

  const handleInsertRow = () => {
    insertRow(cellPosition.row);
    onClose();
  };

  const handleDeleteRow = () => {
    deleteRow(cellPosition.row);
    onClose();
  };

  const handleInsertColumn = () => {
    insertColumn(cellPosition.col);
    onClose();
  };

  const handleDeleteColumn = () => {
    deleteColumn(cellPosition.col);
    onClose();
  };

  const handleMergeCells = () => {
    if (selectedRange) {
      mergeCells(selectedRange);
    }
    onClose();
  };

  const handleUnmergeCells = () => {
    if (selectedRange) {
      unmergeCells(selectedRange);
    }
    onClose();
  };

  const handleFormatBold = () => {
    setCellFormat(cellPosition, { bold: true });
    onClose();
  };

  const handleFormatItalic = () => {
    setCellFormat(cellPosition, { italic: true });
    onClose();
  };

  const handleAddComment = () => {
    const commentText = prompt('Enter comment:');
    if (commentText) {
      addComment(cellPosition, {
        text: commentText,
        author: 'User',
        timestamp: new Date(),
      });
    }
    onClose();
  };

  const menuStyle: React.CSSProperties = {
    position: 'fixed',
    top: position.y,
    left: position.x,
    zIndex: 1000,
  };

  return (
    <div
      ref={menuRef}
      style={menuStyle}
      className="bg-white rounded-md shadow-lg border border-gray-200 py-1 min-w-[200px]"
    >
      <MenuItem icon="📋" label="Copy" shortcut="Ctrl+C" onClick={handleCopy} />
      <MenuItem icon="✂️" label="Cut" shortcut="Ctrl+X" onClick={handleCut} />
      <MenuItem icon="📄" label="Paste" shortcut="Ctrl+V" onClick={handlePaste} />
      <MenuDivider />
      <MenuItem icon="🗑️" label="Delete" shortcut="Del" onClick={handleDelete} />
      <MenuDivider />
      <MenuItem icon="➕" label="Insert Row" onClick={handleInsertRow} />
      <MenuItem icon="➖" label="Delete Row" onClick={handleDeleteRow} />
      <MenuItem icon="➕" label="Insert Column" onClick={handleInsertColumn} />
      <MenuItem icon="➖" label="Delete Column" onClick={handleDeleteColumn} />
      <MenuDivider />
      {selectedRange && selectedRange.startRow !== selectedRange.endRow || selectedRange && selectedRange.startCol !== selectedRange.endCol ? (
        <>
          <MenuItem icon="🔗" label="Merge Cells" onClick={handleMergeCells} />
          <MenuItem icon="⛓️" label="Unmerge Cells" onClick={handleUnmergeCells} />
          <MenuDivider />
        </>
      ) : null}
      <MenuItem icon="B" label="Bold" shortcut="Ctrl+B" onClick={handleFormatBold} />
      <MenuItem icon="I" label="Italic" shortcut="Ctrl+I" onClick={handleFormatItalic} />
      <MenuDivider />
      <MenuItem icon="💬" label="Add Comment" onClick={handleAddComment} />
    </div>
  );
};

interface MenuItemProps {
  icon?: string;
  label: string;
  shortcut?: string;
  onClick: () => void;
  disabled?: boolean;
}

const MenuItem: React.FC<MenuItemProps> = ({ icon, label, shortcut, onClick, disabled }) => {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`w-full px-3 py-2 text-left text-sm flex items-center justify-between hover:bg-gray-100 ${
        disabled ? 'opacity-50 cursor-not-allowed' : 'cursor-pointer'
      }`}
    >
      <span className="flex items-center gap-2">
        {icon && <span className="w-4 text-center">{icon}</span>}
        <span>{label}</span>
      </span>
      {shortcut && <span className="text-xs text-gray-400">{shortcut}</span>}
    </button>
  );
};

const MenuDivider: React.FC = () => {
  return <div className="my-1 border-t border-gray-200" />;
};
