/**
 * Toolbar component providing Excel-like interface controls
 * Features file operations, AI controls, and spreadsheet tools
 */

import React, { useState, useRef } from 'react';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import { useSessionStore } from '../../stores/sessionStore';
import { useAuthStore } from '../../stores/authStore';
import { importExcelToFortuneSheets, fortuneSheetsToBlob } from '../../utils/excelImport';
import { motion } from 'framer-motion';
import { 
  MessageSquare, 
  Download,
  Grid, Plus, Bot, LogOut, Save
} from 'lucide-react';

interface ToolbarProps {
  className?: string;
}

/**
 * Main toolbar component
 */
export const Toolbar: React.FC<ToolbarProps> = ({ className = '' }) => {
  const {
    fortuneSheets,
    toggleSidebar,
    loadFortuneSheets,
  } = useSpreadsheetStore();

  const { 
    currentFilename, 
    currentFileId,
    isSaving,
    saveCurrentFile,
    clearCurrentSession, 
    createSession, 
    setFilename 
  } = useSessionStore();
  const { logout } = useAuthStore();

  const [isEditingTitle, setIsEditingTitle] = useState(false);
  const [editedTitle, setEditedTitle] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      try {
        const sheets = await importExcelToFortuneSheets(file);
        loadFortuneSheets(sheets);
      } catch (error) {
        console.error('Error importing file:', error);
        alert('Failed to import Excel file. Please try again.');
      }
    }
  };

  const handleExport = async () => {
    if (fortuneSheets) {
      try {
        const blob = await fortuneSheetsToBlob(fortuneSheets);
        const filename = (currentFilename || 'spreadsheet').replace(/\.xlsx?$/, '') + '.xlsx';
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
      } catch (error) {
        console.error('Error exporting file:', error);
        alert('Failed to export Excel file. Please try again.');
      }
    }
  };

  const handleSave = async () => {
    if (!currentFileId || !fortuneSheets) {
      alert('No file to save');
      return;
    }
    try {
      await saveCurrentFile();
    } catch (error: any) {
      console.error('Error saving file:', error);
      alert(error.message || 'Failed to save file. Please try again.');
    }
  };

  const handleLogout = () => {
    logout();
  };

  const handleNewSession = async () => {
    try {
      clearCurrentSession();
      await createSession();
    } catch (error) {
      console.error('Failed to create new session:', error);
    }
  };

  const handleTitleClick = () => {
    setIsEditingTitle(true);
    setEditedTitle(currentFilename || 'Untitled spreadsheet');
  };

  const handleTitleSubmit = () => {
    setIsEditingTitle(false);
    if (editedTitle && editedTitle.trim() !== '') {
      setFilename(editedTitle.trim());
    }
  };

  const handleTitleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleTitleSubmit();
    } else if (e.key === 'Escape') {
      setIsEditingTitle(false);
      setEditedTitle('');
    }
  };

  return (
    <div className={`flex flex-col bg-white ${className}`}>
      {/* Top Bar: Logo, Title, Actions */}
      <div className="flex items-center px-4 py-2 border-b border-gray-200">
        {/* Logo */}
        <div className="flex items-center justify-center w-8 h-8 bg-[#0F9D58] rounded text-white mr-4 cursor-pointer">
          <Grid size={20} />
        </div>
        
        <div className="flex flex-col">
          <div className="flex items-center gap-2 mb-1">
            {isEditingTitle ? (
              <input
                type="text"
                value={editedTitle}
                onChange={(e) => setEditedTitle(e.target.value)}
                onBlur={handleTitleSubmit}
                onKeyDown={handleTitleKeyDown}
                className="text-lg text-gray-700 px-1 border border-gray-300 rounded outline-none bg-white"
                autoFocus
              />
            ) : (
              <span 
                className="text-lg text-gray-700 px-1 hover:border hover:border-gray-300 rounded cursor-text"
                onClick={handleTitleClick}
                title="Click to edit filename"
              >
                {currentFilename || 'Untitled spreadsheet'}
              </span>
            )}
            <div className="flex gap-2 ml-2">
                <button className="p-1 hover:bg-gray-100 rounded-full" title="Star">
                    <svg className="w-4 h-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11.049 2.927c.3-.921 1.603-.921 1.902 0l1.519 4.674a1 1 0 00.95.69h4.915c.969 0 1.371 1.24.588 1.81l-3.976 2.888a1 1 0 00-.363 1.118l1.518 4.674c.3.922-.755 1.688-1.538 1.118l-3.976-2.888a1 1 0 00-1.176 0l-3.976 2.888c-.783.57-1.838-.197-1.538-1.118l1.518-4.674a1 1 0 00-.363-1.118l-3.976-2.888c-.784-.57-.38-1.81.588-1.81h4.914a1 1 0 00.951-.69l1.519-4.674z" /></svg>
                </button>
            </div>
          </div>
        </div>

        <div className="ml-auto flex items-center gap-4">
            <motion.button 
              onClick={handleSave}
              disabled={!fortuneSheets || !currentFileId || isSaving}
              className="px-3 py-2 bg-gradient-to-r from-green-600 to-green-700 text-white rounded-lg shadow-lg hover:shadow-xl transition-all duration-200 flex items-center gap-2 font-medium text-sm border border-green-600/20 disabled:opacity-50 disabled:cursor-not-allowed"
              title="Save file to server"
              whileHover={{ scale: (!fortuneSheets || !currentFileId || isSaving) ? 1 : 1.05 }}
              whileTap={{ scale: (!fortuneSheets || !currentFileId || isSaving) ? 1 : 0.95 }}
            >
              <Save size={16} className={isSaving ? 'animate-pulse' : ''} />
              <span className="hidden sm:inline">{isSaving ? 'Saving...' : 'Save'}</span>
            </motion.button>
            <motion.button 
              onClick={handleExport}
              disabled={!fortuneSheets}
              className="px-3 py-2 bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-lg shadow-lg hover:shadow-xl transition-all duration-200 flex items-center gap-2 font-medium text-sm border border-blue-600/20 disabled:opacity-50 disabled:cursor-not-allowed"
              title="Download Excel"
              whileHover={{ scale: !fortuneSheets ? 1 : 1.05 }}
              whileTap={{ scale: !fortuneSheets ? 1 : 0.95 }}
            >
              <Download size={16} />
              <span className="hidden sm:inline">Download</span>
            </motion.button>
            <motion.button 
              onClick={toggleSidebar} 
              className="px-3 py-2 bg-gradient-to-r from-[#107c41] to-[#0c5f31] text-white rounded-lg shadow-lg hover:shadow-xl transition-all duration-200 flex items-center gap-2 font-medium text-sm border border-[#107c41]/20" 
              title="Ask AI Assistant"
              whileHover={{ scale: 1.05 }}
              whileTap={{ scale: 0.95 }}
            >
              <Bot size={16} />
              <span className="hidden sm:inline">AI Assistant</span>
            </motion.button>
            <button className="p-2 hover:bg-gray-100 rounded-full" title="Comment history">
                <MessageSquare size={20} className="text-gray-600" />
            </button>
            <motion.button 
              onClick={handleNewSession}
              className="px-3 py-2 bg-gradient-to-r from-[#107c41] to-[#0c5f31] text-white rounded-lg shadow-lg hover:shadow-xl transition-all duration-200 flex items-center gap-2 font-medium text-sm border border-[#107c41]/20"
              title="New Session"
              whileHover={{ scale: 1.05 }}
              whileTap={{ scale: 0.95 }}
            >
              <Plus size={16} />
              <span className="hidden sm:inline">New</span>
            </motion.button>
            <motion.button 
              onClick={handleLogout}
              className="px-3 py-2 bg-gradient-to-r from-gray-600 to-gray-700 text-white rounded-lg shadow-lg hover:shadow-xl transition-all duration-200 flex items-center gap-2 font-medium text-sm border border-gray-600/20"
              title="Logout"
              whileHover={{ scale: 1.05 }}
              whileTap={{ scale: 0.95 }}
            >
              <LogOut size={16} />
              <span className="hidden sm:inline">Logout</span>
            </motion.button>
        </div>
      </div>
      
      {/* Hidden file input */}
      <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx, .xls"
          className="hidden"
          onChange={handleFileUpload}
      />
    </div>
  );
};