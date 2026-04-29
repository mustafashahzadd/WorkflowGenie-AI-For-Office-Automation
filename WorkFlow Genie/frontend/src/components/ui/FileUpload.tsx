/**
 * File upload component for importing Excel files
 */

import React, { useRef, useState } from 'react';
import { Upload, FileSpreadsheet, X, AlertCircle, CheckCircle } from 'lucide-react';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import { useSessionStore } from '../../stores/sessionStore';
import { motion, AnimatePresence } from 'framer-motion';

interface FileUploadProps {
  className?: string;
}

export const FileUpload: React.FC<FileUploadProps> = ({ className = '' }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  
  const { toggleSidebar, ui } = useSpreadsheetStore();
  const { uploadFile, isUploadingFile, error: sessionError, currentSession, createSession } = useSessionStore();

  const error = localError || sessionError;

  const handleFileSelect = async (file: File) => {
    // Validate file type
    const validExtensions = ['.xlsx', '.xls', '.xlsm'];
    const fileExtension = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));

    if (!validExtensions.includes(fileExtension)) {
      setLocalError('Please upload a valid Excel file (.xlsx, .xls, .xlsm)');
      return;
    }

    // Check file size (max 50MB)
    if (file.size > 50 * 1024 * 1024) {
      setLocalError('File size must be less than 50MB');
      return;
    }

    setLocalError(null);
    setSuccess(false);

    try {
      // Create a new session if one doesn't exist
      if (!currentSession) {
        await createSession();
      }

      // Upload file to backend — loadSession inside uploadFile will load into FortuneSheet
      await uploadFile(file);
      setSuccess(true);
      
      // Open the AI chat sidebar automatically
      if (!ui.sidebarOpen) {
        toggleSidebar();
      }
      
      // Clear success message after 3 seconds
      setTimeout(() => setSuccess(false), 3000);
    } catch (err: any) {
      console.error('Error uploading file:', err);
      setLocalError(err.message || 'Failed to upload file');
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFileSelect(file);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);

    const file = e.dataTransfer.files?.[0];
    if (file) {
      handleFileSelect(file);
    }
  };

  const handleClick = () => {
    fileInputRef.current?.click();
  };

  const clearError = () => {
    setLocalError(null);
  };

  return (
    <div className={`relative ${className}`}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.xlsm"
        onChange={handleFileInputChange}
        className="hidden"
        disabled={isUploadingFile}
      />

      <div
        onClick={handleClick}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={`
          border-2 border-dashed rounded-lg p-8 text-center cursor-pointer
          transition-all duration-200 ease-in-out
          ${isDragging 
            ? 'border-[#107c41] bg-green-50' 
            : 'border-gray-300 hover:border-[#107c41] hover:bg-gray-50'
          }
          ${isUploadingFile ? 'opacity-50 cursor-not-allowed' : ''}
        `}
      >
        {isUploadingFile ? (
          <div className="flex flex-col items-center gap-3">
            <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-[#107c41]"></div>
            <p className="text-gray-600">Uploading file...</p>
            <p className="text-xs text-gray-500">This may take a moment</p>
          </div>
        ) : success ? (
          <motion.div
            initial={{ scale: 0.8, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="flex flex-col items-center gap-3"
          >
            <div className="p-3 bg-green-100 rounded-full">
              <CheckCircle className="w-8 h-8 text-green-600" />
            </div>
            <p className="text-lg font-semibold text-green-700">File uploaded successfully!</p>
          </motion.div>
        ) : (
          <div className="flex flex-col items-center gap-3">
            <div className="p-3 bg-[#107c41] bg-opacity-10 rounded-full">
              {isDragging ? (
                <FileSpreadsheet className="w-8 h-8 text-[#107c41]" />
              ) : (
                <Upload className="w-8 h-8 text-[#107c41]" />
              )}
            </div>
            <div>
              <p className="text-lg font-semibold text-gray-700">
                {isDragging ? 'Drop your Excel file here' : 'Upload Excel File'}
              </p>
              <p className="text-sm text-gray-500 mt-1">
                Drag and drop or click to browse
              </p>
              <p className="text-xs text-gray-400 mt-2">
                Supports .xlsx, .xls, .xlsm (max 50MB)
              </p>
              <p className="text-xs text-amber-600 mt-2 font-medium">
                One file per session
              </p>
            </div>
          </div>
        )}
      </div>

      <AnimatePresence>
        {error && (
          <motion.div
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="mt-4 p-3 bg-red-50 border border-red-200 rounded-lg flex items-start gap-2"
          >
            <AlertCircle className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5" />
            <div className="flex-1">
              <p className="text-sm text-red-700">{error}</p>
            </div>
            <button
              onClick={clearError}
              className="text-red-400 hover:text-red-600 transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
};
