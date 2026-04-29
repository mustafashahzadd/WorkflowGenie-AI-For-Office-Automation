/**
 * File Service
 * Handles file upload, download, and management
 */

import { apiClient } from './apiClient';
import API_CONFIG from '../config/api';

export interface FileUploadResponse {
  file_id: string;
  filename: string;
  filepath: string;
  message: string;
}

export interface FileInfo {
  file_id: string;
  filename: string;
  filepath: string;
  uploaded_at: string;
  session_id: string;
  exists_on_disk: boolean;
}

export interface FileListItem {
  file_id: string;
  filename: string;
  uploaded_at: string;
}

export interface FileSaveResponse {
  message: string;
  file_id: string;
  filename: string;
  size_bytes: number;
  validation?: {
    signature_verified: boolean;
    structure_verified: boolean;
    integrity_verified: boolean;
  };
}

export class FileService {
  /**
   * Upload Excel file to session (ONE per session)
   */
  async uploadFile(file: File, sessionId: string): Promise<FileUploadResponse> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('session_id', sessionId);

    return apiClient.post<FileUploadResponse>(
      API_CONFIG.ENDPOINTS.FILES_UPLOAD,
      formData,
      true // isFormData
    );
  }

  /**
   * Download Excel file
   */
  async downloadFile(fileId: string): Promise<Blob> {
    return apiClient.downloadFile(
      API_CONFIG.ENDPOINTS.FILES_DOWNLOAD(fileId)
    );
  }

  /**
   * Get file info
   */
  async getFileInfo(fileId: string): Promise<FileInfo> {
    return apiClient.get<FileInfo>(
      API_CONFIG.ENDPOINTS.FILES_INFO(fileId)
    );
  }

  /**
   * List all user files
   */
  async listFiles(limit: number = 50): Promise<FileListItem[]> {
    return apiClient.get<FileListItem[]>(
      API_CONFIG.ENDPOINTS.FILES_LIST,
      { limit }
    );
  }

  /**
   * Download file and create download link
   */
  async downloadFileToUser(fileId: string, filename?: string): Promise<void> {
    const blob = await this.downloadFile(fileId);
    
    // Create download link
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename || `file_${fileId}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  }

  /**
   * Get file as ArrayBuffer for xlsx parsing
   * Includes cache-busting to ensure fresh file content
   */
  async getFileAsArrayBuffer(fileId: string): Promise<ArrayBuffer> {
    // Add timestamp to bust browser cache and ensure we get the latest file
    const cacheBuster = Date.now();
    const endpoint = `${API_CONFIG.ENDPOINTS.FILES_DOWNLOAD(fileId)}?_=${cacheBuster}`;
    const blob = await apiClient.get<Blob>(endpoint);
    return blob.arrayBuffer();
  }

  /**
   * Save/update file content for an existing file ID
   */
  async saveFile(fileId: string, file: File): Promise<FileSaveResponse> {
    const formData = new FormData();
    formData.append('file', file);

    return apiClient.put<FileSaveResponse>(
      API_CONFIG.ENDPOINTS.FILES_SAVE(fileId),
      formData,
      true // isFormData
    );
  }
}

// Export singleton instance
export const fileService = new FileService();
