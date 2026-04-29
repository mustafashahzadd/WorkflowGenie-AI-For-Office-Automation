/**
 * Session Store
 * Zustand store for managing session state
 */

import { create } from 'zustand';
import { devtools } from 'zustand/middleware';
import { sessionService, type Session, type SessionListItem } from '../services/sessionService';
import { fileService } from '../services/fileService';
import { chatService, type ChatMessageResponse } from '../services/chatService';
import { importExcelToFortuneSheets, fortuneSheetsToBlob } from '../utils/excelImport';
import { useSpreadsheetStore } from './spreadsheetStore';

interface SessionState {
  // Current session
  currentSession: Session | null;
  sessions: SessionListItem[];
  
  // File data
  currentFileId: string | null;
  currentFilename: string | null;
  
  // Chat history
  chatHistory: Array<{
    role: 'user' | 'assistant';
    content: string;
    timestamp: string;
    operations?: any[];
  }>;
  
  // Loading states
  isLoadingSessions: boolean;
  isLoadingSession: boolean;
  isUploadingFile: boolean;
  isSendingMessage: boolean;
  isSaving: boolean;
  
  // Error state
  error: string | null;

  // Actions
  createSession: () => Promise<string>;
  loadSession: (sessionId: string) => Promise<void>;
  loadSessions: (limit?: number, activeOnly?: boolean) => Promise<void>;
  deleteSession: (sessionId: string) => Promise<void>;
  uploadFile: (file: File) => Promise<void>;
  saveCurrentFile: () => Promise<void>;
  sendMessage: (message: string, options?: {
    sheetName?: string;
    provider?: 'openai' | 'claude' | null;
    useRag?: boolean | null;
    ragTopK?: number;
  }) => Promise<ChatMessageResponse>;
  clearCurrentSession: () => void;
  clearError: () => void;
  setFilename: (filename: string) => void;
}

export const useSessionStore = create<SessionState>()(
  devtools(
    (set, get) => ({
      currentSession: null,
      sessions: [],
      currentFileId: null,
      currentFilename: null,
      chatHistory: [],
      isLoadingSessions: false,
      isLoadingSession: false,
      isSaving: false,
      isUploadingFile: false,
      isSendingMessage: false,
      error: null,

      createSession: async () => {
        set({ error: null });
        try {
          const response = await sessionService.createSession();
          sessionService.setCurrentSession(response.session_id);

          // Set a minimal session object directly — no need to fetch history for a brand new session
          const now = new Date().toISOString();
          set({
            currentSession: {
              session_id: response.session_id,
              user_id: '',
              created_at: now,
              updated_at: now,
              is_active: true,
              summary: null,
              file_id: null,
              filename: null,
            },
            currentFileId: null,
            currentFilename: null,
            chatHistory: [],
          });

          return response.session_id;
        } catch (error: any) {
          set({ error: error.message || 'Failed to create session' });
          throw error;
        }
      },

      loadSession: async (sessionId: string) => {
        set({ isLoadingSession: true, error: null });
        try {
          const [session, history] = await Promise.all([
            sessionService.getSession(sessionId),
            sessionService.getSessionHistory(sessionId),
          ]);

          sessionService.setCurrentSession(sessionId);

          set({
            currentSession: session,
            currentFileId: session.file_id,
            currentFilename: session.filename,
            chatHistory: history.messages,
            isLoadingSession: false,
          });

          // If session has a file, load it into FortuneSheet
          if (session.file_id) {
            try {
              const arrayBuffer = await fileService.getFileAsArrayBuffer(session.file_id);
              const blob = new Blob([arrayBuffer]);
              const file = new File([blob], session.filename || 'session.xlsx');
              const sheets = await importExcelToFortuneSheets(file);
              useSpreadsheetStore.getState().loadFortuneSheets(sheets);
            } catch (fileError) {
              console.error('Failed to load file for session:', fileError);
            }
          }
        } catch (error: any) {
          set({
            error: error.message || 'Failed to load session',
            isLoadingSession: false,
          });
          throw error;
        }
      },

      loadSessions: async (limit = 50, activeOnly = true) => {
        set({ isLoadingSessions: true, error: null });
        try {
          const sessions = await sessionService.listSessions(limit, activeOnly);
          set({ sessions, isLoadingSessions: false });
        } catch (error: any) {
          set({
            error: error.message || 'Failed to load sessions',
            isLoadingSessions: false,
          });
          throw error;
        }
      },

      deleteSession: async (sessionId: string) => {
        set({ error: null });
        try {
          await sessionService.deleteSession(sessionId);
          
          // Remove from sessions list
          set((state) => ({
            sessions: state.sessions.filter((s) => s.session_id !== sessionId),
          }));

          // Clear current session if it was deleted
          if (get().currentSession?.session_id === sessionId) {
            get().clearCurrentSession();
          }
        } catch (error: any) {
          set({ error: error.message || 'Failed to delete session' });
          throw error;
        }
      },

      uploadFile: async (file: File) => {
        const { currentSession } = get();
        if (!currentSession) {
          throw new Error('No active session');
        }

        set({ isUploadingFile: true, error: null });
        try {
          const response = await fileService.uploadFile(file, currentSession.session_id);
          
          set({
            currentFileId: response.file_id,
            currentFilename: response.filename,
            isUploadingFile: false,
          });

          // Reload session to get updated file info
          await get().loadSession(currentSession.session_id);
        } catch (error: any) {
          set({
            error: error.message || 'Failed to upload file',
            isUploadingFile: false,
          });
          throw error;
        }
      },

      saveCurrentFile: async () => {
        const { currentFileId, currentFilename } = get();
        if (!currentFileId) {
          throw new Error('No file to save');
        }

        set({ isSaving: true, error: null });
        try {
          const fortuneSheets = useSpreadsheetStore.getState().fortuneSheets;
          if (!fortuneSheets) {
            throw new Error('No workbook loaded');
          }

          const blob = await fortuneSheetsToBlob(fortuneSheets);
          const filename = currentFilename || 'spreadsheet.xlsx';
          const file = new File([blob], filename, {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          });

          await fileService.saveFile(currentFileId, file);
          set({ isSaving: false });
        } catch (error: any) {
          set({
            error: error.message || 'Failed to save file',
            isSaving: false,
          });
          throw error;
        }
      },

      sendMessage: async (message: string, options: {
        provider?: 'openai' | 'claude' | null;
        useRag?: boolean | null;
        ragTopK?: number;
      } = {}) => {
        const { currentSession, currentFileId } = get();
        if (!currentSession) {
          throw new Error('No active session');
        }

        // Auto-save before sending message to prevent data loss
        if (currentFileId) {
          try {
            await get().saveCurrentFile();
          } catch (saveError) {
            console.error('Failed to auto-save before sending message:', saveError);
            // Continue anyway - user can manually save later
          }
        }

        // Optimistically add user message
        set((state) => ({
          isSendingMessage: true,
          error: null,
          chatHistory: [
            ...state.chatHistory,
            {
              role: 'user' as const,
              content: message,
              timestamp: new Date().toISOString(),
            },
          ]
        }));

        try {
          // Read active sheet dynamically from the live FortuneSheet data
          const liveSheets = useSpreadsheetStore.getState().fortuneSheets;
          const activeSheet = liveSheets?.find((s: any) => s.status === 1)?.name ?? liveSheets?.[0]?.name ?? undefined;

          const response = await chatService.sendMessage(
            currentSession.session_id,
            message,
            {
              fileId: currentFileId,
              activeSheet,
              provider: options.provider,
              useRag: options.useRag,
              ragTopK: options.ragTopK,
            }
          );

          // Add assistant response to chat history
          set((state) => ({
            chatHistory: [
              ...state.chatHistory,
              {
                role: 'assistant' as const,
                content: response.response,
                timestamp: new Date().toISOString(),
                operations: response.operations,
              },
            ],
            isSendingMessage: false,
          }));

          // Reload the file from backend to get the AI-updated version
          if (currentFileId) {
            try {
              const arrayBuffer = await fileService.getFileAsArrayBuffer(currentFileId);
              const blob = new Blob([arrayBuffer]);
              const file = new File([blob], get().currentFilename || 'session.xlsx');
              const sheets = await importExcelToFortuneSheets(file);
              // Preserve whichever sheet was active before the reload
              const prevSheets = useSpreadsheetStore.getState().fortuneSheets;
              const activeSheetName = prevSheets?.find((s: any) => s.status === 1)?.name;
              if (activeSheetName) {
                let found = false;
                sheets.forEach((s: any) => {
                  s.status = (!found && s.name === activeSheetName) ? (found = true, 1) : 0;
                });
                if (!found) sheets[0].status = 1; // fallback
              }
              useSpreadsheetStore.getState().loadFortuneSheets(sheets);
            } catch (fileError) {
              console.error('Failed to reload file after message:', fileError);
            }
          }

          return response;
        } catch (error: any) {
          // Revert optimistic update on failure by removing the last message
          set((state) => ({
            chatHistory: state.chatHistory.slice(0, -1),
            error: error.message || 'Failed to send message',
            isSendingMessage: false,
          }));
          throw error;
        }
      },

      clearCurrentSession: () => {
        sessionService.clearCurrentSession();
        // Reset fortuneSheets to null so Dashboard shows the upload screen
        useSpreadsheetStore.setState({ fortuneSheets: null });
        set({
          currentSession: null,
          currentFileId: null,
          currentFilename: null,
          chatHistory: [],
        });
      },

      clearError: () => set({ error: null }),

      setFilename: (filename: string) => {
        set({ currentFilename: filename });
      },
    }),
    { name: 'SessionStore' }
  )
);
