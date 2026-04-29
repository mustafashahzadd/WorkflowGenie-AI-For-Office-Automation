/**
 * API Configuration
 * Centralized configuration for API endpoints and settings
 */

export const API_CONFIG = {
  BASE_URL: import.meta.env.VITE_API_BASE_URL || '',
  TIMEOUT: 30000, // 30 seconds (default)
  CHAT_TIMEOUT: 120000, // 2 minutes for AI requests
  
  ENDPOINTS: {
    // Auth
    AUTH_REGISTER: '/api/auth/register',
    AUTH_LOGIN: '/api/auth/login',
    AUTH_VERIFY: '/api/auth/verify',
    
    // Sessions
    SESSIONS_CREATE: '/api/sessions/create',
    SESSIONS_GET: (sessionId: string) => `/api/sessions/${sessionId}`,
    SESSIONS_DELETE: (sessionId: string) => `/api/sessions/${sessionId}`,
    
    // Files
    FILES_UPLOAD: '/api/files/upload',
    FILES_DOWNLOAD: (fileId: string) => `/api/files/download/${fileId}`,
    FILES_INFO: (fileId: string) => `/api/files/info/${fileId}`,
    FILES_LIST: '/api/files/list',
    FILES_SAVE: (fileId: string) => `/api/files/save/${fileId}`,
    
    // Chat
    CHAT_MESSAGE: '/api/chat/message',
    
    // History
    HISTORY_SESSIONS: '/api/history/sessions',
    HISTORY_SESSION_FULL: (sessionId: string) => `/api/history/session/${sessionId}/full`,
  },
  
  STORAGE_KEYS: {
    TOKEN: 'workflowgenie_token',
    USER: 'workflowgenie_user',
    CURRENT_SESSION: 'workflowgenie_current_session',
  },
};

export default API_CONFIG;
