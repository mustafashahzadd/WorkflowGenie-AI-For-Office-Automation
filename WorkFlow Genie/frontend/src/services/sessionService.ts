/**
 * Session Service
 * Handles session creation, retrieval, and management
 */

import { apiClient } from './apiClient';
import API_CONFIG from '../config/api';

export interface Session {
  session_id: string;
  user_id: string;
  created_at: string;
  updated_at: string;
  is_active: boolean;
  summary: string | null;
  file_id: string | null;
  filename: string | null;
}

export interface SessionListItem {
  session_id: string;
  created_at: string;
  updated_at: string;
  first_prompt: string | null;
  last_response: string | null;
  summary: string | null;
  filename: string | null;
  message_count: number;
  task_count: number;
  is_active: boolean;
}

export interface SessionFullHistory {
  session_id: string;
  created_at: string;
  updated_at: string;
  summary: string | null;
  file: {
    file_id: string | null;
    filename: string | null;
  };
  messages: Array<{
    id: string;
    role: 'user' | 'assistant';
    content: string;
    timestamp: string;
  }>;
  tasks: any[];
}

export class SessionService {
  /**
   * Create a new session
   */
  async createSession(): Promise<{ session_id: string; message: string }> {
    return apiClient.post(API_CONFIG.ENDPOINTS.SESSIONS_CREATE);
  }

  /**
   * Get session details
   */
  async getSession(sessionId: string): Promise<Session> {
    return apiClient.get<Session>(
      API_CONFIG.ENDPOINTS.SESSIONS_GET(sessionId)
    );
  }

  /**
   * Delete session (soft delete)
   */
  async deleteSession(sessionId: string): Promise<{ message: string; session_id: string }> {
    return apiClient.delete(
      API_CONFIG.ENDPOINTS.SESSIONS_DELETE(sessionId)
    );
  }

  /**
   * Get list of sessions (WhatsApp-style preview)
   */
  async listSessions(limit: number = 50, activeOnly: boolean = true): Promise<SessionListItem[]> {
    return apiClient.get<SessionListItem[]>(
      API_CONFIG.ENDPOINTS.HISTORY_SESSIONS,
      { limit, active_only: activeOnly }
    );
  }

  /**
   * Get complete session history
   */
  async getSessionHistory(sessionId: string): Promise<SessionFullHistory> {
    return apiClient.get<SessionFullHistory>(
      API_CONFIG.ENDPOINTS.HISTORY_SESSION_FULL(sessionId)
    );
  }

  /**
   * Store current session ID
   */
  setCurrentSession(sessionId: string): void {
    localStorage.setItem(API_CONFIG.STORAGE_KEYS.CURRENT_SESSION, sessionId);
  }

  /**
   * Get current session ID
   */
  getCurrentSessionId(): string | null {
    return localStorage.getItem(API_CONFIG.STORAGE_KEYS.CURRENT_SESSION);
  }

  /**
   * Clear current session
   */
  clearCurrentSession(): void {
    localStorage.removeItem(API_CONFIG.STORAGE_KEYS.CURRENT_SESSION);
  }
}

// Export singleton instance
export const sessionService = new SessionService();
