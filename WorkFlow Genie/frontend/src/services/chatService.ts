/**
 * Chat Service
 * Handles chat messages and Excel operations execution
 */

import { apiClient, ApiClient } from './apiClient';
import API_CONFIG from '../config/api';

const chatApiClient = new ApiClient(undefined, API_CONFIG.CHAT_TIMEOUT);

export interface ChatMessageRequest {
  session_id?: string;
  message: string;
  file_id?: string | null;
  active_sheet?: string | null;
  provider?: 'openai' | 'claude' | null;
  use_rag?: boolean | null;
  rag_top_k?: number;
}

export interface OperationResult {
  step: number;
  description: string;
  tool: string;
  status: 'completed' | 'failed';
  result: any;
  message: string;
  error: string | null;
}

export interface ChatMessageResponse {
  session_id: string;
  response: string;
  operations: OperationResult[];
  context: {
    file_id: string;
    sheet_name: string;
    total_operations: number;
    successful: number;
  };
}

export class ChatService {
  /**
   * Send message and execute Excel operations
   */
  async sendMessage(
    sessionId: string,
    message: string,
    options: {
      fileId?: string | null;
      activeSheet?: string | null;
      provider?: 'openai' | 'claude' | null;
      useRag?: boolean | null;
      ragTopK?: number;
    } = {}
  ): Promise<ChatMessageResponse> {
    const { fileId, activeSheet, provider, useRag, ragTopK } = options;

    const request: ChatMessageRequest = {
      session_id: sessionId,
      message,
      file_id: fileId ?? null,
      ...(activeSheet && { active_sheet: activeSheet }),
      ...(provider !== undefined && { provider }),
      ...(useRag !== undefined && { use_rag: useRag }),
      ...(ragTopK !== undefined && { rag_top_k: ragTopK }),
    };

    return chatApiClient.post<ChatMessageResponse>(
      API_CONFIG.ENDPOINTS.CHAT_MESSAGE,
      request
    );
  }
}

// Export singleton instance
export const chatService = new ChatService();
