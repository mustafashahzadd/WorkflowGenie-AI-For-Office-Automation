/**
 * Authentication Service
 * Handles user registration, login, and token verification
 */

import { apiClient } from './apiClient';
import API_CONFIG from '../config/api';

export interface RegisterRequest {
  username: string;
  password: string;
}

export interface LoginRequest {
  username: string;
  password: string;
}

export interface AuthResponse {
  user_id: string;
  username: string;
  token: string;
  message: string;
}

export interface VerifyResponse {
  user_id: string;
  username: string;
  valid: boolean;
}

export interface User {
  user_id: string;
  username: string;
}

export class AuthService {
  /**
   * Register a new user
   */
  async register(data: RegisterRequest): Promise<AuthResponse> {
    const response = await apiClient.post<AuthResponse>(
      API_CONFIG.ENDPOINTS.AUTH_REGISTER,
      data
    );

    // Store token and user info
    this.storeAuthData(response);

    return response;
  }

  /**
   * Login existing user
   */
  async login(data: LoginRequest): Promise<AuthResponse> {
    const response = await apiClient.post<AuthResponse>(
      API_CONFIG.ENDPOINTS.AUTH_LOGIN,
      data
    );

    // Store token and user info
    this.storeAuthData(response);

    return response;
  }

  /**
   * Verify JWT token
   */
  async verify(): Promise<VerifyResponse> {
    return apiClient.get<VerifyResponse>(API_CONFIG.ENDPOINTS.AUTH_VERIFY);
  }

  /**
   * Logout user
   */
  logout(): void {
    localStorage.removeItem(API_CONFIG.STORAGE_KEYS.TOKEN);
    localStorage.removeItem(API_CONFIG.STORAGE_KEYS.USER);
    localStorage.removeItem(API_CONFIG.STORAGE_KEYS.CURRENT_SESSION);
  }

  /**
   * Check if user is authenticated
   */
  isAuthenticated(): boolean {
    return !!localStorage.getItem(API_CONFIG.STORAGE_KEYS.TOKEN);
  }

  /**
   * Get current user from localStorage
   */
  getCurrentUser(): User | null {
    const userStr = localStorage.getItem(API_CONFIG.STORAGE_KEYS.USER);
    if (!userStr) return null;

    try {
      return JSON.parse(userStr);
    } catch {
      return null;
    }
  }

  /**
   * Get current token
   */
  getToken(): string | null {
    return localStorage.getItem(API_CONFIG.STORAGE_KEYS.TOKEN);
  }

  /**
   * Store authentication data
   */
  private storeAuthData(response: AuthResponse): void {
    localStorage.setItem(API_CONFIG.STORAGE_KEYS.TOKEN, response.token);
    localStorage.setItem(
      API_CONFIG.STORAGE_KEYS.USER,
      JSON.stringify({
        user_id: response.user_id,
        username: response.username,
      })
    );
  }
}

// Export singleton instance
export const authService = new AuthService();
