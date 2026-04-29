/**
 * API Client
 * Core HTTP client with authentication and error handling
 */

import API_CONFIG from '../config/api';

export interface ApiError {
  message: string;
  status?: number;
  details?: any;
}

export class ApiClient {
  private baseURL: string;
  private timeout: number;

  constructor(baseURL?: string, timeout?: number) {
    this.baseURL = baseURL || API_CONFIG.BASE_URL;
    this.timeout = timeout || API_CONFIG.TIMEOUT;
  }

  /**
   * Get authorization token from localStorage
   */
  private getAuthToken(): string | null {
    return localStorage.getItem(API_CONFIG.STORAGE_KEYS.TOKEN);
  }

  /**
   * Build headers for request
   */
  private buildHeaders(isFormData: boolean = false): Record<string, string> {
    const headers: Record<string, string> = {};

    if (!isFormData) {
      headers['Content-Type'] = 'application/json';
    }

    const token = this.getAuthToken();
    if (token) {
      headers['Authorization'] = `Bearer ${token}`;
    }

    return headers;
  }

  /**
   * Handle API response
   */
  private async handleResponse<T>(response: Response): Promise<T> {
    if (!response.ok) {
      let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
      let errorDetails: any = null;

      try {
        const errorData = await response.json();
        errorMessage = errorData.message || errorData.detail || errorMessage;
        errorDetails = errorData;
      } catch {
        // If error response is not JSON, use status text
      }

      const error: ApiError = {
        message: errorMessage,
        status: response.status,
        details: errorDetails,
      };

      // Handle 401 Unauthorized - clear auth and redirect to login
      if (response.status === 401) {
        // Clear token from localStorage
        localStorage.removeItem(API_CONFIG.STORAGE_KEYS.TOKEN);
        localStorage.removeItem(API_CONFIG.STORAGE_KEYS.USER);
        
        // Redirect to login page if not already there
        if (window.location.pathname !== '/') {
          window.location.href = '/';
        }
      }

      throw error;
    }

    // Handle empty responses (204 No Content)
    if (response.status === 204) {
      return {} as T;
    }

    // Handle blob responses (file downloads)
    if (response.headers.get('content-type')?.includes('application/octet-stream') ||
        response.headers.get('content-type')?.includes('application/vnd.openxmlformats')) {
      return response.blob() as unknown as T;
    }

    return response.json();
  }

  /**
   * Make GET request
   */
  async get<T>(endpoint: string, queryParams?: Record<string, any>): Promise<T> {
    let url = `${this.baseURL}${endpoint}`;

    if (queryParams) {
      const params = new URLSearchParams();
      Object.entries(queryParams).forEach(([key, value]) => {
        if (value !== undefined && value !== null) {
          params.append(key, String(value));
        }
      });
      url += `?${params.toString()}`;
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), this.timeout);

    try {
      const headers = this.buildHeaders();
      // Disable caching for file downloads to ensure fresh content
      if (endpoint.includes('/files/download/')) {
        headers['Cache-Control'] = 'no-cache, no-store, must-revalidate';
        headers['Pragma'] = 'no-cache';
        headers['Expires'] = '0';
      }

      const response = await fetch(url, {
        method: 'GET',
        headers,
        signal: controller.signal,
      });

      return await this.handleResponse<T>(response);
    } catch (error: any) {
      if (error.name === 'AbortError') {
        throw { message: 'Request timeout', status: 408 } as ApiError;
      }
      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  /**
   * Make POST request
   */
  async post<T>(endpoint: string, data?: any, isFormData: boolean = false): Promise<T> {
    const url = `${this.baseURL}${endpoint}`;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), this.timeout);

    try {
      const body = isFormData ? data : JSON.stringify(data);

      const response = await fetch(url, {
        method: 'POST',
        headers: this.buildHeaders(isFormData),
        body: body,
        signal: controller.signal,
      });

      return await this.handleResponse<T>(response);
    } catch (error: any) {
      if (error.name === 'AbortError') {
        throw { message: 'Request timeout', status: 408 } as ApiError;
      }
      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  /**
   * Make PUT request
   */
  async put<T>(endpoint: string, data?: any, isFormData: boolean = false): Promise<T> {
    const url = `${this.baseURL}${endpoint}`;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), this.timeout);

    try {
      const body = isFormData ? data : JSON.stringify(data);

      const response = await fetch(url, {
        method: 'PUT',
        headers: this.buildHeaders(isFormData),
        body: body,
        signal: controller.signal,
      });

      return await this.handleResponse<T>(response);
    } catch (error: any) {
      if (error.name === 'AbortError') {
        throw { message: 'Request timeout', status: 408 } as ApiError;
      }
      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  /**
   * Make DELETE request
   */
  async delete<T>(endpoint: string): Promise<T> {
    const url = `${this.baseURL}${endpoint}`;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), this.timeout);

    try {
      const response = await fetch(url, {
        method: 'DELETE',
        headers: this.buildHeaders(),
        signal: controller.signal,
      });

      return await this.handleResponse<T>(response);
    } catch (error: any) {
      if (error.name === 'AbortError') {
        throw { message: 'Request timeout', status: 408 } as ApiError;
      }
      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  /**
   * Download file as blob
   */
  async downloadFile(endpoint: string): Promise<Blob> {
    return this.get<Blob>(endpoint);
  }
}

// Export singleton instance
export const apiClient = new ApiClient();
