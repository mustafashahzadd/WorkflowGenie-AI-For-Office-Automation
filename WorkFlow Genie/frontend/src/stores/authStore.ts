/**
 * Authentication Store
 * Zustand store for managing authentication state
 */

import { create } from 'zustand';
import { devtools } from 'zustand/middleware';
import { authService, type User } from '../services/authService';

interface AuthState {
  user: User | null;
  isAuthenticated: boolean;
  isLoading: boolean;
  error: string | null;

  // Actions
  login: (username: string, password: string) => Promise<void>;
  register: (username: string, password: string) => Promise<void>;
  logout: () => void;
  verifyAuth: () => Promise<void>;
  clearError: () => void;
}

export const useAuthStore = create<AuthState>()(
  devtools(
    (set) => ({
      user: authService.getCurrentUser(),
      isAuthenticated: authService.isAuthenticated(),
      isLoading: false, // never block on startup
      error: null,

      login: async (username: string, password: string) => {
        set({ isLoading: true, error: null });
        try {
          const response = await authService.login({ username, password });
          set({
            user: {
              user_id: response.user_id,
              username: response.username,
            },
            isAuthenticated: true,
            isLoading: false,
            error: null,
          });
        } catch (error: any) {
          set({
            error: error.message || 'Login failed',
            isLoading: false,
            isAuthenticated: false,
          });
          throw error;
        }
      },

      register: async (username: string, password: string) => {
        set({ isLoading: true, error: null });
        try {
          const response = await authService.register({ username, password });
          set({
            user: {
              user_id: response.user_id,
              username: response.username,
            },
            isAuthenticated: true,
            isLoading: false,
            error: null,
          });
        } catch (error: any) {
          set({
            error: error.message || 'Registration failed',
            isLoading: false,
            isAuthenticated: false,
          });
          throw error;
        }
      },

      logout: () => {
        authService.logout();
        set({
          user: null,
          isAuthenticated: false,
          error: null,
        });
      },

      verifyAuth: async () => {
        if (authService.isAuthenticated()) {
          const user = authService.getCurrentUser();
          set({ isAuthenticated: true, user, isLoading: false });
          return;
        }
        set({ isAuthenticated: false, user: null, isLoading: false });
      },

      clearError: () => set({ error: null }),
    }),
    { name: 'AuthStore' }
  )
);
