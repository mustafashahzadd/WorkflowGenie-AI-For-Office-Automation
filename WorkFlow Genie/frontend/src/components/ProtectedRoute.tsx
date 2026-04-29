/**
 * Protected Route Component
 * Redirects to login if user is not authenticated
 */

import React, { useEffect } from 'react';
import { Navigate } from 'react-router-dom';
import { useAuthStore } from '../stores/authStore';

interface ProtectedRouteProps {
  children: React.ReactNode;
}

export const ProtectedRoute: React.FC<ProtectedRouteProps> = ({ children }) => {
  const { isAuthenticated, verifyAuth } = useAuthStore();

  useEffect(() => {
    // Verify authentication on mount
    verifyAuth();
  }, [verifyAuth]);

  if (!isAuthenticated) {
    return <Navigate to="/" replace />;
  }

  return <>{children}</>;
};
