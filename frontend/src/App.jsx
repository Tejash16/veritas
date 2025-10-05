import React, { useState, useEffect } from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import toast, { Toaster } from 'react-hot-toast';

// Components
import Header from './components/Common/Header';
import AuditDashboard from './components/Dashboard/AuditDashboard';
import ErrorBoundary from './components/Common/ErrorBoundary';
import LoadingSpinner from './components/Common/LoadingSpinner';

// Enhanced Components
import InteractiveValidation from './components/Validation/InteractiveValidation';
import EnhancedFileUpload from './components/Upload/EnhancedFileUpload';

// Services
import { authService } from './services/auth';

// Styles
import './App.css';

function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [loading, setLoading] = useState(true);
  const [user, setUser] = useState(null);

  useEffect(() => {
    checkAuthStatus();
  }, []);

  const checkAuthStatus = async () => {
    try {
      const token = localStorage.getItem('veritas_token');
      if (token) {
        const userData = await authService.validateToken(token);
        setUser(userData);
        setIsAuthenticated(true);
      }
    } catch (error) {
      // console.error('Auth check failed:', error);
      localStorage.removeItem('veritas_token');
    } finally {
      setLoading(false);
    }
  };

  const handleLogin = async (credentials) => {
    try {
      setLoading(true);
      const { token, user: userData } = await authService.login(credentials);
      localStorage.setItem('veritas_token', token);
      setUser(userData);
      setIsAuthenticated(true);
      toast.success(`Welcome back, ${userData.username}!`);
    } catch (error) {
      toast.error('Login failed: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  const handleLogout = () => {
    localStorage.removeItem('veritas_token');
    setUser(null);
    setIsAuthenticated(false);
    toast.success('Logged out successfully');
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <LoadingSpinner size="xl" text="Loading Veritas AI Auditor..." />
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex flex-col justify-center py-12 sm:px-6 lg:px-8">
        <div className="sm:mx-auto sm:w-full sm:max-w-md">
          <div className="text-center">
            <h1 className="text-4xl font-extrabold text-gray-900">Veritas - Document Auditor</h1>
            <p className="text-lg text-blue-600 font-semibold">Powered by Orbicle Labs </p>
          </div>
          <h2 className="mt-6 text-center text-2xl font-bold text-gray-900">
            Sign in to your account
          </h2>
          <p className="mt-2 text-center text-sm text-gray-600">
            Created with ❤️ by <a href="https://docrack.ai" className="font-medium text-blue-600 hover:text-blue-500">Orbicle Labs</a>
          </p>
        </div>

        <div className="mt-8 sm:mx-auto sm:w-full sm:max-w-md">
          <div className="bg-white py-8 px-4 shadow-xl sm:rounded-lg sm:px-10">
            <LoginForm onLogin={handleLogin} loading={loading} />
            
            {/* Demo Credentials */}
            {/* <div className="mt-6 p-4 bg-blue-50 rounded-md">
              <p className="text-sm font-medium text-blue-900 mb-2">Demo Credentials:</p>
              <div className="text-sm text-blue-800 space-y-1">
                <p><strong>Username:</strong> demo</p>
                <p><strong>Password:</strong> demo123</p>
                <p className="text-xs text-blue-600 mt-2">
                  Or use admin/admin123 for administrator access
                </p>
              </div>
            </div> */}
          </div>
        </div>
        <Toaster position="top-right" />
      </div>
    );
  }

  return (
    <ErrorBoundary>
      <Router>
        <div className="min-h-screen bg-gray-50">
          <Header user={user} onLogout={handleLogout} />
          
          <main className="max-w-7xl mx-auto py-6 sm:px-6 lg:px-8">
            <Routes>
              <Route path="/" element={<Navigate to="/upload" replace />} />
              {/* Enhanced Routes */}
              <Route path="/upload" element={<EnhancedFileUpload />} />
              <Route path="/validation/:sessionId" element={<InteractiveValidation />} />
              <Route path="/audit/:sessionId" element={<AuditDashboard />} />
              <Route path="*" element={<Navigate to="/upload" replace />} />
            </Routes>
          </main> 
          
          <Toaster 
            position="top-right" 
            toastOptions={{
              duration: 4000,
              style: {
                background: '#363636',
                color: '#fff',
              },
              success: {
                duration: 3000,
                iconTheme: {
                  primary: '#4ade80',
                  secondary: '#fff',
                },
              },
              error: {
                duration: 5000,
                iconTheme: {
                  primary: '#ef4444',
                  secondary: '#fff',
                },
              },
            }}
          />
        </div>
      </Router>
    </ErrorBoundary>
  );
}

// Enhanced Login Form Component
const LoginForm = ({ onLogin, loading }) => {
  const [credentials, setCredentials] = useState({ username: '', password: '' });
  const [submitting, setSubmitting] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setSubmitting(true);
    try {
      await onLogin(credentials);
    } finally {
      setSubmitting(false);
    }
  };

  const fillDemoCredentials = () => {
    setCredentials({ username: 'demo', password: 'demo123' });
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <div>
        <label htmlFor="username" className="block text-sm font-medium text-gray-700">
          Username
        </label>
        <div className="mt-1">
          <input
            id="username"
            name="username"
            type="text"
            required
            className="appearance-none block w-full px-3 py-2 border border-gray-300 rounded-md placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
            value={credentials.username}
            onChange={(e) => setCredentials({ ...credentials, username: e.target.value })}
            placeholder="Enter your username"
          />
        </div>
      </div>

      <div>
        <label htmlFor="password" className="block text-sm font-medium text-gray-700">
          Password
        </label>
        <div className="mt-1">
          <input
            id="password"
            name="password"
            type="password"
            required
            className="appearance-none block w-full px-3 py-2 border border-gray-300 rounded-md placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
            value={credentials.password}
            onChange={(e) => setCredentials({ ...credentials, password: e.target.value })}
            placeholder="Enter your password"
          />
        </div>
      </div>

      <div className="flex items-center justify-between">
        <button
          type="button"
          onClick={fillDemoCredentials}
          className="text-sm text-blue-600 hover:text-blue-500"
        >
          Use demo credentials
        </button>
      </div>

      <div>
        <button
          type="submit"
          disabled={submitting || loading}
          className="w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          {submitting ? (
            <LoadingSpinner size="sm" text="Signing in..." />
          ) : (
            'Sign in'
          )}
        </button>
      </div>
    </form>
  );
};

export default App;