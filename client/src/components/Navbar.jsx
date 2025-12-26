import React from 'react';
import { Link, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { Sheet, LogOut, FileSpreadsheet, History } from 'lucide-react';
import Button from './ui/Button';
import ThemeToggle from './ui/ThemeToggle';

const Navbar = () => {
  const { user, logout } = useAuth();
  const location = useLocation();

  if (['/login', '/register'].includes(location.pathname)) return null;

  return (
    <nav className="border-b border-slate-200 dark:border-slate-800 bg-white/80 dark:bg-slate-900/50 backdrop-blur-xl sticky top-0 z-50 transition-colors">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-16">
          <Link to="/" className="flex items-center gap-2 hover:opacity-80 transition-opacity">
            <Sheet className="h-8 w-8 text-blue-500" />
            <span className="text-xl font-bold bg-gradient-to-r from-blue-400 to-indigo-400 bg-clip-text text-transparent">
              SheetPilot
            </span>
          </Link>

          <div className="flex items-center gap-4">
            <button className="p-2 text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 transition-colors" title="History">
               <History className="h-5 w-5" />
            </button>
            <ThemeToggle />
            {user ? (
              <>
                <div className="text-sm text-slate-600 dark:text-slate-400">
                  Welcome, <span className="text-slate-900 dark:text-slate-200 font-medium">{user.username}</span>
                </div>
                <Button variant="ghost" size="sm" onClick={logout}>
                  <LogOut className="h-4 w-4 mr-2" />
                  Logout
                </Button>
              </>
            ) : (
              <>
                <Link to="/login">
                  <Button variant="ghost" size="sm">Login</Button>
                </Link>
                <Link to="/register">
                  <Button size="sm">Get Started</Button>
                </Link>
              </>
            )}
          </div>
        </div>
      </div>
    </nav>
  );
};

export default Navbar;
