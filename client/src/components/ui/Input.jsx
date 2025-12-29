import React from 'react';
import { cn } from '../../utils';

const Input = ({ className, label, error, type, ...props }) => {
  // Determine autocomplete attribute based on input type
  const getAutocomplete = () => {
    // If explicitly provided, use it
    if (props.autoComplete) {
      return props.autoComplete;
    }
    if (type === 'password') {
      // Default to current-password, can be overridden with autoComplete="new-password" for registration
      return 'current-password';
    }
    if (type === 'email') {
      return 'email';
    }
    return 'off';
  };

  return (
    <div className="w-full">
      {label && <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">{label}</label>}
      <input
        type={type}
        autoComplete={getAutocomplete()}
        className={cn(
          'w-full px-3 py-2 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg text-slate-900 dark:text-slate-100 placeholder-slate-400 dark:placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-colors',
          error && 'border-red-500 focus:ring-red-500',
          className
        )}
        {...props}
      />
      {error && <p className="mt-1 text-sm text-red-500">{error}</p>}
    </div>
  );
};

export default Input;
