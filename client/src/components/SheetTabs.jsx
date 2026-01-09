import React from 'react';
import { motion as Motion } from 'framer-motion';

/**
 * SheetTabs - Excel-style sheet tabs component
 * 
 * Features:
 * - Display all sheets from workbook metadata
 * - Active sheet highlighting
 * - Click to switch sheets
 * - Excel-style appearance
 */
const SheetTabs = ({ sheets = [], activeSheetId = 0, onSheetChange }) => {
  if (!sheets || sheets.length === 0) {
    return null;
  }

  return (
    <div className="flex items-end bg-slate-100 dark:bg-slate-800 border-t border-slate-300 dark:border-slate-700 overflow-x-auto">
      <div className="flex items-end min-w-full">
        {sheets.map((sheet) => {
          const isActive = sheet.sheetId === activeSheetId;
          
          return (
            <Motion.button
              key={sheet.sheetId}
              onClick={() => onSheetChange(sheet.sheetId)}
              className={`
                px-4 py-2 text-sm font-medium whitespace-nowrap
                border-t-2 border-l border-r border-slate-300 dark:border-slate-600
                transition-colors duration-150
                ${isActive 
                  ? 'bg-white dark:bg-slate-900 text-blue-600 dark:text-blue-400 border-t-blue-600 dark:border-t-blue-400 z-10' 
                  : 'bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-400 hover:bg-slate-300 dark:hover:bg-slate-600'
                }
              `}
              whileHover={{ y: isActive ? 0 : -2 }}
              whileTap={{ scale: 0.98 }}
            >
              {sheet.name}
            </Motion.button>
          );
        })}
        
        {/* Add sheet button (optional) */}
        <button
          className="px-3 py-2 text-slate-400 dark:text-slate-500 hover:text-slate-600 dark:hover:text-slate-300 transition-colors"
          title="Add new sheet"
          onClick={() => {
            // TODO: Implement add sheet functionality
            console.log('Add sheet clicked');
          }}
        >
          <svg 
            className="w-5 h-5" 
            fill="none" 
            stroke="currentColor" 
            viewBox="0 0 24 24"
          >
            <path 
              strokeLinecap="round" 
              strokeLinejoin="round" 
              strokeWidth={2} 
              d="M12 4v16m8-8H4" 
            />
          </svg>
        </button>
      </div>
    </div>
  );
};

export default SheetTabs;


