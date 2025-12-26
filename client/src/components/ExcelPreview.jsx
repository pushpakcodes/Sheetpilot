import React, { useState, useEffect } from 'react';
import { motion } from 'framer-motion';
import { Plus } from 'lucide-react';

const ExcelPreview = ({ data }) => {
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);

  // Handle data structure (legacy array vs new object with sheets)
  const sheets = data?.sheets || (Array.isArray(data) ? [{ name: 'Sheet1', data: data }] : []);
  
  // If no data, return null
  if (!sheets || sheets.length === 0) return null;

  const currentSheet = sheets[activeSheetIndex] || sheets[0];
  const sheetData = currentSheet.data || [];

  // Determine the maximum number of columns to render
  // Ensure at least 26 columns (A-Z) or more if data requires it
  const maxCols = sheetData.reduce((max, row) => Math.max(max, row.length), 26);
  
  // Create an array for column headers
  const headerIndices = Array.from({ length: maxCols }, (_, i) => i);

  // Helper to generate column letters (A, B, ..., Z, AA, AB...)
  const getColumnLabel = (index) => {
    let label = '';
    let i = index;
    while (i >= 0) {
      label = String.fromCharCode((i % 26) + 65) + label;
      i = Math.floor(i / 26) - 1;
    }
    return label;
  };

  return (
    <div className="flex flex-col h-full bg-white dark:bg-slate-900 overflow-hidden font-sans text-[13px]">
      {/* Formula Bar (Visual Only) */}
      <div className="flex items-center gap-2 px-2 py-1 border-b border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800">
        <div className="text-slate-500 font-bold w-8 text-center italic font-serif">fx</div>
        <div className="flex-1 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-600 px-2 py-1 text-slate-700 dark:text-slate-300 h-7 flex items-center shadow-inner"></div>
      </div>

      <div className="flex-1 overflow-auto relative custom-scrollbar">
        <table className="w-full text-left border-collapse cursor-cell">
          <thead className="bg-slate-50 dark:bg-slate-800 sticky top-0 z-20">
            <tr>
              {/* Corner Header (Select All) */}
              <th className="w-10 min-w-[40px] border-r border-b border-slate-300 dark:border-slate-600 bg-slate-100 dark:bg-slate-700 sticky left-0 z-30">
                <div className="w-0 h-0 border-l-[10px] border-l-transparent border-t-[10px] border-t-slate-400 dark:border-t-slate-500 absolute bottom-1 right-1"></div>
              </th>
              
              {/* Column Headers (A, B, C...) */}
              {headerIndices.map((index) => (
                <th key={index} className="px-1 h-6 font-normal text-center text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 border-r border-b border-slate-300 dark:border-slate-600 min-w-[80px] select-none hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors">
                  {getColumnLabel(index)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
             {/* Data Rows */}
            {sheetData.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {/* Row Number (1, 2, 3...) */}
                <td className="w-10 min-w-[40px] text-center text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 border-r border-b border-slate-300 dark:border-slate-600 sticky left-0 z-10 select-none hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors">
                  {rowIndex + 1}
                </td>
                
                {/* Cell Data */}
                {headerIndices.map((colIndex) => {
                  const cell = row[colIndex];
                  return (
                    <td 
                      key={colIndex} 
                      className="px-2 h-6 border-r border-b border-slate-200 dark:border-slate-700 text-slate-900 dark:text-slate-100 whitespace-nowrap overflow-hidden text-ellipsis max-w-[200px] hover:border-double hover:border-blue-500 focus:border-2 focus:border-green-500 outline-none"
                      contentEditable={false} // Placeholder for future editability
                    >
                       {typeof cell === 'object' && cell !== null ? JSON.stringify(cell) : cell}
                    </td>
                  );
                })}
              </tr>
            ))}
            {/* Empty filler rows to fill space */}
            {Array.from({ length: Math.max(0, 50 - sheetData.length) }).map((_, i) => (
               <tr key={`empty-${i}`}>
                 <td className="w-10 min-w-[40px] text-center text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 border-r border-b border-slate-300 dark:border-slate-600 sticky left-0 z-10 select-none">
                    {sheetData.length + i + 1}
                 </td>
                 {headerIndices.map((_, j) => (
                   <td key={j} className="border-r border-b border-slate-200 dark:border-slate-700 h-6"></td>
                 ))}
               </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Sheet Tabs */}
      <div className="flex items-center gap-0 bg-slate-100 dark:bg-slate-800 border-t border-slate-300 dark:border-slate-700 px-1 h-9 overflow-x-auto custom-scrollbar">
        <div className="flex items-center gap-4 px-2 text-slate-400 flex-shrink-0">
           <div className="flex gap-1">
             <button className="hover:bg-slate-200 rounded p-0.5"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M15 18l-6-6 6-6"/></svg></button>
             <button className="hover:bg-slate-200 rounded p-0.5"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M9 18l6-6-6-6"/></svg></button>
           </div>
        </div>
        
        {sheets.map((sheet, index) => (
          <button 
            key={index}
            onClick={() => setActiveSheetIndex(index)}
            className={`flex items-center gap-2 px-4 h-[calc(100%-4px)] mt-[2px] font-medium text-xs shadow-sm rounded-t-sm border-t-2 transition-colors whitespace-nowrap
              ${activeSheetIndex === index 
                ? 'bg-white dark:bg-slate-900 text-green-700 dark:text-green-500 border-green-600' 
                : 'bg-slate-200 dark:bg-slate-800 text-slate-600 dark:text-slate-400 border-transparent hover:bg-slate-300 dark:hover:bg-slate-700'
              }`}
          >
             {sheet.name}
          </button>
        ))}

        <button className="flex items-center justify-center w-8 h-8 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-500 dark:text-slate-400 rounded-full ml-1 flex-shrink-0">
           <Plus className="h-4 w-4" />
        </button>
      </div>
    </div>
  );
};

export default ExcelPreview;
