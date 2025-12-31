import React from 'react';
import { 
  Clipboard, Scissors, Copy, Type, Bold, Italic, Underline, 
  AlignLeft, AlignCenter, AlignRight, Grid, Table, 
  FileSpreadsheet, Save, Undo, Redo, Share, ChevronDown
} from 'lucide-react';

const Ribbon = ({ onUndo, onRedo, onSave, onDownload }) => {
  const tabs = ['Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Help'];
  
  return (
    <div className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 transition-colors flex-shrink-0">
      {/* Top Bar (File menu and Tabs) */}
      <div className="flex items-center px-2 bg-green-700 text-white justify-between">
        <div className="flex items-center">
          <button className="px-4 py-1 hover:bg-green-800 font-semibold transition-colors">File</button>
          {/* Quick Access Toolbar */}
          <div className="flex items-center gap-1 mx-2 border-l border-green-600 pl-2">
            <button onClick={onUndo} className="p-1 hover:bg-green-600 rounded transition-colors" title="Undo">
              <Undo className="h-4 w-4" />
            </button>
            <button onClick={onRedo} className="p-1 hover:bg-green-600 rounded transition-colors" title="Redo">
              <Redo className="h-4 w-4" />
            </button>
            <button onClick={onSave} className="p-1 hover:bg-green-600 rounded transition-colors" title="Save">
              <Save className="h-4 w-4" />
            </button>
            <button onClick={onDownload} className="p-1 hover:bg-green-600 rounded transition-colors" title="Download">
              <FileSpreadsheet className="h-4 w-4" />
            </button>
          </div>
        </div>
        
        {/* Tabs */}
        <div className="flex items-center overflow-x-auto no-scrollbar">
          {tabs.map(tab => (
            <button 
              key={tab} 
              className={`px-4 py-1 text-sm whitespace-nowrap transition-colors ${tab === 'Home' ? 'bg-white dark:bg-slate-900 text-slate-900 dark:text-white font-medium rounded-t-md mt-1' : 'hover:bg-green-600'}`}
            >
              {tab}
            </button>
          ))}
        </div>
        
        <div className="w-4"></div> {/* Spacer */}
      </div>

      {/* Ribbon Toolbar (Home Tab Content) */}
      <div className="flex items-stretch p-2 gap-2 h-24 overflow-x-auto text-slate-700 dark:text-slate-200">
        
        {/* Clipboard Group */}
        <div className="flex flex-col gap-1 pr-2 border-r border-slate-200 dark:border-slate-700">
          <button className="flex items-center gap-1 p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded">
            <Clipboard className="h-5 w-5" />
            <span className="text-xs">Paste</span>
          </button>
          <div className="flex gap-1">
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded" title="Cut"><Scissors className="h-4 w-4" /></button>
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded" title="Copy"><Copy className="h-4 w-4" /></button>
          </div>
          <span className="text-[10px] text-center text-slate-500 mt-auto">Clipboard</span>
        </div>

        {/* Font Group */}
        <div className="flex flex-col gap-1 px-2 border-r border-slate-200 dark:border-slate-700 min-w-[140px]">
          <div className="flex gap-1 mb-1">
            <div className="bg-slate-100 dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded px-2 py-0.5 text-xs flex items-center justify-between w-24">
              Calibri <ChevronDown className="h-3 w-3" />
            </div>
            <div className="bg-slate-100 dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded px-2 py-0.5 text-xs flex items-center justify-between w-12">
              11 <ChevronDown className="h-3 w-3" />
            </div>
          </div>
          <div className="flex gap-1">
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded font-bold"><Bold className="h-4 w-4" /></button>
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded italic"><Italic className="h-4 w-4" /></button>
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded underline"><Underline className="h-4 w-4" /></button>
             <div className="w-px h-4 bg-slate-300 dark:bg-slate-600 mx-1"></div>
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded"><Grid className="h-4 w-4" /></button>
          </div>
          <span className="text-[10px] text-center text-slate-500 mt-auto">Font</span>
        </div>

        {/* Alignment Group */}
        <div className="flex flex-col gap-1 px-2 border-r border-slate-200 dark:border-slate-700">
           <div className="flex gap-1">
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded"><AlignLeft className="h-4 w-4" /></button>
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded"><AlignCenter className="h-4 w-4" /></button>
             <button className="p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded"><AlignRight className="h-4 w-4" /></button>
           </div>
           <span className="text-[10px] text-center text-slate-500 mt-auto">Alignment</span>
        </div>

        {/* Number Group */}
         <div className="flex flex-col gap-1 px-2 border-r border-slate-200 dark:border-slate-700 min-w-[100px]">
          <div className="bg-slate-100 dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded px-2 py-0.5 text-xs flex items-center justify-between">
              General <ChevronDown className="h-3 w-3" />
          </div>
          <div className="flex gap-2 justify-center mt-1">
             <span className="text-xs font-mono">$</span>
             <span className="text-xs font-mono">%</span>
          </div>
           <span className="text-[10px] text-center text-slate-500 mt-auto">Number</span>
        </div>

        {/* Styles Group */}
         <div className="flex flex-col gap-1 px-2 border-r border-slate-200 dark:border-slate-700">
           <div className="flex gap-2">
             <button className="flex flex-col items-center p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded">
                <Table className="h-5 w-5 text-blue-500" />
                <span className="text-[10px]">Format as Table</span>
             </button>
             <button className="flex flex-col items-center p-1 hover:bg-slate-100 dark:hover:bg-slate-800 rounded">
                <Grid className="h-5 w-5 text-green-500" />
                <span className="text-[10px]">Cell Styles</span>
             </button>
           </div>
           <span className="text-[10px] text-center text-slate-500 mt-auto">Styles</span>
        </div>
      </div>
    </div>
  );
};

export default Ribbon;
