import React, { useState } from 'react';
import { uploadFile, processCommand, getWorkbookMetadata, downloadWorkbook } from '../services/api';
import Navbar from '../components/Navbar';
import FileUploader from '../components/FileUploader';
import ExcelPreview from '../components/ExcelPreview';
import VirtualizedExcelView from '../components/VirtualizedExcelView';
import Spreadsheet from '../components/Spreadsheet';
import SheetTabs from '../components/SheetTabs';
import Ribbon from '../components/Ribbon';
import ChatSidebar from '../components/ChatSidebar';
import { motion as Motion, AnimatePresence } from 'framer-motion';
import { Download, History, FileText, RotateCcw } from 'lucide-react';
import Button from '../components/ui/Button';

const Dashboard = () => {
  const [currentFile, setCurrentFile] = useState(null);
  const [previewData, setPreviewData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [chatHistory, setChatHistory] = useState([]);
  const [fileHistory, setFileHistory] = useState([]); // Undo stack
  const [workbookMetadata, setWorkbookMetadata] = useState(null);
  const [activeSheetId, setActiveSheetId] = useState(null); // Store sheet name, not index
  const [saveSignal, setSaveSignal] = useState(0);

  const handleUpload = async (file) => {
    setLoading(true);
    // Add upload message to chat
    setChatHistory(prev => [...prev, { type: 'user', content: `Uploading ${file.name}...` }]);
    
    try {
      const formData = new FormData();
      formData.append('file', file);
      const { data } = await uploadFile(formData);
      setCurrentFile(data);
      setPreviewData(data.preview);
      setFileHistory([]); // Reset undo stack on new upload
      
      // Fetch workbook metadata for sheet tabs
      try {
        const metadataResponse = await getWorkbookMetadata(data.filePath);
        setWorkbookMetadata(metadataResponse.data);
        console.log('📊 Workbook metadata loaded:', metadataResponse.data);
        
        // Set active sheet to first visible sheet (by name, not index)
        if (metadataResponse.data.sheets && metadataResponse.data.sheets.length > 0) {
          setActiveSheetId(metadataResponse.data.sheets[0].sheetId); // Use sheet name
        }
      } catch (metadataError) {
        console.error('Failed to load workbook metadata:', metadataError);
        // Continue without metadata - sheet tabs won't show
      }
      
      setChatHistory(prev => [...prev, { type: 'bot', content: `Successfully uploaded ${file.name}. What would you like to do with it?` }]);
    } catch (error) {
      console.error(error);
      setChatHistory(prev => [...prev, { type: 'bot', content: 'Upload failed. Please try again.' }]);
    } finally {
      setLoading(false);
    }
  };

  const handleSheetChange = (sheetId) => {
    console.log('📑 Switching to sheet:', sheetId);
    setActiveSheetId(sheetId);
    // Spreadsheet component will handle the sheet change via sheetId prop
  };

  const handleCommand = async (command) => {
    if (!currentFile) {
        setChatHistory(prev => [...prev, { type: 'user', content: command }]);
        setTimeout(() => {
            setChatHistory(prev => [...prev, { type: 'bot', content: 'Please upload a file first so I can help you analyze it.' }]);
        }, 500);
        return;
    }
    
    setLoading(true);
    setChatHistory(prev => [...prev, { type: 'user', content: command }]);
    
    // Save state for undo
    setFileHistory(prev => [...prev, { file: currentFile, preview: previewData }]);

    try {
      const { data } = await processCommand(command, currentFile.filePath, activeSheetId);
      const nextFilePath = data?.filePath;
      const nextPreview = data?.preview;

      if (nextFilePath && nextPreview) {
        setCurrentFile(prev => ({ ...prev, filePath: nextFilePath }));
        setPreviewData(nextPreview);
      } else {
        setFileHistory(prev => prev.slice(0, -1));
      }

      let message = '';
      if (Array.isArray(data?.results) && data.results.length > 0) {
        const successCount = data.results.filter(r => r.success).length;
        const total = data.results.length;
        const header = data.success ? `Done! Completed ${successCount}/${total} step(s).` : `Stopped after ${successCount}/${total} step(s).`;
        const lines = data.results.map((r, i) => {
          if (r.success) return `${i + 1}) ${r.action}: success`;
          return `${i + 1}) ${r.action}: failed${r.message ? ` (${r.message})` : ''}`;
        });
        message = [header, ...lines].join('\n');
      } else if (data?.success && data?.action?.action) {
        message = `Done! ${data.action.action} executed successfully.`;
      } else if (!data?.success && data?.action?.action) {
        message = `Stopped: ${data.action.action} did not complete.`;
      } else {
        message = data?.success ? 'Done!' : 'I could not complete that command.';
      }

      setChatHistory(prev => [...prev, { type: 'bot', content: message }]);
    } catch (error) {
      console.error(error);
      const resp = error.response?.data;
      const errorMsg = resp?.message || error.message;
      const type = resp?.type;
      const prefix = type === 'VALIDATION_ERROR' ? 'Validation error' : 'I encountered an error';
      const detailKey = resp?.details?.key;
      const followUp =
        type === 'VALIDATION_ERROR' && detailKey
          ? (() => {
              if (detailKey === 'filterValue') return 'Which exact value should I match (e.g. Name = "Raj")?';
              if (detailKey === 'filterColumn') return 'Which column should I match against?';
              if (detailKey === 'targetColumn') return 'Which column should I update?';
              if (detailKey === 'operation') return 'Should I SET, add (+), subtract (-), multiply (*), or divide (/)?';
              return 'Can you rephrase with concrete column names and values?';
            })()
          : type === 'AI_ERROR'
            ? 'Can you rephrase with the exact sheet, column names, and values?'
            : '';
      const content = followUp ? `${prefix}: ${errorMsg}\n${followUp}` : `${prefix}: ${errorMsg}`;
      setChatHistory(prev => [...prev, { type: 'bot', content }]);
      // Revert undo stack if failed
      setFileHistory(prev => prev.slice(0, -1));
    } finally {
      setLoading(false);
    }
  };

  const handleUndo = () => {
    if (fileHistory.length === 0) return;
    const lastState = fileHistory[fileHistory.length - 1];
    setCurrentFile(lastState.file);
    setPreviewData(lastState.preview);
    setFileHistory(prev => prev.slice(0, -1));
    setChatHistory(prev => [...prev, { type: 'user', content: 'Undo last action' }, { type: 'bot', content: 'Action undone.' }]);
  };
  
  const handleSave = () => {
    setSaveSignal((s) => s + 1);
  };
  
  const handleDownload = async () => {
    if (!currentFile?.filePath) return;
    try {
      const blob = await downloadWorkbook(currentFile.filePath);
      const url = URL.createObjectURL(new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = 'workbook.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error('Download failed', e);
    }
  };

  const suggestions = [
    "Sort by Revenue descending",
    "Highlight rows where Sales > 50000",
    "Add a Profit column = Revenue - Cost",
    "Add 10 empty rows"
  ];

  return (
    <div className="h-screen flex flex-col bg-slate-50 dark:bg-slate-950 font-sans transition-colors overflow-hidden">
      <Navbar />
      
      <div className="flex-1 flex overflow-hidden">
        <div className="flex-1 flex flex-col min-w-0">
          <Ribbon onUndo={handleUndo} onRedo={() => {}} onSave={handleSave} onDownload={handleDownload} />
          
          {/* Left Side - Excel Grid */}
          <main className="flex-1 flex flex-col relative overflow-hidden bg-white dark:bg-slate-900">
            <AnimatePresence mode="wait">
              {!currentFile ? (
                <Motion.div
                  key="upload"
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 0.95 }}
                  className="absolute inset-0 flex items-center justify-center bg-slate-50/50 dark:bg-slate-900/50 backdrop-blur-sm z-10 p-8"
                >
                  <div className="max-w-xl w-full">
                    <h1 className="text-3xl font-bold text-center mb-2 text-slate-900 dark:text-white">
                      Start with <span className="text-green-600">SheetPilot</span>
                    </h1>
                    <p className="text-slate-600 dark:text-slate-400 text-center mb-8">
                      Upload an Excel file to unlock AI-powered analysis.
                    </p>
                    <FileUploader onUpload={handleUpload} />
                  </div>
                </Motion.div>
              ) : (
                 // Use Handsontable Spreadsheet for Excel-like experience
                 currentFile?.filePath ? (
                   <>
                     <div className="flex-1 overflow-hidden">
                       {activeSheetId && (
                         <Spreadsheet 
                           filePath={currentFile.filePath} 
                           sheetId={activeSheetId}
                           saveSignal={saveSignal}
                           onDataChange={(changes) => {
                             // Handle cell changes - can sync to backend here
                             console.log('Cell changes:', changes);
                           }}
                         />
                       )}
                     </div>
                     {workbookMetadata && workbookMetadata.sheets && (
                       <SheetTabs
                         sheets={workbookMetadata.sheets}
                         activeSheetId={activeSheetId}
                         onSheetChange={handleSheetChange}
                       />
                     )}
                   </>
                 ) : (
                   <ExcelPreview data={previewData} />
                 )
              )}
            </AnimatePresence>
          </main>
        </div>

        {/* Right Side - AI Chat */}
        <ChatSidebar 
          onCommand={handleCommand} 
          loading={loading} 
          history={chatHistory} 
          suggestions={suggestions}
        />
      </div>
    </div>
  );
};

export default Dashboard;
