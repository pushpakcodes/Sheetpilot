import React, { useState, useEffect } from 'react';
import { uploadFile, processCommand, getFiles } from '../services/api';
import Navbar from '../components/Navbar';
import FileUploader from '../components/FileUploader';
import ExcelPreview from '../components/ExcelPreview';
import Ribbon from '../components/Ribbon';
import ChatSidebar from '../components/ChatSidebar';
import { motion, AnimatePresence } from 'framer-motion';
import { Download, History, FileText, RotateCcw } from 'lucide-react';
import Button from '../components/ui/Button';

const Dashboard = () => {
  const [currentFile, setCurrentFile] = useState(null);
  const [previewData, setPreviewData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [chatHistory, setChatHistory] = useState([]);
  const [fileHistory, setFileHistory] = useState([]); // Undo stack

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
      setChatHistory(prev => [...prev, { type: 'bot', content: `Successfully uploaded ${file.name}. What would you like to do with it?` }]);
    } catch (error) {
      console.error(error);
      setChatHistory(prev => [...prev, { type: 'bot', content: 'Upload failed. Please try again.' }]);
    } finally {
      setLoading(false);
    }
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
      const { data } = await processCommand(command, currentFile.filePath);
      if (data.success) {
        setCurrentFile(prev => ({ ...prev, filePath: data.filePath }));
        setPreviewData(data.preview);
        setChatHistory(prev => [...prev, { type: 'bot', content: `Done! ${data.action.action} executed successfully.` }]);
      }
    } catch (error) {
      console.error(error);
      const errorMsg = error.response?.data?.message || error.message;
      setChatHistory(prev => [...prev, { type: 'bot', content: `I encountered an error: ${errorMsg}` }]);
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

  const suggestions = [
    "Sort by Revenue descending",
    "Highlight rows where Sales > 50000",
    "Add a Profit column = Revenue - Cost",
    "Add 10 empty rows"
  ];

  return (
    <div className="h-screen flex flex-col bg-slate-50 dark:bg-slate-950 font-sans transition-colors overflow-hidden">
      <Navbar />
      <Ribbon onUndo={handleUndo} />
      
      <div className="flex-1 flex overflow-hidden">
        {/* Left Side - Excel Grid */}
        <main className="flex-1 flex flex-col relative overflow-hidden bg-white dark:bg-slate-900">
          <AnimatePresence mode="wait">
            {!currentFile ? (
              <motion.div
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
              </motion.div>
            ) : (
               <ExcelPreview data={previewData} />
            )}
          </AnimatePresence>
        </main>

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
