import React, { useState, useRef, useEffect } from 'react';
import { Send, Sparkles, Bot, User } from 'lucide-react';
import { motion as Motion } from 'framer-motion';

const ChatSidebar = ({ onCommand, loading, history = [], suggestions = [] }) => {
  const [input, setInput] = useState('');
  const messagesEndRef = useRef(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [history, loading]);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!input.trim() || loading) return;
    onCommand(input);
    setInput('');
  };

  return (
    <div className="flex flex-col h-full bg-slate-50 dark:bg-slate-900 border-l border-slate-200 dark:border-slate-800 w-80 lg:w-96 flex-shrink-0 transition-colors">
      {/* Header */}
      <div className="p-4 border-b border-slate-200 dark:border-slate-800 flex items-center gap-2 bg-white dark:bg-slate-900">
        <Sparkles className="h-5 w-5 text-blue-500" />
        <h2 className="font-semibold text-slate-800 dark:text-slate-100">AI Copilot</h2>
      </div>

      {/* Messages Area */}
      <div className="flex-1 overflow-y-auto p-4 space-y-4 bg-slate-50 dark:bg-slate-950/50">
        {history.length === 0 && (
          <div className="text-center text-slate-500 mt-10">
            <Bot className="h-12 w-12 mx-auto mb-2 opacity-50" />
            <p className="text-sm">Hi! I'm SheetPilot.</p>
            <p className="text-xs mt-1">Ask me to analyze your data.</p>
          </div>
        )}

        {history.map((msg, index) => (
          <Motion.div
            key={index}
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            className={`flex gap-3 ${msg.type === 'user' ? 'flex-row-reverse' : ''}`}
          >
            <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 ${msg.type === 'user' ? 'bg-indigo-500 text-white' : 'bg-green-600 text-white'}`}>
              {msg.type === 'user' ? <User className="h-4 w-4" /> : <Bot className="h-4 w-4" />}
            </div>
            <div className={`max-w-[80%] rounded-lg p-3 text-sm ${
              msg.type === 'user' 
                ? 'bg-indigo-100 dark:bg-indigo-900/30 text-indigo-900 dark:text-indigo-100' 
                : 'bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 border border-slate-200 dark:border-slate-700 shadow-sm'
            }`}>
              {msg.content}
              {msg.timestamp && (
                <div className="text-[10px] opacity-50 mt-1 text-right">
                  {new Date(msg.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </div>
              )}
            </div>
          </Motion.div>
        ))}

        {loading && (
          <div className="flex gap-3">
             <div className="w-8 h-8 rounded-full bg-green-600 text-white flex items-center justify-center flex-shrink-0">
               <Bot className="h-4 w-4" />
             </div>
             <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg p-3 shadow-sm">
                <div className="flex gap-1">
                  <div className="w-2 h-2 bg-slate-400 rounded-full animate-bounce" style={{ animationDelay: '0ms' }} />
                  <div className="w-2 h-2 bg-slate-400 rounded-full animate-bounce" style={{ animationDelay: '150ms' }} />
                  <div className="w-2 h-2 bg-slate-400 rounded-full animate-bounce" style={{ animationDelay: '300ms' }} />
                </div>
             </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input Area */}
      <div className="p-4 border-t border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900">
        <form onSubmit={handleSubmit} className="relative">
          <input
            type="text"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Ask anything..."
            className="w-full bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-slate-100 border-none rounded-xl py-3 pl-4 pr-12 focus:ring-2 focus:ring-blue-500 placeholder-slate-400 dark:placeholder-slate-500"
            disabled={loading}
          />
          <button
            type="submit"
            disabled={!input.trim() || loading}
            className="absolute right-2 top-1/2 -translate-y-1/2 p-1.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:hover:bg-blue-600 transition-colors"
          >
            <Send className="h-4 w-4" />
          </button>
        </form>
        
        {/* Quick Suggestions */}
        <div className="mt-3 flex gap-2 overflow-x-auto pb-1 no-scrollbar">
          {suggestions.map((s, i) => (
             <button
               key={i}
               onClick={() => onCommand(s)}
               className="flex-shrink-0 text-xs text-slate-600 dark:text-slate-400 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 px-3 py-1.5 rounded-full border border-slate-200 dark:border-slate-700 transition-colors whitespace-nowrap"
             >
               {s.length > 20 ? s.substring(0, 20) + '...' : s}
             </button>
          ))}
        </div>
      </div>
    </div>
  );
};

export default ChatSidebar;
