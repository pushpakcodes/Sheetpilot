import React, { useState } from 'react';
import { Send, Sparkles } from 'lucide-react';
import Button from './ui/Button';

const CommandBar = ({ onCommand, loading }) => {
  const [command, setCommand] = useState('');

  const suggestions = [
    "Sort by Revenue descending",
    "Highlight rows where Sales > 50000",
    "Add a Profit column = Revenue - Cost",
    "Add 10 empty rows"
  ];

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!command.trim() || loading) return;
    onCommand(command);
    setCommand('');
  };

  return (
    <div className="w-full max-w-3xl mx-auto">
      <form onSubmit={handleSubmit} className="relative group">
        <div className="absolute -inset-1 bg-gradient-to-r from-blue-500 to-indigo-500 rounded-xl blur opacity-25 group-hover:opacity-50 transition duration-1000 group-hover:duration-200"></div>
        <div className="relative flex items-center bg-slate-900 rounded-xl border border-slate-700 shadow-2xl p-2">
            <div className="pl-4 pr-2">
                <Sparkles className="h-5 w-5 text-blue-400" />
            </div>
          <input
            type="text"
            value={command}
            onChange={(e) => setCommand(e.target.value)}
            placeholder="Describe what you want to do with your data..."
            className="flex-1 bg-transparent border-none focus:ring-0 text-slate-100 placeholder-slate-500 text-lg py-2 px-2 outline-none"
            disabled={loading}
          />
          <Button
            type="submit"
            size="sm"
            disabled={!command.trim() || loading}
            className="ml-2"
          >
            {loading ? 'Processing...' : <Send className="h-4 w-4" />}
          </Button>
        </div>
      </form>
      
      <div className="mt-4 flex flex-wrap gap-2 justify-center">
        {suggestions.map((s, i) => (
           <button
             key={i}
             type="button"
             onClick={() => setCommand(s)}
             className="text-xs text-slate-400 bg-slate-800/50 hover:bg-slate-700 hover:text-white px-3 py-1.5 rounded-full border border-slate-700 transition-colors"
           >
             {s}
           </button>
        ))}
      </div>
    </div>
  );
};

export default CommandBar;
