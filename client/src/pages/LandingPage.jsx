import React from 'react';
import { motion as Motion } from 'framer-motion';
import { useNavigate } from 'react-router-dom';
import { ArrowRight, Table, Sparkles, Shield, Zap } from 'lucide-react';
import Button from '../components/ui/Button';
import ThemeToggle from '../components/ui/ThemeToggle';

const LandingPage = () => {
  const navigate = useNavigate();

  return (
    <div className="min-h-screen font-sans selection:bg-blue-500/30">
      {/* Navbar */}
      <nav className="border-b border-slate-200 dark:border-slate-800/50 bg-white/80 dark:bg-slate-900/50 backdrop-blur-md fixed w-full z-50 top-0 transition-colors">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-gradient-to-tr from-green-500 to-emerald-300 p-2 rounded-lg">
              <Table className="h-6 w-6 text-white dark:text-slate-950" />
            </div>
            <span className="text-xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-slate-900 to-slate-600 dark:from-white dark:to-slate-400">
              SheetPilot
            </span>
          </div>
          <div className="flex items-center gap-4">
             <ThemeToggle />
             <Button variant="ghost" onClick={() => navigate('/login')}>Sign In</Button>
             <Button onClick={() => navigate('/register')}>Get Started</Button>
          </div>
        </div>
      </nav>

      {/* Hero Section */}
      <section className="relative pt-32 pb-20 overflow-hidden">
        <div className="absolute inset-0 bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:24px_24px]"></div>
        <div className="absolute left-0 right-0 top-0 -z-10 m-auto h-[310px] w-[310px] rounded-full bg-blue-500 opacity-20 blur-[100px]"></div>

        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 relative z-10 text-center">
          <Motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
          >
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-500/10 border border-blue-500/20 text-blue-600 dark:text-blue-400 text-sm mb-6">
              <Sparkles className="h-3 w-3" />
              <span>AI-Powered Excel Automation</span>
            </div>
            <h1 className="text-5xl md:text-7xl font-bold tracking-tight mb-8 text-slate-900 dark:text-white">
              Talk to your <br />
              <span className="text-transparent bg-clip-text bg-gradient-to-r from-green-500 to-emerald-600 dark:from-green-400 dark:to-emerald-600">
                Spreadsheets
              </span>
            </h1>
            <p className="text-xl text-slate-600 dark:text-slate-400 max-w-2xl mx-auto mb-10 leading-relaxed">
              No more complex formulas. Just tell SheetPilot what you need, and watch your data transform instantly.
            </p>
            
            <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
              <Button 
                size="lg" 
                className="text-lg px-8 py-6 rounded-full shadow-xl shadow-blue-500/20"
                onClick={() => navigate('/register')}
              >
                Try SheetPilot Free <ArrowRight className="ml-2 h-5 w-5" />
              </Button>
            </div>
          </Motion.div>

          {/* Visual Preview */}
          <Motion.div 
            initial={{ opacity: 0, y: 40 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.2, duration: 0.8 }}
            className="mt-20 relative mx-auto max-w-5xl"
          >
            <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-2 shadow-2xl">
              <div className="flex items-center gap-2 px-4 py-3 border-b border-slate-100 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/50">
                <div className="flex gap-1.5">
                  <div className="w-3 h-3 rounded-full bg-red-500/20 border border-red-500/50" />
                  <div className="w-3 h-3 rounded-full bg-yellow-500/20 border border-yellow-500/50" />
                  <div className="w-3 h-3 rounded-full bg-green-500/20 border border-green-500/50" />
                </div>
                <div className="ml-4 text-xs text-slate-500 font-mono">finance_report_2024.xlsx</div>
              </div>
              <div className="p-4 grid grid-cols-5 gap-px bg-slate-200 dark:bg-slate-800 border border-slate-200 dark:border-slate-800 overflow-hidden rounded-lg">
                {/* Header */}
                {['Date', 'Region', 'Product', 'Revenue', 'Profit'].map((h, i) => (
                  <div key={i} className="bg-slate-50 dark:bg-slate-900 p-3 text-xs font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                    {h}
                  </div>
                ))}
                {/* Mock Rows */}
                {[...Array(5)].map((_, rowIdx) => (
                  <React.Fragment key={rowIdx}>
                    <div className="bg-white dark:bg-slate-950/50 p-3 text-sm text-slate-700 dark:text-slate-300 border-t border-slate-100 dark:border-slate-800">2024-01-0{rowIdx + 1}</div>
                    <div className="bg-white dark:bg-slate-950/50 p-3 text-sm text-slate-700 dark:text-slate-300 border-t border-slate-100 dark:border-slate-800">{['North', 'South', 'East', 'West', 'North'][rowIdx]}</div>
                    <div className="bg-white dark:bg-slate-950/50 p-3 text-sm text-slate-700 dark:text-slate-300 border-t border-slate-100 dark:border-slate-800">Widget {String.fromCharCode(65 + rowIdx)}</div>
                    <div className="bg-white dark:bg-slate-950/50 p-3 text-sm text-green-600 dark:text-green-400 font-mono border-t border-slate-100 dark:border-slate-800">${(1000 + rowIdx * 150).toLocaleString()}</div>
                    <div className="bg-white dark:bg-slate-950/50 p-3 text-sm text-blue-600 dark:text-blue-400 font-mono border-t border-slate-100 dark:border-slate-800 relative">
                        ${(500 + rowIdx * 100).toLocaleString()}
                        {rowIdx === 2 && (
                            <Motion.div 
                                initial={{ width: 0 }}
                                animate={{ width: '100%' }}
                                className="absolute inset-0 bg-blue-500/20 border border-blue-500/50 z-10"
                            />
                        )}
                    </div>
                  </React.Fragment>
                ))}
              </div>
              
              {/* Command Simulation */}
              <div className="absolute -bottom-6 left-1/2 -translate-x-1/2 w-3/4 bg-white/90 dark:bg-slate-800/90 backdrop-blur-xl border border-slate-200 dark:border-slate-700 p-4 rounded-xl shadow-2xl flex items-center gap-3">
                 <div className="p-2 bg-blue-500/20 rounded-lg">
                    <Sparkles className="h-5 w-5 text-blue-600 dark:text-blue-400" />
                 </div>
                 <div className="text-sm text-slate-700 dark:text-slate-200">
                    "Highlight rows where <span className="text-blue-600 dark:text-blue-400 font-mono">Profit &gt; $600</span>"
                 </div>
              </div>
            </div>
          </Motion.div>
        </div>
      </section>

      {/* Features Grid */}
      <section className="py-24 bg-slate-50 dark:bg-slate-950 relative transition-colors">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="grid md:grid-cols-3 gap-8">
                {[
                    { icon: Zap, title: "Lightning Fast", desc: "Process thousands of rows in seconds with our optimized Excel engine." },
                    { icon: Shield, title: "Secure by Design", desc: "Your data is encrypted and processed in a secure, isolated environment." },
                    { icon: Sparkles, title: "Smart AI", desc: "Our model understands complex data transformation requests effortlessly." }
                ].map((f, i) => (
                    <div key={i} className="bg-white dark:bg-slate-900/50 p-8 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-blue-500/50 transition-colors group shadow-sm">
                        <div className="w-12 h-12 bg-slate-100 dark:bg-slate-800 rounded-xl flex items-center justify-center mb-6 group-hover:bg-blue-500/20 transition-colors">
                            <f.icon className="h-6 w-6 text-slate-500 dark:text-slate-400 group-hover:text-blue-500 dark:group-hover:text-blue-400" />
                        </div>
                        <h3 className="text-xl font-semibold mb-3 text-slate-900 dark:text-white">{f.title}</h3>
                        <p className="text-slate-600 dark:text-slate-400 leading-relaxed">{f.desc}</p>
                    </div>
                ))}
            </div>
        </div>
      </section>

      {/* Corner Button */}
      <div className="fixed bottom-8 right-8 z-50">
          <Motion.button 
            whileHover={{ scale: 1.05 }}
            whileTap={{ scale: 0.95 }}
            onClick={() => navigate('/register')}
            className="bg-gradient-to-r from-blue-600 to-indigo-600 text-white px-6 py-3 rounded-full shadow-lg shadow-blue-600/30 font-semibold flex items-center gap-2"
          >
            Try SheetPilot <ArrowRight className="h-4 w-4" />
          </Motion.button>
      </div>

      {/* Bottom Button */}
      <div className="bg-white dark:bg-slate-900 py-8 border-t border-slate-200 dark:border-slate-800 text-center transition-colors">
        <Button variant="outline" onClick={() => navigate('/dashboard')}>
           Go to Website (Dashboard)
        </Button>
        <p className="text-slate-500 dark:text-slate-600 text-xs mt-4">© 2024 SheetPilot. All rights reserved.</p>
      </div>
    </div>
  );
};

export default LandingPage;
