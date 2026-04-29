import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { useNavigate } from 'react-router-dom';
import {
  ArrowRight, Table,
  Sparkles, Zap, BarChart3, BrainCircuit, PieChart, AlertCircle, Mail, Wand2
} from 'lucide-react';
import { useAuthStore } from '../stores/authStore';

export const AuthPage: React.FC = () => {
  const [isLogin, setIsLogin] = useState(true);
  const navigate = useNavigate();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [activeFeature, setActiveFeature] = useState(0);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [selectedProduct, setSelectedProduct] = useState<'excelerate' | 'mailacer'>('excelerate');
  
  const { login, register, isAuthenticated, isLoading, error, clearError } = useAuthStore();

  const excelerateFeatures = [
    {
      title: "AI-Powered Analysis",
      description: "Transform raw data into actionable insights with our advanced AI engine.",
      icon: BrainCircuit,
      color: "bg-[#0a5c2e]",
      textColor: "text-[#052e16]"
    },
    {
      title: "Smart Automation",
      description: "Say goodbye to manual entry. Let AI handle the repetitive work for you.",
      icon: Zap,
      color: "bg-[#084a24]",
      textColor: "text-[#052e16]"
    },
    {
      title: "Visual Intelligence",
      description: "Create stunning visualizations and dashboards in seconds, not hours.",
      icon: BarChart3,
      color: "bg-[#052e16]",
      textColor: "text-[#052e16]"
    }
  ];

  const mailacerFeatures = [
    {
      title: "Smart Email Processing",
      description: "Automatically categorize, prioritize, and respond to your emails with AI.",
      icon: Mail,
      color: "bg-[#b8860b]",
      textColor: "text-[#5c3d00]"
    },
    {
      title: "Workflow Automation",
      description: "Streamline your email workflows with intelligent automation.",
      icon: Wand2,
      color: "bg-[#8b4513]",
      textColor: "text-white"
    },
    {
      title: "Intelligent Responses",
      description: "Generate personalized email responses in seconds with AI assistance.",
      icon: Sparkles,
      color: "bg-[#654321]",
      textColor: "text-white"
    }
  ];

  const features = selectedProduct === 'excelerate' ? excelerateFeatures : mailacerFeatures;

  useEffect(() => {
    const interval = setInterval(() => {
      setActiveFeature((prev) => (prev + 1) % features.length);
    }, 5000);
    return () => clearInterval(interval);
  }, []);

  // Redirect if already authenticated
  useEffect(() => {
    if (isAuthenticated) {
      navigate('/spreadsheet');
    }
  }, [isAuthenticated, navigate]);

  // Clear errors when switching between login and register
  useEffect(() => {
    clearError();
    setErrorMessage(null);
  }, [isLogin, clearError]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMessage(null);
    clearError();

    // Validation
    if (!username.trim() || !password.trim()) {
      setErrorMessage('Please fill in all fields');
      return;
    }

    if (username.length < 3 || username.length > 50) {
      setErrorMessage('Username must be between 3 and 50 characters');
      return;
    }

    if (password.length < 6) {
      setErrorMessage('Password must be at least 6 characters');
      return;
    }

    try {
      if (isLogin) {
        await login(username, password);
      } else {
        await register(username, password);
      }
      // Navigation will happen automatically via useEffect when isAuthenticated changes
    } catch (err: any) {
      setErrorMessage(err.message || 'An error occurred');
    }
  };

  return (
    <div className="min-h-screen w-full flex font-sans bg-white overflow-hidden">
      {/* Left Side - Visual Experience (60%) */}
      <div className={`hidden lg:flex w-[60%] relative ${selectedProduct === 'excelerate' ? 'bg-gradient-to-br from-[#0a5c2e]/10 via-[#084a24]/5 to-[#052e16]/20' : 'bg-gradient-to-br from-[#b8860b]/10 via-[#8b4513]/5 to-[#5c3d00]/20'} ${selectedProduct === 'excelerate' ? 'text-[#052e16]' : 'text-[#5c3d00]'} overflow-hidden flex-col justify-between p-16`}>

        {/* Animated Background */}
        <div className="absolute inset-0">
          <div className="absolute inset-0 bg-[url('https://grainy-gradients.vercel.app/noise.svg')] opacity-20"></div>
          {selectedProduct === 'excelerate' ? (
            <>
              <div className="absolute top-[-20%] left-[-20%] w-[80%] h-[80%] rounded-full bg-[#0a5c2e]/20 blur-[120px] animate-pulse-slow" />
              <div className="absolute bottom-[-20%] right-[-20%] w-[80%] h-[80%] rounded-full bg-[#084a24]/15 blur-[120px] animate-pulse-slow" style={{ animationDelay: '2s' }} />
            </>
          ) : (
            <>
              <div className="absolute top-[-20%] left-[-20%] w-[80%] h-[80%] rounded-full bg-[#b8860b]/20 blur-[120px] animate-pulse-slow" />
              <div className="absolute bottom-[-20%] right-[-20%] w-[80%] h-[80%] rounded-full bg-[#8b4513]/15 blur-[120px] animate-pulse-slow" style={{ animationDelay: '2s' }} />
            </>
          )}
        </div>

        {/* Brand */}
        <div className="relative z-10 flex items-center gap-3">
          <div className={`p-2.5 backdrop-blur-md rounded-xl border ${selectedProduct === 'excelerate' ? 'bg-[#0a5c2e]/20 border-[#0a5c2e]/30' : 'bg-[#b8860b]/20 border-[#b8860b]/30'}`}>
            {selectedProduct === 'excelerate' ? (
              <Table className="w-6 h-6 text-[#0a5c2e]" />
            ) : (
              <Mail className="w-6 h-6 text-[#5c3d00]" />
            )}
          </div>
          <div className="flex flex-col">
            <span className="text-lg font-bold tracking-wide">Workflow Genie</span>
            {selectedProduct === 'excelerate' ? (
              <span className="text-sm text-[#052e16] capitalize">
                <span className="font-medium">Excel</span>
                <span className="italic text-[#0a5c2e] font-serif">erate</span>
              </span>
            ) : (
              <span className="text-sm text-[#5c3d00] capitalize">
                <span className="font-medium">Mail</span>
                <span className="italic text-[#8b4513] font-serif">acer</span>
              </span>
            )}
          </div>
        </div>

        {/* Central Visual - Glassmorphism Card */}
        <div className="relative z-10 flex-1 flex items-center justify-center perspective-[1000px]">
          <motion.div
            initial={{ opacity: 0, rotateX: 20, rotateY: -20 }}
            animate={{ opacity: 1, rotateX: 10, rotateY: -10, y: [0, -20, 0] }}
            transition={{
              y: { duration: 6, repeat: Infinity, ease: "easeInOut" },
              default: { duration: 1 }
            }}
            className={`w-[500px] h-[350px] rounded-2xl border shadow-2xl overflow-hidden relative ${
              selectedProduct === 'excelerate' 
                ? 'bg-[#0a5c2e]/10 backdrop-blur-xl border-[#0a5c2e]/20' 
                : 'bg-[#8b4513]/10 backdrop-blur-xl border-[#b8860b]/20'
            }`}
          >
            <AnimatePresence mode="wait">
              {activeFeature === 0 && (
                <motion.div
                  key="analysis"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  transition={{ duration: 0.5 }}
                  className="w-full h-full p-6 flex flex-col gap-4"
                >
                  {/* Header */}
                  <div className={`flex items-center gap-2 mb-2 border-b pb-4 ${selectedProduct === 'excelerate' ? 'border-white/10' : 'border-slate-200'}`}>
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-red-500/50' : 'bg-red-400'}`} />
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-yellow-500/50' : 'bg-yellow-400'}`} />
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-green-500/50' : 'bg-green-400'}`} />
                    <div className={`ml-auto text-xs font-mono ${selectedProduct === 'excelerate' ? 'text-slate/40' : 'text-slate-500'}`}>
                      {selectedProduct === 'excelerate' ? 'AI reads your entire excel' : 'emails summarized'}
                    </div>
                  </div>
                  {selectedProduct === 'excelerate' ? (
                    /* Excelerate Grid */
                    <div className="grid grid-cols-4 gap-3 h-full relative">
                      {[...Array(12)].map((_, i) => (
                        <motion.div
                          key={i}
                          initial={{ opacity: 0.3, scale: 0.8 }}
                          animate={{
                            opacity: [0.3, 1, 0.3],
                            scale: [0.9, 1, 0.9],
                            backgroundColor: ["rgba(255,255,255,0.05)", "rgba(16, 185, 129, 0.2)", "rgba(255,255,255,0.05)"]
                          }}
                          transition={{
                            duration: 3,
                            repeat: Infinity,
                            delay: Math.random() * 2,
                            ease: "easeInOut"
                          }}
                          className="rounded-lg border border-[#0a5c2e]/20 flex items-center justify-center"
                        >
                          <div className="w-8 h-1 bg-[#0a5c2e]/30 rounded-full" />
                        </motion.div>
                      ))}
                      {/* Scanning Line */}
                      <motion.div
                        animate={{ top: ["0%", "100%"], opacity: [0, 1, 0] }}
                        transition={{ duration: 2, repeat: Infinity, ease: "linear" }}
                        className="absolute left-0 right-0 h-1 bg-[#0a5c2e] blur-md shadow-[0_0_20px_rgba(10,92,46,0.8)]"
                      />
                    </div>
                  ) : (
                    /* Mailacer Email Inbox */
                    <div className="space-y-3 h-full overflow-hidden">
                      {[...Array(6)].map((_, i) => (
                        <motion.div
                          key={i}
                          initial={{ opacity: 0, x: -20 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: i * 0.1 }}
                          className="flex items-center gap-3 p-3 rounded-lg bg-[#8b4513]/10 border border-[#b8860b]/20 hover:bg-[#b8860b]/10 transition-colors"
                        >
                          <div className="w-8 h-8 rounded-full bg-[#b8860b]/20 flex items-center justify-center">
                            <Mail className="w-4 h-4 text-[#5c3d00]" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center justify-between">
                              <span className="text-sm font-medium text-slate-900 truncate">sender{i+1}@example.com</span>
                              <span className="text-xs text-slate-500">2m ago</span>
                            </div>
                            <div className="text-xs text-slate-600 truncate">Subject line for email {i+1}</div>
                          </div>
                          <motion.div
                            animate={{ scale: [1, 1.1, 1] }}
                            transition={{ duration: 2, repeat: Infinity, delay: i * 0.5 }}
                            className="w-2 h-2 rounded-full bg-[#b7245c]"
                          />
                        </motion.div>
                      ))}
                    </div>
                  )}
                  {/* Floating Badge */}
                    <motion.div
                    initial={{ scale: 0.8, opacity: 0 }}
                    animate={{ scale: 1, opacity: 1 }}
                    className={`absolute bottom-6 right-6 ${selectedProduct === 'excelerate' ? 'bg-[#0a5c2e]/20 border-[#0a5c2e]' : 'bg-[#b8860b]/20 border-[#b8860b]'} backdrop-blur-md px-4 py-2 rounded-full flex items-center gap-2 shadow-xl`}
                    >
                    <Sparkles className={`w-4 h-4 ${selectedProduct === 'excelerate' ? 'text-[#052e16]' : 'text-[#5c3d00]'}`} />
                    {selectedProduct === 'excelerate' ? <span className="text-xs font-medium text-[#052e16]">
                      Excel Processed
                    </span> : <span className="text-xs font-medium text-[#5c3d00]">
                      Email Processed
                    </span>}
                    </motion.div>
                </motion.div>
              )}

              {activeFeature === 1 && (
                <motion.div
                  key="automation"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  transition={{ duration: 0.5 }}
                  className="w-full h-full p-6 flex flex-col gap-4"
                >
                  {/* Header */}
                  <div className={`flex items-center gap-2 mb-2 border-b pb-4 ${selectedProduct === 'excelerate' ? 'border-white/10' : 'border-slate-200'}`}>
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-red-500/50' : 'bg-red-400'}`} />
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-yellow-500/50' : 'bg-yellow-400'}`} />
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-green-500/50' : 'bg-green-400'}`} />
                    <div className={`ml-auto text-xs font-mono ${selectedProduct === 'excelerate' ? 'text-slate/40' : 'text-slate-500'}`}>
                      {selectedProduct === 'excelerate' ? 'AI performs any task you want' : 'automatically send emails'}
                    </div>
                  </div>
                  {selectedProduct === 'excelerate' ? (
                    /* Excelerate Progress Bars */
                    <div className="space-y-3">
                      {[1, 2, 3, 4].map((i) => (
                        <div key={i} className="flex items-center gap-3 p-3 rounded-lg bg-[#0a5c2e]/10 border border-[#0a5c2e]/20">
                          <div className="w-8 h-8 rounded bg-[#084a24]/20 flex items-center justify-center text-xs text-[#052e16] font-mono">0{i}</div>
                          <div className="flex-1 h-2 bg-[#084a24]/20 rounded-full overflow-hidden">
                            <motion.div
                              initial={{ width: "0%" }}
                              animate={{ width: "100%" }}
                              transition={{ duration: 1.5, delay: i * 0.4, ease: "easeInOut" }}
                              className="h-full bg-[#0a5c2e]"
                            />
                          </div>
                          <motion.div
                            initial={{ opacity: 0, scale: 0.5 }}
                            animate={{ opacity: 1, scale: 1 }}
                            transition={{ delay: i * 0.4 + 1, duration: 0.3 }}
                          >
                            <div className="px-2 py-1 rounded text-[10px] font-bold bg-[#0a5c2e]/20 text-[#052e16] border border-[#0a5c2e]/30">
                              DONE
                            </div>
                          </motion.div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    /* Mailacer Email Sending */
                    <div className="flex flex-col items-center justify-center h-full gap-6">
                      <motion.div
                        animate={{ y: [0, -10, 0] }}
                        transition={{ duration: 2, repeat: Infinity, ease: "easeInOut" }}
                        className="relative"
                      >
                        <div className="w-16 h-12 bg-[#b8860b]/20 rounded-lg border border-[#b8860b]/30 flex items-center justify-center">
                          <Mail className="w-6 h-6 text-[#5c3d00]" />
                        </div>
                        <motion.div
                          animate={{ x: [0, 50, 100], opacity: [1, 1, 0] }}
                          transition={{ duration: 1.5, repeat: Infinity, delay: 0.5 }}
                          className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2"
                        >
                          <div className="w-2 h-2 bg-blue-400 rounded-full" />
                        </motion.div>
                      </motion.div>
                      <div className="text-center">
                        <motion.h3
                          animate={{ opacity: [0.5, 1, 0.5] }}
                          transition={{ duration: 2, repeat: Infinity }}
                          className="text-lg font-semibold text-slate-900 mb-2"
                        >
                          Sending Emails...
                        </motion.h3>
                        <div className="flex gap-2 justify-center">
                          {[1, 2, 3].map((i) => (
                            <motion.div
                              key={i}
                              animate={{ scale: [1, 1.2, 1] }}
                              transition={{ duration: 1, repeat: Infinity, delay: i * 0.2 }}
                              className="w-2 h-2 bg-[#b8860b] rounded-full"
                            />
                          ))}
                        </div>
                      </div>
                    </div>
                  )}
                </motion.div>
              )}

              {activeFeature === 2 && (
                <motion.div
                  key="visual"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  transition={{ duration: 0.5 }}
                  className="w-full h-full p-6 flex flex-col"
                >
                  {/* Header */}
                  <div className={`flex items-center gap-2 mb-6 border-b pb-4 ${selectedProduct === 'excelerate' ? 'border-white/10' : 'border-slate-200'}`}>
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-red-500/50' : 'bg-red-400'}`} />
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-yellow-500/50' : 'bg-yellow-400'}`} />
                    <div className={`w-3 h-3 rounded-full ${selectedProduct === 'excelerate' ? 'bg-green-500/50' : 'bg-green-400'}`} />
                    <div className={`ml-auto text-xs font-mono ${selectedProduct === 'excelerate' ? 'text-slate/40' : 'text-slate-500'}`}>
                      {selectedProduct === 'excelerate' ? 'perform complex tasks' : 'view analytics dashboard'}
                    </div>
                  </div>

                  {selectedProduct === 'excelerate' ? (
                    /* Excelerate Charts */
                    <>
                      <div className="flex items-end justify-between h-40 gap-4 px-2 mb-6">
                        {[40, 70, 50, 90, 60].map((h, i) => (
                          <motion.div
                            key={i}
                            initial={{ height: 0 }}
                            animate={{ height: `${h}%` }}
                            transition={{ duration: 0.8, delay: i * 0.1, type: "spring" }}
                            className="w-full bg-gradient-to-t from-[#084a24] to-[#0a5c2e] rounded-t-md opacity-90 relative group"
                          >
                          </motion.div>
                        ))}
                      </div>

                      <div className="flex gap-4">
                        <motion.div
                          initial={{ x: -20, opacity: 0 }}
                          animate={{ x: 0, opacity: 1 }}
                          transition={{ delay: 0.5 }}
                          className="flex-1 h-20 bg-[#0a5c2e]/10 rounded-xl border border-[#0a5c2e]/20 p-3 flex items-center gap-3"
                        >
                          <div className="w-10 h-10 rounded-full border-4 border-[#084a24] border-t-transparent animate-spin" />
                          <div className="space-y-2">
                            <div className="w-16 h-2 bg-[#084a24]/30 rounded-full" />
                            <div className="w-10 h-2 bg-[#052e16]/20 rounded-full" />
                          </div>
                        </motion.div>
                        <motion.div
                          initial={{ x: 20, opacity: 0 }}
                          animate={{ x: 0, opacity: 1 }}
                          transition={{ delay: 0.7 }}
                          className="flex-1 h-20 bg-[#0a5c2e]/10 rounded-xl border border-[#0a5c2e]/20 p-3 flex items-center gap-3"
                        >
                          <PieChart className="w-8 h-8 text-[#0a5c2e]" />
                          <div className="space-y-2">
                            <div className="w-16 h-2 bg-[#084a24]/30 rounded-full" />
                            <div className="w-10 h-2 bg-[#052e16]/20 rounded-full" />
                          </div>
                        </motion.div>
                      </div>
                    </>
                  ) : (
                    /* Mailacer Email Stats */
                    <div className="flex flex-col h-full gap-6">
                      <div className="grid grid-cols-2 gap-4">
                        <motion.div
                          initial={{ scale: 0.8, opacity: 0 }}
                          animate={{ scale: 1, opacity: 1 }}
                          transition={{ delay: 0.2 }}
                          className="bg-[#b8860b]/10 rounded-xl border border-[#b8860b]/20 p-4 text-center"
                        >
                          <div className="text-2xl font-bold text-[#5c3d00] mb-1">247</div>
                          <div className="text-xs text-[#5c3d00]/80">Emails Today</div>
                        </motion.div>
                        <motion.div
                          initial={{ scale: 0.8, opacity: 0 }}
                          animate={{ scale: 1, opacity: 1 }}
                          transition={{ delay: 0.4 }}
                          className="bg-[#8b4513]/10 rounded-xl border border-[#8b4513]/20 p-4 text-center"
                        >
                          <div className="text-2xl font-bold text-[#5c3d00] mb-1">89%</div>
                          <div className="text-xs text-[#5c3d00]/80">Response Rate</div>
                        </motion.div>
                      </div>
                      <motion.div
                        initial={{ y: 20, opacity: 0 }}
                        animate={{ y: 0, opacity: 1 }}
                        transition={{ delay: 0.6 }}
                        className="flex-1 bg-[#654321]/10 rounded-xl border border-[#654321]/20 p-4 flex items-center justify-center"
                      >
                        <div className="text-center">
                          <PieChart className="w-12 h-12 text-[#b8860b] mx-auto mb-2" />
                          <div className="text-sm font-medium text-[#5c3d00]">Email Analytics</div>
                          <div className="text-xs text-[#5c3d00]/80">Real-time insights</div>
                        </div>
                      </motion.div>
                    </div>
                  )}
                </motion.div>
              )}
            </AnimatePresence>
          </motion.div>
        </div>

        {/* Feature Carousel */}
        <div className="relative z-10 h-32">
          <AnimatePresence mode="wait">
            <motion.div
              key={activeFeature}
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              transition={{ duration: 0.5 }}
              className="max-w-lg"
            >
              <div className="flex items-center gap-3 mb-3">
                <div className={`p-2 rounded-lg ${features[activeFeature].color} bg-opacity-20`}>
                  {React.createElement(features[activeFeature].icon, { className: `w-5 h-5 ${features[activeFeature].textColor}` })}
                </div>
                <h3 className="text-xl font-semibold">{features[activeFeature].title}</h3>
              </div>
              <p className={`leading-relaxed ${selectedProduct === 'excelerate' ? 'text-slate-700' : 'text-amber-800'}`}>
                {features[activeFeature].description}
              </p>
            </motion.div>
          </AnimatePresence>

          {/* Indicators */}
          <div className="flex gap-2 mt-6">
            {features.map((_, idx) => (
              <div
                key={idx}
                className={`h-1 rounded-full transition-all duration-500 ${
                  idx === activeFeature 
                    ? `w-8 ${selectedProduct === 'excelerate' ? 'bg-[#0a5c2e]' : 'bg-[#b8860b]'}`
                    : `w-2 ${selectedProduct === 'excelerate' ? 'bg-[#0a5c2e]/60' : 'bg-[#b8860b]/60'}`
                }`}
              />
            ))}
          </div>
        </div>
      </div>

      {/* Right Side - Form (40%) */}
      <div className="w-full lg:w-[40%] flex flex-col justify-center px-8 lg:px-24 bg-white relative">
        <div className="max-w-[400px] mx-auto w-full">
          {/* Product Tabs */}
          <div className="mb-8">
            <h1 className="text-2xl font-bold text-slate-900 text-center mb-6">Workflow Genie</h1>
            <div className={`flex rounded-xl p-1 ${selectedProduct === 'excelerate' ? 'bg-slate-100' : 'bg-[#b8860b]/20'}`}>
              <button
                onClick={() => setSelectedProduct('excelerate')}
                className={`flex-1 py-2 px-4 rounded-lg font-medium transition-all ${
                  selectedProduct === 'excelerate'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Excelerate
              </button>
              <button
                onClick={() => setSelectedProduct('mailacer')}
                className={`flex-1 py-2 px-4 rounded-lg font-medium transition-all ${
                  selectedProduct === 'mailacer'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <span className="font-medium">Mail</span>
                <span className="italic text-amber-700 font-serif">acer</span>
              </button>
            </div>
          </div>

          <div className="mb-10">
            <h2 className="text-3xl font-bold text-slate-900 mb-3">
              {selectedProduct === 'excelerate' 
                ? (isLogin ? 'Welcome back' : 'Get started')
                : (isLogin ? 'Welcome to Mailacer' : 'Join Mailacer')
              }
            </h2>
            <p className="text-slate-500">
              {selectedProduct === 'excelerate'
                ? (isLogin ? 'Enter your details to access your workspace.' : 'Create your account now.')
                : (isLogin ? 'Sign in to your Mailacer account.' : 'Create your Mailacer account.')
              }
            </p>
          </div>

          <form onSubmit={handleSubmit} className="space-y-5">
                {/* Error Message */}
                <AnimatePresence>
                  {(errorMessage || error) && (
                    <motion.div
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -10 }}
                      className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-xl flex items-center gap-2 text-sm"
                    >
                      <AlertCircle className="w-4 h-4 flex-shrink-0" />
                      <span>{errorMessage || error}</span>
                    </motion.div>
                  )}
                </AnimatePresence>

                <div className="space-y-1.5">
                  <label className="text-sm font-medium text-slate-700">Username</label>
                  <input
                    type="text"
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    className={`w-full px-4 py-3 rounded-xl bg-slate-50 border border-slate-200 focus:border-[${selectedProduct === 'excelerate' ? '#0a5c2e' : '#b8860b'}] focus:ring-2 focus:ring-[${selectedProduct === 'excelerate' ? 'rgba(10,92,46,0.1)' : 'rgba(184,134,11,0.1)'}] outline-none transition-all text-slate-900 placeholder:text-slate-400`}
                    placeholder="johndoe"
                    disabled={isLoading}
                    autoComplete="username"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-sm font-medium text-slate-700">Password</label>
                  <input
                    type="password"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    className={`w-full px-4 py-3 rounded-xl bg-slate-50 border border-slate-200 focus:border-[${selectedProduct === 'excelerate' ? '#0a5c2e' : '#b8860b'}] focus:ring-2 focus:ring-[${selectedProduct === 'excelerate' ? 'rgba(10,92,46,0.1)' : 'rgba(184,134,11,0.1)'}] outline-none transition-all text-slate-900 placeholder:text-slate-400`}
                    placeholder="••••••••"
                    disabled={isLoading}
                    autoComplete={isLogin ? 'current-password' : 'new-password'}
                  />
                  {!isLogin && (
                    <p className="text-xs text-slate-500 mt-1">Must be at least 6 characters</p>
                  )}
                </div>

                <button
                  type="submit"
                  disabled={isLoading}
                  className={`w-full py-3.5 ${selectedProduct === 'excelerate' ? 'bg-[#0a5c2e] hover:bg-[#084a24] shadow-[#0a5c2e]/20' : 'bg-[#b8860b] hover:bg-[#a0780a] shadow-[#b8860b]/20'} text-white font-bold rounded-xl shadow-lg transition-all transform hover:scale-[1.01] active:scale-[0.99] flex items-center justify-center gap-2 mt-2`}
                >
                  {isLoading ? (
                    <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                  ) : (
                    <>
                      {isLogin ? 'Sign In' : 'Create Account'}
                      <ArrowRight className="w-4 h-4" />
                    </>
                  )}
                </button>
              </form>

              <div className="mt-8 text-center">
                <p className="text-slate-500 text-sm">
                  {isLogin 
                    ? `New to ${selectedProduct === 'excelerate' ? 'Excelerate' : 'Mailacer'}?`
                    : "Already have an account?"
                  }{' '}
                  <button
                    onClick={() => setIsLogin(!isLogin)}
                    className={`font-bold hover:underline transition-all ${selectedProduct === 'excelerate' ? 'text-[#0a5c2e]' : 'text-[#b8860b]'}`}
                  >
                    {isLogin ? 'Create an account' : 'Sign in'}
                  </button>
                </p>
              </div>
        </div>
      </div>
    </div>
  );
};
