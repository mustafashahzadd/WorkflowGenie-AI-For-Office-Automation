/**
 * AI Prompt Interface Component
 * Provides a state-of-the-art, Cursor-like interface for users to interact with AI agents
 * Features real-time operation display, chat-like interaction, and smooth animations
 */

import React, { useState, useRef, useEffect } from 'react';
import { useSessionStore } from '../../stores/sessionStore';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';
import {
  Send,
  Bot,
  User,
  AlertCircle,
  CheckCircle,
  Loader,
  Sparkles,
  Zap,
  BarChart3,
  Table,
  ArrowRight,
  X,
  ChevronDown,
  ChevronRight,
  Terminal,
  Info,
  SlidersHorizontal,
  Plus,
  Minus
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';

interface AIPromptInterfaceProps {
  className?: string;
}

/**
 * Operation status icon component with animations
 */
const OperationStatusIcon: React.FC<{ status: string }> = ({ status }) => {
  switch (status) {
    case 'completed':
      return <CheckCircle className="w-3.5 h-3.5 text-emerald-500" />;
    case 'failed':
      return <AlertCircle className="w-3.5 h-3.5 text-red-500" />;
    default:
      return <Loader className="w-3.5 h-3.5 text-blue-500 animate-spin" />;
  }
};

/**
 * Individual operation item component
 */
interface OperationItemProps {
  operation: any;
  index: number;
}

const OperationItem: React.FC<OperationItemProps> = ({ operation, index }) => {
  const status = operation.status || 'completed';

  // Extract detailed changes from operation result
  const getDetailedChanges = () => {
    if (!operation.result) return [];

    const result = operation.result;

    // Handle bulk_update_all with updated_rows
    if (result.updated_rows && Array.isArray(result.updated_rows)) {
      return result.updated_rows.map((change: any) => ({
        type: 'update',
        description: `Row ${change.row}: ${change.old_value} → ${change.new_value}`,
        row: change.row,
        oldValue: change.old_value,
        newValue: change.new_value
      }));
    }

    // Handle add_row
    if (result.row_number && result.data) {
      return [{
        type: 'add',
        description: `Added row ${result.row_number} with data: ${Array.isArray(result.data) ? result.data.join(', ') : result.data}`,
        row: result.row_number
      }];
    }

    // Handle single update
    if (result.updated_count === 1) {
      return [{
        type: 'update',
        description: operation.description || result.message || 'Updated 1 item'
      }];
    }

    return [];
  };

  const detailedChanges = getDetailedChanges();

  return (
    <motion.div
      initial={{ opacity: 0, x: -10 }}
      animate={{ opacity: 1, x: 0 }}
      transition={{ delay: index * 0.05 }}
      className="py-2 px-1 group"
    >
      <div className="flex items-start gap-3">
        <div className="mt-0.5 opacity-70 group-hover:opacity-100 transition-opacity">
          <OperationStatusIcon status={status} />
        </div>
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span className="text-xs font-mono text-gray-500 uppercase tracking-wider">
              {operation.tool?.replace(/_/g, ' ') || 'System'}
            </span>
          </div>
          <p className="text-xs text-gray-700 mt-0.5 font-mono">
            {operation.description || operation.message}
          </p>

          {/* Show detailed changes if available */}
          {detailedChanges.length > 0 && (
            <div className="mt-2 ml-4 space-y-1">
              {detailedChanges.map((change: any, idx: number) => (
                <div
                  key={idx}
                  className="text-[11px] font-mono text-gray-600 flex items-start gap-2 py-0.5"
                >
                  <span className="text-gray-400">•</span>
                  <span>{change.description}</span>
                </div>
              ))}
            </div>
          )}

          {operation.error && (
            <p className="text-xs text-red-500 mt-1 font-mono bg-red-50 p-1 rounded border border-red-100 inline-block">
              Error: {operation.error}
            </p>
          )}
        </div>
      </div>
    </motion.div>
  );
};

/**
 * Tool usage summary modal
 */
interface ToolUsageSummaryProps {
  operations: any[];
  onClose: () => void;
}

const ToolUsageSummary: React.FC<ToolUsageSummaryProps> = ({ operations, onClose }) => {
  // Calculate tool usage statistics
  const toolStats = operations.reduce((acc: any, op: any) => {
    const tool = op.tool || 'unknown';
    if (!acc[tool]) {
      acc[tool] = { count: 0, operations: [] };
    }
    acc[tool].count++;
    acc[tool].operations.push(op);
    return acc;
  }, {});

  const totalOperations = operations.length;
  const successfulOperations = operations.filter(op => op.status === 'completed').length;
  const failedOperations = operations.filter(op => op.status === 'failed').length;

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4"
      onClick={onClose}
    >
      <motion.div
        initial={{ scale: 0.95, opacity: 0 }}
        animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-lg shadow-xl max-w-md w-full max-h-[80vh] overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="p-4 border-b border-gray-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Info className="w-4 h-4 text-blue-500" />
            <h3 className="text-sm font-semibold text-gray-900">Operation Details</h3>
          </div>
          <button
            onClick={onClose}
            className="p-1 hover:bg-gray-100 rounded-md text-gray-400 hover:text-gray-600 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-4 overflow-y-auto max-h-[calc(80vh-100px)]">
          {/* Summary Stats */}
          <div className="grid grid-cols-3 gap-3 mb-4">
            <div className="bg-gray-50 rounded-lg p-3 text-center">
              <div className="text-2xl font-bold text-gray-900">{totalOperations}</div>
              <div className="text-xs text-gray-500 mt-1">Total</div>
            </div>
            <div className="bg-emerald-50 rounded-lg p-3 text-center">
              <div className="text-2xl font-bold text-emerald-600">{successfulOperations}</div>
              <div className="text-xs text-gray-500 mt-1">Success</div>
            </div>
            <div className="bg-red-50 rounded-lg p-3 text-center">
              <div className="text-2xl font-bold text-red-600">{failedOperations}</div>
              <div className="text-xs text-gray-500 mt-1">Failed</div>
            </div>
          </div>

          {/* Tool Usage Breakdown */}
          <div className="space-y-3">
            <h4 className="text-xs font-semibold text-gray-700 uppercase tracking-wider">Tools Used</h4>
            {Object.entries(toolStats).map(([tool, stats]: [string, any]) => (
              <div key={tool} className="bg-gray-50 rounded-lg p-3">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-sm font-mono text-gray-700">
                    {tool.replace(/_/g, ' ')}
                  </span>
                  <span className="text-xs font-semibold bg-gray-200 text-gray-700 px-2 py-0.5 rounded-full">
                    {stats.count}x
                  </span>
                </div>
                <div className="space-y-1">
                  {stats.operations.map((op: any, idx: number) => (
                    <div key={idx} className="text-xs text-gray-600 flex items-start gap-1.5">
                      <OperationStatusIcon status={op.status} />
                      <span className="flex-1 font-mono">{op.description}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </motion.div>
    </motion.div>
  );
};

/**
 * Chat message component
 */
interface ChatMessageProps {
  message: any;
  isHistory?: boolean;
}

const ChatMessage: React.FC<ChatMessageProps> = ({ message, isHistory = false }) => {
  const isUser = message.role === 'user';
  const [isExpanded, setIsExpanded] = useState(true);
  const [showToolUsage, setShowToolUsage] = useState(false);

  // Auto-collapse operations if message is history
  useEffect(() => {
    if (isHistory) setIsExpanded(false);
  }, [isHistory]);

  if (isUser) {
    return (
      <motion.div
        initial={{ opacity: 0, y: 10 }}
        animate={{ opacity: isHistory ? 0.5 : 1, y: 0 }}
        className="flex justify-end mb-6"
      >
        <div className="bg-[#052e16] text-white px-4 py-2.5 rounded-2xl rounded-tr-sm max-w-[85%] shadow-sm">
          <p className="text-sm leading-relaxed">{message.content}</p>
        </div>
      </motion.div>
    );
  }

  return (
    <motion.div
      initial={{ opacity: 0 }}
      animate={{ opacity: isHistory ? 0.7 : 1 }}
      className="mb-8 pl-1"
    >
      {/* AI Message Content with Info Button */}
      <div className="flex items-start gap-2 mb-3">
        <div className="flex-1 bg-[#0a5c2e]/10 text-[#052e16] px-4 py-3 rounded-2xl rounded-tl-sm shadow-sm text-sm leading-relaxed whitespace-pre-wrap font-medium">
          {message.content}
        </div>
        {message.operations && message.operations.length > 0 && (
          <button
            onClick={() => setShowToolUsage(true)}
            className="flex-shrink-0 p-1.5 hover:bg-gray-100 rounded-md text-gray-400 hover:text-blue-600 transition-colors group"
            title="View operation details"
          >
            <Info className="w-3.5 h-3.5" />
          </button>
        )}
      </div>

      {/* Tool Usage Summary Modal */}
      <AnimatePresence>
        {showToolUsage && (
          <ToolUsageSummary
            operations={message.operations}
            onClose={() => setShowToolUsage(false)}
          />
        )}
      </AnimatePresence>

      {/* Operations / Thinking Process */}
      {message.operations && message.operations.length > 0 && (
        <div className="mt-2">
          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="flex items-center gap-2 text-xs text-gray-400 hover:text-gray-600 transition-colors mb-2 select-none"
          >
            {isExpanded ? <ChevronDown className="w-3 h-3" /> : <ChevronRight className="w-3 h-3" />}
            <span className="font-medium uppercase tracking-wider text-[10px]">Process Log</span>
            <span className="bg-gray-100 text-gray-500 px-1.5 py-0.5 rounded text-[9px] font-mono">
              {message.operations.length} steps
            </span>
          </button>

          <AnimatePresence>
            {isExpanded && (
              <motion.div
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: 'auto', opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                className="overflow-hidden"
              >
                <div className="border-l-2 border-gray-100 pl-3 py-1 space-y-1">
                  {message.operations.map((op: any, idx: number) => (
                    <OperationItem key={idx} operation={op} index={idx} />
                  ))}
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      )}
    </motion.div>
  );
};

const ThinkingIndicator = () => (
  <motion.div
    initial={{ opacity: 0 }}
    animate={{ opacity: 1 }}
    className="flex items-center gap-2 text-gray-400 pl-1"
  >
    <Loader className="w-3.5 h-3.5 animate-spin" />
    <span className="text-xs font-medium">Thinking...</span>
  </motion.div>
);

// Example suggestions for the empty state
const suggestions = [
  { text: 'Analyze sales trends', icon: BarChart3 },
  { text: 'Format as currency', icon: Zap },
  { text: 'Add total row', icon: Table },
];

/**
 * Main AI Interface Component
 */
export const AIPromptInterface: React.FC<AIPromptInterfaceProps> = ({ className = '' }) => {
  const {
    chatHistory,
    sendMessage,
    isSendingMessage,
    currentSession,
    currentFileId
  } = useSessionStore();

  const { toggleSidebar } = useSpreadsheetStore();

  const [inputValue, setInputValue] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  const [provider, setProvider] = useState<'openai' | 'claude' | null>(null);
  const [useRag, setUseRag] = useState<boolean | null>(null);
  const [ragTopK, setRagTopK] = useState(4);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [chatHistory, isSendingMessage]);

  // Focus input on mount
  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!inputValue.trim() || isSendingMessage) return;

    if (!currentSession) {
      setError('No active session');
      return;
    }

    if (!currentFileId) {
      setError('Please upload a file first');
      return;
    }

    const prompt = inputValue;
    setInputValue('');
    setError(null);

    try {
      await sendMessage(prompt, { provider, useRag, ragTopK });
    } catch (err: any) {
      setError(err.message || 'Failed to send message');
      setInputValue(prompt);
    }
  };

  return (
    <div className={`h-full flex flex-col bg-white ${className}`}>
      {/* Header */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-gray-100">
        <div className="flex items-center gap-2">
          <Sparkles className="w-4 h-4 text-emerald-600" />
          <span className="text-sm font-semibold text-gray-800">AI Assistant</span>
        </div>
        <button
          onClick={toggleSidebar}
          className="p-1 hover:bg-gray-100 rounded-md text-gray-400 hover:text-gray-600 transition-colors"
        >
          <X className="w-4 h-4" />
        </button>
      </div>

      {/* Messages Area */}
      <div className="flex-1 overflow-y-auto p-5 custom-scrollbar">
        {chatHistory.length === 0 ? (
          <div className="h-full flex flex-col items-center justify-center text-center opacity-0 animate-fade-in" style={{ animationFillMode: 'forwards' }}>
            <div className="w-12 h-12 bg-gray-50 rounded-xl flex items-center justify-center mb-4">
              <Bot className="w-6 h-6 text-gray-400" />
            </div>
            <h3 className="text-sm font-medium text-gray-900 mb-1">How can I help?</h3>
            <p className="text-xs text-gray-500 max-w-[200px] mb-6">
              I can analyze data, create charts, and edit your spreadsheet.
            </p>

            {currentFileId && (
              <div className="flex flex-wrap justify-center gap-2 max-w-[250px]">
                {suggestions.map((s, i) => (
                  <button
                    key={i}
                    onClick={() => setInputValue(s.text)}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-gray-200 hover:border-emerald-500/50 hover:bg-emerald-50/50 rounded-full text-xs text-gray-600 transition-all cursor-pointer"
                  >
                    <s.icon className="w-3 h-3" />
                    {s.text}
                  </button>
                ))}
              </div>
            )}
          </div>
        ) : (
          <div className="space-y-6">
            {chatHistory.map((message, index) => {
              const isLastMessage = index === chatHistory.length - 1;
              const isSecondToLast = index === chatHistory.length - 2;
              const lastMessage = chatHistory[chatHistory.length - 1];

              let isHistory = true;
              if (lastMessage.role === 'assistant') {
                if (isLastMessage || (isSecondToLast && message.role === 'user')) {
                  isHistory = false;
                }
              } else {
                if (isLastMessage) {
                  isHistory = false;
                }
              }

              return (
                <ChatMessage
                  key={message.timestamp}
                  message={message}
                  isHistory={isHistory}
                />
              );
            })}

            {isSendingMessage && <ThinkingIndicator />}
            <div ref={messagesEndRef} />
          </div>
        )}
      </div>

      {/* Error display */}
      {error && (
        <div className="px-4 pb-2">
          <div className="bg-red-50 border border-red-100 rounded-md p-2 flex items-center gap-2">
            <AlertCircle className="w-3 h-3 text-red-500 flex-shrink-0" />
            <p className="text-xs text-red-600">{error}</p>
          </div>
        </div>
      )}

      {/* Input Area */}
      <div className="p-4 border-t border-gray-100 bg-white">
        {/* Settings Panel */}
        <AnimatePresence>
          {showSettings && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="overflow-hidden"
            >
              <div className="mb-3 p-3 bg-gray-50 rounded-xl border border-gray-200 space-y-3">
                {/* Provider */}
                <div className="flex items-center justify-between gap-3">
                  <span className="text-[11px] font-medium text-gray-500 uppercase tracking-wider w-16 flex-shrink-0">Provider</span>
                  <div className="flex items-center bg-white rounded-lg border border-gray-200 p-0.5 gap-0.5">
                    {([['auto', null], ['openai', 'openai'], ['claude', 'claude']] as [string, 'openai' | 'claude' | null][]).map(([label, val]) => (
                      <button
                        key={label}
                        type="button"
                        onClick={() => setProvider(val)}
                        className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-all ${
                          provider === val
                            ? 'bg-gray-900 text-white shadow-sm'
                            : 'text-gray-500 hover:text-gray-800'
                        }`}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Use RAG */}
                <div className="flex items-center justify-between gap-3">
                  <span className="text-[11px] font-medium text-gray-500 uppercase tracking-wider w-16 flex-shrink-0">RAG</span>
                  <div className="flex items-center bg-white rounded-lg border border-gray-200 p-0.5 gap-0.5">
                    {([['auto', null], ['on', true], ['off', false]] as [string, boolean | null][]).map(([label, val]) => (
                      <button
                        key={label}
                        type="button"
                        onClick={() => setUseRag(val)}
                        className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-all ${
                          useRag === val
                            ? 'bg-gray-900 text-white shadow-sm'
                            : 'text-gray-500 hover:text-gray-800'
                        }`}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                </div>

                {/* RAG Top-K */}
                <div className="flex items-center justify-between gap-3">
                  <span className="text-[11px] font-medium text-gray-500 uppercase tracking-wider w-16 flex-shrink-0">Top-K</span>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setRagTopK(k => Math.max(1, k - 1))}
                      className="w-6 h-6 flex items-center justify-center rounded-md border border-gray-200 bg-white text-gray-500 hover:text-gray-800 hover:border-gray-300 transition-all"
                    >
                      <Minus className="w-2.5 h-2.5" />
                    </button>
                    <span className="w-6 text-center text-xs font-mono font-semibold text-gray-700">{ragTopK}</span>
                    <button
                      type="button"
                      onClick={() => setRagTopK(k => Math.min(20, k + 1))}
                      className="w-6 h-6 flex items-center justify-center rounded-md border border-gray-200 bg-white text-gray-500 hover:text-gray-800 hover:border-gray-300 transition-all"
                    >
                      <Plus className="w-2.5 h-2.5" />
                    </button>
                    <span className="text-[10px] text-gray-400">(1–20)</span>
                  </div>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        <form onSubmit={handleSubmit} className="relative">
          <div className="relative flex items-center bg-gray-50 hover:bg-white focus-within:bg-white rounded-xl border border-gray-200 focus-within:border-emerald-500 focus-within:ring-1 focus-within:ring-emerald-500/20 transition-all shadow-sm">
            <input
              ref={inputRef}
              type="text"
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              placeholder={currentFileId ? "Ask AI to edit..." : "Upload a file first..."}
              className="w-full pl-4 pr-10 py-3 bg-transparent border-none focus:outline-none text-sm text-gray-800 placeholder-gray-400"
              disabled={isSendingMessage || !currentFileId}
            />
            <div className="absolute right-2 flex items-center">
              <button
                type="submit"
                disabled={!inputValue.trim() || isSendingMessage || !currentFileId}
                className="p-1.5 bg-gray-900 text-white rounded-lg hover:bg-gray-800 disabled:opacity-30 disabled:cursor-not-allowed transition-all"
              >
                {isSendingMessage ? (
                  <Loader className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <ArrowRight className="w-3.5 h-3.5" />
                )}
              </button>
            </div>
          </div>
          <div className="flex justify-between items-center mt-2 px-1">
            <button
              type="button"
              onClick={() => setShowSettings(s => !s)}
              className={`flex items-center gap-1 text-[10px] font-medium transition-colors ${
                showSettings ? 'text-emerald-600' : 'text-gray-400 hover:text-gray-600'
              }`}
            >
              <SlidersHorizontal className="w-3 h-3" />
              Options
              {(provider !== null || useRag !== null || ragTopK !== 4) && (
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 ml-0.5" />
              )}
            </button>
            <span className="text-[10px] text-gray-300 font-mono">
              Return to send
            </span>
          </div>
        </form>
      </div>
    </div>
  );
};
