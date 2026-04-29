import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useSpreadsheetStore } from '../stores/spreadsheetStore';
import { useSessionStore } from '../stores/sessionStore';
import { useAuthStore } from '../stores/authStore';
import { LuckysheetSpreadsheet } from '../components/spreadsheet/LuckysheetSpreadsheet';
import { Toolbar } from '../components/ui/Toolbar';
import { AIPromptInterface } from '../components/ai-interface/AIPromptInterface';
import { FileUpload } from '../components/ui/FileUpload';
import { AnimatePresence, motion } from 'framer-motion';
import { Loader, FileSpreadsheet, Clock, MessageSquare, Plus } from 'lucide-react';

export const Dashboard: React.FC = () => {
  const navigate = useNavigate();
  const [showSessionSelector, setShowSessionSelector] = useState(false);
  const [sessionsFetched, setSessionsFetched] = useState(false);
  const [isCreating, setIsCreating] = useState(false);
  const [lastChartType, setLastChartType] = useState<'bar' | 'line' | 'pie' | null>(null);
  const [lastChartCell, setLastChartCell] = useState<string | null>(null);
  const [lastChartSheet, setLastChartSheet] = useState<string | null>(null);

  const { fortuneSheets, ui } = useSpreadsheetStore();
  const {
    currentSession,
    sessions,
    currentFileId,
    chatHistory,
    isLoadingSessions,
    createSession,
    loadSession,
    loadSessions,
    clearCurrentSession,
  } = useSessionStore();
  const { isAuthenticated } = useAuthStore();

  // Redirect if not logged in
  useEffect(() => {
    if (!isAuthenticated) navigate('/');
  }, [isAuthenticated, navigate]);

  // Load sessions once on mount — never block UI
  useEffect(() => {
    if (!isAuthenticated) return;
    loadSessions(50, true)
      .catch(() => {})
      .finally(() => setSessionsFetched(true));
  }, [isAuthenticated, loadSessions]);

  // After sessions load: if there are previous sessions show the picker,
  // otherwise show the "no session" screen where user clicks Create Session.
  // NO auto-create here — let the user click.
  useEffect(() => {
    if (!sessionsFetched || currentSession) return;
    if (sessions.length > 0) setShowSessionSelector(true);
    // sessions.length === 0 → just show the "no active session" buttons, don't auto-do anything
  }, [sessionsFetched]); // intentionally only runs once when fetch completes

  // Detect chart type from latest AI response
  useEffect(() => {
    const last = chatHistory[chatHistory.length - 1];
    if (!last || last.role !== 'assistant') return;
    for (const op of (last.operations ?? []) as any[]) {
      const t: string = op.tool ?? '';
      const type = t.includes('bar') ? 'bar' : t.includes('line') ? 'line' : t.includes('pie') ? 'pie' : null;
      if (type) {
        setLastChartType(type as 'bar' | 'line' | 'pie');
        setLastChartCell(op.params?.position ?? op.result?.position ?? null);
        // Sheet name: from params, result, or data_range prefix (e.g. "Sales!A1:C6")
        const sheetFromRange = (op.params?.data_range as string | undefined)?.split('!')?.[0]?.replace(/'/g, '');
        setLastChartSheet(op.params?.sheet_name ?? op.result?.sheet_name ?? sheetFromRange ?? null);
        return;
      }
    }
  }, [chatHistory]);

  // ── handlers ──────────────────────────────────────────────────────────────

  const handleCreateSession = async () => {
    if (isCreating) return;
    setIsCreating(true);
    try {
      clearCurrentSession();
      await createSession();
      setShowSessionSelector(false);
    } catch (err: any) {
      console.error('Failed to create session:', err);
      if (err?.status === 401) {
        window.location.href = '/';
      } else {
        alert('Could not create session. Please try again.');
      }
    } finally {
      setIsCreating(false);
    }
  };

  const handleSessionSelect = async (sessionId: string) => {
    try {
      await loadSession(sessionId);
      setShowSessionSelector(false);
    } catch (err) {
      console.error('Failed to load session:', err);
    }
  };

  // ── render ────────────────────────────────────────────────────────────────

  return (
    <div className="h-screen w-screen flex flex-col bg-[#f8f9fa] fixed inset-0 overflow-hidden">

      {/* Session picker modal — only shown when user has previous sessions */}
      <AnimatePresence>
        {showSessionSelector && (
          <motion.div
            initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4"
            onClick={() => setShowSessionSelector(false)}
          >
            <motion.div
              initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full max-h-[80vh] overflow-hidden"
              onClick={e => e.stopPropagation()}
            >
              <div className="p-6 border-b border-gray-200">
                <h2 className="text-2xl font-bold text-gray-900">Your Sessions</h2>
                <p className="text-gray-600 mt-1">Continue where you left off, or start a new session</p>
              </div>

              <div className="p-6 max-h-[60vh] overflow-y-auto">
                {/* Create new */}
                <button
                  onClick={handleCreateSession}
                  disabled={isCreating}
                  className="w-full mb-6 p-4 border-2 border-dashed border-gray-300 rounded-xl hover:border-[#107c41] hover:bg-green-50 transition-all group disabled:opacity-60 disabled:cursor-not-allowed"
                >
                  <div className="flex items-center gap-3">
                    <div className="w-12 h-12 bg-[#107c41] rounded-xl flex items-center justify-center">
                      {isCreating
                        ? <Loader className="w-5 h-5 text-white animate-spin" />
                        : <Plus className="w-6 h-6 text-white" />}
                    </div>
                    <div className="text-left">
                      <h3 className="font-semibold text-gray-900">
                        {isCreating ? 'Creating…' : 'Create New Session'}
                      </h3>
                      <p className="text-sm text-gray-500">Start fresh with a new file</p>
                    </div>
                  </div>
                </button>

                {/* Previous sessions */}
                <div className="border-t border-gray-200 pt-6">
                  <h3 className="text-lg font-semibold text-gray-900 mb-4">Previous Sessions</h3>
                  <div className="space-y-3">
                    {isLoadingSessions ? (
                      <div className="text-center py-8">
                        <Loader className="w-8 h-8 text-gray-400 animate-spin mx-auto" />
                      </div>
                    ) : sessions.length === 0 ? (
                      <div className="text-center py-8 text-gray-500">
                        <MessageSquare className="w-12 h-12 mx-auto mb-2 opacity-40" />
                        <p>No previous sessions</p>
                      </div>
                    ) : sessions.map(session => (
                      <button
                        key={session.session_id}
                        onClick={() => handleSessionSelect(session.session_id)}
                        className="w-full p-4 bg-gray-50 hover:bg-gray-100 rounded-xl transition-all text-left group border border-gray-200 hover:border-[#107c41]"
                      >
                        <div className="flex items-start gap-3">
                          <div className="w-12 h-12 bg-white rounded-lg flex items-center justify-center border border-gray-200 group-hover:border-[#107c41] transition-colors">
                            <FileSpreadsheet className="w-6 h-6 text-[#107c41]" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center justify-between mb-1">
                              <h3 className="font-semibold text-gray-900 truncate">
                                {session.filename || session.summary || 'Untitled Session'}
                              </h3>
                              <span className="text-xs text-gray-500 flex items-center gap-1 ml-2 shrink-0">
                                <Clock className="w-3 h-3" />
                                {new Date(session.updated_at).toLocaleDateString()}
                              </span>
                            </div>
                            <p className="text-sm text-gray-500 line-clamp-2">
                              {session.last_response || session.first_prompt || 'No messages yet'}
                            </p>
                            <div className="flex items-center gap-3 mt-2 text-xs text-gray-400">
                              <span>{session.message_count} messages</span>
                              {session.filename && (
                                <span className="flex items-center gap-1">
                                  <FileSpreadsheet className="w-3 h-3" />{session.filename}
                                </span>
                              )}
                            </div>
                          </div>
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <Toolbar />

      <div className="flex-1 flex min-h-0 relative">
        <div className="flex-1 flex flex-col bg-white min-w-0">

          {/* ── No session yet ── */}
          {!currentSession ? (
            <div className="flex-1 flex items-center justify-center p-8">
              <div className="text-center max-w-md">
                {isCreating ? (
                  <>
                    <Loader className="w-12 h-12 text-[#107c41] animate-spin mx-auto mb-4" />
                    <h2 className="text-xl font-semibold text-gray-700">Creating your session…</h2>
                  </>
                ) : !sessionsFetched ? (
                  <>
                    <Loader className="w-12 h-12 text-gray-300 animate-spin mx-auto mb-4" />
                    <p className="text-gray-400 text-sm">Loading…</p>
                  </>
                ) : (
                  <>
                    <FileSpreadsheet className="w-16 h-16 text-gray-200 mx-auto mb-5" />
                    <h2 className="text-xl font-semibold text-gray-800 mb-2">Welcome to WorkflowGenie</h2>
                    <p className="text-gray-500 mb-6">Create a session to upload an Excel file and start working with AI.</p>
                    <div className="flex justify-center gap-3">
                      <button
                        onClick={handleCreateSession}
                        className="px-5 py-2.5 bg-[#107c41] text-white rounded-lg font-medium hover:bg-[#0c5f31] transition"
                      >
                        + New Session
                      </button>
                      {sessions.length > 0 && (
                        <button
                          onClick={() => setShowSessionSelector(true)}
                          className="px-5 py-2.5 border border-gray-300 rounded-lg text-gray-700 font-medium hover:border-[#107c41] hover:text-[#107c41] transition"
                        >
                          Open Previous
                        </button>
                      )}
                    </div>
                  </>
                )}
              </div>
            </div>

          /* ── Session active, no file yet ── */
          ) : !currentFileId ? (
            <div className="flex-1 flex items-center justify-center p-8">
              <div className="max-w-2xl w-full">
                <div className="text-center mb-8">
                  <h1 className="text-3xl font-bold text-gray-800 mb-2">Upload Your Excel File</h1>
                  <p className="text-gray-500">Drop an .xlsx file to start working with AI assistance</p>
                </div>
                <FileUpload className="mb-6" />
                <div className="text-center mt-4 flex justify-center gap-4 text-sm">
                  <button onClick={() => setShowSessionSelector(true)} className="text-[#107c41] hover:underline font-medium">
                    Switch session
                  </button>
                  <span className="text-gray-300">|</span>
                  <button onClick={handleCreateSession} disabled={isCreating} className="text-[#107c41] hover:underline font-medium disabled:opacity-50">
                    New session
                  </button>
                </div>
              </div>
            </div>

          /* ── File loaded but FortuneSheet data not ready yet ── */
          ) : !fortuneSheets ? (
            <div className="flex-1 flex items-center justify-center p-8">
              <div className="text-center">
                <FileSpreadsheet className="w-16 h-16 text-[#107c41] mx-auto mb-4" />
                <h2 className="text-xl font-semibold text-gray-800 mb-2">File ready</h2>
                <p className="text-gray-500">Open the AI Assistant panel to start sending commands.</p>
              </div>
            </div>

          /* ── Spreadsheet view ── */
          ) : (
            <div className="flex-1 min-h-0 relative">
              <LuckysheetSpreadsheet lastChartType={lastChartType} chartCell={lastChartCell} chartSheet={lastChartSheet} />
            </div>
          )}
        </div>

        {/* AI sidebar */}
        <AnimatePresence>
          {ui.sidebarOpen && (
            <motion.div
              initial={{ width: 0, opacity: 0 }}
              animate={{ width: 400, opacity: 1, zIndex: 4000 }}
              exit={{ width: 0, opacity: 0 }}
              transition={{ type: 'spring', stiffness: 300, damping: 30 }}
              className="h-full border-l border-gray-300 bg-white shadow-xl z-40 flex flex-col"
            >
              <AIPromptInterface />
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
};
