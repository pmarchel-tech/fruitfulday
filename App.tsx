import React, { useState, useEffect, useMemo, useRef, Component, ErrorInfo, ReactNode } from 'react';
import { Menu, Plus, Search, Filter, Calendar, ChevronDown, ChevronLeft, ChevronRight, Save, AlertCircle, AlertTriangle, CheckCircle2, Check, X, LayoutList, FileText, Loader2, LogOut, Users, UserPlus, Shield, Mail, User as UserIcon, CheckSquare, Square, Settings, Lock, ArrowUpDown, ArrowUp, ArrowDown, History, Layers, ArrowUpRight, LogIn, Trash2, LayoutDashboard, Trophy, Flame, TrendingUp, Target, Sparkles, Award, Globe, Sun, Moon, Clock } from 'lucide-react';
import { Task, TaskUpdate, TaskStatus, FilterTime, SortOrder, User, UserRole, Tag, AiTokenRecord } from './types';
import { STATUS_COLORS, CATEGORIES, STATUSES, INDONESIAN_HOLIDAYS } from './constants';
import { 
  getTasks, 
  saveTasks,
  saveSingleTask, 
  deleteTask, 
  getUpdates, 
  saveSingleUpdate, 
  deleteUpdate, 
  loginUser, 
  registerUser, 
  logoutUser, 
  getCurrentUser, 
  getUsers, 
  addTeamMembers, 
  updateUser,
  subscribeToTasks,
  subscribeToUpdates,
  subscribeToUsers,
  subscribeToTags,
  saveSingleTag,
  importRestoreDatabase,
  saveAiTokenRecord,
  subscribeToAiTokenRecords,
  backfillTaskUpdates
} from './services/supabaseService';
import SwipeableCard from './components/SwipeableCard';
import { ErrorBoundary } from './components/ErrorBoundary';
import { auth, loginWithGoogle, loginAnonymously, getCachedAccessToken } from './firebase';
import { createExportSpreadsheet, populateSpreadsheetData, importSpreadsheetData, parseCSV } from './services/googleSheetsService';
import { supabase } from './services/supabaseClient';
import * as XLSX from 'xlsx';
import firebaseConfig from './firebase-applet-config.json';

// --- Helpers ---
const isValidUuid = (str: string): boolean => {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(str);
};

const generateUUID = (): string => {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
};

const getLocalDate = (d: Date = new Date()) => {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const parseExcelDate = (val: any): string => {
  if (!val) return getLocalDate();
  if (val instanceof Date) {
    return getLocalDate(val);
  }
  
  const num = Number(val);
  if (!isNaN(num) && num > 0) {
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + num * 24 * 60 * 60 * 1000);
    return getLocalDate(date);
  }
  
  const dateStr = String(val).trim();
  const parsedTime = Date.parse(dateStr);
  if (!isNaN(parsedTime)) {
    return getLocalDate(new Date(parsedTime));
  }
  
  return getLocalDate();
};

const parseExcelTimestamp = (val: any): number => {
  if (!val) return Date.now();
  if (val instanceof Date) {
    return val.getTime();
  }
  
  const num = Number(val);
  if (!isNaN(num) && num > 0) {
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + num * 24 * 60 * 60 * 1000);
    return date.getTime();
  }
  
  const parsedTime = Date.parse(String(val).trim());
  if (!isNaN(parsedTime)) {
    return parsedTime;
  }
  
  return Date.now();
};

const findSheetName = (workbook: any, keywords: string[], fallbackIdx: number): string => {
  const names = workbook.SheetNames;
  for (const name of names) {
    const lowerName = name.toLowerCase();
    if (keywords.some(kw => lowerName.includes(kw))) {
      return name;
    }
  }
  return names[fallbackIdx] || names[0];
};

const getMostRecentScheduledSlot = (now: Date): Date => {
  const check = new Date(now);
  check.setHours(17, 0, 0, 0);
  
  while (true) {
    const day = check.getDay();
    if ((day === 1 || day === 3 || day === 5) && check.getTime() <= now.getTime()) {
      return check;
    }
    check.setDate(check.getDate() - 1);
    check.setHours(17, 0, 0, 0);
  }
};

const getNextScheduledSlot = (now: Date): Date => {
  const check = new Date(now);
  check.setHours(17, 0, 0, 0);
  if (check.getTime() <= now.getTime()) {
    check.setDate(check.getDate() + 1);
  }
  while (true) {
    const day = check.getDay();
    if (day === 1 || day === 3 || day === 5) {
      return check;
    }
    check.setDate(check.getDate() + 1);
    check.setHours(17, 0, 0, 0);
  }
};

const formatDDMMM = (dateStr: string) => {
  if (!dateStr) return '';
  try {
    const [year, month, day] = dateStr.split('-').map(Number);
    const date = new Date(year, month - 1, day);
    if (isNaN(date.getTime())) return dateStr;
    const dd = day.toString().padStart(2, '0');
    const mmm = date.toLocaleDateString('en-US', { month: 'short' });
    return `${dd}-${mmm}`;
  } catch (e) {
    return dateStr;
  }
};

const getDaysRemaining = (targetDate: string) => {
  let target: Date;
  try { const [y, m, d] = targetDate.split('-').map(Number); target = new Date(y, m - 1, d); } catch(e) { target = new Date(targetDate); }
  const today = new Date(); today.setHours(0,0,0,0);
  const diffTime = target.getTime() - today.getTime();
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
};

const getDaysLabel = (days: number) => {
  if (days < 0) return <span className="text-red-600 font-bold">{Math.abs(days)}d overdue</span>;
  if (days === 0) return <span className="text-amber-600 font-bold">Due Today</span>;
  return <span className="text-emerald-600">{days} days left</span>;
};

const getDaysLabelLocalized = (days: number, lang: 'EN' | 'ID') => {
  if (lang === 'ID') {
    if (days < 0) return <span className="text-rose-500 font-bold">{Math.abs(days)} hari terlambat</span>;
    if (days === 0) return <span className="text-amber-500 font-bold">Hari Ini</span>;
    return <span className="text-emerald-500">{days} hari tersisa</span>;
  }
  if (days < 0) return <span className="text-rose-500 font-bold">{Math.abs(days)}d overdue</span>;
  if (days === 0) return <span className="text-amber-500 font-bold">Due Today</span>;
  return <span className="text-emerald-600">{days} days left</span>;
};

const TRANSLATIONS = {
  EN: {
    fruitfulDay: "Fruitful Day",
    tasks: "Tasks",
    doneThisWeek: "Done this week",
    tagUsage: "Inputs & Actions",
    week: "Week",
    month: "Month",
    year: "Year",
    dailyUpdateCounter: "Daily update counter",
    priorityTasks: "Priority Tasks",
    viewAll: "View All",
    lastUpdated: "Last updated",
    completed: "completed",
    active: "Active",
    overdue: "Overdue",
    daysRemaining: "days remaining",
    dayRemaining: "day remaining",
    today: "Today",
    tomorrow: "Tomorrow",
    yesterday: "Yesterday",
    overdueLabel: "overdue",
    addReport: "Add Report / Update",
    markComplete: "Mark Complete",
    markIncomplete: "Mark Incomplete",
    language: "Language",
    theme: "Theme",
    signOut: "Sign Out",
    themeLight: "Light",
    themeDark: "Dark"
  },
  ID: {
    fruitfulDay: "Hari Produktif",
    tasks: "Tugas Aktif",
    doneThisWeek: "Selesai minggu ini",
    tagUsage: "Input & Tindakan",
    week: "Minggu",
    month: "Bulan",
    year: "Tahun",
    dailyUpdateCounter: "Jumlah Update Harian",
    priorityTasks: "Tugas Prioritas",
    viewAll: "Lihat Semua",
    lastUpdated: "Terakhir kali diperbarui",
    completed: "selesai",
    active: "Aktif",
    overdue: "Terlambat",
    daysRemaining: "hari tersisa",
    dayRemaining: "hari tersisa",
    today: "Hari Ini",
    tomorrow: "Besok",
    yesterday: "Kemarin",
    overdueLabel: "terlambat",
    addReport: "Tambah Laporan / Update",
    markComplete: "Tandai Selesai",
    markIncomplete: "Tandai Belum Selesai",
    language: "Bahasa",
    theme: "Tema",
    signOut: "Keluar",
    themeLight: "Terang",
    themeDark: "Gelap"
  }
};

// --- App Component ---

type TaskSortOption = 'DATE_ASC' | 'DATE_DESC' | 'STATUS_ASC' | 'STATUS_DESC' | 'CATEGORY_ASC' | 'CATEGORY_DESC';

const App: React.FC = () => {
  // --- Auth State ---
  const [currentUser, setCurrentUser] = useState<User | null>(() => {
    try {
      const cached = localStorage.getItem('dailypulse_cached_user');
      return cached ? JSON.parse(cached) : null;
    } catch {
      return null;
    }
  });
  const [authMode, setAuthMode] = useState<'LOGIN' | 'REGISTER'>('LOGIN');
  const [authUsername, setAuthUsername] = useState('');
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authRole, setAuthRole] = useState<UserRole>('MEMBER');
  const [authError, setAuthError] = useState('');
  const [authLoading, setAuthLoading] = useState(false);

  // --- App State ---
  const [tasks, setTasks] = useState<Task[]>([]);
  const [updates, setUpdates] = useState<TaskUpdate[]>([]);
  const [aiTokenRecords, setAiTokenRecords] = useState<AiTokenRecord[]>([]);
  const [tokenFilterType, setTokenFilterType] = useState<'day' | 'weekly' | 'monthly' | 'year'>('day');
  const [selectedTokenDay, setSelectedTokenDay] = useState<string>(() => {
    const d = new Date();
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  });
  const [isLoadingData, setIsLoadingData] = useState(false);
  
  const [selectedTaskId, setSelectedTaskId] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'overview' | 'dashboard' | 'detail' | 'profile' | 'updates' | 'ai'>('dashboard');
  const [dashboardFilter, setDashboardFilter] = useState<'ALL' | 'ACTIVE' | 'OVERDUE'>('OVERDUE');
  const [tagPeriodFilter, setTagPeriodFilter] = useState<'week' | 'month' | 'year'>('week');

  // Updates Feed Search and Filter States
  const [updatesQuery, setUpdatesQuery] = useState('');
  const [updatesProjectFilter, setUpdatesProjectFilter] = useState('ALL');
  const [updatesStatusFilter, setUpdatesStatusFilter] = useState('ALL');

  // AI Chat Bot State
  const [aiInputText, setAiInputText] = useState('');
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [aiClearConfirm, setAiClearConfirm] = useState(false);
  const [aiMessages, setAiMessages] = useState<any[]>(() => {
    const cached = localStorage.getItem('taskflow_ai_messages');
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }
    return [
      {
        id: 'welcome',
        role: 'assistant',
        content: "Hello! I'm your TaskFlow assistant. I can help you organize your schedule, summarize your pending items, or suggest improvements to your workflow. How can I assist you today?",
        timestamp: Date.now()
      }
    ];
  });

  useEffect(() => {
    localStorage.setItem('taskflow_ai_messages', JSON.stringify(aiMessages));
  }, [aiMessages]);

  useEffect(() => {
    if (aiClearConfirm) {
      const timer = setTimeout(() => {
        setAiClearConfirm(false);
      }, 4000);
      return () => clearTimeout(timer);
    }
  }, [aiClearConfirm]);

  const executeAiAction = async (action: any) => {
    if (!action) return;
    try {
      const targetIdStr = action.taskId ? String(action.taskId) : '';
      switch (action.type) {
        case 'CREATE_TASK': {
          const newTask: Task = {
            id: Date.now().toString(),
            userId: currentUser?.id || 'temp',
            title: action.title || 'New AI Task',
            category: action.category || 'General',
            status: TaskStatus.NOT_YET,
            targetDate: action.targetDate || action.newDate || getLocalDate(),
            createdAt: Date.now(),
            tags: []
          };
          setTasks([newTask, ...tasks]);
          setSelectedTaskId(newTask.id);
          setActiveTab('detail');
          try {
            await saveSingleTask(newTask);
            setAiMessages(prev => [...prev, {
              id: Date.now().toString(),
              role: 'assistant',
              content: `Successfully created new task: "${newTask.title}"! 🚀`,
              timestamp: Date.now()
            }]);
          } catch (e) {}
          break;
        }
        case 'COMPLETE_TASK': {
          const taskToUpdate = tasks.find(t => 
            String(t.id) === targetIdStr || 
            (action.title && t.title.toLowerCase() === String(action.title).toLowerCase())
          );
          if (taskToUpdate) {
            const updatedTask = { ...taskToUpdate, status: TaskStatus.DONE };
            setTasks(prev => prev.map(t => t.id === taskToUpdate.id ? updatedTask : t));
            try {
              await saveSingleTask(updatedTask);
              setAiMessages(prev => [...prev, {
                id: Date.now().toString(),
                role: 'assistant',
                content: `Marked "${updatedTask.title}" as completed! Done with this workflow.`,
                timestamp: Date.now()
              }]);
            } catch (e) {}
            // Select it so the user can see it
            setSelectedTaskId(taskToUpdate.id);
            setActiveTab('detail');
          }
          break;
        }
        case 'RESCHEDULE_TASK': {
          const targetDateVal = action.newDate || action.targetDate;
          if (targetDateVal) {
            const taskToUpdate = tasks.find(t => 
              String(t.id) === targetIdStr || 
              (action.title && t.title.toLowerCase() === String(action.title).toLowerCase())
            );
            if (taskToUpdate) {
              const updatedTask = { ...taskToUpdate, targetDate: targetDateVal };
              setTasks(prev => prev.map(t => t.id === taskToUpdate.id ? updatedTask : t));
              try {
                await saveSingleTask(updatedTask);
                setAiMessages(prev => [...prev, {
                  id: Date.now().toString(),
                  role: 'assistant',
                  content: `Rescheduled "${updatedTask.title}" to ${targetDateVal}.`,
                  timestamp: Date.now()
                }]);
              } catch (e) {}
              // Select it so the user can see it
              setSelectedTaskId(taskToUpdate.id);
              setActiveTab('detail');
            }
          }
          break;
        }
        case 'BULK_RESCHEDULE_TASKS': {
          const targetDateVal = action.newDate || action.targetDate;
          const taskIds = action.taskIds || [];
          if (targetDateVal && Array.isArray(taskIds) && taskIds.length > 0) {
            const tasksToUpdate = tasks.filter(t => taskIds.includes(String(t.id)));
            if (tasksToUpdate.length > 0) {
              const updatedTasks = tasksToUpdate.map(t => ({ ...t, targetDate: targetDateVal }));
              setTasks(prev => prev.map(t => {
                const found = updatedTasks.find(ut => ut.id === t.id);
                return found ? found : t;
              }));
              try {
                await saveTasks(updatedTasks);
                setAiMessages(prev => [...prev, {
                  id: Date.now().toString(),
                  role: 'assistant',
                  content: `Successfully rescheduled ${updatedTasks.length} task(s) to ${targetDateVal}! 🎉`,
                  options: [
                    {
                      label: "View Dashboard",
                      action: { type: "NAVIGATE", tab: "dashboard" }
                    }
                  ],
                  timestamp: Date.now()
                }]);
              } catch (e: any) {
                console.error("Bulk reschedule save failed:", e);
                setAiMessages(prev => [...prev, {
                  id: Date.now().toString(),
                  role: 'assistant',
                  content: `⚠️ Failed to update tasks in the database: ${e.message || e}`,
                  timestamp: Date.now()
                }]);
              }
            } else {
              setAiMessages(prev => [...prev, {
                id: Date.now().toString(),
                role: 'assistant',
                content: `Could not find any matching active tasks to reschedule.`,
                timestamp: Date.now()
              }]);
            }
          } else {
            setAiMessages(prev => [...prev, {
              id: Date.now().toString(),
              role: 'assistant',
              content: `Missing or invalid task IDs / target date for bulk reschedule.`,
              timestamp: Date.now()
            }]);
          }
          break;
        }
        case 'NAVIGATE': {
          if (action.tab) {
            setActiveTab(action.tab);
          }
          break;
        }
        default:
          console.warn("Unknown AI action type:", action.type);
      }
    } catch (err) {
      console.error("Failed to execute AI action:", err);
    }
  };

  const sendAiMessage = async (text: string) => {
    if (!text.trim() || isAiLoading) return;
    
    const userMessage = {
      id: Date.now().toString(),
      role: 'user',
      content: text,
      timestamp: Date.now()
    };
    
    setAiMessages(prev => [...prev, userMessage]);
    setAiInputText('');
    setIsAiLoading(true);
    
    try {
      const activeTasksContext = tasks.map(t => ({
        id: t.id,
        title: t.title,
        category: t.category,
        status: t.status,
        targetDate: t.targetDate
      }));

      const response = await fetch("/api/ai/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: text,
          history: aiMessages.slice(-10).map(m => ({ role: m.role, content: m.content })),
          tasks: activeTasksContext,
          currentUser: currentUser
        }),
      });

      if (!response.ok) {
        throw new Error("Failed to contact the Task AI service.");
      }

      const data = await response.json();
      
      const botMessage = {
        id: (Date.now() + 1).toString(),
        role: 'assistant',
        content: data.text || "I was unable to formulate a response.",
        options: data.options || null,
        usage: data.usage || null,
        timestamp: Date.now()
      };
      
      setAiMessages(prev => [...prev, botMessage]);

      // Track AI Token record in database
      const currentUserId = currentUser?.id;
      if (currentUserId) {
        const usage = data.usage || { promptTokens: 110, candidatesTokens: 55 };
        const record: AiTokenRecord = {
          id: Date.now().toString(),
          userId: currentUserId,
          date: new Date().toISOString().split('T')[0],
          timestamp: Date.now(),
          usedFor: text.length > 50 ? text.slice(0, 47) + "..." : text,
          inputTokens: usage.promptTokens,
          outputTokens: usage.candidatesTokens
        };
        saveAiTokenRecord(record).catch(err => console.error("Error saving token record:", err));
      }
    } catch (error: any) {
      console.error("AI chat error:", error);
      const errorMessage = {
        id: (Date.now() + 1).toString(),
        role: 'assistant',
        content: `⚠️ Sorry, I encountered an error: ${error.message || "Failed to connect to the AI service. Verify your internet connection or check your API key settings."}`,
        timestamp: Date.now()
      };
      setAiMessages(prev => [...prev, errorMessage]);
    } finally {
      setIsAiLoading(false);
    }
  };

  const renderAiMessageContent = (content: string) => {
    const lines = content.split('\n');
    return (
      <div className="space-y-1.5 font-medium text-slate-800">
        {lines.map((line, idx) => {
          const matchCheck = line.match(/^-\s*\[\s*\]\s*(.*)$/i);
          if (matchCheck) {
            const taskName = matchCheck[1].replace(/["']/g, '');
            const matchingTask = tasks.find(t => t.title.toLowerCase().trim() === taskName.toLowerCase().trim());
            return (
              <div 
                key={idx} 
                onClick={async () => {
                  if (matchingTask) {
                    const updatedTask = { ...matchingTask, status: TaskStatus.DONE };
                    setTasks(prev => prev.map(t => t.id === matchingTask.id ? updatedTask : t));
                    try {
                      await saveSingleTask(updatedTask);
                    } catch (e) {}
                    setAiMessages(prev => [...prev, {
                      id: Date.now().toString(),
                      role: 'assistant',
                      content: `Marked "${matchingTask.title}" as complete! 🎉`,
                      timestamp: Date.now()
                    }]);
                  }
                }}
                className="flex items-center gap-2.5 py-1.5 px-3 my-1 rounded-xl bg-white border border-stone-200 hover:border-[#0038FF]/40 hover:bg-[#EFF2FC]/40 cursor-pointer shadow-sm transition-all text-stone-700 animate-fade-in"
              >
                <div className="w-[18px] h-[18px] rounded-full border-2 border-slate-300 flex items-center justify-center shrink-0">
                  <Check size={10} className="text-[#0038FF] stroke-[4]" />
                </div>
                <span className="text-xs font-extrabold truncate">{taskName}</span>
              </div>
            );
          }
          
          const matchCheckX = line.match(/^-\s*\[\s*x\s*\]\s*(.*)$/i);
          if (matchCheckX) {
            const taskName = matchCheckX[1].replace(/["']/g, '');
            return (
              <div key={idx} className="flex items-center gap-2.5 py-1.5 px-3 my-1 rounded-xl bg-stone-50 border border-stone-150 opacity-60 text-stone-400 line-through text-xs font-semibold select-none">
                <div className="w-[18px] h-[18px] rounded-full bg-[#EFF2FC] border border-[#0038FF]/10 flex items-center justify-center shrink-0">
                  <Check size={10} className="text-[#0038FF] stroke-[4]" />
                </div>
                <span className="truncate">{taskName}</span>
              </div>
            );
          }

          if (line.trim().startsWith('- ') || line.trim().startsWith('* ')) {
            return (
              <div key={idx} className="flex items-start gap-1.5 pl-1.5">
                <span className="text-[#0038FF] font-bold mt-0.5 shrink-0">•</span>
                <p className="text-xs md:text-sm text-slate-700 leading-relaxed font-bold">{line.replace(/^[-*]\s*/, '')}</p>
              </div>
            );
          }

          if (line.trim() === '') return <div key={idx} className="h-1" />;
          
          return <p key={idx} className="text-xs md:text-sm leading-relaxed text-slate-700 font-bold">{line}</p>;
        })}
      </div>
    );
  };

  // Filters & Sorting
  const [filterCategory, setFilterCategory] = useState<string>('ALL');
  const [filterStatus, setFilterStatus] = useState<string>('ACTIVE'); // Default to ACTIVE TASKS
  
  const [filterTime, setFilterTime] = useState<FilterTime>('ALL');
  const [filterUser, setFilterUser] = useState<string>('ALL_TEAM'); 
  const [taskSort, setTaskSort] = useState<TaskSortOption>('DATE_ASC');
  const [searchQuery, setSearchQuery] = useState('');
  const [overviewSubTab, setOverviewSubTab] = useState<'search' | 'calendar'>('calendar');
  const [calendarSelectedDate, setCalendarSelectedDate] = useState<string>(getLocalDate());
  const [calendarPivot, setCalendarPivot] = useState<Date>(new Date());

  // Team Management State
  const [showTeamModal, setShowTeamModal] = useState(false);
  const [selectedTeamMemberIds, setSelectedTeamMemberIds] = useState<string[]>([]);
  const [allUsers, setAllUsers] = useState<User[]>([]); 
  
  const [showAdminConfirmModal, setShowAdminConfirmModal] = useState(false);
  const [showTaskDeleteConfirm, setShowTaskDeleteConfirm] = useState(false);

  const [updateDate, setUpdateDate] = useState<string>(getLocalDate());
  const [updateContent, setUpdateContent] = useState('');
  const [historySort, setHistorySort] = useState<SortOrder>('NEWEST');
  const [editUpdateId, setEditUpdateId] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  
  const [newPin, setNewPin] = useState('');
  const [confirmPin, setConfirmPin] = useState('');
  const [profileMessage, setProfileMessage] = useState<{type: 'success'|'error', text: string} | null>(null);
  
  const [lastStatus, setLastStatus] = useState<TaskStatus>(TaskStatus.NOT_YET);

  const [isAnonDisabled, setIsAnonDisabled] = useState(false);
  const [firebaseNetworkError, setFirebaseNetworkError] = useState(false);

  // Tags Master & Input State
  const [tagMaster, setTagMaster] = useState<Tag[]>([]);
  const [newTagInput, setNewTagInput] = useState('');

  // Language & Theme State
  const [currentLanguage, setCurrentLanguage] = useState<'EN' | 'ID'>(() => {
    return (localStorage.getItem('dailypulse_lang') as 'EN' | 'ID') || 'EN';
  });
  const [currentTheme, setCurrentTheme] = useState<'light' | 'dark'>(() => {
    return (localStorage.getItem('dailypulse_theme') as 'light' | 'dark') || 'light';
  });
  const [isProfileDropdownOpen, setIsProfileDropdownOpen] = useState(false);

  // Google Sheets Export State
  const [isExporting, setIsExporting] = useState(false);
  const [exportUrl, setExportUrl] = useState<string | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);
  const [showExportModal, setShowExportModal] = useState(false);

  // Automatic Google Sheets Export Scheduler State
  const [autoExportState, setAutoExportState] = useState<{
    lastCheckedSlot: string | null;
    lastExportTime: string | null;
    status: 'SUCCESS' | 'FAILED' | 'PENDING_AUTH' | 'IDLE' | 'RUNNING';
    error: string | null;
  }>(() => {
    return {
      lastCheckedSlot: localStorage.getItem('last_auto_export_checked_slot'),
      lastExportTime: localStorage.getItem('last_auto_export_time'),
      status: (localStorage.getItem('last_auto_export_status') as any) || 'IDLE',
      error: localStorage.getItem('last_auto_export_error')
    };
  });

  const handleExportToSheets = () => {
    setShowExportModal(true);
  };

  const executeExportToSheets = async () => {
    setShowExportModal(false);
    setIsExporting(true);
    setExportError(null);
    setExportUrl(null);

    try {
      let token = getCachedAccessToken();
      
      if (!token) {
        const { user: firebaseUser, error: loginError } = await loginWithGoogle();
        if (loginError) {
          throw new Error(loginError);
        }
        token = getCachedAccessToken();
      }

      if (!token) {
        throw new Error(
          currentLanguage === 'ID'
            ? 'Autentikasi Google berhasil tetapi gagal mendapatkan token akses spreadsheet. Silakan coba lagi.'
            : 'Google authentication succeeded but failed to acquire spreadsheets access token. Please try again.'
        );
      }

      const timestamp = new Date().toISOString().slice(0, 10);
      const title = `Grow Daily Workspace Data - ${timestamp}`;
      
      const sheetInfo = await createExportSpreadsheet(token, title);
      await populateSpreadsheetData(token, sheetInfo.spreadsheetId, {
        tasks,
        updates,
        users: allUsers,
        currentUser: currentUser!
      });

      setExportUrl(sheetInfo.spreadsheetUrl);
    } catch (err: any) {
      console.error('Exporting error:', err);
      let errMsg = err?.message || 'Failed to export workspace data to Google Sheets.';
      if (typeof errMsg === 'string' && (
        errMsg.includes('popup-blocked') || 
        errMsg.includes('popup-closed-by-user') || 
        errMsg.includes('cancelled-popup-request') ||
        errMsg.includes('auth/popup-blocked') ||
        errMsg.includes('auth/popup-closed-by-user')
      )) {
        errMsg = currentLanguage === 'ID'
          ? 'Popup autentikasi Google diblokir atau ditutup karena aplikasi berjalan di dalam frame (iframe). Silakan buka aplikasi di tab baru (gunakan tombol "Open in new tab" di kanan atas) atau aktifkan izin popup di peramban Anda, lalu coba lagi.'
          : 'Google authentication popup was blocked or closed because the app is running inside an iframe. Please open the app in a new browser tab/window (click the "Open in new tab" button on the top right), enable popup permissions in your browser, and try again.';
      }
      setExportError(errMsg);
    } finally {
      setIsExporting(false);
    }
  };

  // Google Sheets Import States
  const [importUrlOrId, setImportUrlOrId] = useState('');
  const [isImporting, setIsImporting] = useState(false);
  const [importError, setImportError] = useState<string | null>(null);
  const [importSuccess, setImportSuccess] = useState<string | null>(null);
  const [showImportPreview, setShowImportPreview] = useState(false);
  const [previewImportData, setPreviewImportData] = useState<{
    tasks: Task[];
    updates: TaskUpdate[];
    users: User[];
  } | null>(null);
  const tasksFileInputRef = useRef<HTMLInputElement>(null);
  const updatesFileInputRef = useRef<HTMLInputElement>(null);

  const handleImportClick = async () => {
    setImportError(null);
    setImportSuccess(null);
    if (!importUrlOrId.trim()) {
      setImportError(
        currentLanguage === 'ID'
          ? 'Silakan masukkan URL atau ID Google Spreadsheet.'
          : 'Please enter a Google Spreadsheet URL or ID.'
      );
      return;
    }

    let spreadsheetId = importUrlOrId.trim();
    const sheetUrlRegex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
    const match = spreadsheetId.match(sheetUrlRegex);
    if (match && match[1]) {
      spreadsheetId = match[1];
    }

    setIsImporting(true);
    try {
      let token = getCachedAccessToken();
      
      if (!token) {
        const { user: firebaseUser, error: loginError } = await loginWithGoogle();
        if (loginError) {
          throw new Error(loginError);
        }
        token = getCachedAccessToken();
      }

      if (!token) {
        throw new Error(
          currentLanguage === 'ID'
            ? 'Autentikasi Google berhasil tetapi gagal mendapatkan token akses spreadsheet.'
            : 'Google authentication completed but failed to retrieve access token.'
        );
      }

      const imported = await importSpreadsheetData(token, spreadsheetId, currentUser?.id);
      
      if (imported.tasks.length === 0 && imported.updates.length === 0 && imported.users.length === 0) {
        throw new Error(
          currentLanguage === 'ID'
            ? 'Tidak ditemukan data valid pada tab "Tasks Checklist", "Updates Timeline", atau "Team Directory". Pastikan format data sesuai.'
            : 'No valid data found on worksheets "Tasks Checklist", "Updates Timeline", or "Team Directory". Please check format.'
        );
      }

      setPreviewImportData(imported);
      setShowImportPreview(true);
    } catch (err: any) {
      console.error('Import error:', err);
      let errMsg = err?.message || 'Failed to import Google Sheets content.';
      if (typeof errMsg === 'string' && (
        errMsg.includes('popup-blocked') || 
        errMsg.includes('popup-closed-by-user') || 
        errMsg.includes('cancelled-popup-request') ||
        errMsg.includes('auth/popup-blocked') ||
        errMsg.includes('auth/popup-closed-by-user')
      )) {
        errMsg = currentLanguage === 'ID'
          ? 'Popup autentikasi Google diblokir atau ditutup. Silakan buka aplikasi di tab baru (tombol "Open in new tab" di kanan atas) dan coba lagi.'
          : 'Google authentication popup was blocked or closed. Please open the app in a new browser tab/window (click the "Open in new tab" button on the top right) and try again.';
      }
      setImportError(errMsg);
    } finally {
      setIsImporting(false);
    }
  };

  const handleConfirmImport = async () => {
    if (!previewImportData) return;
    setIsImporting(true);
    setImportError(null);
    try {
      await importRestoreDatabase(previewImportData);
      
      const taskCount = previewImportData.tasks.length;
      const updateCount = previewImportData.updates.length;
      const userCount = previewImportData.users.length;
      
      let msg = '';
      if (currentLanguage === 'ID') {
        msg = 'Impor berhasil!';
        if (taskCount > 0) msg += ` Memuat ${taskCount} Tugas.`;
        if (updateCount > 0) msg += ` Memuat ${updateCount} Pembaruan.`;
        if (userCount > 0) msg += ` Memuat ${userCount} Profil Tim.`;
      } else {
        msg = 'Successfully imported!';
        if (taskCount > 0) msg += ` Loaded ${taskCount} Tasks.`;
        if (updateCount > 0) msg += ` Loaded ${updateCount} Updates.`;
        if (userCount > 0) msg += ` Loaded ${userCount} Team Users.`;
      }

      setImportSuccess(msg);
      setImportUrlOrId('');
      setShowImportPreview(false);
      setPreviewImportData(null);
    } catch (err: any) {
      console.error('Restore db error:', err);
      setImportError(err?.message || 'Failed to apply imported data to Supabase.');
      setShowImportPreview(false);
      setPreviewImportData(null);
    } finally {
      setIsImporting(false);
    }
  };

  const handleTasksFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setImportError(null);
    setImportSuccess(null);
    const file = e.target.files?.[0];
    if (!file) return;

    const tempUsersToCreate: User[] = [];
    const tempUserMap = new Map<string, string>();

    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    
    if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });

          const parsedTasks: Task[] = [];
          const sheetName = findSheetName(workbook, ['task', 'checklist', 'tugas'], 0);
          const tasksSheet = workbook.Sheets[sheetName];
          
          if (tasksSheet) {
            const rows: any[][] = XLSX.utils.sheet_to_json(tasksSheet, { header: 1 });
            if (rows.length > 1) {
              const header = rows[0].map(h => String(h).trim().toLowerCase());
              const idIdx = header.indexOf('task id');
              const titleIdx = header.indexOf('task title');
              const categoryIdx = header.indexOf('category');
              const statusIdx = header.indexOf('current status');
              const targetDateIdx = header.indexOf('target due date');
              const createdAtIdx = header.indexOf('created at time');
              const tagsIdx = header.indexOf('associated tags');
              const ownerIdx = header.indexOf('assignee / owner');

              if (titleIdx !== -1) {
                for (let i = 1; i < rows.length; i++) {
                  const row = rows[i];
                  if (!row || row.length === 0 || !row[titleIdx]) continue;

                  const title = String(row[titleIdx]).trim();
                  const id = idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : `task_${Date.now()}_${i}`;
                  const category = categoryIdx !== -1 && row[categoryIdx] ? String(row[categoryIdx]).trim() : 'General';
                  
                  let status = TaskStatus.NOT_YET;
                  if (statusIdx !== -1 && row[statusIdx]) {
                    const statusInput = String(row[statusIdx]).trim().toUpperCase();
                    if (Object.values(TaskStatus).includes(statusInput as TaskStatus)) {
                      status = statusInput as TaskStatus;
                    }
                  }

                  const targetDate = targetDateIdx !== -1 && row[targetDateIdx] !== undefined && row[targetDateIdx] !== null
                    ? parseExcelDate(row[targetDateIdx])
                    : new Date().toISOString().slice(0, 10);
                  
                  const createdAt = createdAtIdx !== -1 && row[createdAtIdx] !== undefined && row[createdAtIdx] !== null
                    ? parseExcelTimestamp(row[createdAtIdx])
                    : Date.now();

                  const tags = tagsIdx !== -1 && row[tagsIdx] 
                    ? String(row[tagsIdx]).split(',').map(s => s.trim()).filter(Boolean) 
                    : [];

                  let userId = (currentUser && isValidUuid(currentUser.id)) ? currentUser.id : '';
                  if (!userId) {
                    throw new Error(
                      currentLanguage === 'ID'
                        ? 'Sesi tidak valid. Silakan keluar dan masuk kembali.'
                        : 'Invalid session. Please sign out and sign in again.'
                    );
                  }

                  if (ownerIdx !== -1 && row[ownerIdx]) {
                    const ownerName = String(row[ownerIdx]).trim();
                    const testOwner = ownerName.toLowerCase();
                    const matchedUser = allUsers.find(u => u.username.toLowerCase() === testOwner);
                    if (matchedUser) {
                      userId = matchedUser.id;
                    } else {
                      let generatedId = tempUserMap.get(testOwner);
                      if (!generatedId) {
                        generatedId = generateUUID();
                        tempUserMap.set(testOwner, generatedId);
                        tempUsersToCreate.push({
                          id: generatedId,
                          username: ownerName,
                          pin: '1234',
                          role: 'MEMBER',
                          teamMemberIds: []
                        });
                      }
                      userId = generatedId;
                    }
                  }

                  parsedTasks.push({
                    id,
                    userId,
                    title,
                    category,
                    status,
                    targetDate,
                    createdAt,
                    tags,
                  });
                }
              }
            }
          }

          if (parsedTasks.length === 0) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'Tidak ada data tugas valid yang ditemukan.'
                : 'No valid task records found in the Excel file.'
            );
          }

          setPreviewImportData({
            tasks: parsedTasks,
            updates: [],
            users: tempUsersToCreate,
          });
          setShowImportPreview(true);
        } catch (err: any) {
          setImportError(err.message || 'Failed to parse Excel file.');
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      // Existing CSV parser fallback
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const text = event.target?.result as string;
          if (!text) return;

          const rows = parseCSV(text);
          if (rows.length < 2) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'File CSV kosong atau tidak memiliki baris data.'
                : 'CSV file is empty or does not contain data.'
            );
          }

          const header = rows[0].map(h => h.trim().toLowerCase());
          const idIdx = header.indexOf('task id');
          const titleIdx = header.indexOf('task title');
          const categoryIdx = header.indexOf('category');
          const statusIdx = header.indexOf('current status');
          const targetDateIdx = header.indexOf('target due date');
          const createdAtIdx = header.indexOf('created at time');
          const tagsIdx = header.indexOf('associated tags');
          const ownerIdx = header.indexOf('assignee / owner');

          if (titleIdx === -1) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'Format CSV tidak valid. Harus terdapat kolom "Task Title".'
                : 'Invalid CSV format. Missing required column "Task Title".'
            );
          }

          const parsedTasks: Task[] = [];
          
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0 || !row[titleIdx]) continue;

            const title = row[titleIdx]?.trim() || '';
            if (!title) continue;

            const id = idIdx !== -1 && row[idIdx]?.trim() ? row[idIdx].trim() : `task_${Date.now()}_${i}`;
            const category = categoryIdx !== -1 && row[categoryIdx]?.trim() ? row[categoryIdx].trim() : 'General';
            
            let status = TaskStatus.NOT_YET;
            if (statusIdx !== -1 && row[statusIdx]?.trim()) {
              const statusInput = row[statusIdx].trim().toUpperCase();
              if (Object.values(TaskStatus).includes(statusInput as TaskStatus)) {
                status = statusInput as TaskStatus;
              }
            }

            const targetDate = targetDateIdx !== -1 && row[targetDateIdx] !== undefined && row[targetDateIdx] !== null
              ? parseExcelDate(row[targetDateIdx])
              : new Date().toISOString().slice(0, 10);
            
            const createdAt = createdAtIdx !== -1 && row[createdAtIdx] !== undefined && row[createdAtIdx] !== null
              ? parseExcelTimestamp(row[createdAtIdx])
              : Date.now();

            const tags = tagsIdx !== -1 && row[tagsIdx]?.trim() 
              ? row[tagsIdx].split(',').map(s => s.trim()).filter(Boolean) 
              : [];

            let userId = (currentUser && isValidUuid(currentUser.id)) ? currentUser.id : '';
            if (!userId) {
              throw new Error(
                currentLanguage === 'ID'
                  ? 'Sesi tidak valid. Silakan keluar dan masuk kembali.'
                  : 'Invalid session. Please sign out and sign in again.'
              );
            }

            if (ownerIdx !== -1 && row[ownerIdx]?.trim()) {
              const ownerName = row[ownerIdx].trim();
              const testOwner = ownerName.toLowerCase();
              const matchedUser = allUsers.find(u => u.username.toLowerCase() === testOwner);
              if (matchedUser) {
                userId = matchedUser.id;
              } else {
                let generatedId = tempUserMap.get(testOwner);
                if (!generatedId) {
                  generatedId = generateUUID();
                  tempUserMap.set(testOwner, generatedId);
                  tempUsersToCreate.push({
                    id: generatedId,
                    username: ownerName,
                    pin: '1234',
                    role: 'MEMBER',
                    teamMemberIds: []
                  });
                }
                userId = generatedId;
              }
            }

            parsedTasks.push({
              id,
              userId,
              title,
              category,
              status,
              targetDate,
              createdAt,
              tags,
            });
          }

          if (parsedTasks.length === 0) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'Tidak ada baris tugas valid yang berhasil dipetakan.'
                : 'No valid task records mapped from the selected file.'
            );
          }

          setPreviewImportData({
            tasks: parsedTasks,
            updates: [],
            users: tempUsersToCreate,
          });
          setShowImportPreview(true);
        } catch (err: any) {
          setImportError(err.message || 'Failed to parse uploaded backup file.');
        }
      };
      reader.readAsText(file);
    }
  };

  const handleUpdatesFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setImportError(null);
    setImportSuccess(null);
    const file = e.target.files?.[0];
    if (!file) return;

    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    
    if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });

          const parsedUpdates: TaskUpdate[] = [];
          const sheetName = findSheetName(workbook, ['update', 'timeline', 'log', 'pembaruan', 'history'], 1);
          const updatesSheet = workbook.Sheets[sheetName];
          
          if (updatesSheet) {
            const rows: any[][] = XLSX.utils.sheet_to_json(updatesSheet, { header: 1 });
            if (rows.length > 1) {
              const header = rows[0].map(h => String(h).trim().toLowerCase());
              const idIdx = header.indexOf('update id');
              const taskIdIdx = header.indexOf('related task id');
              const dateIdx = header.indexOf('date logged');
              const contentIdx = header.indexOf('content / note');
              const transitionIdx = header.indexOf('status transition');
              const timestampIdx = header.indexOf('logged timestamp');

              if (taskIdIdx !== -1 && contentIdx !== -1) {
                for (let i = 1; i < rows.length; i++) {
                  const row = rows[i];
                  if (!row || row.length === 0 || !row[taskIdIdx] || !row[contentIdx]) continue;

                  const id = idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : `update_${Date.now()}_${i}`;
                  const taskId = String(row[taskIdIdx]).trim();
                  const date = dateIdx !== -1 && row[dateIdx] !== undefined && row[dateIdx] !== null
                    ? parseExcelDate(row[dateIdx])
                    : new Date().toISOString().slice(0, 10);
                  const content = String(row[contentIdx]).trim();

                  let statusChange;
                  if (transitionIdx !== -1 && row[transitionIdx]) {
                    const transitionInput = String(row[transitionIdx]).trim();
                    if (transitionInput && transitionInput !== 'None') {
                      const parts = transitionInput.split(/➔|->/).map(s => s.trim().toUpperCase());
                      if (parts.length === 2) {
                        const from = parts[0] as TaskStatus;
                        const to = parts[1] as TaskStatus;
                        if (Object.values(TaskStatus).includes(from) && Object.values(TaskStatus).includes(to)) {
                          statusChange = { from, to };
                        }
                      }
                    }
                  }

                  const timestamp = timestampIdx !== -1 && row[timestampIdx] !== undefined && row[timestampIdx] !== null
                    ? parseExcelTimestamp(row[timestampIdx])
                    : Date.now();

                  const updateDoc: TaskUpdate = {
                    id,
                    taskId,
                    date,
                    content,
                    timestamp,
                    isArchived: false,
                  };

                  if (statusChange) {
                    updateDoc.statusChange = statusChange;
                  }

                  parsedUpdates.push(updateDoc);
                }
              }
            }
          }

          if (parsedUpdates.length === 0) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'Tidak ada data pembaruan valid yang ditemukan.'
                : 'No valid update records found in the Excel file.'
            );
          }

          setPreviewImportData({
            tasks: [],
            updates: parsedUpdates,
            users: [],
          });
          setShowImportPreview(true);
        } catch (err: any) {
          setImportError(err.message || 'Failed to parse Excel file.');
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      // Fallback CSV parser for updates
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const text = event.target?.result as string;
          if (!text) return;

          const rows = parseCSV(text);
          if (rows.length < 2) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'File CSV kosong atau tidak memiliki baris data.'
                : 'CSV file is empty or does not contain data.'
            );
          }

          const header = rows[0].map(h => h.trim().toLowerCase());
          const idIdx = header.indexOf('update id');
          const taskIdIdx = header.indexOf('related task id');
          const dateIdx = header.indexOf('date logged');
          const contentIdx = header.indexOf('content / note');
          const transitionIdx = header.indexOf('status transition');
          const timestampIdx = header.indexOf('logged timestamp');

          if (taskIdIdx === -1 || contentIdx === -1) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'Format CSV tidak valid. Harus terdapat kolom "Related Task ID" dan "Content / Note".'
                : 'Invalid CSV format. Missing required columns "Related Task ID" or "Content / Note".'
            );
          }

          const parsedUpdates: TaskUpdate[] = [];
          
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0 || !row[taskIdIdx] || !row[contentIdx]) continue;

            const id = idIdx !== -1 && row[idIdx]?.trim() ? row[idIdx].trim() : `update_${Date.now()}_${i}`;
            const taskId = row[taskIdIdx].trim();
            const date = dateIdx !== -1 && row[dateIdx] !== undefined && row[dateIdx] !== null
              ? parseExcelDate(row[dateIdx])
              : new Date().toISOString().slice(0, 10);
            const content = row[contentIdx].trim();

            let statusChange;
            if (transitionIdx !== -1 && row[transitionIdx]?.trim()) {
              const transitionInput = row[transitionIdx].trim();
              if (transitionInput && transitionInput !== 'None') {
                const parts = transitionInput.split(/➔|->/).map(s => s.trim().toUpperCase());
                if (parts.length === 2) {
                  const from = parts[0] as TaskStatus;
                  const to = parts[1] as TaskStatus;
                  if (Object.values(TaskStatus).includes(from) && Object.values(TaskStatus).includes(to)) {
                    statusChange = { from, to };
                  }
                }
              }
            }

            const timestamp = timestampIdx !== -1 && row[timestampIdx] !== undefined && row[timestampIdx] !== null
              ? parseExcelTimestamp(row[timestampIdx])
              : Date.now();

            const updateDoc: TaskUpdate = {
              id,
              taskId,
              date,
              content,
              timestamp,
              isArchived: false,
            };

            if (statusChange) {
              updateDoc.statusChange = statusChange;
            }

            parsedUpdates.push(updateDoc);
          }

          if (parsedUpdates.length === 0) {
            throw new Error(
              currentLanguage === 'ID'
                ? 'Tidak ada baris pembaruan valid yang berhasil dipetakan.'
                : 'No valid update records mapped from the selected file.'
            );
          }

          setPreviewImportData({
            tasks: [],
            updates: parsedUpdates,
            users: [],
          });
          setShowImportPreview(true);
        } catch (err: any) {
          setImportError(err.message || 'Failed to parse CSV file.');
        }
      };
      reader.readAsText(file);
    }
  };


  useEffect(() => {
    localStorage.setItem('dailypulse_lang', currentLanguage);
  }, [currentLanguage]);

  useEffect(() => {
    localStorage.setItem('dailypulse_theme', currentTheme);
  }, [currentTheme]);

  useEffect(() => {
    if (currentUser) {
      localStorage.setItem('dailypulse_cached_user', JSON.stringify(currentUser));
    } else {
      localStorage.removeItem('dailypulse_cached_user');
    }
  }, [currentUser]);

  const handleRetryAuth = async () => {
    setAuthLoading(true);
    setFirebaseNetworkError(false);
    setIsAnonDisabled(false);
    setAuthError('');
    try {
      const user = await getCurrentUser();
      if (user) {
        setCurrentUser(user);
      }
    } catch (e) {
      console.error("Retry login failed:", e);
    } finally {
      setAuthLoading(false);
    }
  };

  useEffect(() => {
    const initSession = async () => {
      try {
        const user = await getCurrentUser();
        if (user) {
          setCurrentUser(user);
        }
      } catch (error) {
        console.error("Error initializing Supabase session:", error);
      }
    };
    initSession();

    const { data: { subscription } } = supabase.auth.onAuthStateChange(async (event, session) => {
      if (session?.user) {
        const user = await getCurrentUser();
        if (user) {
          setCurrentUser(user);
        }
      } else {
        setCurrentUser(null);
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  // --- Backfill Legacy updates ---
  const hasBackfilledRef = useRef(false);
  useEffect(() => {
    if (currentUser && tasks.length > 0 && !hasBackfilledRef.current) {
      hasBackfilledRef.current = true;
      backfillTaskUpdates(currentUser.id, tasks);
    }
  }, [currentUser, tasks]);

  useEffect(() => {
    if (!currentUser) return;

    const unsubTasks = subscribeToTasks(setTasks, currentUser);
    const unsubUpdates = subscribeToUpdates(setUpdates, currentUser);
    const unsubUsers = subscribeToUsers(setAllUsers);
    const unsubTags = subscribeToTags(setTagMaster);
    const unsubAiTokens = subscribeToAiTokenRecords(setAiTokenRecords, currentUser.role === 'ADMIN' ? undefined : currentUser.id);

    return () => {
      unsubTasks();
      unsubUpdates();
      unsubUsers();
      unsubTags();
      unsubAiTokens();
    };
  }, [currentUser]);

  // Automatic Google Sheets Export Scheduler Effect
  useEffect(() => {
    if (!currentUser) return;

    const checkScheduler = async () => {
      const now = new Date();
      const mostRecentSlot = getMostRecentScheduledSlot(now);
      const slotStr = mostRecentSlot.toISOString();
      const storedCheckedSlot = localStorage.getItem('last_auto_export_checked_slot');

      if (storedCheckedSlot !== slotStr) {
        // Trigger silent automatic export!
        const token = getCachedAccessToken();
        
        if (!token) {
          console.warn('Scheduled auto-export: No Google cached access token available.');
          localStorage.setItem('last_auto_export_checked_slot', slotStr);
          localStorage.setItem('last_auto_export_status', 'PENDING_AUTH');
          localStorage.setItem('last_auto_export_error', 'Requires Google Sign-In to run automatic background export');
          setAutoExportState({
            lastCheckedSlot: slotStr,
            lastExportTime: localStorage.getItem('last_auto_export_time'),
            status: 'PENDING_AUTH',
            error: 'Google Sign-In is required to run scheduled exports.'
          });
          return;
        }

        // We have a token! Perform silent export
        setAutoExportState({
          lastCheckedSlot: slotStr,
          lastExportTime: localStorage.getItem('last_auto_export_time'),
          status: 'RUNNING',
          error: null
        });

        try {
          const timestamp = now.toISOString().slice(0, 10);
          const title = `Scheduled Auto-Export - ${timestamp}`;
          
          const sheetInfo = await createExportSpreadsheet(token, title);
          await populateSpreadsheetData(token, sheetInfo.spreadsheetId, {
            tasks,
            updates,
            users: allUsers,
            currentUser: currentUser!
          });

          const nowStr = new Date().toISOString();
          localStorage.setItem('last_auto_export_checked_slot', slotStr);
          localStorage.setItem('last_auto_export_time', nowStr);
          localStorage.setItem('last_auto_export_status', 'SUCCESS');
          localStorage.setItem('last_auto_export_error', '');

          setAutoExportState({
            lastCheckedSlot: slotStr,
            lastExportTime: nowStr,
            status: 'SUCCESS',
            error: null
          });
        } catch (err: any) {
          console.error('Scheduled auto-export failed:', err);
          const errMsg = err?.message || 'Failed to populate sheet';
          localStorage.setItem('last_auto_export_checked_slot', slotStr);
          localStorage.setItem('last_auto_export_status', 'FAILED');
          localStorage.setItem('last_auto_export_error', errMsg);

          setAutoExportState({
            lastCheckedSlot: slotStr,
            lastExportTime: localStorage.getItem('last_auto_export_time'),
            status: 'FAILED',
            error: errMsg
          });
        }
      }
    };

    // Run check on mount or when dependencies change
    checkScheduler();

    // Set stable interval check every 30 seconds
    const intervalId = setInterval(checkScheduler, 30000);
    return () => clearInterval(intervalId);
  }, [tasks, updates, allUsers, currentUser]);

  useEffect(() => {
    if (selectedTaskId) {
      setUpdateContent('');
      setUpdateDate(getLocalDate());
      setEditUpdateId(null);
      
      const t = tasks.find(x => x.id === selectedTaskId);
      if (t) {
        setLastStatus(t.status);
      }
    }
  }, [selectedTaskId]);

  useEffect(() => {
    if (activeTab === 'detail') {
      const timer = setTimeout(() => {
        window.scrollTo({ top: 0, behavior: 'instant' as ScrollBehavior });
        
        // Scroll the container labeled as detail-tab-container
        const tabContainer = document.getElementById('detail-tab-container');
        if (tabContainer) {
          tabContainer.scrollTo({ top: 0, behavior: 'instant' as ScrollBehavior });
        }
        
        // Also scroll any other overflow-y-auto elements to provide consistent scroll focus
        const overflows = document.querySelectorAll('.overflow-y-auto');
        overflows.forEach(el => {
          el.scrollTo({ top: 0, behavior: 'instant' as ScrollBehavior });
        });
      }, 50);
      return () => clearTimeout(timer);
    }
  }, [selectedTaskId, activeTab]);

  const loadData = async (userOverride?: User) => {
    // Data is now handled by real-time subscriptions
    setIsLoadingData(true);
    const userToUse = userOverride || currentUser;
    try {
      const [fetchedTasks, fetchedUpdates, fetchedUsers] = await Promise.all([
          getTasks(userToUse || undefined),
          getUpdates(),
          getUsers()
      ]);
      setTasks(fetchedTasks);
      setUpdates(fetchedUpdates);
      setAllUsers(fetchedUsers);
    } catch (error) {
      console.error("Error loading data:", error);
    } finally {
      setIsLoadingData(false);
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthLoading(true);
    setAuthError('');
    try {
      const user = await loginUser(authEmail, authPassword);
      if (user) {
        setCurrentUser(user);
        await loadData(user);
        setAuthEmail('');
        setAuthPassword('');
      } else {
        setAuthError(currentLanguage === 'ID' ? 'Gagal masuk: Akun tidak ditemukan.' : 'Login failed: Account not found.');
      }
    } catch (error: any) {
      console.error("Login error:", error);
      setAuthError(error.message || (currentLanguage === 'ID' ? 'Gagal masuk.' : 'Login failed.'));
    } finally {
      setAuthLoading(false);
    }
  };

  const handleRegister = async (e: React.FormEvent) => {
    e.preventDefault();
    if (authPassword.length < 6) {
      setAuthError(currentLanguage === 'ID' ? 'Kata sandi harus minimal 6 karakter' : 'Password must be at least 6 characters');
      return;
    }
    setAuthLoading(true);
    setAuthError('');
    try {
      const user = await registerUser(authUsername, authEmail, authPassword, authRole);
      if (user) {
        setCurrentUser(user);
        await loadData(user);
        setAuthUsername('');
        setAuthEmail('');
        setAuthPassword('');
      } else {
        setAuthError(currentLanguage === 'ID' ? 'Gagal mendaftar.' : 'Registration failed.');
      }
    } catch (error: any) {
      console.error("Registration error:", error);
      setAuthError(error.message || (currentLanguage === 'ID' ? 'Gagal mendaftar.' : 'Registration failed.'));
    } finally {
      setAuthLoading(false);
    }
  };

  const handleGoogleLogin = async () => {
    setAuthLoading(true);
    setAuthError('');
    try {
      const { data, error } = await supabase.auth.signInWithOAuth({
        provider: 'google',
        options: {
          redirectTo: window.location.origin
        }
      });
      
      if (error) {
        setAuthError(`Google login failed: ${error.message}`);
        return;
      }
    } catch (error: any) {
      console.error("Google login error:", error);
      setAuthError(`An error occurred: ${error.message || error}`);
    } finally {
      setAuthLoading(false);
    }
  };

  const handleLogout = () => {
    logoutUser();
    setCurrentUser(null);
    setTasks([]);
    setUpdates([]);
    setSelectedTaskId(null);
    setActiveTab('overview');
    setFilterUser('ALL_TEAM');
  };

  const handleUpdateProfile = async () => {
    if (!currentUser) return;
    setProfileMessage(null);
    if (newPin.length < 6) {
      setProfileMessage({ type: 'error', text: currentLanguage === 'ID' ? 'Kata sandi harus minimal 6 karakter.' : 'Password must be at least 6 characters.' });
      return;
    }
    if (newPin !== confirmPin) {
      setProfileMessage({ type: 'error', text: currentLanguage === 'ID' ? 'Kata sandi tidak cocok.' : 'Passwords do not match.' });
      return;
    }
    const updatedUser = { ...currentUser, pin: newPin };
    const success = await updateUser(updatedUser);
    if (success) {
      setCurrentUser(updatedUser);
      setProfileMessage({ type: 'success', text: currentLanguage === 'ID' ? 'Kata sandi berhasil diperbarui!' : 'Password updated successfully!' });
      setNewPin('');
      setConfirmPin('');
    } else {
      setProfileMessage({ type: 'error', text: currentLanguage === 'ID' ? 'Gagal memperbarui kata sandi.' : 'Failed to update password.' });
    }
  };

  const handleToggleTeamMember = (userId: string) => {
    setSelectedTeamMemberIds(prev => 
      prev.includes(userId) ? prev.filter(id => id !== userId) : [...prev, userId]
    );
  };

  const handleAddSelectedMembers = async () => {
    if (!currentUser || currentUser.role !== 'ADMIN') return;
    if (selectedTeamMemberIds.length === 0) return;
    const success = await addTeamMembers(currentUser.id, selectedTeamMemberIds);
    if (success) {
      const updatedUser = await getCurrentUser();
      if (updatedUser) setCurrentUser(updatedUser);
      const u = await getUsers();
      setAllUsers(u);
      setSelectedTeamMemberIds([]);
      setShowTeamModal(false);
    }
  };

  const tokenDayOptions = useMemo(() => {
    const options = [];
    const now = new Date();
    for (let i = 0; i < 6; i++) {
      const d = new Date();
      d.setDate(now.getDate() - i);
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const dateStr = `${year}-${month}-${day}`;
      
      const englishMonths = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
      const indonesianMonths = ["januari", "februari", "maret", "april", "mei", "juni", "juli", "agustus", "september", "oktober", "november", "desember"];
      
      const isIndo = currentLanguage === 'ID';
      const mLabel = isIndo ? indonesianMonths[d.getMonth()] : englishMonths[d.getMonth()];
      
      options.push({
        value: dateStr,
        label: `${mLabel} ${d.getDate()}`
      });
    }
    return options;
  }, [currentLanguage]);

  const filteredTokenRecords = useMemo(() => {
    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    
    return aiTokenRecords.filter(record => {
      if (tokenFilterType === 'day') {
        return record.date === selectedTokenDay;
      } else if (tokenFilterType === 'weekly') {
        const sevenDaysAgo = todayStart - (6 * 24 * 60 * 60 * 1000);
        return record.timestamp >= sevenDaysAgo;
      } else if (tokenFilterType === 'monthly') {
        const recordDate = new Date(record.timestamp);
        return recordDate.getMonth() === now.getMonth() && recordDate.getFullYear() === now.getFullYear();
      } else if (tokenFilterType === 'year') {
        const recordDate = new Date(record.timestamp);
        return recordDate.getFullYear() === now.getFullYear();
      }
      return true;
    });
  }, [aiTokenRecords, tokenFilterType, selectedTokenDay]);

  const tokenStats = useMemo(() => {
    let inputs = 0;
    let outputs = 0;
    filteredTokenRecords.forEach(r => {
      inputs += r.inputTokens || 0;
      outputs += r.outputTokens || 0;
    });
    return {
      inputs,
      outputs,
      total: inputs + outputs
    };
  }, [filteredTokenRecords]);

  const allCategories = useMemo(() => {
    const taskCategories = tasks.map(t => t.category);
    return Array.from(new Set([...CATEGORIES, ...taskCategories])).sort();
  }, [tasks]);

  const selectedTask = useMemo(() => 
    tasks.find(t => t.id === selectedTaskId), 
  [tasks, selectedTaskId]);

  const taskHistory = useMemo(() => {
    if (!selectedTaskId) return [];
    return updates
      .filter(u => u.taskId === selectedTaskId && !u.isArchived)
      .sort((a, b) => {
        return historySort === 'NEWEST' ? b.timestamp - a.timestamp : a.timestamp - b.timestamp;
      });
  }, [updates, selectedTaskId, historySort]);

  const groupedUpdatesByDate = useMemo(() => {
    if (!currentUser) return {};
    const groups: Record<string, TaskUpdate[]> = {};
    const relevantUpdates = updates.filter(u => {
        const task = tasks.find(t => t.id === u.taskId);
        if (!task) return false;
        
        let isOwned = false;
        if (currentUser.role === 'ADMIN') {
            isOwned = task.userId === currentUser.id || currentUser.teamMemberIds.includes(task.userId);
        } else {
            isOwned = task.userId === currentUser.id;
        }
        if (!isOwned) return false;

        // Search text matching: update content or task title
        if (updatesQuery.trim()) {
          const q = updatesQuery.toLowerCase();
          const matchesContent = u.content.toLowerCase().includes(q);
          const matchesTitle = task.title.toLowerCase().includes(q);
          if (!matchesContent && !matchesTitle) return false;
        }

        // Project/Category filter
        if (updatesProjectFilter !== 'ALL') {
          if (task.category !== updatesProjectFilter) return false;
        }

        // Status filter
        if (updatesStatusFilter !== 'ALL') {
          const status = u.statusChange?.to || task.status;
          if (status !== updatesStatusFilter) return false;
        }

        return true;
    });

    relevantUpdates.forEach(u => {
      if (!groups[u.date]) groups[u.date] = [];
      groups[u.date].push(u);
    });

    // Sort individual updates by timestamp descending
    Object.keys(groups).forEach(date => {
      groups[date].sort((a, b) => b.timestamp - a.timestamp);
    });

    return groups;
  }, [updates, tasks, currentUser, updatesQuery, updatesProjectFilter, updatesStatusFilter]);

  const sortedUpdateDates = useMemo(() => {
    return Object.keys(groupedUpdatesByDate).sort((a, b) => b.localeCompare(a));
  }, [groupedUpdatesByDate]);

  const userMap = useMemo(() => {
    return allUsers.reduce((acc, user) => {
      acc[user.id] = user;
      return acc;
    }, {} as Record<string, User>);
  }, [allUsers]);

  const dashboardStats = useMemo(() => {
    if (!currentUser) return null;
    
    const userTasks = tasks.filter(task => {
      if (currentUser.role === 'ADMIN') {
        if (filterUser === 'ME') return task.userId === currentUser.id;
        if (filterUser === 'ALL_TEAM') return true; // Admin can see all tasks in the system under 'All Team'
        return task.userId === filterUser;
      }
      return task.userId === currentUser.id;
    });

    const total = userTasks.length;
    const completed = userTasks.filter(t => t.status === TaskStatus.DONE).length;
    const inProgress = userTasks.filter(t => t.status === TaskStatus.PROGRESS).length;
    const pending = userTasks.filter(t => [TaskStatus.NOT_YET, TaskStatus.PENDING, TaskStatus.FOLLOW_UP, TaskStatus.REPETITIVE].includes(t.status)).length;
    const completionRate = total > 0 ? Math.round((completed / total) * 100) : 0;

    const now = new Date();
    now.setHours(0,0,0,0);
    const overdue = userTasks.filter(task => {
      if (task.status === TaskStatus.DONE || task.status === TaskStatus.CANCEL) return false;
      try {
        const [ty, tm, td] = task.targetDate.split('-').map(Number);
        const target = new Date(ty, tm - 1, td);
        return target.getTime() < now.getTime();
      } catch {
        return new Date(task.targetDate).getTime() < now.getTime();
      }
    }).length;

    const updateLogsCount = updates.filter(u => userTasks.map(t => t.id).includes(u.taskId)).length;
    const gamificationPoints = (completed * 100) + (inProgress * 35) + (updateLogsCount * 15);
    
    const ptsPerLevel = 250;
    const currentLevel = Math.max(1, Math.floor(gamificationPoints / ptsPerLevel) + 1);
    const xpInCurrentLevel = gamificationPoints % ptsPerLevel;
    const levelProgressPercent = Math.round((xpInCurrentLevel / ptsPerLevel) * 100);

    let levelTitle = 'Seedling';
    if (currentLevel === 2) levelTitle = 'Sprout';
    else if (currentLevel === 3) levelTitle = 'Budding';
    else if (currentLevel === 4) levelTitle = 'Blossom';
    else if (currentLevel === 5) levelTitle = 'Citrus Scholar';
    else if (currentLevel === 6) levelTitle = 'Fruitful Pro';
    else if (currentLevel >= 7) levelTitle = 'Abundant Harvest Grandmaster';

    const isInPeriod = (dateStr: string, period: 'week' | 'month' | 'year') => {
      try {
        const parts = dateStr.split('-').map(Number);
        if (parts.length !== 3) return false;
        const [y, m, d] = parts;
        const target = new Date(y, m - 1, d);
        const today = new Date();
        
        if (period === 'year') {
          return target.getFullYear() === today.getFullYear();
        }
        if (period === 'month') {
          return target.getFullYear() === today.getFullYear() && target.getMonth() === today.getMonth();
        }
        if (period === 'week') {
          const currentDay = today.getDay();
          const distanceToMonday = currentDay === 0 ? -6 : 1 - currentDay;
          const startOfWeek = new Date(today);
          startOfWeek.setDate(today.getDate() + distanceToMonday);
          startOfWeek.setHours(0,0,0,0);
          
          const endOfWeek = new Date(startOfWeek);
          endOfWeek.setDate(startOfWeek.getDate() + 6);
          endOfWeek.setHours(23,59,59,999);
          
          return target.getTime() >= startOfWeek.getTime() && target.getTime() <= endOfWeek.getTime();
        }
      } catch {
        return false;
      }
      return false;
    };

    const categoryBreakdown = CATEGORIES.reduce((acc, cat) => {
      const catTaskIds = userTasks.filter(t => t.category === cat).map(t => t.id);
      const count = updates.filter(u => catTaskIds.includes(u.taskId) && isInPeriod(u.date, tagPeriodFilter)).length;
      acc[cat] = { count, percent: 0 };
      return acc;
    }, {} as Record<string, { count: number, percent: number }>);

    const maxCategoryCount = Math.max(...Object.values(categoryBreakdown).map(x => x.count), 1);
    CATEGORIES.forEach(cat => {
      if (categoryBreakdown[cat]) {
        categoryBreakdown[cat].percent = Math.round((categoryBreakdown[cat].count / maxCategoryCount) * 100);
      }
    });

    const last7Days = Array.from({ length: 7 }, (_, i) => {
      const d = new Date();
      d.setDate(d.getDate() - i);
      return getLocalDate(d);
    });

    const activeDatesWithUpdates = new Set(
      updates
        .filter(u => userTasks.map(t => t.id).includes(u.taskId))
        .map(u => u.date)
    );

    let streakCount = 0;
    for (const dateStr of last7Days) {
      if (activeDatesWithUpdates.has(dateStr)) {
        streakCount++;
      } else {
        if (dateStr === getLocalDate(new Date())) {
          continue;
        }
        break;
      }
    }

    const achievements = [
      { id: 'first_task', name: 'First Milestone', desc: 'Complete any task to score points', achieved: completed > 0, icon: '🎯' },
      { id: 'tri_force', name: 'Tri-Force', desc: 'Complete tasks in 3 distinct categories', achieved: new Set(userTasks.filter(t => t.status === TaskStatus.DONE).map(t => t.category)).size >= 3, icon: '🌟' },
      { id: 'streak_3', name: 'Habit Former', desc: 'Maintain updates across at least 3 distinct days', achieved: streakCount >= 3, icon: '🔥' },
      { id: 'high_flyer', name: 'Citrus Master', desc: 'Reach Level 3 to blossom fully', achieved: currentLevel >= 3, icon: '👑' },
    ];

    return {
      total,
      completed,
      inProgress,
      pending,
      completionRate,
      overdue,
      gamificationPoints,
      currentLevel,
      levelTitle,
      xpInCurrentLevel,
      levelProgressPercent,
      ptsPerLevel,
      categoryBreakdown,
      streakCount,
      achievements,
      updateLogsCount
    };
  }, [tasks, updates, currentUser, filterUser, tagPeriodFilter]);

  const dailyUpdateCounts = useMemo(() => {
    const labelsMap: Record<'EN' | 'ID', string[]> = {
      EN: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
      ID: ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab']
    };
    
    // Always from today and trace back 7 days
    const result = Array.from({ length: 7 }, (_, i) => {
      const d = new Date();
      d.setDate(d.getDate() - (6 - i));
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const dateStr = `${year}-${month}-${day}`;
      const weekdayIdx = d.getDay();
      
      const langKey = currentLanguage === 'ID' ? 'ID' : 'EN';
      const label = labelsMap[langKey][weekdayIdx];
      
      return {
        dateStr,
        label: label.substring(0, 1).toUpperCase(),
        count: 0
      };
    });

    updates.forEach(u => {
      const found = result.find(item => item.dateStr === u.date);
      if (found) {
        found.count++;
      }
    });

    return result;
  }, [updates, currentLanguage]);

  const dashboardFilteredTasks = useMemo(() => {
    if (!currentUser) return [];
    let list = tasks.filter(task => {
      if (currentUser.role === 'ADMIN') {
        if (filterUser === 'ME') return task.userId === currentUser.id;
        if (filterUser === 'ALL_TEAM') return true; // Admin can see all tasks in the system under 'All Team'
        return task.userId === filterUser;
      }
      return task.userId === currentUser.id;
    });

    if (dashboardFilter === 'ACTIVE') {
      list = list.filter(t => t.status === TaskStatus.PROGRESS);
    } else if (dashboardFilter === 'OVERDUE') {
      const now = new Date();
      now.setHours(0,0,0,0);
      list = list.filter(task => {
        if (task.status === TaskStatus.DONE || task.status === TaskStatus.CANCEL) return false;
        try {
          const [ty, tm, td] = task.targetDate.split('-').map(Number);
          const target = new Date(ty, tm - 1, td);
          return target.getTime() <= now.getTime();
        } catch {
          return new Date(task.targetDate).getTime() <= now.getTime();
        }
      });
    }

    return list.sort((a, b) => {
      const remainingA = getDaysRemaining(a.targetDate);
      const remainingB = getDaysRemaining(b.targetDate);
      const aOverdue = remainingA < 0 && a.status !== TaskStatus.DONE;
      const bOverdue = remainingB < 0 && b.status !== TaskStatus.DONE;
      if (aOverdue && !bOverdue) return -1;
      if (!aOverdue && bOverdue) return 1;
      return new Date(a.targetDate).getTime() - new Date(b.targetDate).getTime();
    });
  }, [tasks, currentUser, filterUser, dashboardFilter]);

  const teamMembers = useMemo(() => {
    if (!currentUser || currentUser.role !== 'ADMIN') return [];
    return currentUser.teamMemberIds.map(id => userMap[id]).filter(Boolean) as User[];
  }, [currentUser, userMap]); 

  const availableUsersToAdd = useMemo(() => {
    if (!currentUser || currentUser.role !== 'ADMIN') return [];
    return allUsers.filter((u: User) => u.id !== currentUser.id && !currentUser.teamMemberIds.includes(u.id));
  }, [currentUser, allUsers]);

  const filteredTasks = useMemo(() => {
    if (!currentUser) return [];
    const now = new Date();
    now.setHours(0,0,0,0);
    const filtered = tasks.filter(task => {
      let isVisible = false;
      if (currentUser.role === 'ADMIN') {
        if (filterUser === 'ME') isVisible = task.userId === currentUser.id;
        else if (filterUser === 'ALL_TEAM') isVisible = true; // Admin can see all tasks in the system under 'All Team'
        else isVisible = task.userId === filterUser;
      } else isVisible = task.userId === currentUser.id;
      if (!isVisible) return false;
      const matchesSearch = task.title.toLowerCase().includes(searchQuery.toLowerCase()) || 
        (task.tags && task.tags.some(tag => tag.toLowerCase().includes(searchQuery.toLowerCase().replace('#', ''))));
      const matchesCategory = filterCategory === 'ALL' || task.category === filterCategory;
      const matchesStatus = filterStatus === 'ALL' ? true : filterStatus === 'ACTIVE' ? [TaskStatus.NOT_YET, TaskStatus.PROGRESS, TaskStatus.FOLLOW_UP, TaskStatus.REPETITIVE].includes(task.status) : task.status === filterStatus;
      let matchesTime = true;
      let targetTime = 0;
      try {
        const [ty, tm, td] = task.targetDate.split('-').map(Number);
        const target = new Date(ty, tm - 1, td);
        targetTime = target.getTime();
      } catch(e) { targetTime = new Date(task.targetDate).getTime(); }
      const diffDays = (targetTime - now.getTime()) / (1000 * 3600 * 24);
      if (filterTime === 'OVERDUE') matchesTime = diffDays < -0.9;
      else if (filterTime === 'TODAY') matchesTime = diffDays >= -0.9 && diffDays <= 0.9;
      else if (filterTime === 'THIS_WEEK') matchesTime = diffDays >= -0.9 && diffDays <= 7;
      else if (filterTime === 'FUTURE') matchesTime = diffDays > 7;
      return matchesSearch && matchesCategory && matchesStatus && matchesTime;
    });
    return filtered.sort((a, b) => {
        const getT = (dStr: string) => {
            try { const [y, m, d] = dStr.split('-').map(Number); return new Date(y, m-1, d).getTime(); } catch { return 0; }
        };
        switch (taskSort) {
            case 'DATE_ASC': return getT(a.targetDate) - getT(b.targetDate);
            case 'DATE_DESC': return getT(b.targetDate) - getT(a.targetDate);
            case 'STATUS_ASC': return STATUSES.indexOf(a.status) - STATUSES.indexOf(b.status);
            case 'STATUS_DESC': return STATUSES.indexOf(b.status) - STATUSES.indexOf(a.status);
            case 'CATEGORY_ASC': return a.category.localeCompare(b.category);
            case 'CATEGORY_DESC': return b.category.localeCompare(a.category);
            default: return 0;
        }
    });
  }, [tasks, searchQuery, filterCategory, filterStatus, filterTime, filterUser, currentUser, taskSort]);

  const tasksToShow = useMemo(() => {
    if (overviewSubTab === 'calendar') {
      return filteredTasks.filter(t => t.targetDate === calendarSelectedDate);
    }
    return filteredTasks;
  }, [filteredTasks, overviewSubTab, calendarSelectedDate]);

  const formatListDate = (dateStr: string) => {
    if (!dateStr) return '';
    try {
      const [year, month, day] = dateStr.split('-').map(Number);
      const date = new Date(year, month - 1, day);
      if (isNaN(date.getTime())) return dateStr;
      const ddd = date.toLocaleDateString('en-US', { weekday: 'short' });
      const dd = day.toString().padStart(2, '0');
      const mmm = date.toLocaleDateString('en-US', { month: 'short' });
      const yy = year.toString().slice(-2);
      return `${ddd}, ${dd}-${mmm}-${yy}`;
    } catch (e) { return dateStr; }
  };

  const getCalendarDays = () => {
    const year = calendarPivot.getFullYear();
    const month = calendarPivot.getMonth();
    
    // First day of current month (0 = Sunday, 1 = Monday ...)
    const firstDayIndex = new Date(year, month, 1).getDay();
    // Total days in current month
    const totalDays = new Date(year, month + 1, 0).getDate();
    
    // Backfill from previous month
    const prevMonthTotalDays = new Date(year, month, 0).getDate();
    const days: { dateStr: string; dayNum: number; isCurrentMonth: boolean }[] = [];
    
    for (let i = firstDayIndex - 1; i >= 0; i--) {
      const prevMonthYear = month === 0 ? year - 1 : year;
      const prevMonth = month === 0 ? 11 : month - 1;
      const dayNum = prevMonthTotalDays - i;
      days.push({
        dateStr: `${prevMonthYear}-${String(prevMonth + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`,
        dayNum,
        isCurrentMonth: false,
      });
    }
    
    // Current month's days
    for (let i = 1; i <= totalDays; i++) {
      days.push({
        dateStr: `${year}-${String(month + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`,
        dayNum: i,
        isCurrentMonth: true,
      });
    }
    
    // Pad the end with next month's days to make complete grid rows (multiple of 7)
    const remaining = 42 - days.length;
    for (let i = 1; i <= remaining; i++) {
      const nextMonthYear = month === 11 ? year + 1 : year;
      const nextMonth = month === 11 ? 0 : month + 1;
      days.push({
        dateStr: `${nextMonthYear}-${String(nextMonth + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`,
        dayNum: i,
        isCurrentMonth: false,
      });
    }
    
    return days;
  };

  const handleTaskSelect = (id: string) => { 
    setSelectedTaskId(id); 
    const task = tasks.find(t => t.id === id);
    if (task) {
      setLastStatus(task.status);
    }
    setActiveTab('detail'); 
  };
  const handleDeleteTask = () => {
    if (!selectedTask) return;
    setShowTaskDeleteConfirm(true);
  };
  const handleConfirmDeleteTask = async () => {
    if (!selectedTask) return;
    const targetTaskId = selectedTask.id;
    setShowTaskDeleteConfirm(false);
    
    // Perform local state updates immediately for instantaneous response
    setTasks(prev => prev.filter(t => t.id !== targetTaskId));
    setUpdates(prev => prev.filter(u => u.taskId !== targetTaskId));
    setSelectedTaskId(null);
    setActiveTab('overview');
    
    try {
      await deleteTask(targetTaskId);
    } catch (err: any) {
      console.error("Failed to delete task in Firestore:", err);
    }
  };
  const handleSaveUpdate = async () => {
    if (!selectedTask || !updateContent.trim()) return;
    if (currentUser?.role === 'ADMIN' && selectedTask.userId !== currentUser.id) { setShowAdminConfirmModal(true); return; }
    await executeSaveUpdate(false);
  };
  const handleConfirmAdminUpdate = async () => { setShowAdminConfirmModal(false); await executeSaveUpdate(true); };
  const executeSaveUpdate = async (isAdminOverride: boolean) => {
    if (!selectedTask || !updateContent.trim()) return;
    let finalContent = updateContent;
    if (isAdminOverride && currentUser) finalContent = `${finalContent}\n(Updated by ${currentUser.username})`;
    
    setIsSaving(true);
    try {
      if (!currentUser) throw new Error("No user logged in");

      let statusChangeData = undefined;
      if (lastStatus !== selectedTask.status) {
        statusChangeData = { from: lastStatus, to: selectedTask.status };
        const updatedTask = { ...selectedTask };
        setTasks(prev => prev.map(t => t.id === selectedTask.id ? updatedTask : t));
        await saveSingleTask(updatedTask);
      }
      
      if (editUpdateId) {
        const updatedUpdate = updates.find(u => u.id === editUpdateId);
        if (updatedUpdate) {
          const newUpdateObj: TaskUpdate = { 
            ...updatedUpdate, 
            userId: updatedUpdate.userId || currentUser.id,
            content: finalContent, 
            date: updateDate
          };
          
          if (statusChangeData) {
            newUpdateObj.statusChange = statusChangeData;
          } else if (!updatedUpdate.statusChange) {
            // Ensure we don't carry over an undefined field
            delete (newUpdateObj as any).statusChange;
          }
          
          setUpdates(prev => prev.map(u => u.id === editUpdateId ? newUpdateObj : u));
          await saveSingleUpdate(newUpdateObj);
          setEditUpdateId(null);
        }
      } else {
      const newUpdate: TaskUpdate = { 
        id: Date.now().toString(), 
        taskId: selectedTask.id, 
        userId: currentUser.id,
        date: updateDate, 
        content: finalContent, 
        timestamp: Date.now(), 
        isArchived: false
      };
      
      if (statusChangeData) {
        newUpdate.statusChange = statusChangeData;
      }
      
      setUpdates(prev => [newUpdate, ...prev]);
      await saveSingleUpdate(newUpdate);
      }
      
      if (statusChangeData) {
        setLastStatus(selectedTask.status);
      }
      
      setUpdateContent(''); 
      setUpdateDate(getLocalDate());
      
      // Verification log
      console.log("Update saved successfully. Content cleared.");
    } catch (error: any) {
      console.error("Error saving update:", error);
      let errorMsg = "Failed to save update.";
      try {
        const parsed = JSON.parse(error.message);
        if (parsed.error) errorMsg = `Error: ${parsed.error}`;
      } catch {
        errorMsg = `Error: ${error.message || 'Unknown error'}`;
      }
      alert(errorMsg);
    } finally {
      setIsSaving(false);
    }
  };
  const handleCancelEdit = () => { setEditUpdateId(null); setUpdateContent(''); setUpdateDate(getLocalDate()); };
  const handleTaskChange = async (field: keyof Task, value: any) => {
    if (!selectedTask) return;
    const updatedTask: Task = {
      id: selectedTask.id,
      userId: selectedTask.userId || currentUser?.id || '',
      title: selectedTask.title || 'Untitled Task',
      category: selectedTask.category || CATEGORIES[0],
      status: selectedTask.status,
      targetDate: selectedTask.targetDate || getLocalDate(),
      createdAt: selectedTask.createdAt || Date.now(),
      tags: selectedTask.tags || [],
      [field]: value
    };
    setTasks(prev => prev.map(t => t.id === selectedTask.id ? updatedTask : t));
    try {
      await saveSingleTask(updatedTask);
    } catch (err: any) {
      console.error("Failed to save task update:", err);
      alert("Failed to save task changes: " + (err.message || err));
    }
  };
  const handleToggleTaskStatus = async (task: Task, e: React.MouseEvent) => {
    e.stopPropagation();
    const newStatus = task.status === TaskStatus.DONE ? TaskStatus.PROGRESS : TaskStatus.DONE;
    const updatedTask: Task = {
      id: task.id,
      userId: task.userId || currentUser?.id || '',
      title: task.title || 'Untitled Task',
      category: task.category || CATEGORIES[0],
      status: newStatus,
      targetDate: task.targetDate || getLocalDate(),
      createdAt: task.createdAt || Date.now(),
      tags: task.tags || []
    };
    setTasks(prev => prev.map(t => t.id === task.id ? updatedTask : t));
    try {
      await saveSingleTask(updatedTask);
    } catch (err: any) {
      console.error("Failed to toggle task status:", err);
      alert("Failed to update status on server: " + (err.message || err));
    }
  };
  const handleAddToCalendar = (update: TaskUpdate) => {
    if (!selectedTask) return;
    const title = encodeURIComponent(`${selectedTask.title} - Update`);
    const details = encodeURIComponent(`Task: ${selectedTask.title}\nUpdate: ${update.content}`);
    const startDateStr = update.date.replace(/-/g, '');
    const dateObj = new Date(update.date);
    dateObj.setDate(dateObj.getDate() + 1);
    const endDateStr = dateObj.toISOString().split('T')[0].replace(/-/g, '');
    const dates = `${startDateStr}/${endDateStr}`;
    const url = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${title}&details=${details}&dates=${dates}`;
    window.open(url, '_blank');
  };
  const handleDeleteUpdate = async (update: TaskUpdate) => {
    setUpdates(prev => prev.filter(u => u.id !== update.id));
    await deleteUpdate(update.id);
    if (editUpdateId === update.id) handleCancelEdit();
  };
  const handleEditUpdateStart = (update: TaskUpdate) => { setUpdateContent(update.content); setUpdateDate(update.date); setEditUpdateId(update.id); };
  const createNewTask = async () => {
    if (!currentUser) return;
    const newTask: Task = { id: Date.now().toString(), userId: currentUser.id, title: 'New Task', category: CATEGORIES[0], status: TaskStatus.NOT_YET, targetDate: getLocalDate(), createdAt: Date.now(), tags: [] };
    setTasks([newTask, ...tasks]); setSelectedTaskId(newTask.id); setActiveTab('detail'); setLastStatus(TaskStatus.NOT_YET);
    await saveSingleTask(newTask);
  };

  if (!currentUser) {
    return (
      <div className="flex h-[100dvh] w-full bg-beige items-center justify-center p-6">
        <div className="w-full max-w-md bg-white rounded-2xl shadow-xl overflow-hidden border border-stone-200">
           <div className="bg-accent p-8 text-center">
              <svg className="w-16 h-16 mx-auto mb-4 drop-shadow-md" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect width="100" height="100" rx="24" fill="#1c1917" />
                <path d="M50 72V48" stroke="#a8a29e" strokeWidth="6" strokeLinecap="round" />
                <path d="M50 58C45 54 40 54 38 50" stroke="#a8a29e" strokeWidth="4" strokeLinecap="round" />
                <path d="M50 54C55 50 60 50 62 46" stroke="#a8a29e" strokeWidth="4" strokeLinecap="round" />
                <circle cx="50" cy="36" r="16" fill="#86efac" fillOpacity="0.9" />
                <circle cx="41" cy="43" r="12" fill="#4ade80" fillOpacity="0.9" />
                <circle cx="59" cy="43" r="12" fill="#22c55e" fillOpacity="0.85" />
                <circle cx="44" cy="32" r="3" fill="#ef4444" />
                <circle cx="56" cy="36" r="3" fill="#f97316" />
                <circle cx="37" cy="42" r="3" fill="#ef4444" />
                <circle cx="63" cy="43" r="3" fill="#f97316" />
                <circle cx="49" cy="44" r="3" fill="#ef4444" />
              </svg>
              <h1 className="text-2xl font-bold text-secondary">DailyPulse</h1>
              <p className="text-secondary/80 font-medium uppercase tracking-tight">Daily Update Status</p>
           </div>
           <div className="p-8">
              <div className="flex bg-stone-100 p-1 rounded-xl mb-6">
                 <button onClick={() => { setAuthMode('LOGIN'); setAuthError(''); }} className={`flex-1 py-2 text-sm font-bold rounded-lg transition-colors ${authMode === 'LOGIN' ? 'bg-white text-secondary shadow-sm' : 'text-stone-400'}`}>Login</button>
                 <button onClick={() => { setAuthMode('REGISTER'); setAuthError(''); }} className={`flex-1 py-2 text-sm font-bold rounded-lg transition-colors ${authMode === 'REGISTER' ? 'bg-white text-secondary shadow-sm' : 'text-stone-400'}`}>Register</button>
              </div>
              <form onSubmit={authMode === 'LOGIN' ? handleLogin : handleRegister} className="flex flex-col gap-4">
                 {authMode === 'REGISTER' ? (
                   <>
                     <div>
                        <label className="text-xs font-bold text-stone-500 uppercase ml-1">
                          Username
                        </label>
                        <div className="relative mt-1">
                          <UserIcon size={18} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                          <input type="text" value={authUsername} onChange={e => setAuthUsername(e.target.value)} className="w-full pl-10 pr-4 py-3 bg-stone-50 border border-stone-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary/50 text-secondary font-medium" placeholder={currentLanguage === 'ID' ? 'Masukkan username' : 'Enter username'} required disabled={authLoading} />
                        </div>
                     </div>
                     <div>
                        <label className="text-xs font-bold text-stone-500 uppercase ml-1">Email</label>
                        <div className="relative mt-1">
                          <Mail size={18} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                          <input type="email" value={authEmail} onChange={e => setAuthEmail(e.target.value)} className="w-full pl-10 pr-4 py-3 bg-stone-50 border border-stone-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary/50 text-secondary font-medium" placeholder={currentLanguage === 'ID' ? 'Masukkan email' : 'Enter email address'} required disabled={authLoading} />
                        </div>
                     </div>
                   </>
                 ) : (
                   <div>
                      <label className="text-xs font-bold text-stone-500 uppercase ml-1">
                        {currentLanguage === 'ID' ? 'Username atau Email' : 'Username or Email'}
                      </label>
                      <div className="relative mt-1">
                        <UserIcon size={18} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                        <input type="text" value={authEmail} onChange={e => setAuthEmail(e.target.value)} className="w-full pl-10 pr-4 py-3 bg-stone-50 border border-stone-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary/50 text-secondary font-medium" placeholder={currentLanguage === 'ID' ? 'Masukkan username atau email' : 'Enter username or email'} required disabled={authLoading} />
                      </div>
                   </div>
                 )}
                 <div>
                    <label className="text-xs font-bold text-stone-500 uppercase ml-1">
                      {currentLanguage === 'ID' ? 'Kata Sandi' : 'Password'}
                    </label>
                    <div className="relative mt-1">
                      <Shield size={18} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                      <input type="password" value={authPassword} onChange={e => setAuthPassword(e.target.value)} className="w-full pl-10 pr-4 py-3 bg-stone-50 border border-stone-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary/50 text-secondary font-medium tracking-wide" placeholder={currentLanguage === 'ID' ? 'Min. 6 karakter' : 'Min. 6 characters'} required disabled={authLoading} />
                    </div>
                 </div>
                 {authMode === 'REGISTER' && (
                   <div>
                     <label className="text-xs font-bold text-stone-500 uppercase ml-1">Role</label>
                     <div className="flex gap-4 mt-2">
                        <label className="flex items-center gap-2 cursor-pointer">
                           <input type="radio" name="role" checked={authRole === 'MEMBER'} onChange={() => setAuthRole('MEMBER')} className="accent-secondary" />
                           <span className="text-sm font-bold text-stone-600">Member</span>
                        </label>
                        <label className="flex items-center gap-2 cursor-pointer">
                           <input type="radio" name="role" checked={authRole === 'ADMIN'} onChange={() => setAuthRole('ADMIN')} className="accent-secondary" />
                           <span className="text-sm font-bold text-stone-600">Admin</span>
                        </label>
                     </div>
                   </div>
                 )}
                 {firebaseNetworkError && (
                    <div className="mb-4 bg-rose-50 border border-rose-200 p-4 rounded-xl">
                      <div className="flex items-center gap-2 text-rose-700 font-bold text-sm mb-1">
                        <AlertTriangle size={16} />
                        Firebase Connection Issue
                      </div>
                      <p className="text-xs text-rose-600 leading-relaxed mb-2">
                        Could not establish a connection with Google auth server. You can still try accessing with your account, or retry the connection below.
                      </p>
                      <button
                        type="button"
                        onClick={handleRetryAuth}
                        className="text-xs font-bold text-rose-700 underline hover:text-rose-900 cursor-pointer flex items-center gap-1"
                        disabled={authLoading}
                      >
                        {authLoading ? 'Connecting...' : '🔄 Retry Connection'}
                      </button>
                    </div>
                  )}
                 {isAnonDisabled && (
                    <div className="mb-4 bg-amber-50 border border-amber-200 p-4 rounded-xl">
                      <div className="flex items-center gap-2 text-amber-700 font-bold text-sm mb-1">
                        <AlertCircle size={16} />
                        Anonymous Auth Disabled
                      </div>
                      <p className="text-xs text-amber-600 leading-relaxed">
                        Username/PIN login requires Anonymous Auth. Please enable it in the 
                        <a href="https://console.firebase.google.com/" target="_blank" rel="noopener noreferrer" className="underline ml-1">
                          Firebase Console
                        </a> (Authentication &gt; Sign-in method).
                      </p>
                    </div>
                  )}
                 {authError && <div className="bg-red-50 text-red-600 text-xs font-bold p-3 rounded-xl flex items-center gap-2"><AlertCircle size={14} /> {authError}</div>}
                 <button type="submit" className="mt-4 bg-secondary text-accent py-3 rounded-xl font-bold hover:bg-stone-800 transition-transform active:scale-95 shadow-lg shadow-stone-200 flex justify-center items-center gap-2 w-full disabled:opacity-50" disabled={authLoading || isAnonDisabled}>{authLoading ? <Loader2 size={18} className="animate-spin"/> : (authMode === 'LOGIN' ? 'Access Dashboard' : 'Create Account')}</button>

                  <div className="relative my-4">
                    <div className="absolute inset-0 flex items-center"><div className="w-full border-t border-stone-200"></div></div>
                    <div className="relative flex justify-center text-xs uppercase"><span className="bg-white px-2 text-stone-400 font-bold">Or continue with</span></div>
                  </div>

                  <button 
                    type="button"
                    onClick={handleGoogleLogin}
                    disabled={authLoading}
                    className="w-full py-3 bg-white border border-stone-200 text-secondary rounded-xl font-bold flex items-center justify-center gap-3 hover:bg-stone-50 transition-all active:scale-95 disabled:opacity-50"
                  >
                    <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google" className="w-5 h-5" referrerPolicy="no-referrer" />
                    Google Account
                  </button>

                  {auth.currentUser && (
                    <div className="mt-6 pt-6 border-t border-stone-100 text-center">
                      <p className="text-[10px] text-stone-400 mb-2">Logged in as: {auth.currentUser.email || 'Anonymous'}</p>
                      
                      {!currentUser && (
                        <button 
                          type="button"
                          onClick={handleGoogleLogin}
                          className="w-full py-2 mb-4 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs font-bold hover:bg-amber-100 transition-all"
                        >
                          Profile not found. Click to Sync Profile
                        </button>
                      )}

                      <button 
                        type="button"
                        onClick={() => auth.signOut()}
                        className="text-[10px] text-red-400 hover:text-red-600 font-bold underline uppercase tracking-wider"
                      >
                        Sign out of this session
                      </button>
                    </div>
                  )}
              </form>
           </div>
           <div className="px-8 pb-8 text-center opacity-30 hover:opacity-100 transition-opacity">
              <p className="text-[9px] text-stone-400 font-mono leading-tight">
                 Database: Supabase<br/>
                 URL: https://czzpkodqlnfjycljixyh.supabase.co
              </p>
           </div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-[100dvh] bg-beige overflow-hidden relative font-sans">
      {isLoadingData && <div className="absolute inset-0 z-[60] bg-white/50 backdrop-blur-sm flex items-center justify-center"><div className="bg-white p-4 rounded-xl shadow-xl flex items-center gap-3"><Loader2 className="animate-spin text-primary" size={24} /><span className="font-bold text-secondary">Loading Data...</span></div></div>}
      {showAdminConfirmModal && (
        <div className="absolute inset-0 z-[80] bg-secondary/60 backdrop-blur-sm flex items-center justify-center p-6">
           <div className="bg-white rounded-2xl w-full max-w-sm p-6 shadow-2xl animate-in zoom-in-95 duration-200">
              <div className="w-12 h-12 bg-amber-100 rounded-full flex items-center justify-center mb-4 mx-auto text-amber-600"><AlertTriangle size={24} /></div>
              <h3 className="text-lg font-bold text-secondary text-center mb-2">Confirm Update</h3>
              <p className="text-stone-500 text-sm text-center mb-6">Are you sure? This task is not yours. Your update will be tagged with your username.</p>
              <div className="flex gap-3">
                 <button onClick={() => setShowAdminConfirmModal(false)} className="flex-1 py-3 text-sm font-bold text-stone-500 bg-stone-100 hover:bg-stone-200 rounded-xl transition-colors">Cancel</button>
                 <button onClick={handleConfirmAdminUpdate} className="flex-1 py-3 text-sm font-bold text-white bg-secondary hover:bg-black rounded-xl transition-colors shadow-lg">Confirm Update</button>
              </div>
           </div>
        </div>
      )}
      {showTaskDeleteConfirm && selectedTask && (
        <div className="absolute inset-0 z-[80] bg-secondary/60 backdrop-blur-sm flex items-center justify-center p-6 animate-fade-in">
           <div className="bg-white rounded-2xl w-full max-w-sm p-6 shadow-2xl animate-in zoom-in-95 duration-200">
              <div className="w-12 h-12 bg-red-100 text-red-600 rounded-full flex items-center justify-center mb-4 mx-auto"><Trash2 size={24} /></div>
              <h3 className="text-lg font-bold text-secondary text-center mb-1">Delete Task</h3>
              <p className="text-stone-500 text-xs font-semibold text-center mb-5">
                Are you sure you want to delete task <span className="text-slate-800 font-bold">"{selectedTask.title}"</span>? 
                {updates.some(u => u.taskId === selectedTask.id) && " This task has historic actions/updates which will be permanently deleted."}
              </p>
              <div className="flex gap-2.5">
                 <button onClick={() => setShowTaskDeleteConfirm(false)} className="flex-1 py-2.5 text-xs font-bold text-stone-500 bg-stone-100 hover:bg-stone-200 rounded-xl transition-colors">Cancel</button>
                 <button onClick={handleConfirmDeleteTask} className="flex-1 py-2.5 text-xs font-bold text-white bg-red-600 hover:bg-red-700 rounded-xl transition-all shadow-md shadow-red-200 dark:shadow-none hover:shadow-lg active:scale-95">Yes, Delete</button>
              </div>
           </div>
        </div>
      )}
      {showTeamModal && (
        <div className="absolute inset-0 z-50 bg-secondary/50 backdrop-blur-sm flex items-center justify-center p-4">
           <div className="bg-white rounded-2xl w-full max-w-sm p-6 shadow-2xl animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
              <div className="flex justify-between items-center mb-4 shrink-0"><h3 className="text-lg font-bold text-secondary">Add Line Members</h3><button onClick={() => { setShowTeamModal(false); setSelectedTeamMemberIds([]); }} className="p-1 hover:bg-stone-100 rounded-full"><X size={20}/></button></div>
              <p className="text-stone-500 text-sm mb-4 shrink-0">Select users to add to your reporting line.</p>
              <div className="flex-1 overflow-y-auto mb-4 border border-stone-200 rounded-xl min-h-[150px]">{availableUsersToAdd.length === 0 ? <div className="h-full flex flex-col items-center justify-center text-stone-400 p-4"><Users size={32} className="mb-2 opacity-30" /><p className="text-sm text-center">No other members available to add.</p></div> : <div className="divide-y divide-stone-100">{availableUsersToAdd.map(user => <label key={user.id} className="flex items-center p-3 hover:bg-stone-50 cursor-pointer group transition-colors"><div className="relative flex items-center justify-center mr-3"><input type="checkbox" checked={selectedTeamMemberIds.includes(user.id)} onChange={() => handleToggleTeamMember(user.id)} className="peer appearance-none w-5 h-5 border-2 border-stone-300 rounded checked:bg-primary checked:border-primary transition-all"/><CheckSquare size={14} className="absolute text-secondary opacity-0 peer-checked:opacity-100 pointer-events-none transform scale-0 peer-checked:scale-100 transition-transform" /></div><div className="flex-1"><div className="font-bold text-stone-700">{user.username}</div><div className="text-[10px] font-bold text-stone-400 uppercase">{user.role}</div></div></label>)}</div>}</div>
              <div className="flex gap-2 shrink-0"><button onClick={() => { setShowTeamModal(false); setSelectedTeamMemberIds([]); }} className="flex-1 py-3 font-bold text-stone-500 hover:bg-stone-50 rounded-xl">Cancel</button><button onClick={handleAddSelectedMembers} disabled={selectedTeamMemberIds.length === 0} className="flex-1 py-3 bg-secondary text-accent font-bold rounded-xl hover:bg-stone-800 disabled:opacity-50 disabled:cursor-not-allowed transition-all active:scale-95">Add Selected ({selectedTeamMemberIds.length})</button></div>
           </div>
        </div>
      )}
      {showExportModal && (
        <div className="absolute inset-0 z-[80] bg-secondary/60 backdrop-blur-sm flex items-center justify-center p-6 animate-fade-in">
           <div className="bg-white rounded-2xl w-full max-w-sm p-6 shadow-2xl animate-in zoom-in-95 duration-200">
              <div className="w-12 h-12 bg-[#EFF2FC] text-[#0038FF] rounded-full flex items-center justify-center mb-4 mx-auto"><Globe size={24} /></div>
              <h3 className="text-lg font-black text-secondary text-center mb-2" style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                {currentLanguage === 'ID' ? 'Ekspor Data Workspace' : 'Export Workspace Data'}
              </h3>
              <p className="text-stone-500 text-xs font-semibold text-center mb-6 leading-relaxed">
                {currentLanguage === 'ID'
                  ? 'Apakah Anda yakin ingin mengekspor seluruh data tugas, timeline pembaruan, dan direktori tim ke Google Spreadsheet baru?'
                  : 'Are you sure you want to export all tasks, updates timeline, and team directory data to a new Google Spreadsheet?'}
              </p>
              <div className="flex gap-2.5">
                 <button 
                   onClick={() => setShowExportModal(false)} 
                   className="flex-1 py-2.5 text-xs font-bold text-stone-500 bg-stone-100 hover:bg-stone-200 rounded-xl transition-colors"
                 >
                   {currentLanguage === 'ID' ? 'Batal' : 'Cancel'}
                 </button>
                 <button 
                   onClick={executeExportToSheets} 
                   className="flex-1 py-2.5 text-xs font-bold text-white bg-[#0038FF] hover:bg-[#002fcf] rounded-xl transition-all shadow-md shadow-blue-100 active:scale-95"
                 >
                   {currentLanguage === 'ID' ? 'Ya, Ekspor' : 'Yes, Export'}
                 </button>
              </div>
           </div>
        </div>
      )}
      {showImportPreview && previewImportData && (
        <div className="absolute inset-0 z-[80] bg-secondary/60 backdrop-blur-sm flex items-center justify-center p-6 animate-fade-in">
           <div className="bg-white rounded-2xl w-full max-w-sm p-6 shadow-2xl animate-in zoom-in-95 duration-200">
              <div className="w-12 h-12 bg-rose-50 text-rose-600 rounded-full flex items-center justify-center mb-4 mx-auto"><AlertTriangle size={24} /></div>
              <h3 className="text-lg font-black text-secondary text-center mb-2" style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                {currentLanguage === 'ID' ? 'Peringatan Impor' : 'Import Warning'}
              </h3>
              
              <div className="bg-rose-50 border border-rose-100 rounded-xl p-3 mb-4 text-center">
                <p className="text-rose-600 font-extrabold text-xs uppercase tracking-wider mb-1">
                  {currentLanguage === 'ID' ? 'Perhatian!' : 'Attention!'}
                </p>
                <p className="text-rose-700 text-xs font-bold leading-relaxed">
                  {currentLanguage === 'ID'
                    ? 'Semua data Anda saat ini akan digantikan dengan file impor ini.'
                    : 'All your data will be replaced with this import file.'}
                </p>
              </div>

              <p className="text-stone-500 text-xs font-semibold text-center mb-4 leading-relaxed">
                {currentLanguage === 'ID'
                  ? 'Berikut rincian data baru yang diekstrak dan siap diimpor:'
                  : 'Here is the summary of the new data found in your file:'}
              </p>
              
              <div className="bg-stone-50 border border-stone-200 rounded-xl p-4 mb-6 flex flex-col gap-2 mx-auto max-w-[280px]">
                <div className="flex justify-between items-center text-xs">
                  <span className="font-bold text-stone-500">{currentLanguage === 'ID' ? 'Total Tugas:' : 'Total Tasks:'}</span>
                  <span className="font-black text-secondary bg-stone-200 px-2.5 py-0.5 rounded-md">{previewImportData.tasks.length}</span>
                </div>
                <div className="flex justify-between items-center text-xs">
                  <span className="font-bold text-stone-500">{currentLanguage === 'ID' ? 'Log Pembaruan:' : 'Total Updates:'}</span>
                  <span className="font-black text-secondary bg-stone-200 px-2.5 py-0.5 rounded-md">{previewImportData.updates.length}</span>
                </div>
                <div className="flex justify-between items-center text-xs">
                  <span className="font-bold text-stone-500">{currentLanguage === 'ID' ? 'Profil Tim:' : 'Team Profiles:'}</span>
                  <span className="font-black text-secondary bg-stone-200 px-2.5 py-0.5 rounded-md">{previewImportData.users.length}</span>
                </div>
              </div>

              <div className="flex gap-2.5">
                 <button 
                   onClick={() => { setShowImportPreview(false); setPreviewImportData(null); }} 
                   className="flex-1 py-2.5 text-xs font-bold text-stone-500 bg-stone-100 hover:bg-stone-200 rounded-xl transition-colors"
                 >
                   {currentLanguage === 'ID' ? 'Tidak / Batal' : 'No / Cancel'}
                 </button>
                 <button 
                   onClick={handleConfirmImport} 
                   className="flex-1 py-2.5 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl transition-all shadow-md shadow-rose-100 active:scale-95"
                 >
                   {currentLanguage === 'ID' ? 'Ya, Impor' : 'Yes, Import'}
                 </button>
              </div>
           </div>
        </div>
      )}
      <div className="flex-1 relative overflow-hidden flex flex-col">
          <div className={`flex flex-col bg-slate-50 border-b border-slate-200 min-h-0 overflow-y-auto ${activeTab === 'overview' ? 'flex flex-1 animate-fade-in' : 'hidden md:flex'} md:relative md:h-[55vh]`}>
            
            {/* BEGIN: Styled Modern Header */}
            <header className={`sticky top-0 z-30 px-5 py-3 shrink-0 transition-colors duration-200 border-b ${
              currentTheme === 'dark' 
                ? 'bg-[#111827] border-[#1F2937]' 
                : 'bg-white border-slate-150'
            }`}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  {currentUser?.role === 'ADMIN' && (
                    <button 
                      onClick={() => setShowTeamModal(true)} 
                      className="mr-1.5 p-1 text-slate-800 hover:bg-slate-100 dark:text-slate-205 dark:hover:bg-slate-800 rounded-full transition-all flex items-center justify-center shrink-0 active:scale-95"
                      title="Management / Team line"
                    >
                      <Menu size={18} className="text-[#0038FF]" />
                    </button>
                  )}
                  <div className="w-2.5 h-2.5 bg-[#0038FF] rounded-full animate-pulse" />
                  <h1 className={`text-sm font-black uppercase tracking-widest transition-colors ${
                    currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                  }`} style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                    {TRANSLATIONS[currentLanguage].tasks}
                  </h1>
                </div>
                
                <div className="flex items-center gap-3">
                  <span className={`text-[10px] font-bold px-2.5 py-1 rounded-lg border hidden sm:block ${
                    currentTheme === 'dark' 
                      ? 'text-slate-400 bg-slate-800/40 border-slate-700/60' 
                      : 'text-slate-400 bg-slate-50 border-slate-150'
                  }`}>
                    {new Date().toLocaleDateString(currentLanguage === 'ID' ? 'id-ID' : 'en-US', { weekday: 'short', month: 'short', day: 'numeric' })}
                  </span>

                  {/* Profile Dropdown Menu */}
                  <div className="relative">
                    <button 
                      onClick={() => setIsProfileDropdownOpen(!isProfileDropdownOpen)}
                      className={`flex items-center gap-1.5 p-1 rounded-full border transition-all ${
                        currentTheme === 'dark' 
                          ? 'bg-slate-800 border-slate-700 hover:bg-slate-750 text-white shadow-md' 
                          : 'bg-slate-50 border-slate-150 hover:bg-slate-100 text-slate-850 shadow-sm'
                      }`}
                    >
                      <div className="w-6 h-6 rounded-full bg-[#0038FF] text-white flex items-center justify-center text-xs font-black uppercase shadow-inner">
                        {currentUser.username[0].toUpperCase()}
                      </div>
                      <span className="text-[11px] font-extrabold pr-1 md:block hidden">{currentUser.username}</span>
                      <ChevronDown size={11} className={`opacity-60 transition-transform ${isProfileDropdownOpen ? 'rotate-180' : ''}`} />
                    </button>

                    {isProfileDropdownOpen && (
                      <>
                        <div className="fixed inset-0 z-40" onClick={() => setIsProfileDropdownOpen(false)} />
                        <div className={`absolute right-0 mt-2 w-52 rounded-2xl border p-3.5 shadow-xl z-50 animate-in fade-in slide-in-from-top-1 duration-150 ${
                          currentTheme === 'dark' 
                            ? 'bg-slate-900 border-slate-850 text-slate-205' 
                            : 'bg-white border-slate-150 text-slate-800'
                        }`} style={{ background: currentTheme === 'dark' ? '#1F2937' : '#FFFFFF', borderColor: currentTheme === 'dark' ? '#374151' : '#E2E8F0' }}>
                          <div className="border-b pb-2 mb-2 border-dashed border-slate-200/55 dark:border-slate-700/55">
                            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">{currentLanguage === 'ID' ? 'MASUK SEBAGAI' : 'SIGNED IN AS'}</p>
                            <p className="text-xs font-black text-[#0038FF] truncate uppercase">{currentUser.username}</p>
                            <span className={`inline-block mt-0.5 text-[8.5px] font-extrabold px-1.5 py-0.5 rounded tracking-wide ${
                              currentTheme === 'dark' ? 'bg-slate-950 text-slate-300' : 'bg-[#EFF2FC] text-[#0038FF]'
                            }`}>
                              {currentUser.role}
                            </span>
                          </div>

                          {/* Language Submenu */}
                          <div className="mb-2.5">
                            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                              <Globe size={10} />
                              {TRANSLATIONS[currentLanguage].language}
                            </p>
                            <div className={`grid grid-cols-2 gap-0.5 p-0.5 rounded-lg border ${
                              currentTheme === 'dark' ? 'bg-slate-950 border-slate-800' : 'bg-slate-100 border-slate-150'
                            }`}>
                              <button
                                onClick={() => {
                                  setCurrentLanguage('EN');
                                  setIsProfileDropdownOpen(false);
                                }}
                                className={`py-1 text-[9px] font-black rounded-md transition-all flex items-center justify-center gap-1 ${
                                  currentLanguage === 'EN'
                                    ? 'bg-[#0038FF] text-white shadow'
                                    : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                }`}
                              >
                                EN
                              </button>
                              <button
                                onClick={() => {
                                  setCurrentLanguage('ID');
                                  setIsProfileDropdownOpen(false);
                                }}
                                className={`py-1 text-[9px] font-black rounded-md transition-all flex items-center justify-center gap-1 ${
                                  currentLanguage === 'ID'
                                    ? 'bg-[#0038FF] text-white shadow'
                                    : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                }`}
                              >
                                ID
                              </button>
                            </div>
                          </div>

                          {/* Theme Submenu */}
                          <div className="mb-2.5">
                            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                              {currentTheme === 'dark' ? <Moon size={10} /> : <Sun size={10} />}
                              {TRANSLATIONS[currentLanguage].theme}
                            </p>
                            <div className={`grid grid-cols-2 gap-0.5 p-0.5 rounded-lg border ${
                              currentTheme === 'dark' ? 'bg-slate-950 border-slate-800' : 'bg-slate-100 border-slate-150'
                            }`}>
                              <button
                                onClick={() => {
                                  setCurrentTheme('light');
                                  setIsProfileDropdownOpen(false);
                                }}
                                className={`py-1 text-[9px] font-black rounded-md transition-all flex items-center justify-center gap-1 ${
                                  currentTheme === 'light'
                                    ? 'bg-[#0038FF] text-white shadow'
                                    : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                }`}
                              >
                                <Sun size={9} />
                                <span>{TRANSLATIONS[currentLanguage].themeLight}</span>
                              </button>
                              <button
                                onClick={() => {
                                  setCurrentTheme('dark');
                                  setIsProfileDropdownOpen(false);
                                }}
                                className={`py-1 text-[9px] font-black rounded-md transition-all flex items-center justify-center gap-1 ${
                                  currentTheme === 'dark'
                                    ? 'bg-[#0038FF] text-white shadow'
                                    : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                }`}
                              >
                                <Moon size={9} />
                                <span>{TRANSLATIONS[currentLanguage].themeDark}</span>
                              </button>
                            </div>
                          </div>

                          {/* Sign Out Trigger */}
                          <button
                            onClick={() => {
                              handleLogout();
                              setIsProfileDropdownOpen(false);
                            }}
                            className="w-full mt-2 pt-2 border-t border-slate-150 dark:border-slate-800 flex items-center justify-center gap-1 py-1.5 text-[10px] font-black text-rose-500 hover:bg-rose-50/50 dark:hover:bg-rose-950/20 rounded-lg transition-all"
                          >
                            <LogOut size={11} />
                            <span>{TRANSLATIONS[currentLanguage].signOut}</span>
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                </div>
              </div>
            </header>
            {/* END: Styled Modern Header */}

            {/* BEGIN: Sub-Tab Toggle */}
            <div className="px-6 mb-2 shrink-0 animate-fade-in">
              <div className="bg-[#EFF2FC] p-1 rounded-2xl flex items-center justify-between shadow-inner">
                <button 
                  type="button"
                  onClick={() => setOverviewSubTab('search')}
                  className={`flex-1 py-1.5 rounded-xl font-bold text-xs transition-all duration-200 flex items-center justify-center gap-2 ${
                    overviewSubTab === 'search' 
                      ? 'bg-[#0038FF] text-white shadow-md shadow-blue-600/15' 
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <Search size={14} />
                  <span>Search</span>
                </button>
                <button 
                  type="button"
                  onClick={() => setOverviewSubTab('calendar')}
                  className={`flex-1 py-1.5 rounded-xl font-bold text-xs transition-all duration-200 flex items-center justify-center gap-2 ${
                    overviewSubTab === 'calendar' 
                      ? 'bg-[#0038FF] text-white shadow-md shadow-blue-600/15' 
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <Calendar size={14} />
                  <span>Calendar</span>
                </button>
              </div>
            </div>
            {/* END: Sub-Tab Toggle */}

            {/* BEGIN: Search / Custom Calendar Bar */}
            {overviewSubTab === 'search' ? (
              <div className="px-6 mb-2 shrink-0 animate-fade-in">
                <div className="relative">
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                  <input 
                    type="text" 
                    placeholder="Search tasks" 
                    value={searchQuery} 
                    onChange={e => setSearchQuery(e.target.value)} 
                    className="w-full pl-10 pr-9 py-2 bg-[#EFF2FC]/80 border-none rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-[#0038FF]/20 text-slate-800 placeholder:text-slate-400 font-semibold"
                  />
                  {searchQuery && (
                    <button onClick={() => setSearchQuery('')} className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                      <X size={14} />
                    </button>
                  )}
                </div>
              </div>
            ) : (
              <div className="px-6 mb-2 shrink-0 animate-fade-in">
                <div className="bg-white p-2.5 rounded-2xl border border-slate-100 shadow-lg shadow-blue-900/5 transition-all">
                  
                  {/* Calendar Header */}
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-2">
                      <h3 className="font-bold text-xs md:text-sm text-slate-800 font-sans tracking-tight">
                        {calendarPivot.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}
                      </h3>
                      <button
                        type="button"
                        onClick={() => {
                          const todayStr = getLocalDate();
                          setCalendarSelectedDate(todayStr);
                          setCalendarPivot(new Date());
                          setFilterTime('ALL');
                        }}
                        className="text-[9px] font-bold text-[#0038FF] hover:underline bg-[#EFF2FC] px-1.5 py-0.5 rounded transition-all active:scale-95"
                      >
                        Today
                      </button>
                    </div>
                    <div className="flex gap-1.5">
                      <button 
                        type="button"
                        onClick={() => {
                          const prev = new Date(calendarPivot);
                          prev.setMonth(prev.getMonth() - 1);
                          setCalendarPivot(prev);
                        }}
                        className="w-5 h-5 rounded-full flex items-center justify-center hover:bg-slate-100 active:scale-95 text-slate-600 transition-all border border-slate-100"
                        title="Previous Month"
                      >
                        <ChevronLeft size={11} />
                      </button>
                      <button 
                        type="button"
                        onClick={() => {
                          const next = new Date(calendarPivot);
                          next.setMonth(next.getMonth() + 1);
                          setCalendarPivot(next);
                        }}
                        className="w-5 h-5 rounded-full flex items-center justify-center hover:bg-slate-100 active:scale-95 text-slate-600 transition-all border border-slate-100"
                        title="Next Month"
                      >
                        <ChevronRight size={11} />
                      </button>
                    </div>
                  </div>

                  {/* Weekday labels */}
                  <div className="grid grid-cols-7 text-center mb-1">
                    {['S', 'M', 'T', 'W', 'T', 'F', 'S'].map((day, idx) => (
                      <span key={idx} className="text-[9px] font-bold text-slate-450 uppercase tracking-wider">
                        {day}
                      </span>
                    ))}
                  </div>

                  {/* Grid Days */}
                  <div className="grid grid-cols-7 gap-y-0.5 gap-x-0.5 text-center">
                    {getCalendarDays().map(({ dateStr, dayNum, isCurrentMonth }) => {
                      const isSelected = dateStr === calendarSelectedDate;
                      const isToday = dateStr === getLocalDate();
                      const holidayName = INDONESIAN_HOLIDAYS[dateStr];
                      const isHoliday = !!holidayName;
                      
                      const activeTasksOnDate = tasks.filter(t => t.targetDate === dateStr && t.status !== TaskStatus.DONE);
                      const hasTasks = activeTasksOnDate.length > 0;

                      return (
                        <button
                          key={dateStr}
                          type="button"
                          title={holidayName ? `Libur Nasional: ${holidayName}` : undefined}
                          onClick={() => {
                            setCalendarSelectedDate(dateStr);
                            setFilterTime('ALL');
                          }}
                          className="flex flex-col items-center justify-center relative focus:outline-none group py-0.5"
                        >
                          <div className={`w-[22px] h-[22px] rounded-full flex items-center justify-center text-[10px] font-bold transition-all relative ${
                            !isCurrentMonth 
                              ? isHoliday 
                                ? 'text-rose-300 font-medium' 
                                : 'text-slate-300' 
                              : isSelected 
                                ? isHoliday
                                  ? 'bg-rose-600 text-white shadow-sm shadow-rose-600/30'
                                  : 'bg-[#0038FF] text-white shadow-sm shadow-blue-600/30' 
                                : isToday
                                  ? isHoliday
                                    ? 'bg-rose-100 text-rose-600 border border-rose-350'
                                    : 'bg-[#EFF2FC] text-[#0038FF] border border-[#0038FF]/20'
                                  : isHoliday
                                    ? 'bg-rose-50 text-rose-600 border border-rose-200/50 hover:bg-rose-100 font-extrabold'
                                    : 'text-slate-700 hover:bg-slate-100'
                          }`}>
                            {dayNum}
                          </div>
                          
                          {/* Task count/dot indicator */}
                          {hasTasks && (
                            <span 
                              className={`w-1 h-1 rounded-full absolute bottom-px bg-[#0038FF]`} 
                            />
                          )}
                        </button>
                      );
                    })}
                  </div>

                </div>
              </div>
            )}
            {/* END: Search / Custom Calendar Bar */}

            {/* BEGIN: Filters Bar */}
            <div className={`grid ${currentUser?.role === 'ADMIN' ? 'grid-cols-4 gap-1.5' : 'grid-cols-3 gap-2'} px-4 py-1.5 shrink-0 pb-2 border-b border-slate-100`}>
              {currentUser.role === 'ADMIN' && (
                <div className="relative w-full">
                  <select 
                    value={filterUser} 
                    onChange={e => setFilterUser(e.target.value)} 
                    className="appearance-none w-full bg-slate-800 text-white text-[11px] md:text-xs font-bold rounded-xl pl-2 pr-6 py-1.5 cursor-pointer text-center truncate"
                  >
                    <option value="ALL_TEAM">All Team</option>
                    <option value="ME">My Tasks</option>
                    {teamMembers.map(m => (
                      <option key={m.id} value={m.id}>{m.username}</option>
                    ))}
                  </select>
                  <ChevronDown size={11} className="absolute right-1.5 top-1/2 -translate-y-1/2 text-white pointer-events-none" />
                </div>
              )}

              <div className="relative w-full">
                <select 
                  value={filterCategory} 
                  onChange={e => setFilterCategory(e.target.value)} 
                  className="appearance-none w-full bg-[#EFF2FC] border border-transparent text-[11px] md:text-xs font-bold text-slate-600 rounded-xl pl-2 pr-6 py-1.5 cursor-pointer text-center truncate"
                >
                  <option value="ALL">Categories</option>
                  {allCategories.map(cat => (
                    <option key={cat} value={cat}>{cat}</option>
                  ))}
                </select>
                <ChevronDown size={11} className="absolute right-1.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
              </div>

              <div className="relative w-full">
                <select 
                  value={filterStatus} 
                  onChange={e => setFilterStatus(e.target.value)} 
                  className="appearance-none w-full bg-[#EFF2FC] border border-transparent text-[11px] md:text-xs font-bold text-slate-600 rounded-xl pl-2 pr-6 py-1.5 cursor-pointer text-center truncate"
                >
                  <option value="ALL">Status</option>
                  <option value="ACTIVE">Tasks</option>
                  {STATUSES.map(status => (
                    <option key={status} value={status}>{status.replace('_', ' ')}</option>
                  ))}
                </select>
                <ChevronDown size={11} className="absolute right-1.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
              </div>

              <div className="relative w-full">
                <select 
                  value={filterTime} 
                  onChange={e => setFilterTime(e.target.value as any)} 
                  className="appearance-none w-full bg-[#EFF2FC] border border-transparent text-[11px] md:text-xs font-bold text-slate-600 rounded-xl pl-2 pr-6 py-1.5 cursor-pointer text-center truncate"
                >
                  <option value="ALL">Time</option>
                  <option value="OVERDUE">Overdue</option>
                  <option value="TODAY">Due Today</option>
                  <option value="THIS_WEEK">Due Week</option>
                  <option value="FUTURE">Future</option>
                </select>
                <ChevronDown size={11} className="absolute right-1.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
              </div>
            </div>
            {/* END: Filters Bar */}

            {/* BEGIN: Daily Sprint Info & Sort */}
            <div className="px-6 py-2 flex items-center justify-between shrink-0 bg-transparent">
              <div>
                <div className="flex items-center flex-wrap gap-2">
                  <h2 className="text-xl font-bold text-[#1C2038] tracking-tight">
                    Daily Sprint
                  </h2>
                  {overviewSubTab === 'calendar' && INDONESIAN_HOLIDAYS[calendarSelectedDate] && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-black uppercase tracking-wider text-rose-600 bg-rose-50 px-2.5 py-0.5 rounded-full border border-rose-200 animate-pulse">
                      🎉 {INDONESIAN_HOLIDAYS[calendarSelectedDate]}
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-400 font-semibold mt-0.5">
                  {tasksToShow.length === 1 ? 'You have 1 task' : `You have ${tasksToShow.length} tasks`}{overviewSubTab === 'calendar' ? ' for chosen day' : ' for today'}
                </p>
              </div>
              
              <button 
                onClick={() => {
                  setTaskSort(prev => {
                    if (prev === 'DATE_ASC') return 'DATE_DESC';
                    if (prev === 'DATE_DESC') return 'STATUS_ASC';
                    if (prev === 'STATUS_ASC') return 'STATUS_DESC';
                    return 'DATE_ASC';
                  });
                }}
                className="flex items-center gap-1.5 text-[#0038FF] font-bold text-xs hover:opacity-85 transition-opacity"
              >
                <ArrowUpDown size={14} />
                <span>
                  {taskSort === 'DATE_ASC' ? 'Newest First' : 
                   taskSort === 'DATE_DESC' ? 'Oldest First' : 
                   taskSort === 'STATUS_ASC' ? 'Status Asc' : 'Status Desc'}
                </span>
              </button>
            </div>
            {/* END: Daily Sprint Info & Sort */}

            {/* BEGIN: Task Cards List */}
            <div className="px-6 pb-28 md:pb-8 space-y-4">
              {tasksToShow.length === 0 ? (
                <div className="flex flex-col items-center justify-center p-12 text-slate-450 bg-white rounded-3xl border border-slate-100 shadow-sm">
                  <div className="w-12 h-12 bg-[#EFF2FC] rounded-full flex items-center justify-center mb-3 text-[#0038FF]">
                    <LayoutList size={22} />
                  </div>
                  <p className="text-sm font-semibold text-slate-600">No tasks found</p>
                  <p className="text-xs text-slate-450 mt-1">Try resetting search query or filters.</p>
                </div>
              ) : (
                tasksToShow.map(task => {
                  const isCompleted = task.status === TaskStatus.DONE;
                  const daysRemaining = getDaysRemaining(task.targetDate);
                  const isOverdue = daysRemaining < 0 && !isCompleted;
                  const taskActionsCount = updates.filter(u => u.taskId === task.id).length;
                  
                  return (
                    <div 
                      key={task.id} 
                      onClick={() => handleTaskSelect(task.id)}
                      className={`bg-white p-5 rounded-[24px] flex items-center justify-between border-2 transition-all cursor-pointer ${
                        selectedTaskId === task.id 
                          ? 'border-[#0038FF] ring-4 ring-[#0038FF]/5 shadow-md shadow-blue-500/5' 
                          : 'border-slate-100/80 hover:border-slate-200 hover:shadow-sm'
                      }`}
                    >
                      {/* Left details */}
                      <div className="flex-1 min-w-0 pr-4">
                        <h3 className={`font-bold text-base leading-tight truncate transition-all duration-300 ${
                          isCompleted ? 'text-[#8E9BB2] line-through font-semibold' : 'text-[#1C2038]'
                        }`}>
                          {task.title}
                        </h3>
                        
                        <div className="flex flex-wrap items-center gap-1.5 mt-2">
                          <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-[#E6ECFF] text-[#0038FF] uppercase tracking-wider">
                            {task.category}
                          </span>
                          
                          {isOverdue ? (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-[#FFEBEC] text-[#D83F52] uppercase tracking-wider">
                              Overdue
                            </span>
                          ) : isCompleted ? (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-[#E6F4EA] text-[#137333] uppercase tracking-wider">
                              Done
                            </span>
                          ) : (
                            <span className={`inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-slate-100 text-[#5F6368] uppercase tracking-wider`}>
                              In Progress
                            </span>
                          )}

                          <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-slate-100 text-slate-600 uppercase tracking-wider" title="Due Date">
                            {formatDDMMM(task.targetDate)}
                          </span>

                          <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-purple-50 text-purple-700 border border-purple-100/60 uppercase tracking-wider" title="Actions Count">
                            {taskActionsCount}
                          </span>

                          {task.tags?.map(t => (
                            <span key={t} className="inline-flex items-center px-1.5 py-0.5 rounded-md text-[8.5px] font-extrabold bg-blue-50/70 text-[#0038FF] border border-blue-100/30 uppercase tracking-tight">
                              #{t}
                            </span>
                          ))}

                          {/* Show owner small badge for admin */}
                          {currentUser.role === 'ADMIN' && task.userId !== currentUser.id && userMap[task.userId] && (
                            <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-amber-50 text-amber-700 border border-amber-100" title="Task Owner">
                              {userMap[task.userId].username}
                            </span>
                          )}
                        </div>
                      </div>
                      
                      {/* Right toggle completion checkmark */}
                      <button 
                        type="button"
                        onClick={(e) => handleToggleTaskStatus(task, e)}
                        className={`p-1.5 flex items-center justify-center shrink-0 transition-all active:scale-[0.98] hover:scale-105 ${
                          isCompleted 
                            ? 'text-[#0038FF]' 
                            : 'text-slate-300 hover:text-[#0038FF]'
                        }`}
                        title={isCompleted ? 'Mark Incomplete' : 'Mark Complete'}
                      >
                        <Check size={18} className="stroke-[3.5]" />
                      </button>
                    </div>
                  );
                })
              )}
            </div>
            {/* END: Task Cards List */}

            {/* Dynamic FAB Floating Action Button */}
            <button 
              onClick={createNewTask}
              className="absolute bottom-28 right-6 w-14 h-14 bg-[#0038FF] hover:bg-blue-700 text-white rounded-full flex items-center justify-center shadow-xl shadow-blue-600/30 active:scale-95 hover:scale-105 transition-all z-40"
              title="Create New Task"
            >
              <Plus size={28} />
            </button>
            
          </div>

          <div id="detail-tab-container" className={`flex-col bg-stone-50 border-t border-stone-200 shadow-[0_-4px_20px_rgba(0,0,0,0.05)] z-10 min-h-0 ${['detail', 'profile', 'updates', 'dashboard', 'ai'].includes(activeTab) ? 'flex flex-1 animate-fade-in' : 'hidden'} ${activeTab === 'ai' ? 'overflow-hidden' : 'overflow-y-auto'} md:overflow-hidden md:relative md:flex md:h-[45vh] md:flex-row md:z-auto`}>
            {activeTab === 'updates' ? (
              <div className="w-full h-full flex flex-col bg-white animate-fade-in overflow-hidden">
                 <header className={`sticky top-0 z-30 px-5 py-3 shrink-0 transition-colors duration-200 border-b ${
                   currentTheme === 'dark' 
                     ? 'bg-[#111827] border-[#1F2937]' 
                     : 'bg-[#F9FAFB] border-slate-150'
                 }`}>
                   <div className="flex items-center justify-between">
                     <div className="flex items-center gap-2">
                       <div className="w-2.5 h-2.5 bg-[#0038FF] rounded-full animate-pulse" />
                       <h1 className={`text-sm font-black uppercase tracking-widest transition-colors ${
                         currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                       }`} style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                         {currentLanguage === 'ID' ? 'RIWAYAT UPDATE' : 'UPDATES FEED'}
                       </h1>
                     </div>
                     <div className="flex items-center gap-3">
                       <span className={`text-[10px] font-bold px-2.5 py-1 rounded-lg border ${
                         currentTheme === 'dark' 
                           ? 'text-slate-400 bg-slate-800/40 border-slate-700/60' 
                           : 'text-[#0038FF] bg-[#EFF2FC] border-[#0038FF]/20 font-extrabold'
                       }`}>
                         {updates.length} {currentLanguage === 'ID' ? 'TOTAL ENTRI' : 'TOTAL ENTRIES'}
                       </span>
                     </div>
                   </div>
                 </header>

                 {/* SEARCH & FILTERS ROW FOR UPDATES FEED */}
                 <div className={`px-5 py-3 border-b flex flex-col md:flex-row gap-3 ${
                   currentTheme === 'dark' ? 'bg-[#0D1524] border-[#1F2937]' : 'bg-slate-50 border-slate-150'
                 }`}>
                   {/* Search input with search icon + action button */}
                   <div className="relative flex-1">
                     <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-405" size={13} />
                     <input 
                       type="text" 
                       placeholder={currentLanguage === 'ID' ? 'Cari teks ritme update...' : 'Search update text...'} 
                       value={updatesQuery} 
                       onChange={e => setUpdatesQuery(e.target.value)} 
                       className={`w-full pl-8 pr-8 py-1.5 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-[#0038FF]/20 border ${
                         currentTheme === 'dark' 
                           ? 'bg-slate-900 border-slate-800 text-white placeholder:text-slate-500' 
                           : 'bg-white border-slate-200 text-slate-800 placeholder:text-slate-400'
                       } font-semibold`}
                     />
                     {updatesQuery && (
                       <button onClick={() => setUpdatesQuery('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                         <X size={12} />
                       </button>
                     )}
                   </div>
                   
                   {/* Dropdown Filters for Category(Project) and Status */}
                   <div className="flex gap-2 shrink-0">
                     {/* Project (Category) filter */}
                     <span className="relative flex-1 md:w-36">
                       <select 
                         value={updatesProjectFilter} 
                         onChange={e => setUpdatesProjectFilter(e.target.value)} 
                         className={`appearance-none w-full border text-xs font-bold rounded-xl pl-3 pr-7 py-1.5 cursor-pointer truncate ${
                           currentTheme === 'dark' 
                             ? 'bg-slate-900 border-slate-800 text-slate-300' 
                             : 'bg-white border-slate-200 text-slate-600'
                         }`}
                       >
                         <option value="ALL">{currentLanguage === 'ID' ? 'Proyek: Semua' : 'Project: All'}</option>
                         {allCategories.map(cat => (
                           <option key={cat} value={cat}>{cat}</option>
                         ))}
                       </select>
                       <ChevronDown size={11} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                     </span>

                     {/* Status filter */}
                     <span className="relative flex-1 md:w-36">
                       <select 
                         value={updatesStatusFilter} 
                         onChange={e => setUpdatesStatusFilter(e.target.value)} 
                         className={`appearance-none w-full border text-xs font-bold rounded-xl pl-3 pr-7 py-1.5 cursor-pointer truncate ${
                           currentTheme === 'dark' 
                             ? 'bg-slate-900 border-[#1F2937] text-slate-300' 
                             : 'bg-white border-slate-203 text-slate-600'
                         }`}
                       >
                         <option value="ALL">{currentLanguage === 'ID' ? 'Status: Semua' : 'Status: All'}</option>
                         {STATUSES.map(status => (
                           <option key={status} value={status}>{status.replace('_', ' ')}</option>
                         ))}
                       </select>
                       <ChevronDown size={11} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                     </span>
                   </div>
                 </div>

                 <div className="flex-1 overflow-y-auto p-4 md:p-6 space-y-6 bg-beige/30">
                    {sortedUpdateDates.length === 0 ? (
                        <div className="h-full flex flex-col items-center justify-center text-stone-400 opacity-40">
                            <Layers size={48} className="mb-4" />
                            <p className="font-bold">No updates found.</p>
                        </div>
                    ) : (
                        sortedUpdateDates.map(date => {
                            const dayUpdatesCount = (groupedUpdatesByDate[date] || []).length;
                            return (
                                <div key={date} className="relative pl-6 border-l-2 border-primary/40 pb-2">
                                    <div className="absolute -left-[9px] top-0 w-4 h-4 rounded-full bg-primary border-4 border-white shadow-sm"></div>
                                    <div className="flex items-center gap-2 mb-3">
                                        <h3 className="text-[10px] font-black text-secondary/60 bg-primary/20 inline-block px-3 py-1 rounded-full uppercase tracking-widest">{formatListDate(date)}</h3>
                                        <span className="text-[10px] font-black bg-[#0038FF] text-white px-2 py-0.5 rounded-full shadow-xs shrink-0 flex items-center gap-1">
                                            <span>{dayUpdatesCount}</span>
                                            <span className="text-[8px] tracking-wider opacity-90">{dayUpdatesCount === 1 ? 'UPDATE' : 'UPDATES'}</span>
                                        </span>
                                    </div>
                                    <div className="space-y-2.5">
                                    {(groupedUpdatesByDate[date] || []).map((entry: TaskUpdate) => {
                                        const taskObj = tasks.find(t => t.id === entry.taskId);
                                        const taskTitle = taskObj?.title || 'Unknown Task';
                                        const status = entry.statusChange?.to || taskObj?.status || TaskStatus.NOT_YET;
                                        return (
                                            <div key={entry.id} onClick={() => handleTaskSelect(entry.taskId)} className="group bg-white rounded-xl border border-stone-200 p-3 shadow-sm hover:ring-1 hover:ring-primary/50 cursor-pointer transition-all active:scale-[0.99] hover:shadow-md">
                                                <div className="flex justify-between items-center gap-3">
                                                    <div className="flex items-center gap-2 min-w-0">
                                                        <div className="w-1 h-3 bg-accent rounded-full shrink-0"></div>
                                                        <h4 className="text-xs font-bold text-secondary truncate">{taskTitle}</h4>
                                                    </div>
                                                    <div className="flex items-center gap-2 shrink-0">
                                                        <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded-full border whitespace-nowrap ${STATUS_COLORS[status as TaskStatus]}`}>{status.replace('_', ' ')}</span>
                                                        {taskObj && (
                                                            <span className={`text-[8.5px] font-extrabold px-1.5 py-0.5 rounded-full border tracking-wide uppercase whitespace-nowrap ${
                                                                currentTheme === 'dark'
                                                                    ? 'bg-slate-800/60 border-slate-700/80 text-slate-350'
                                                                    : 'bg-[#EFF2FC] border-[#0038FF]/15 text-[#0038FF]'
                                                            }`}>
                                                                {taskObj.category}
                                                            </span>
                                                        )}
                                                    </div>
                                                </div>
                                                <div className="mt-1.5 border-l-2 border-stone-100 ml-1 pl-2.5">
                                                    <div className="text-[11px] text-stone-600 leading-relaxed italic">"{entry.content}"</div>
                                                    {currentUser?.role === 'ADMIN' && taskObj && taskObj.userId !== currentUser.id && (
                                                        <div className="text-[9px] text-stone-400 mt-1 font-bold">
                                                            {currentLanguage === 'ID' ? 'Oleh: ' : 'By: '}
                                                            {userMap[taskObj.userId]?.username || 'Team Member'}
                                                        </div>
                                                    )}
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                            </div>
                        )})
                    )}
                 </div>
              </div>
            ) : activeTab === 'ai' ? (
              <div className="w-full h-full flex flex-col bg-white animate-fade-in overflow-hidden relative">
                <header className={`sticky top-0 z-30 px-5 py-3 shrink-0 transition-colors duration-200 border-b ${
                  currentTheme === 'dark' 
                    ? 'bg-[#111827] border-[#1F2937] text-white' 
                    : 'bg-[#F9FAFB] border-slate-150'
                }`}>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <div className="w-2.5 h-2.5 bg-[#0038FF] rounded-full animate-pulse" />
                      <h1 className={`text-sm font-black uppercase tracking-widest transition-colors ${
                        currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                      }`} style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                        {currentLanguage === 'ID' ? 'ASISTEN TUGAS AI' : 'TASK AI ASSISTANT'}
                      </h1>
                    </div>
                    <button 
                      onClick={() => {
                        if (aiClearConfirm) {
                          setAiMessages([
                            {
                              id: 'welcome',
                              role: 'assistant',
                              content: "Hello! I'm your TaskFlow assistant. I can help you organize your schedule, summarize your pending items, or suggest improvements to your workflow. How can I assist you today?",
                              timestamp: Date.now()
                            }
                          ]);
                          setAiClearConfirm(false);
                        } else {
                          setAiClearConfirm(true);
                        }
                      }}
                      className={`text-[10px] font-extrabold px-3 py-1.5 rounded-lg border transition-all duration-200 flex items-center gap-1 hover:shadow-xs active:scale-95 ${
                        aiClearConfirm 
                          ? 'bg-red-50 border-red-300 text-red-600 hover:bg-red-100 hover:border-red-400 font-bold' 
                          : 'bg-white border-stone-200 text-stone-400 hover:text-[#0038FF] hover:border-[#0038FF]/30'
                      }`}
                      title={aiClearConfirm ? "Confirm Clear" : "Clear History"}
                    >
                      <Trash2 size={11} /> {aiClearConfirm ? (currentLanguage === 'ID' ? 'YAKIN?' : 'SURE?') : (currentLanguage === 'ID' ? 'KOSONGKAN' : 'CLEAR')}
                    </button>
                  </div>
                </header>

                <div className="flex-1 overflow-hidden p-4 flex flex-col justify-between bg-stone-50/40">
                  <div className="flex-1 overflow-y-auto space-y-4 max-h-[calc(100vh-290px)] md:max-h-[30vh] p-1 custom-scrollbar">
                    {aiMessages.map((msg) => (
                      <div key={msg.id} className={`flex flex-col ${msg.role === 'user' ? 'items-end' : 'items-start'} max-w-[85%] ${msg.role === 'user' ? 'ml-auto' : 'mr-auto'}`}>
                        {msg.role === 'assistant' && (
                          <div className="flex items-center gap-1.5 mb-1 pl-1 text-[10px] font-black text-slate-500 uppercase tracking-widest leading-none">
                            <div className="w-[18px] h-[18px] rounded-full bg-[#EFF2FC] text-[#0038FF] border border-[#0038FF]/10 flex items-center justify-center">
                              <Sparkles size={10} className="stroke-[2.5]" />
                            </div>
                            Task AI
                          </div>
                        )}
                        
                        <div 
                          className={`p-3.5 rounded-2xl text-xs md:text-sm shadow-sm transition-all duration-200 leading-relaxed ${
                            msg.role === 'user' 
                              ? 'bg-[#0038FF] text-white rounded-tr-none font-bold' 
                              : 'bg-[#EFF2FC]/75 text-slate-800 border border-slate-200/50 border-l-[3.5px] border-l-[#0038FF] rounded-tl-none'
                          }`}
                        >
                          {msg.role === 'user' ? (
                            <p className="whitespace-pre-wrap">{msg.content}</p>
                          ) : (
                            renderAiMessageContent(msg.content)
                          )}
                          
                          {msg.options && msg.options.length > 0 && (
                            <div className="mt-3.5 flex flex-wrap gap-2 pt-2 border-t border-[#0038FF]/10">
                              {msg.options.map((opt: any, i: number) => (
                                <button
                                  key={i}
                                  id={`ai-option-btn-${i}`}
                                  onClick={() => {
                                    if (opt.action) {
                                      executeAiAction(opt.action);
                                      setAiMessages(prev => [...prev, {
                                        id: Date.now().toString(),
                                        role: 'assistant',
                                        content: `Processing auto-action: ${opt.label}...`,
                                        timestamp: Date.now()
                                      }]);
                                    } else {
                                      sendAiMessage(opt.label);
                                    }
                                  }}
                                  className="px-3 py-1.5 rounded-full text-[10px] font-black bg-[#EBF0FF] text-[#0038FF] border border-[#0038FF]/20 hover:border-[#0038FF] hover:bg-[#0038FF] hover:text-white active:scale-95 transition-all duration-150 shadow-sm flex items-center gap-1.5 cursor-pointer"
                                >
                                  <Sparkles size={10} className="stroke-[2.5]" />
                                  {opt.label}
                                </button>
                              ))}
                            </div>
                          )}
                        </div>
                        
                        <div className="w-full flex items-center justify-between gap-4 mt-1 px-1">
                          <span className="text-[9px] text-[#A3A3A3] font-mono select-none">
                            {new Date(msg.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                          </span>
                          {msg.role === 'assistant' && msg.usage && (
                            <span className="text-[9px] text-slate-500 font-mono bg-[#EFF2FC]/80 dark:bg-stone-800/40 px-2 py-0.5 rounded-md flex items-center gap-2 select-none border border-slate-200/50">
                              <span>📥 Input: <strong className="text-slate-800 font-black">{msg.usage.promptTokens}</strong></span>
                              <span className="opacity-40">|</span>
                              <span>📤 Output: <strong className="text-slate-800 font-black">{msg.usage.candidatesTokens}</strong></span>
                            </span>
                          )}
                        </div>
                      </div>
                    ))}
                    
                    {isAiLoading && (
                      <div className="flex flex-col items-start max-w-[80%] mr-auto animate-pulse">
                        <div className="flex items-center gap-1.5 mb-1 pl-1 text-[10px] font-black text-slate-500 uppercase tracking-widest leading-none">
                          <div className="w-[18px] h-[18px] rounded-full bg-[#EFF2FC] text-[#0038FF] border border-[#0038FF]/10 flex items-center justify-center">
                            <Sparkles size={10} className="stroke-[2.5]" />
                          </div>
                          Task AI
                        </div>
                        <div className="bg-[#EFF2FC]/60 p-3 rounded-2xl rounded-tl-none border-l-[3.5px] border-l-[#0038FF] text-xs font-medium text-slate-500 flex items-center gap-2">
                          <Loader2 size={12} className="animate-spin text-[#0038FF]" /> Processing schedules & task graphs...
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="space-y-2 mt-2 pt-2 border-t border-stone-100 pb-20 md:pb-2">
                    <div className="flex gap-1.5 overflow-x-auto pb-1 text-[10px] no-scrollbar">
                      <button 
                        type="button"
                        onClick={() => sendAiMessage("Analyze my Overdue tasks and list them")}
                        className="px-3 py-1.5 rounded-full border border-stone-200 bg-white hover:bg-[#EFF2FC] hover:text-[#0038FF] hover:border-[#0038FF]/30 transition-all cursor-pointer font-extrabold shrink-0 shadow-sm"
                      >
                        📋 Overdue Tasks
                      </button>
                      <button 
                        type="button"
                        onClick={() => sendAiMessage("Suggest improvements for my pending tasks")}
                        className="px-3 py-1.5 rounded-full border border-stone-200 bg-white hover:bg-[#EFF2FC] hover:text-[#0038FF] hover:border-[#0038FF]/30 transition-all cursor-pointer font-extrabold shrink-0 shadow-sm"
                      >
                        ⚡ Workflow Optimization
                      </button>
                      <button 
                        type="button"
                        onClick={() => sendAiMessage("Summarize my active task counts by category")}
                        className="px-3 py-1.5 rounded-full border border-stone-200 bg-white hover:bg-[#EFF2FC] hover:text-[#0038FF] hover:border-[#0038FF]/30 transition-all cursor-pointer font-extrabold shrink-0 shadow-sm"
                      >
                        📊 Category Balance
                      </button>
                    </div>

                    <form 
                      onSubmit={(e) => {
                        e.preventDefault();
                        sendAiMessage(aiInputText);
                      }}
                      className="flex gap-1.5 items-center"
                    >
                      <input
                        type="text"
                        value={aiInputText}
                        onChange={(e) => setAiInputText(e.target.value)}
                        placeholder={currentLanguage === 'ID' ? 'Tanyakan Asisten AI... (misal: "Apa target sore ini?")' : 'Ask Task AI... (e.g. "What does my afternoon look like?")'}
                        disabled={isAiLoading}
                        className="flex-1 px-4.5 py-2.5 rounded-full border border-stone-200 bg-white text-xs md:text-sm font-extrabold text-secondary focus:outline-none focus:ring-2 focus:ring-[#0038FF]/30 shadow-inner"
                      />
                      <button
                        type="submit"
                        disabled={isAiLoading || !aiInputText.trim()}
                        className={`w-10 h-10 rounded-full flex items-center justify-center text-white transition-all ${
                          aiInputText.trim() && !isAiLoading
                            ? 'bg-[#0038FF] hover:bg-blue-700 shadow-md shadow-blue-600/30 active:scale-95'
                            : 'bg-stone-200 text-stone-400 cursor-not-allowed'
                        }`}
                      >
                        <Plus size={18} className="rotate-45" />
                      </button>
                    </form>
                  </div>
                </div>
              </div>
            ) : activeTab === 'profile' ? (
              <div className="w-full h-full flex flex-col bg-white animate-fade-in overflow-hidden">
                <header className={`sticky top-0 z-30 px-5 py-3 shrink-0 transition-colors duration-200 border-b ${
                  currentTheme === 'dark' 
                    ? 'bg-[#111827] border-[#1F2937]' 
                    : 'bg-[#F9FAFB] border-slate-150'
                }`}>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <div className="w-2.5 h-2.5 bg-[#0038FF] rounded-full animate-pulse" />
                      <h1 className={`text-sm font-black uppercase tracking-widest transition-colors ${
                        currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                      }`} style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                        {currentLanguage === 'ID' ? 'PENGATURAN PROFIL' : 'PROFILE SETTINGS'}
                      </h1>
                    </div>
                  </div>
                </header>
                <div className="flex-1 overflow-y-auto p-6 md:p-8">
                  <div className="max-w-xl mx-auto">
                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-stone-200 mb-6 flex items-center gap-5">
                      <div className="w-16 h-16 bg-accent rounded-full flex items-center justify-center text-secondary text-2xl font-bold shadow-md">{currentUser.username[0].toUpperCase()}</div>
                      <div>
                        <div className="font-bold text-xl text-secondary">{currentUser.username}</div>
                        <div className="text-stone-500 text-xs font-bold uppercase tracking-wider bg-stone-100 px-2 py-0.5 rounded-md inline-block mt-1">{currentUser.role}</div>
                      </div>
                    </div>
                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-stone-200 mb-6">
                      <h3 className="font-bold text-stone-700 mb-4 flex items-center gap-2"><Lock size={16}/> {currentLanguage === 'ID' ? 'Ubah Kata Sandi' : 'Update Password'}</h3>
                      <div className="flex flex-col gap-4">
                        <div>
                          <label className="text-xs font-bold text-stone-400 uppercase ml-1 mb-1 block">{currentLanguage === 'ID' ? 'Kata Sandi Baru' : 'New Password'}</label>
                          <input type="password" value={newPin} onChange={e => setNewPin(e.target.value)} placeholder={currentLanguage === 'ID' ? 'Masukkan kata sandi baru' : 'Enter new password'} className="w-full px-4 py-3 bg-stone-50 border border-stone-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary/50 text-secondary tracking-wide" />
                        </div>
                        <div>
                          <label className="text-xs font-bold text-stone-400 uppercase ml-1 mb-1 block">{currentLanguage === 'ID' ? 'Konfirmasi Kata Sandi Baru' : 'Confirm New Password'}</label>
                          <input type="password" value={confirmPin} onChange={e => setConfirmPin(e.target.value)} placeholder={currentLanguage === 'ID' ? 'Konfirmasi kata sandi baru' : 'Confirm new password'} className="w-full px-4 py-3 bg-stone-50 border border-stone-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary/50 text-secondary tracking-wide" />
                        </div>
                        {profileMessage && (
                          <div className={`text-xs font-bold p-3 rounded-lg flex items-center gap-2 ${profileMessage.type === 'success' ? 'bg-emerald-50 text-emerald-600' : 'bg-red-50 text-red-600'}`}>
                            {profileMessage.type === 'success' ? <CheckCircle2 size={14} /> : <AlertCircle size={14} />}
                            {profileMessage.text}
                          </div>
                        )}
                        <button onClick={handleUpdateProfile} className="bg-secondary text-accent font-bold py-3 rounded-xl hover:bg-stone-800 transition-colors shadow-lg shadow-stone-200 mt-2">{currentLanguage === 'ID' ? 'Perbarui Kata Sandi' : 'Update Password'}</button>
                      </div>
                    </div>
                    {/* Google Sheets Export Card */}
                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-stone-200 mb-6 mt-6">
                      <h3 className="font-black text-secondary mb-2 flex items-center gap-2 text-sm tracking-tight" style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                        <Globe size={18} className="text-[#0038FF]" /> 
                        {currentLanguage === 'ID' ? 'EKSPOR DATA KE GOOGLE SHEET' : 'EXPORT DATA TO GOOGLE SPREADSHEET'}
                      </h3>
                      <p className="text-xs text-stone-500 font-medium mb-5 leading-relaxed">
                        {currentLanguage === 'ID' 
                          ? 'Dapatkan salinan data tugas, rincian update harian, dan tim Anda dalam bentuk Google Spreadsheet multi-tab yang tertata rapi secara real-time.'
                          : 'Produce a beautifully formatted, multi-tab Google Spreadsheet file containing all tasks, updates line feed, and user settings in real-time.'}
                      </p>

                      <div className="flex flex-col gap-3">
                        <button
                          onClick={handleExportToSheets}
                          disabled={isExporting}
                          className="w-full bg-[#0038FF] hover:bg-[#0038FF]/95 disabled:bg-slate-400 font-black text-white py-3.5 px-4 rounded-xl transition-all shadow-md active:scale-[0.99] flex items-center justify-center gap-2 group text-xs uppercase tracking-wider"
                        >
                          {isExporting ? (
                            <>
                              <Loader2 size={16} className="animate-spin" />
                              <span>{currentLanguage === 'ID' ? 'SEDANG MENGEKSPOR...' : 'EXPORTING DATA...'}</span>
                            </>
                          ) : (
                            <>
                              <ArrowUpRight size={16} className="group-hover:translate-x-0.5 group-hover:-translate-y-0.5 transition-transform" />
                              <span>{currentLanguage === 'ID' ? 'EKSPOR SEKARANG' : 'EXPORT NOW'}</span>
                            </>
                          )}
                        </button>

                        {exportUrl && (
                          <div className="mt-3 p-4 bg-[#EFF2FC] border border-[#0038FF]/20 rounded-xl flex items-center justify-between gap-3 animate-fade-in group">
                            <div className="min-w-0">
                              <p className="text-[11px] font-black text-[#0038FF] uppercase tracking-wider mb-0.5">
                                {currentLanguage === 'ID' ? 'Pengeksportan Berhasil!' : 'Export Successful!'}
                              </p>
                              <p className="text-[10px] text-slate-500 truncate">
                                {currentLanguage === 'ID' ? 'Data berhasil dipindahkan ke sheet baru.' : 'Your data is synced with a fresh Google sheet.'}
                              </p>
                            </div>
                            <a
                              href={exportUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="bg-white border border-[#0038FF]/25 hover:bg-slate-50 text-[#0038FF] font-black text-[10px] px-3.5 py-2 rounded-lg transition-all shadow-sm shrink-0 uppercase tracking-widest"
                            >
                              {currentLanguage === 'ID' ? 'Buka Sheet' : 'Open Sheet'}
                            </a>
                          </div>
                        )}

                        {exportError && (
                          <div className="mt-3 p-4 bg-rose-50 border border-rose-100 rounded-xl text-rose-600 text-xs font-bold leading-relaxed flex items-start gap-2 animate-fade-in">
                            <AlertCircle size={15} className="mt-0.5 shrink-0" />
                            <div>
                              <p className="font-extrabold uppercase tracking-wide text-[10px] mb-0.5">Export Failed</p>
                              <p className="text-[11px] opacity-90 font-medium">{exportError}</p>
                            </div>
                          </div>
                        )}

                        {/* Automated scheduler information */}
                        <div className="mt-4 pt-4 border-t border-slate-100 flex flex-col gap-2.5">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-1.5">
                              <Clock size={13} className="text-[#0038FF]" />
                              <span className="text-[10px] font-extrabold uppercase tracking-widest text-[#1C2038]">
                                {currentLanguage === 'ID' ? 'Jadwal Ekspor Otomatis' : 'AUTOMATED EXPORT SCHEDULE'}
                              </span>
                            </div>
                            <span className="text-[9px] font-bold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full uppercase tracking-wider">
                              {currentLanguage === 'ID' ? 'Sen, Rab, Jum @ 17:00' : 'MON, WED, FRI @ 5 PM'}
                            </span>
                          </div>
                          
                          <div className="bg-stone-50/50 rounded-xl p-3 border border-stone-100 text-[11px] leading-relaxed">
                            <div className="flex justify-between items-center mb-1">
                              <span className="text-stone-400 font-bold">
                                {currentLanguage === 'ID' ? 'Ekspor Berikutnya:' : 'Next Auto-Export:'}
                              </span>
                              <span className="font-extrabold text-[#1C2038]">
                                {new Date(getNextScheduledSlot(new Date())).toLocaleString('en-US', {
                                  weekday: 'short',
                                  month: 'short',
                                  day: '2-digit',
                                  hour: 'numeric',
                                  minute: '2-digit',
                                  hour12: true
                                })}
                              </span>
                            </div>

                            {autoExportState.lastExportTime && (
                              <div className="flex justify-between items-center mb-2">
                                <span className="text-stone-400 font-bold">
                                  {currentLanguage === 'ID' ? 'Terakhir Diekspor:' : 'Last Exported:'}
                                </span>
                                <span className="font-semibold text-stone-600">
                                  {new Date(autoExportState.lastExportTime).toLocaleString('en-US', {
                                    month: 'short',
                                    day: '2-digit',
                                    hour: 'numeric',
                                    minute: '2-digit',
                                    hour12: true
                                  })}
                                </span>
                              </div>
                            )}

                            {/* Status block */}
                            {autoExportState.status === 'RUNNING' && (
                              <div className="mt-2.5 p-2 bg-blue-50 border border-blue-100 text-blue-600 rounded-lg flex items-center gap-2 animate-pulse text-[10px] font-bold uppercase tracking-wider">
                                <Loader2 size={12} className="animate-spin text-blue-500" />
                                <span>{currentLanguage === 'ID' ? 'Sedang mengekspor data secara otomatis...' : 'Running background auto-export...'}</span>
                              </div>
                            )}

                            {autoExportState.status === 'PENDING_AUTH' && (
                              <div className="mt-2.5 p-2 bg-orange-50 border border-orange-100 text-orange-600 rounded-lg flex flex-col gap-1.5">
                                <div className="flex items-start gap-1.5">
                                  <AlertCircle size={14} className="mt-0.5 shrink-0 text-orange-500" />
                                  <div>
                                    <p className="font-extrabold uppercase tracking-wide text-[9px] text-orange-700">
                                      {currentLanguage === 'ID' ? 'Butuh Otorisasi Google' : 'Google Authorization Required'}
                                    </p>
                                    <p className="text-[10px] text-orange-600/90 font-medium">
                                      {currentLanguage === 'ID' 
                                        ? 'Jadwal ekspor tertunda. Masuk dengan Google untuk mengaktifkan sinkronisasi terjadwal.' 
                                        : 'Auto-export is pending activation. Authorize Google Sheets to allow automated exports.'}
                                    </p>
                                  </div>
                                </div>
                                <button
                                  type="button"
                                  onClick={async () => {
                                    try {
                                      const { error } = await loginWithGoogle();
                                      if (error) throw new Error(error);
                                      // Trigger check
                                      const token = getCachedAccessToken();
                                      if (token) {
                                        const now = new Date();
                                        const mostRecentSlot = getMostRecentScheduledSlot(now);
                                        const slotStr = mostRecentSlot.toISOString();
                                        setAutoExportState(prev => ({ ...prev, status: 'RUNNING' }));
                                        const timestamp = now.toISOString().slice(0, 10);
                                        const title = `Scheduled Auto-Export - ${timestamp}`;
                                        const sheetInfo = await createExportSpreadsheet(token, title);
                                        await populateSpreadsheetData(token, sheetInfo.spreadsheetId, {
                                          tasks,
                                          updates,
                                          users: allUsers,
                                          currentUser: currentUser!
                                        });
                                        const nowStr = new Date().toISOString();
                                        localStorage.setItem('last_auto_export_checked_slot', slotStr);
                                        localStorage.setItem('last_auto_export_time', nowStr);
                                        localStorage.setItem('last_auto_export_status', 'SUCCESS');
                                        localStorage.setItem('last_auto_export_error', '');
                                        setAutoExportState({
                                          lastCheckedSlot: slotStr,
                                          lastExportTime: nowStr,
                                          status: 'SUCCESS',
                                          error: null
                                        });
                                      }
                                    } catch (err: any) {
                                      console.error('Scheduled export login error:', err);
                                    }
                                  }}
                                  className="self-start text-[10px] font-black uppercase tracking-wider text-orange-750 hover:text-orange-900 bg-white border border-orange-200 hover:border-orange-300 px-3 py-1.5 rounded-lg transition-colors cursor-pointer"
                                >
                                  {currentLanguage === 'ID' ? 'Masuk & Sinkronkan Sekarang' : 'Sign-In & Sync Now'}
                                </button>
                              </div>
                            )}

                            {autoExportState.status === 'SUCCESS' && (
                              <div className="mt-2.5 p-2 bg-emerald-50 border border-emerald-100 text-emerald-600 rounded-lg flex items-center gap-1.5 text-[10px] font-bold">
                                <CheckCircle2 size={13} className="text-emerald-500 shrink-0" />
                                <span className="flex-1">
                                  {currentLanguage === 'ID' 
                                    ? 'Ekspor terjadwal terakhir berhasil diproses.' 
                                    : 'Last automated export executed successfully.'}
                                </span>
                              </div>
                            )}

                            {autoExportState.status === 'FAILED' && (
                              <div className="mt-2.5 p-2 bg-rose-50 border border-rose-100 text-rose-600 rounded-lg flex items-start gap-1.5 text-[10px]">
                                <AlertTriangle size={13} className="text-rose-500 shrink-0 mt-0.5" />
                                <div className="flex-1">
                                  <p className="font-extrabold uppercase tracking-wider text-[9px] text-rose-700">
                                    {currentLanguage === 'ID' ? 'Ekspor Terjadwal Gagal' : 'Scheduled Export Failed'}
                                  </p>
                                  <p className="opacity-90 font-medium">{autoExportState.error}</p>
                                </div>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Google Sheets Import Card */}
                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-stone-200 mb-6 mt-4">
                      <h3 className="font-black text-secondary mb-2 flex items-center gap-2 text-sm tracking-tight" style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                        <Globe size={18} className="text-[#00C48C]" /> 
                        {currentLanguage === 'ID' ? 'IMPOR DATA DARI SPREADSHEET' : 'IMPORT DATA FROM SPREADSHEET'}
                      </h3>
                      <p className="text-xs text-stone-500 font-medium mb-5 leading-relaxed">
                        {currentLanguage === 'ID' 
                          ? 'Impor data tugas, log pembaruan harian, dan anggota tim kembali dari URL Google Spreadsheet pilihan Anda atau unggah file CSV ekspor.'
                          : 'Restore task data, updates feed, and team configurations back from an active Google Spreadsheet URL or upload your exported CSV sheet.'}
                      </p>

                      <div className="flex flex-col gap-4">
                        {/* URL/ID Input */}
                        <div>
                          <label className="text-[10px] font-bold text-stone-400 uppercase tracking-wider ml-1 mb-1.5 block">
                            {currentLanguage === 'ID' ? 'URL ATAU ID GOOGLE SPREADSHEET' : 'GOOGLE SPREADSHEET URL OR ID'}
                          </label>
                          <div className="flex gap-2">
                            <input 
                              type="text" 
                              value={importUrlOrId} 
                              onChange={e => setImportUrlOrId(e.target.value)} 
                              placeholder={currentLanguage === 'ID' ? 'https://docs.google.com/spreadsheets/d/...' : 'https://docs.google.com/spreadsheets/d/...'}
                              className="flex-1 px-4 py-3 bg-stone-50 border border-stone-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#0038FF]/40 text-xs font-semibold text-secondary min-w-0" 
                            />
                            <button
                              onClick={handleImportClick}
                              disabled={isImporting}
                              className="bg-[#0038FF] hover:bg-[#0038FF]/95 disabled:bg-slate-300 font-bold text-white px-5 rounded-xl text-xs flex items-center justify-center gap-1.5 transition-all active:scale-[0.98] whitespace-nowrap shadow-sm"
                            >
                              {isImporting ? <Loader2 size={14} className="animate-spin" /> : null}
                              <span>{currentLanguage === 'ID' ? 'IMPOR' : 'IMPORT'}</span>
                            </button>
                          </div>
                        </div>

                        {/* File Upload Options - Split by Table */}
                        <div className="border-t border-stone-100 pt-4 space-y-3">
                          <div>
                            <label className="text-[10px] font-bold text-stone-400 uppercase tracking-wider ml-1 mb-1.5 block">
                              {currentLanguage === 'ID' ? '1. UNGGAH FILE TUGAS (TASKS)' : '1. UPLOAD TASKS FILE'}
                            </label>
                            <input 
                              type="file" 
                              ref={tasksFileInputRef} 
                              onChange={handleTasksFileUpload} 
                              accept=".csv,.xlsx,.xls" 
                              className="hidden" 
                            />
                            <button
                              onClick={() => tasksFileInputRef.current?.click()}
                              className="w-full bg-stone-50 hover:bg-stone-100 text-stone-600 border border-dashed border-stone-300 font-bold py-3 px-4 rounded-xl text-xs flex items-center justify-center gap-2 transition-all active:scale-[0.99]"
                            >
                              <ArrowUpDown size={14} className="text-stone-400" />
                              <span>{currentLanguage === 'ID' ? 'PILIH FILE TUGAS (EXCEL / CSV)' : 'SELECT TASKS FILE (EXCEL / CSV)'}</span>
                            </button>
                          </div>

                          <div>
                            <label className="text-[10px] font-bold text-stone-400 uppercase tracking-wider ml-1 mb-1.5 block">
                              {currentLanguage === 'ID' ? '2. UNGGAH FILE TIMELINE (UPDATES)' : '2. UPLOAD TIMELINE FILE'}
                            </label>
                            <input 
                              type="file" 
                              ref={updatesFileInputRef} 
                              onChange={handleUpdatesFileUpload} 
                              accept=".csv,.xlsx,.xls" 
                              className="hidden" 
                            />
                            <button
                              onClick={() => updatesFileInputRef.current?.click()}
                              className="w-full bg-stone-50 hover:bg-stone-100 text-stone-600 border border-dashed border-stone-300 font-bold py-3 px-4 rounded-xl text-xs flex items-center justify-center gap-2 transition-all active:scale-[0.99]"
                            >
                              <ArrowUpDown size={14} className="text-stone-400" />
                              <span>{currentLanguage === 'ID' ? 'PILIH FILE TIMELINE (EXCEL / CSV)' : 'SELECT TIMELINE FILE (EXCEL / CSV)'}</span>
                            </button>
                          </div>
                        </div>

                        {/* Messages Panel */}
                        {importSuccess && (
                          <div className="p-4 bg-emerald-50 border border-emerald-100 rounded-xl text-emerald-600 text-xs font-bold leading-relaxed flex items-start gap-2 animate-fade-in">
                            <CheckCircle2 size={15} className="mt-0.5 shrink-0" />
                            <div>
                              <p className="font-extrabold uppercase tracking-wide text-[10px] mb-0.5">
                                {currentLanguage === 'ID' ? 'Impor Berhasil' : 'Import Successful'}
                              </p>
                              <p className="text-[11px] opacity-90 font-medium">{importSuccess}</p>
                            </div>
                          </div>
                        )}

                        {importError && (
                          <div className="p-4 bg-rose-50 border border-rose-100 rounded-xl text-rose-600 text-xs font-bold leading-relaxed flex items-start gap-2 animate-fade-in">
                            <AlertCircle size={15} className="mt-0.5 shrink-0" />
                            <div>
                              <p className="font-extrabold uppercase tracking-wide text-[10px] mb-0.5">Import Failed</p>
                              <p className="text-[11px] opacity-90 font-medium">{importError}</p>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>

                    {currentUser.role === 'ADMIN' && (
                      <div className="bg-stone-100 border border-stone-200 p-6 rounded-2xl flex items-center justify-between">
                        <div>
                          <h3 className="font-bold text-secondary">Team Management</h3>
                          <p className="text-xs text-stone-500 font-medium">Add or remove members from your reporting line.</p>
                        </div>
                        <button onClick={() => setShowTeamModal(true)} className="bg-white border border-stone-200 hover:bg-stone-50 text-secondary p-3 rounded-xl shadow-sm transition-colors"><Users size={20} /></button>
                      </div>
                    )}

                    {/* AI Token Record Tracker Section */}
                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-stone-200 mb-6 mt-6">
                      <h3 className="font-black text-secondary mb-1 flex items-center gap-2 text-sm tracking-tight" style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                        <Sparkles size={18} className="text-[#0038FF]" /> 
                        {currentLanguage === 'ID' ? 'REKAMAN PENGGUNAAN TOKEN AI' : 'AI TOKEN USAGE TRACKER'}
                      </h3>
                      <p className="text-xs text-stone-500 font-medium mb-5 leading-relaxed">
                        {currentLanguage === 'ID' 
                          ? 'Daftar riwayat dan filter konsumsi token untuk kueri asisten AI (Task AI) Anda.'
                          : 'Monitor list logs and filter token consumption for your AI assistant integrations.'}
                      </p>

                      {/* Filter Bar */}
                      <div className="flex flex-wrap items-center justify-between gap-3 mb-6 bg-stone-50 p-3 rounded-xl border border-stone-100">
                        {/* Selector type */}
                        <div className="flex gap-1">
                          {[
                            { value: 'day', label: currentLanguage === 'ID' ? 'Hari' : 'Day' },
                            { value: 'weekly', label: currentLanguage === 'ID' ? 'Minggu' : 'Weekly' },
                            { value: 'monthly', label: currentLanguage === 'ID' ? 'Bulan' : 'Monthly' },
                            { value: 'year', label: currentLanguage === 'ID' ? 'Tahun' : 'Year' }
                          ].map(option => (
                            <button
                              key={option.value}
                              type="button"
                              onClick={() => setTokenFilterType(option.value as any)}
                              className={`px-3 py-1.5 text-[10px] font-extrabold uppercase tracking-widest rounded-lg transition-all ${
                                tokenFilterType === option.value
                                  ? 'bg-secondary text-white shadow-sm'
                                  : 'text-stone-400 hover:text-stone-700 bg-transparent'
                              }`}
                            >
                              {option.label}
                            </button>
                          ))}
                        </div>

                        {/* Combobox for days back */}
                        {tokenFilterType === 'day' && (
                          <div className="flex items-center gap-2">
                            <label className="text-[10px] font-black uppercase text-stone-400 tracking-wider">
                              {currentLanguage === 'ID' ? 'Pilih Hari:' : 'Select Day:'}
                            </label>
                            <select
                              value={selectedTokenDay}
                              onChange={e => setSelectedTokenDay(e.target.value)}
                              className="px-3 py-1.5 bg-white border border-stone-200 rounded-lg text-[11px] font-bold text-secondary focus:outline-none focus:ring-1 focus:ring-[#0038FF]/30"
                            >
                              {tokenDayOptions.map(opt => (
                                <option key={opt.value} value={opt.value}>
                                  {opt.label}
                                </option>
                              ))}
                            </select>
                          </div>
                        )}
                      </div>

                      {/* Tokens stats summary cards */}
                      <div className="grid grid-cols-3 gap-3 mb-6">
                        <div className="bg-[#EFF2FC] border border-[#0038FF]/10 p-3.5 rounded-xl text-center">
                          <p className="text-[9px] font-extrabold text-[#0038FF] uppercase tracking-widest mb-1">
                            {currentLanguage === 'ID' ? 'INPUT TOKEN' : 'INPUT TOKENS'}
                          </p>
                          <p className="font-mono text-base font-black text-secondary leading-none">
                            {tokenStats.inputs.toLocaleString()}
                          </p>
                        </div>
                        <div className="bg-[#EAFBF3] border border-emerald-500/10 p-3.5 rounded-xl text-center">
                          <p className="text-[9px] font-extrabold text-[#00C48C] uppercase tracking-widest mb-1">
                            {currentLanguage === 'ID' ? 'OUTPUT TOKEN' : 'OUTPUT TOKENS'}
                          </p>
                          <p className="font-mono text-base font-black text-secondary leading-none">
                            {tokenStats.outputs.toLocaleString()}
                          </p>
                        </div>
                        <div className="bg-stone-50 border border-stone-200/50 p-3.5 rounded-xl text-center">
                          <p className="text-[9px] font-extrabold text-stone-400 uppercase tracking-widest mb-1">
                            TOTAL
                          </p>
                          <p className="font-mono text-base font-black text-secondary leading-none">
                            {tokenStats.total.toLocaleString()}
                          </p>
                        </div>
                      </div>

                      {/* Token usage table list */}
                      <div className="overflow-x-auto border border-stone-150 rounded-xl">
                        <table className="w-full text-left border-collapse min-w-[340px]">
                          <thead>
                            <tr className="bg-stone-50 text-stone-400 font-extrabold text-[9px] uppercase tracking-wider border-b border-stone-150">
                              <th className="py-2.5 px-3">{currentLanguage === 'ID' ? 'Tanggal' : 'Date'}</th>
                              <th className="py-2.5 px-3">{currentLanguage === 'ID' ? 'Keterangan Penggunaan' : 'Token Used For'}</th>
                              <th className="py-2.5 px-3 text-right">{currentLanguage === 'ID' ? 'Input' : 'Input'}</th>
                              <th className="py-2.5 px-3 text-right">{currentLanguage === 'ID' ? 'Output' : 'Output'}</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-stone-100 text-xs text-stone-600">
                            {filteredTokenRecords.length === 0 ? (
                              <tr>
                                <td colSpan={4} className="py-6 text-center text-[11px] text-stone-400 font-bold bg-white">
                                  {currentLanguage === 'ID' ? 'Belum ada rekaman token AI dalam filter ini' : 'No AI token records found within this filter'}
                                </td>
                              </tr>
                            ) : (
                              filteredTokenRecords.map((record) => (
                                <tr key={record.id} className="hover:bg-stone-50/50 transition-colors">
                                  <td className="py-2.5 px-3 font-mono text-[10px] text-stone-500 font-semibold whitespace-nowrap">
                                    {record.date}
                                  </td>
                                  <td className="py-2.5 px-3 font-medium text-secondary truncate max-w-[150px]" title={record.usedFor}>
                                    {record.usedFor}
                                  </td>
                                  <td className="py-2.5 px-3 font-mono text-right text-stone-500 font-semibold">
                                    {record.inputTokens}
                                  </td>
                                  <td className="py-2.5 px-3 font-mono text-right text-stone-500 font-semibold">
                                    {record.outputTokens}
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            ) : activeTab === 'dashboard' && dashboardStats ? (
              <div className="relative w-full h-full overflow-hidden flex flex-col">
                <div id="dashboard-view-container" className={`w-full h-full flex flex-col transition-colors duration-200 ${
                  currentTheme === 'dark' ? 'bg-[#0B0F19]' : 'bg-slate-50'
                } animate-fade-in overflow-y-auto pb-36 md:pb-20 custom-scrollbar`}>
                {/* BEGIN: MainHeader */}
                <header className={`sticky top-0 z-30 px-5 py-3 shrink-0 transition-colors duration-200 border-b ${
                  currentTheme === 'dark' 
                    ? 'bg-[#111827] border-[#1F2937]' 
                    : 'bg-white border-slate-150'
                }`}>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <div className="w-2.5 h-2.5 bg-[#0038FF] rounded-full animate-pulse" />
                      <h1 className={`text-sm font-black uppercase tracking-widest transition-colors ${
                        currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                      }`} style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                        {TRANSLATIONS[currentLanguage].fruitfulDay}
                      </h1>
                    </div>
                    
                    <div className="flex items-center gap-3">
                      <span className={`text-[10px] font-bold px-2.5 py-1 rounded-lg border hidden sm:block ${
                        currentTheme === 'dark' 
                          ? 'text-slate-400 bg-slate-800/40 border-slate-700/60' 
                          : 'text-slate-400 bg-slate-50 border-slate-150'
                      }`}>
                        {new Date().toLocaleDateString(currentLanguage === 'ID' ? 'id-ID' : 'en-US', { weekday: 'short', month: 'short', day: 'numeric' })}
                      </span>

                      {/* Profile Dropdown Menu */}
                      <div className="relative">
                        <button 
                          onClick={() => setIsProfileDropdownOpen(!isProfileDropdownOpen)}
                          className={`flex items-center gap-1.5 p-1 rounded-full border transition-all ${
                            currentTheme === 'dark' 
                              ? 'bg-slate-800 border-slate-700 hover:bg-slate-750 text-white shadow-md' 
                              : 'bg-slate-50 border-slate-150 hover:bg-slate-100 text-slate-850 shadow-sm'
                          }`}
                        >
                          <div className="w-6 h-6 rounded-full bg-[#0038FF] text-white flex items-center justify-center text-xs font-black uppercase shadow-inner">
                            {currentUser.username[0].toUpperCase()}
                          </div>
                          <span className="text-[11px] font-extrabold pr-1 md:block hidden">{currentUser.username}</span>
                          <ChevronDown size={11} className={`opacity-60 transition-transform ${isProfileDropdownOpen ? 'rotate-180' : ''}`} />
                        </button>

                        {isProfileDropdownOpen && (
                          <>
                            <div className="fixed inset-0 z-40" onClick={() => setIsProfileDropdownOpen(false)} />
                            <div className={`absolute right-0 mt-2 w-52 rounded-2xl border p-3.5 shadow-xl z-50 animate-in fade-in slide-in-from-top-1 duration-150 ${
                              currentTheme === 'dark' 
                                ? 'bg-slate-900 border-slate-850 text-slate-205' 
                                : 'bg-white border-slate-150 text-slate-800'
                            }`} style={{ background: currentTheme === 'dark' ? '#1F2937' : '#FFFFFF', borderColor: currentTheme === 'dark' ? '#374151' : '#E2E8F0' }}>
                              <div className="border-b pb-2 mb-2 border-dashed border-slate-200/55 dark:border-slate-700/55">
                                <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">{currentLanguage === 'ID' ? 'MASUK SEBAGAI' : 'SIGNED IN AS'}</p>
                                <p className="text-xs font-black text-[#0038FF] truncate uppercase">{currentUser.username}</p>
                                <span className={`inline-block mt-0.5 text-[8.5px] font-extrabold px-1.5 py-0.5 rounded tracking-wide ${
                                  currentTheme === 'dark' ? 'bg-slate-950 text-slate-300' : 'bg-[#EFF2FC] text-[#0038FF]'
                                }`}>
                                  {currentUser.role}
                                </span>
                              </div>

                              {/* Language Submenu */}
                              <div className="mb-2.5">
                                <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                                  <Globe size={10} />
                                  {TRANSLATIONS[currentLanguage].language}
                                </p>
                                <div className={`grid grid-cols-2 gap-0.5 p-0.5 rounded-lg border ${
                                  currentTheme === 'dark' ? 'bg-slate-950 border-slate-800' : 'bg-slate-100 border-slate-150'
                                }`}>
                                  <button
                                    onClick={() => {
                                      setCurrentLanguage('EN');
                                      setIsProfileDropdownOpen(false);
                                    }}
                                    className={`py-1 text-[10px] font-black rounded-md transition-all ${
                                      currentLanguage === 'EN'
                                        ? 'bg-[#0038FF] text-white shadow'
                                        : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                    }`}
                                  >
                                    EN
                                  </button>
                                  <button
                                    onClick={() => {
                                      setCurrentLanguage('ID');
                                      setIsProfileDropdownOpen(false);
                                    }}
                                    className={`py-1 text-[10px] font-black rounded-md transition-all ${
                                      currentLanguage === 'ID'
                                        ? 'bg-[#0038FF] text-white shadow'
                                        : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                    }`}
                                  >
                                    ID
                                  </button>
                                </div>
                              </div>

                              {/* Theme Submenu */}
                              <div className="mb-2">
                                <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                                  {currentTheme === 'dark' ? <Moon size={10} /> : <Sun size={10} />}
                                  {TRANSLATIONS[currentLanguage].theme}
                                </p>
                                <div className={`grid grid-cols-2 gap-0.5 p-0.5 rounded-lg border ${
                                  currentTheme === 'dark' ? 'bg-slate-950 border-slate-800' : 'bg-slate-100 border-slate-150'
                                }`}>
                                  <button
                                    onClick={() => {
                                      setCurrentTheme('light');
                                      setIsProfileDropdownOpen(false);
                                    }}
                                    className={`py-1 text-[10px] font-black rounded-md transition-all flex items-center justify-center gap-1 ${
                                      currentTheme === 'light'
                                        ? 'bg-[#0038FF] text-white shadow'
                                        : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                    }`}
                                  >
                                    <Sun size={9} />
                                    <span>{TRANSLATIONS[currentLanguage].themeLight}</span>
                                  </button>
                                  <button
                                    onClick={() => {
                                      setCurrentTheme('dark');
                                      setIsProfileDropdownOpen(false);
                                    }}
                                    className={`py-1 text-[10px] font-black rounded-md transition-all flex items-center justify-center gap-1 ${
                                      currentTheme === 'dark'
                                        ? 'bg-[#0038FF] text-white shadow'
                                        : 'text-slate-400 hover:text-slate-600 dark:hover:text-slate-205'
                                    }`}
                                  >
                                    <Moon size={9} />
                                    <span>{TRANSLATIONS[currentLanguage].themeDark}</span>
                                  </button>
                                </div>
                              </div>

                              {/* Sign Out Trigger */}
                              <button
                                onClick={() => {
                                  handleLogout();
                                  setIsProfileDropdownOpen(false);
                                }}
                                className="w-full mt-2 pt-2 border-t border-slate-150 dark:border-slate-800 flex items-center justify-center gap-1 py-1.5 text-[10px] font-black text-rose-500 hover:bg-rose-50/50 dark:hover:bg-rose-950/20 rounded-lg transition-all"
                              >
                                <LogOut size={11} />
                                <span>{TRANSLATIONS[currentLanguage].signOut}</span>
                              </button>
                            </div>
                          </>
                        )}
                      </div>
                    </div>
                  </div>
                </header>
                {/* END: MainHeader */}

                <main className="p-4 space-y-6 flex-1 min-h-0">
                  {/* BEGIN: KPI Section */}
                  <section className="grid grid-cols-2 gap-4" data-purpose="kpi-metrics">
                    {/* Total Active Tasks */}
                    <div className={`rounded-2xl shadow-sm border flex flex-col justify-between p-3 h-24 hover:shadow transition-shadow ${
                      currentTheme === 'dark' 
                        ? 'bg-[#111827] border-[#1F2937]' 
                        : 'bg-white border-slate-100'
                    }`}>
                      <div className="flex justify-between items-start">
                        <div className="bg-blue-50 text-brand-primary rounded-lg p-1.5">
                          <LayoutList size={16} />
                        </div>
                        <span className="text-emerald-500 text-xs font-bold">+12%</span>
                      </div>
                      <div>
                        <p className={`font-bold text-xl leading-none ${
                          currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                        }`}>
                          {dashboardStats.total - dashboardStats.completed}
                        </p>
                        <p className={`text-xs font-medium mt-1 ${
                          currentTheme === 'dark' ? 'text-slate-400' : 'text-slate-500'
                        }`}>{TRANSLATIONS[currentLanguage].tasks}</p>
                      </div>
                    </div>

                    {/* Weekly Completion */}
                    <div className={`rounded-2xl shadow-sm border flex flex-col justify-between p-3 h-24 hover:shadow transition-shadow ${
                      currentTheme === 'dark' 
                        ? 'bg-[#111827] border-[#1F2937]' 
                        : 'bg-white border-slate-100'
                    }`}>
                      <div className="flex justify-between items-start">
                        <div className="bg-emerald-50 text-emerald-600 rounded-lg p-1.5">
                          <CheckCircle2 size={16} />
                        </div>
                        <span className="text-emerald-500 text-xs font-bold">
                          {dashboardStats.completionRate}%
                        </span>
                      </div>
                      <div>
                        <p className={`font-bold text-xl leading-none ${
                          currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                        }`}>
                          {dashboardStats.completed}
                        </p>
                        <p className={`text-xs font-medium mt-1 ${
                          currentTheme === 'dark' ? 'text-slate-400' : 'text-slate-500'
                        }`}>{TRANSLATIONS[currentLanguage].doneThisWeek}</p>
                      </div>
                    </div>

                    {/* Tag Usage */}
                    <section className={`rounded-2xl shadow-sm border p-4 col-span-2 ${
                      currentTheme === 'dark' 
                        ? 'bg-[#111827] border-[#1F2937]' 
                        : 'bg-white border-slate-100'
                    }`}>
                      <div className="flex items-center justify-between mb-4">
                        <h2 className={`text-md font-bold ${
                          currentTheme === 'dark' ? 'text-white' : 'text-slate-850'
                        }`}>PROJECTS</h2>
                        <div className={`flex p-1 rounded-lg ${
                          currentTheme === 'dark' ? 'bg-slate-950' : 'bg-slate-100'
                        }`}>
                          <button 
                            onClick={() => setTagPeriodFilter('week')}
                            className={`px-3 py-1 text-xs font-bold rounded-md transition-all ${
                              tagPeriodFilter === 'week'
                                ? 'bg-[#0038FF] text-white shadow-sm'
                                : currentTheme === 'dark'
                                  ? 'text-slate-400 hover:text-slate-200'
                                  : 'text-slate-500 hover:text-slate-700'
                            }`}
                          >
                            {TRANSLATIONS[currentLanguage].week}
                          </button>
                          <button 
                            onClick={() => setTagPeriodFilter('month')}
                            className={`px-3 py-1 text-xs font-bold rounded-md transition-all ${
                              tagPeriodFilter === 'month'
                                ? 'bg-[#0038FF] text-white shadow-sm'
                                : currentTheme === 'dark'
                                  ? 'text-slate-400 hover:text-slate-200'
                                  : 'text-slate-500 hover:text-slate-700'
                            }`}
                          >
                            {TRANSLATIONS[currentLanguage].month}
                          </button>
                          <button 
                            onClick={() => setTagPeriodFilter('year')}
                            className={`px-3 py-1 text-xs font-bold rounded-md transition-all ${
                              tagPeriodFilter === 'year'
                                ? 'bg-[#0038FF] text-white shadow-sm'
                                : currentTheme === 'dark'
                                  ? 'text-slate-400 hover:text-slate-200'
                                  : 'text-slate-500 hover:text-slate-700'
                            }`}
                          >
                            {TRANSLATIONS[currentLanguage].year}
                          </button>
                        </div>
                      </div>
                      <div className="space-y-4">
                        {[...CATEGORIES]
                          .map(cat => ({
                            name: cat,
                            info: dashboardStats.categoryBreakdown[cat] || { count: 0, percent: 0 }
                          }))
                          .sort((a, b) => b.info.count - a.info.count)
                          .map(({ name: cat, info: catInfo }) => {
                            return (
                              <div key={cat} className="space-y-1.5">
                                <div className="flex justify-between items-center text-xs font-bold">
                                  <span className={currentTheme === 'dark' ? 'text-slate-350' : 'text-slate-600'}>#{cat}</span>
                                  <span className={currentTheme === 'dark' ? 'text-slate-205' : 'text-slate-800'}>
                                    {catInfo.count} <span className="text-[10px] font-normal text-slate-400">{currentLanguage === 'ID' ? 'input' : (catInfo.count === 1 ? 'input' : 'inputs')}</span>
                                  </span>
                                </div>
                                <div className={`w-full h-2 rounded-full overflow-hidden ${
                                  currentTheme === 'dark' ? 'bg-slate-950' : 'bg-slate-100'
                                }`}>
                                  <div 
                                    className="bg-[#0038FF] h-full rounded-full transition-all duration-500 ease-out" 
                                    style={{ width: `${catInfo.percent}%` }}
                                  ></div>
                                </div>
                              </div>
                            );
                          })}
                      </div>
                    </section>

                    {/* Daily task counter */}
                    <div className="bg-[#0038FF] p-4 rounded-xl shadow-md flex flex-col justify-between h-52 col-span-2">
                       <div className="flex flex-col h-full justify-between pb-1">
                          <h3 className="text-white font-bold text-sm mt-0.5">{TRANSLATIONS[currentLanguage].dailyUpdateCounter}</h3>
                          <div className="flex items-end justify-between px-2 h-28">
                             {dailyUpdateCounts.map((item, idx) => {
                                const count = item.count;
                                const maxVal = Math.max(...dailyUpdateCounts.map(x => x.count), 5);
                                const hPct = Math.min(100, Math.max(10, (count / maxVal) * 100));
                                const barGradients = [
                                  'from-cyan-400 to-sky-500 shadow-cyan-500/10',
                                  'from-rose-400 to-rose-600 shadow-rose-500/10',
                                  'from-emerald-400 to-green-500 shadow-emerald-500/10',
                                  'from-amber-400 to-orange-500 shadow-amber-500/10',
                                  'from-indigo-400 to-violet-600 shadow-indigo-500/10',
                                  'from-fuchsia-400 to-pink-600 shadow-fuchsia-500/10',
                                  'from-teal-400 to-emerald-500 shadow-teal-500/10'
                                ];
                                return (
                                  <div key={idx} className="flex flex-col items-center gap-1.5 flex-1 relative group">
                                    <div className="w-5 bg-white/10 h-20 rounded-full flex items-end justify-center hover:bg-white/15 transition-all cursor-pointer overflow-hidden p-[1px]">
                                      <div 
                                        className={`w-full bg-gradient-to-t ${barGradients[idx]} rounded-full transition-all duration-500 ease-out`}
                                        style={{ height: `${hPct}%` }}
                                      ></div>
                                    </div>
                                    <span className="text-[10px] font-black text-white/70 uppercase tracking-wider">{item.label}</span>
                                    <span className="text-[10px] font-black text-white leading-none">{count}</span>
                                  </div>
                                );
                             })}
                          </div>
                       </div>
                    </div>
                  </section>
                  {/* END: KPI Section */}

                  {/* BEGIN: Quick Filters */}
                  <section className="flex items-center gap-3 overflow-x-auto pb-2 custom-scrollbar shrink-0" data-purpose="quick-actions">
                    {['ALL', 'ACTIVE', 'OVERDUE'].map((filterVal) => {
                      const isSelected = dashboardFilter === filterVal;
                      let label = filterVal === 'ALL' ? 'All Tasks' : filterVal.charAt(0) + filterVal.slice(1).toLowerCase();
                      if (currentLanguage === 'ID') {
                        if (filterVal === 'ALL') label = 'Semua Tugas';
                        else if (filterVal === 'ACTIVE') label = 'Aktif';
                        else if (filterVal === 'OVERDUE') label = 'Terlambat';
                      }
                      return (
                        <button 
                          key={filterVal}
                          onClick={() => setDashboardFilter(filterVal as any)}
                          className={`px-4 py-2 text-sm font-semibold rounded-full whitespace-nowrap transition-all ${
                            isSelected 
                              ? 'bg-[#0038FF] text-white shadow-sm' 
                              : currentTheme === 'dark'
                                ? 'bg-[#111827] text-slate-350 border border-slate-800 hover:bg-slate-800'
                                : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-50'
                          }`}
                        >
                          {label}
                        </button>
                      );
                    })}
                  </section>
                  {/* END: Quick Filters */}

                  {/* BEGIN: Task List Container */}
                  <section className="space-y-2.5 pb-24 md:pb-6" data-purpose="task-list">
                    <div className="flex items-center justify-between mb-2">
                      <h2 className={`text-lg font-bold ${
                        currentTheme === 'dark' ? 'text-white' : 'text-slate-808'
                      }`}>{TRANSLATIONS[currentLanguage].priorityTasks}</h2>
                    </div>

                    {dashboardFilteredTasks.length === 0 ? (
                      <div className={`rounded-2xl p-8 border text-center flex flex-col items-center justify-center ${
                        currentTheme === 'dark' 
                          ? 'bg-[#111827] border-[#1F2937] text-slate-500' 
                          : 'bg-white border-slate-150 text-slate-400'
                      }`}>
                        <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 ${
                          currentTheme === 'dark' ? 'bg-[#1E293B] text-slate-400' : 'bg-slate-50 text-slate-350'
                        }`}>
                          <CheckCircle2 size={24} />
                        </div>
                        <p className="text-sm font-semibold">{currentLanguage === 'ID' ? 'Tidak ada tugas yang cocok' : 'No tasks match this filter'}</p>
                        <p className="text-xs mt-1 text-slate-400">{currentLanguage === 'ID' ? 'Buat tugas baru untuk mengisi halaman ini.' : 'Create or update a task to populate this workspace.'}</p>
                      </div>
                    ) : (
                      dashboardFilteredTasks.map((task) => {
                        const daysRemaining = getDaysRemaining(task.targetDate);
                        const isOverdue = daysRemaining < 0 && task.status !== TaskStatus.DONE;
                        
                        let borderStyle = 'border-l-blue-500';
                        let badgeBg = 'bg-blue-50 text-blue-600';
                        let badgeLabel = task.status.replace('_', ' ');

                        if (isOverdue) {
                          borderStyle = 'border-l-rose-500';
                          badgeBg = 'bg-rose-50 text-rose-600';
                          badgeLabel = currentLanguage === 'ID' ? 'Terlambat' : 'Overdue';
                        } else if (task.status === TaskStatus.DONE) {
                          borderStyle = 'border-l-emerald-500';
                          badgeBg = 'bg-emerald-50 text-emerald-600';
                          badgeLabel = currentLanguage === 'ID' ? 'Selesai' : 'Done';
                        } else if (task.status === TaskStatus.FOLLOW_UP) {
                          borderStyle = 'border-l-amber-500';
                          badgeBg = 'bg-amber-50 text-amber-600';
                        } else if (task.status === TaskStatus.PENDING) {
                          borderStyle = 'border-l-orange-500';
                          badgeBg = 'bg-orange-50 text-orange-600';
                        }

                        // Translate status labels specifically for ID
                        if (currentLanguage === 'ID' && !isOverdue) {
                          if (task.status === TaskStatus.NOT_YET) badgeLabel = 'Belum Mulai';
                          else if (task.status === TaskStatus.PROGRESS) badgeLabel = 'Proses';
                          else if (task.status === TaskStatus.PENDING) badgeLabel = 'Tertunda';
                          else if (task.status === TaskStatus.FOLLOW_UP) badgeLabel = 'Tindak Lanjut';
                        }

                        return (
                          <div 
                            key={task.id} 
                            onClick={() => {
                              setSelectedTaskId(task.id);
                              setActiveTab('detail');
                            }}
                            className={`p-2.5 rounded-xl border-l-[3px] ${borderStyle} border hover:ring-1 hover:ring-[#0038FF]/20 cursor-pointer transition-all active:scale-[0.995] ${
                              currentTheme === 'dark' 
                                ? 'bg-[#111827] border-[#1F2937] hover:bg-slate-800' 
                                : 'bg-white border-slate-100 hover:ring-[#0038FF]/25'
                            }`}
                          >
                            <div className="flex justify-between items-center mb-1">
                              <span className={`px-1.5 py-0.5 ${badgeBg} text-[9px] font-extrabold rounded uppercase tracking-wider`}>
                                {badgeLabel}
                              </span>
                              <div className="flex items-center gap-1 shrink-0">
                                <span className="text-[9px] font-bold text-slate-400">#{task.category}</span>
                                <button 
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleToggleTaskStatus(task, e);
                                  }}
                                  className={`p-0.5 flex items-center justify-center shrink-0 transition-all active:scale-[0.98] hover:scale-105 ${
                                    task.status === TaskStatus.DONE 
                                      ? 'text-[#0038FF]' 
                                      : 'text-slate-300 hover:text-[#0038FF]'
                                  }`}
                                  title={task.status === TaskStatus.DONE ? TRANSLATIONS[currentLanguage].markIncomplete : TRANSLATIONS[currentLanguage].markComplete}
                                >
                                  <Check size={12} className="stroke-[3.5]" />
                                </button>
                              </div>
                            </div>
                            <h3 className={`font-bold text-sm leading-snug line-clamp-1 ${
                              currentTheme === 'dark' ? 'text-slate-105' : 'text-slate-800'
                            }`}>
                              {task.title}
                            </h3>
                            {task.tags && task.tags.length > 0 && (
                              <div className="flex flex-wrap gap-1 mt-1">
                                {task.tags.map(t => (
                                  <span key={t} className={`px-1 py-0.5 text-[8px] font-black border border-blue-100/30 rounded uppercase tracking-tight ${
                                    currentTheme === 'dark' 
                                      ? 'bg-blue-950/40 text-blue-400' 
                                      : 'bg-blue-50/70 text-[#0038FF]'
                                  }`}>
                                    #{t}
                                  </span>
                                ))}
                              </div>
                            )}
                            <div className="mt-2 pt-1.5 border-t border-slate-50 dark:border-slate-800/50 flex items-center justify-between">
                              <div className="flex items-center gap-3">
                                <div className="flex items-center text-slate-400 gap-1">
                                  <Calendar size={11} />
                                  <span className="text-[10px] font-semibold">{formatListDate(task.targetDate)}</span>
                                </div>
                                <div className={`text-[10px] font-bold ${isOverdue ? 'text-rose-600' : 'text-slate-500'}`}>
                                  {getDaysLabelLocalized(daysRemaining, currentLanguage)}
                                </div>
                              </div>
                              <div className="flex -space-x-1">
                                <div className="w-5 h-5 rounded-full border border-white bg-[#0038FF] flex items-center justify-center text-[7px] text-white font-extrabold uppercase shadow-sm">
                                  {userMap[task.userId]?.username.substring(0, 1).toUpperCase() || 'U'}
                                </div>
                              </div>
                            </div>
                          </div>
                        );
                      })
                    )}
                  </section>
                  {/* END: Task List Container */}
                </main>
              </div>
              
              {/* Dynamic FAB Floating Action Button for Dashboard */}
              <button 
                onClick={createNewTask}
                className="absolute bottom-28 md:bottom-6 right-6 w-14 h-14 bg-[#0038FF] hover:bg-blue-700 text-white rounded-full flex items-center justify-center shadow-xl shadow-blue-600/30 active:scale-95 hover:scale-105 transition-all z-40"
                title={currentLanguage === 'ID' ? 'Tambah Tugas Baru' : 'Create New Task'}
              >
                <Plus size={28} />
              </button>
            </div>
          ) : selectedTask ? (
              <>
                <div className="flex-[3] md:flex-1 md:w-1/2 p-3 md:p-6 lg:p-8 flex flex-col bg-white border-b md:border-b-0 md:border-r border-stone-200 md:overflow-y-auto animate-fade-in relative md:min-h-0">
                  {/* BEGIN: Detail View Header */}
                  <header className={`sticky top-0 z-30 px-5 py-3 -mx-3 md:-mx-6 lg:-mx-8 border-b mb-4 transition-colors duration-200 ${
                    currentTheme === 'dark' 
                      ? 'bg-[#111827] border-[#1F2937]' 
                      : 'bg-white border-slate-150'
                  }`}>
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <div className="w-2.5 h-2.5 bg-[#0038FF] rounded-full animate-pulse" />
                        <h1 className={`text-sm font-black uppercase tracking-widest transition-colors ${
                          currentTheme === 'dark' ? 'text-white' : 'text-slate-800'
                        }`} style={{ fontFamily: '"Outfit", "Space Grotesk", sans-serif' }}>
                          {currentLanguage === 'ID' ? 'DETAIL TUGAS' : 'TASK MASTER DETAIL'}
                        </h1>
                      </div>
                      <button 
                        onClick={() => setActiveTab('overview')} 
                        className="text-[10px] font-black uppercase tracking-wider text-[#0038FF] bg-[#EFF2FC] hover:bg-[#0038FF]/10 px-3 py-1.5 rounded-full transition-all md:hidden flex items-center gap-1 shrink-0 shadow-sm"
                      >
                        ← Back
                      </button>
                    </div>
                  </header>
                  {/* END: Detail View Header */}

                  <div className="flex flex-col mb-3 space-y-2 shrink-0">
                    <div className="flex items-center justify-between gap-2">
                      <input 
                        type="text" 
                        value={selectedTask.title} 
                        onChange={(e) => handleTaskChange('title', e.target.value)} 
                        className="text-lg md:text-2xl font-bold text-secondary focus:outline-none focus:underline decoration-primary decoration-2 underline-offset-4 bg-transparent flex-1 min-w-0"
                      />
                      <button 
                        onClick={handleDeleteTask} 
                        className="p-2 bg-stone-100 hover:bg-red-50 text-stone-400 hover:text-red-600 rounded-lg transition-colors shrink-0" 
                        title="Delete Task"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                    
                    {/* Row 1: Category, Status, +7, +Fri */}
                    <div className="flex flex-wrap sm:flex-nowrap gap-1.5 text-sm">
                      <div className="flex-1 min-w-[80px] relative">
                        <select 
                          value={selectedTask.category} 
                          onChange={(e) => handleTaskChange('category', e.target.value)} 
                          className="w-full font-bold text-stone-600 bg-stone-50 border border-stone-200 rounded-lg px-2 py-2 hover:border-primary focus:border-primary outline-none transition-all appearance-none text-sm"
                        >
                          {allCategories.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                        <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
                      </div>
                      
                      <div className="flex-1 min-w-[80px] relative">
                        <select 
                          value={selectedTask.status} 
                          onChange={(e) => handleTaskChange('status', e.target.value)} 
                          className="w-full font-bold text-stone-600 bg-stone-50 border border-stone-200 rounded-lg px-2 py-2 hover:border-primary focus:border-primary outline-none transition-all appearance-none text-sm"
                        >
                          {STATUSES.map(s => <option key={s} value={s}>{s.replace('_', ' ')}</option>)}
                        </select>
                        <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
                      </div>

                      <div className="flex gap-1 shrink-0 text-sm">
                        <button 
                          onClick={() => { const d = new Date(); d.setDate(d.getDate() + 7); handleTaskChange('targetDate', getLocalDate(d)); }} 
                          className="px-2.5 py-2 bg-white hover:bg-stone-50 text-stone-700 font-bold rounded-lg border border-stone-200 transition-colors shadow-sm text-sm"
                          title="+7 Days"
                        >
                          +7d
                        </button>
                        <button 
                          onClick={() => { const d = new Date(); const day = d.getDay(); const diff = (5 - day + 7) % 7; const add = diff === 0 ? 7 : diff; d.setDate(d.getDate() + add); handleTaskChange('targetDate', getLocalDate(d)); }} 
                          className="px-2.5 py-2 bg-white hover:bg-stone-50 text-stone-700 font-bold rounded-lg border border-stone-200 transition-colors shadow-sm text-sm"
                          title="Next Friday"
                        >
                          +Fri
                        </button>
                      </div>
                    </div>
                    
                    {/* Row 2: Target Date & Update Date */}
                    <div className="grid grid-cols-2 gap-2">
                      <div className="flex items-center gap-1.5 bg-stone-50 border border-stone-200 rounded-lg px-2 py-2 hover:border-primary transition-colors group">
                        <span className="text-[10px] font-bold text-stone-400 uppercase tracking-tighter shrink-0">Target:</span>
                        <input 
                          type="date" 
                          value={selectedTask.targetDate} 
                          onChange={(e) => handleTaskChange('targetDate', e.target.value)} 
                          className="w-full text-sm font-mono text-stone-600 bg-transparent outline-none cursor-pointer"
                        />
                      </div>

                      <div className="flex items-center gap-1.5 bg-accent/5 border border-stone-100 rounded-lg px-2 py-2 hover:border-primary transition-colors group">
                        <span className="text-[10px] font-bold text-stone-400 uppercase tracking-tighter shrink-0">Log At:</span>
                        <input 
                          type="date" 
                          value={updateDate} 
                          onChange={(e) => setUpdateDate(e.target.value)} 
                          className="w-full text-sm font-mono font-bold text-secondary bg-transparent outline-none cursor-pointer"
                        />
                      </div>
                    </div>

                    {editUpdateId && (
                      <div className="flex items-center justify-between bg-amber-50 px-3 py-1.5 rounded-lg border border-amber-100 mt-1.5 text-xs">
                        <span className="text-[10px] text-amber-700 font-bold uppercase tracking-wider">Editing Existing Entry</span>
                        <button onClick={handleCancelEdit} className="text-amber-500 hover:bg-amber-100 p-0.5 rounded-md transition-colors"><X size={14}/></button>
                      </div>
                    )}
                  </div>

                  <div className="flex flex-col">
                    <textarea 
                      className="w-full p-4 bg-stone-50 border border-stone-200 rounded-2xl resize-none focus:outline-none focus:ring-1 focus:ring-primary/20 focus:bg-white transition-all text-sm leading-relaxed text-secondary placeholder:text-stone-300 shadow-inner h-32 md:h-auto md:flex-1" 
                      placeholder="What happened today?..." 
                      rows={5}
                      value={updateContent} 
                      onChange={e => setUpdateContent(e.target.value)}
                    />

                    {/* Hashtag Field - placed below the textarea */}
                    <div className="flex flex-col gap-1.5 mt-2 bg-stone-100/50 border border-stone-200 rounded-2xl p-2.5">
                      {/* Existing tags of the selected tasks */}
                      {selectedTask.tags && selectedTask.tags.length > 0 && (
                        <div className="flex flex-wrap gap-1 px-1">
                          {selectedTask.tags.map(t => (
                            <span 
                              key={t}
                              className="inline-flex items-center gap-0.5 px-2 py-0.5 rounded-full text-xs font-extrabold bg-blue-50 text-[#0038FF] border border-blue-100 hover:bg-rose-50 hover:text-rose-600 hover:border-rose-100 cursor-pointer transition-all uppercase tracking-tight"
                              title="Click to remove tag"
                              onClick={() => {
                                const newTags = (selectedTask.tags || []).filter(existing => existing !== t);
                                handleTaskChange('tags', newTags);
                              }}
                            >
                              #{t}
                              <X size={10} className="stroke-[3]" />
                            </span>
                          ))}
                        </div>
                      )}

                      {/* Input with Autocomplete / Suggested master tags */}
                      <div className="flex gap-1.5 pl-1 pr-1">
                        <div className="flex-1 relative">
                          <input 
                            type="text" 
                            placeholder="Type to add tag..." 
                            value={newTagInput}
                            onChange={(e) => {
                              // Force tag-compliant letters, numbers, underscores or hyphens
                              setNewTagInput(e.target.value.toLowerCase().replace(/[^a-z0-9_\s-]/g, ''));
                            }}
                            onKeyDown={async (e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                const cleanTag = newTagInput.trim().toLowerCase().replace(/[^a-z0-9_]/g, '');
                                if (!cleanTag) return;
                                setNewTagInput('');
                                
                                // 1. Add to tagMaster in Firestore if not existing
                                if (!tagMaster.some(tag => tag.name === cleanTag)) {
                                  const newTagObj = { id: cleanTag, name: cleanTag };
                                  await saveSingleTag(newTagObj);
                                }
                                
                                // 2. Add to selectedTask tags if not existing
                                const currentTags = selectedTask.tags || [];
                                if (!currentTags.includes(cleanTag)) {
                                  handleTaskChange('tags', [...currentTags, cleanTag]);
                                }
                              }
                            }}
                            className="w-full text-sm font-semibold text-slate-700 bg-white border border-stone-200 rounded-lg px-2.5 py-1.5 focus:border-[#0038FF] focus:outline-none transition-all placeholder:text-stone-300"
                            list="master-tags-list"
                          />
                          <datalist id="master-tags-list">
                            {tagMaster.map(t => (
                              <option key={t.id} value={t.name} />
                            ))}
                          </datalist>
                        </div>
                        <button 
                          type="button"
                          disabled={!newTagInput.trim()}
                          onClick={async () => {
                            const cleanTag = newTagInput.trim().toLowerCase().replace(/[^a-z0-9_]/g, '');
                            if (!cleanTag) return;
                            setNewTagInput('');
                            
                            // 1. Add to tagMaster in Firestore if not existing
                            if (!tagMaster.some(tag => tag.name === cleanTag)) {
                              const newTagObj = { id: cleanTag, name: cleanTag };
                              await saveSingleTag(newTagObj);
                            }
                            
                            // 2. Add to selectedTask tags if not existing
                            const currentTags = selectedTask.tags || [];
                            if (!currentTags.includes(cleanTag)) {
                              handleTaskChange('tags', [...currentTags, cleanTag]);
                            }
                          }}
                          className="px-3 bg-[#0038FF] hover:bg-blue-700 disabled:bg-stone-200 text-white text-sm font-bold rounded-lg transition-all shadow-sm active:scale-95 flex items-center justify-center shrink-0"
                        >
                          <Plus size={14} className="stroke-[3]" />
                        </button>
                      </div>
                    </div>
                  </div>

                  <div className="mt-2 flex flex-col shrink-0">
                    <button 
                      onClick={handleSaveUpdate} 
                      className="w-full bg-[#0038FF] hover:bg-blue-700 disabled:bg-slate-350 text-white py-3 rounded-xl text-sm font-black uppercase tracking-wider flex items-center justify-center gap-2 transition-all active:scale-[0.98] shadow-md shadow-blue-600/15" 
                      disabled={isSaving}
                    >
                      {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
                      {editUpdateId ? 'UPDATE LOG' : 'SAVE UPDATE'}
                    </button>
                  </div>
                </div>

                <div className="flex-[2] md:flex-1 md:w-1/2 p-4 md:p-8 bg-stone-50 flex flex-col md:overflow-hidden min-h-[60vh] md:min-h-0 animate-fade-in relative">
                  <div className="flex justify-between items-center mb-4 shrink-0">
                    <h2 className="text-lg font-bold text-secondary flex items-center gap-2">
                      History 
                      <span className="bg-stone-200 text-stone-600 text-[10px] px-2 py-0.5 rounded-full font-bold">{taskHistory.length}</span>
                    </h2>
                    <button onClick={() => setHistorySort(prev => prev === 'NEWEST' ? 'OLDEST' : 'NEWEST')} className="px-3 py-1.5 text-[10px] font-bold rounded-lg bg-white border border-stone-200 hover:bg-stone-50 text-stone-600 flex items-center gap-1.5 transition-colors shadow-sm">
                      {historySort === 'OLDEST' ? <ArrowUp size={12}/> : <ArrowDown size={12}/>}
                      {historySort === 'OLDEST' ? 'OLDEST FIRST' : 'NEWEST FIRST'}
                    </button>
                  </div>
                  <div className="flex-1 md:overflow-y-auto pr-1 -mr-1 pb-36 md:pb-20">
                    {taskHistory.length === 0 ? (
                      <div className="flex flex-col items-center justify-center py-12 text-stone-400">
                        <div className="w-12 h-12 bg-stone-100 rounded-full flex items-center justify-center mb-3">
                          <AlertCircle size={20} className="opacity-30" />
                        </div>
                        <p className="text-sm">No updates recorded yet.</p>
                      </div>
                    ) : (
                      <div className="space-y-4">
                        {taskHistory.map(update => (
                          <SwipeableCard 
                            key={update.id} 
                            update={update} 
                            onEdit={handleEditUpdateStart} 
                            onCalendar={handleAddToCalendar} 
                            onDelete={handleDeleteUpdate}
                          />
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </>
            ) : (
              <div className="w-full h-full flex flex-col items-center justify-center text-stone-400 bg-stone-50/50 p-6 text-center animate-fade-in"><CheckCircle2 size={48} className="mb-4 opacity-20 text-primary" /><p className="text-lg font-bold text-stone-500">No Task Selected</p><p className="text-sm opacity-60 mt-2">Select a task from the list to view details.</p><button onClick={() => setActiveTab('overview')} className="mt-4 px-6 py-2 bg-white border border-stone-200 rounded-full text-sm font-bold text-primary shadow-sm md:hidden">Go to Overview</button></div>
            )}
          </div>
      </div>

      {/* BEGIN: Raised Floating Capsule Tab Bar */}
      <div className="fixed bottom-5 left-4 right-4 bg-white/95 backdrop-blur-md border border-slate-200/80 rounded-[28px] py-2 px-3 shadow-xl z-50 flex items-center justify-around md:hidden">
        {/* Tab 1: Dashboard */}
        <button 
          onClick={() => setActiveTab('dashboard')} 
          className={`flex-1 flex flex-col items-center justify-center transition-all ${
            activeTab === 'dashboard' 
              ? 'relative -translate-y-1' 
              : 'text-slate-400 hover:text-slate-600'
          }`}
        >
          {activeTab === 'dashboard' ? (
            <div className="w-11 h-11 rounded-full bg-[#0038FF] text-white flex items-center justify-center shadow-lg shadow-blue-600/35">
              <LayoutDashboard size={20} />
            </div>
          ) : (
            <div className="p-2">
              <LayoutDashboard size={22} className="opacity-80" />
            </div>
          )}
        </button>

        {/* Tab 2: Overview */}
        <button 
          onClick={() => setActiveTab('overview')} 
          className={`flex-1 flex flex-col items-center justify-center transition-all ${
            activeTab === 'overview' 
              ? 'relative -translate-y-1' 
              : 'text-slate-400 hover:text-slate-600'
          }`}
        >
          {activeTab === 'overview' ? (
            <div className="w-11 h-11 rounded-full bg-[#0038FF] text-white flex items-center justify-center shadow-lg shadow-blue-600/35">
              <LayoutList size={20} />
            </div>
          ) : (
            <div className="p-2">
              <LayoutList size={22} className="opacity-80" />
            </div>
          )}
        </button>

        {/* Tab 3: Detail */}
        <button 
          onClick={() => setActiveTab('detail')} 
          className={`flex-1 flex flex-col items-center justify-center transition-all ${
            activeTab === 'detail' 
              ? 'relative -translate-y-1' 
              : 'text-slate-400 hover:text-slate-600'
          }`}
        >
          {activeTab === 'detail' ? (
            <div className="w-11 h-11 rounded-full bg-[#0038FF] text-white flex items-center justify-center shadow-lg shadow-blue-600/35">
              <FileText size={20} />
            </div>
          ) : (
            <div className="p-2">
              <FileText size={22} className="opacity-80" />
            </div>
          )}
        </button>

        {/* Tab 4: Updates */}
        <button 
          onClick={() => setActiveTab('updates')} 
          className={`flex-1 flex flex-col items-center justify-center transition-all ${
            activeTab === 'updates' 
              ? 'relative -translate-y-1' 
              : 'text-slate-400 hover:text-slate-600'
          }`}
        >
          {activeTab === 'updates' ? (
            <div className="w-11 h-11 rounded-full bg-[#0038FF] text-white flex items-center justify-center shadow-lg shadow-blue-600/35">
              <History size={20} />
            </div>
          ) : (
            <div className="p-2">
              <History size={22} className="opacity-80" />
            </div>
          )}
        </button>

        {/* Tab: AI */}
        <button 
          onClick={() => setActiveTab('ai')} 
          className={`flex-1 flex flex-col items-center justify-center transition-all ${
            activeTab === 'ai' 
              ? 'relative -translate-y-1' 
              : 'text-slate-400 hover:text-slate-600'
          }`}
        >
          {activeTab === 'ai' ? (
            <div className="w-11 h-11 rounded-full bg-[#0038FF] text-white flex items-center justify-center shadow-lg shadow-blue-600/35 animate-pulse">
              <Sparkles size={20} />
            </div>
          ) : (
            <div className="p-2">
              <Sparkles size={22} className="text-[#0038FF] opacity-90 animate-bounce-slow" />
            </div>
          )}
        </button>

        {/* Tab 5: Profile */}
        <button 
          onClick={() => setActiveTab('profile')} 
          className={`flex-1 flex flex-col items-center justify-center transition-all ${
            activeTab === 'profile' 
              ? 'relative -translate-y-1' 
              : 'text-slate-400 hover:text-slate-600'
          }`}
        >
          {activeTab === 'profile' ? (
            <div className="w-11 h-11 rounded-full bg-[#0038FF] text-white flex items-center justify-center shadow-lg shadow-blue-600/35">
              <UserIcon size={20} />
            </div>
          ) : (
            <div className="p-2">
              <UserIcon size={22} className="opacity-80" />
            </div>
          )}
        </button>
      </div>
    </div>
  );
};

export default App;
