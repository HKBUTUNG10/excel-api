import React, { useState, useMemo, useEffect, useRef } from 'react';
import ExcelJS from 'exceljs';
import { toPng } from 'html-to-image';
import Login from './Login';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line
} from 'recharts';
import { 
  Mail,
  ShieldAlert, 
  Search, 
  Filter, 
  CheckCircle, 
  XCircle, 
  AlertTriangle,
  FileSpreadsheet,
  PlusCircle,
  Clock,
  User,
  Calendar,
  RefreshCw,
  Link as LinkIcon,
  Database,
  LayoutDashboard,
  Settings,
  Menu,
  Bell,
  ChevronRight,
  LogOut,
  X,
  Download,
  Smartphone,
  Lock,
  Eye,
  EyeOff,
  ShieldCheck,
  Trash2,
  SquareCheck,
  CheckSquare,
  BarChart2,
  TrendingUp,
  Users,
  Cloud,
  Wifi
} from 'lucide-react';

// Firebase Imports for Cloud Synchronization across all devices/teachers
import { initializeApp } from 'firebase/app';
import { 
  getAuth, 
  signInWithCustomToken, 
  signInAnonymously, 
  onAuthStateChanged 
} from 'firebase/auth';
import { 
  getFirestore, 
  collection, 
  doc, 
  setDoc, 
  deleteDoc, 
  onSnapshot,
  query,
  where
} from 'firebase/firestore';

// Initialize Firebase Service
const firebaseConfig = {
  apiKey: "AIzaSyBZWTL8zHzyXThABjgtcSsE-YIz-C1ojYo",
  authDomain: "wkc-system-discipline.firebaseapp.com",
  projectId: "wkc-system-discipline",
  storageBucket: "wkc-system-discipline.firebasestorage.app",
  messagingSenderId: "917375802225",
  appId: "1:917375802225:web:84ba777aea965c491634d0",
  measurementId: "G-BJYEPHSTLS"
};
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const appId = typeof __app_id !== 'undefined' ? __app_id : 'wkc-discipline-system';

// Utility to format dates for HK locale
const formatDate = (dateStr) => {
  const safeDate = dateStr.includes('T') ? dateStr : `${dateStr}T12:00:00`;
  return new Date(safeDate).toLocaleDateString('zh-HK', {
    year: 'numeric', month: 'short', day: 'numeric'
  });
};

const formatTime = (timestamp) => {
  return new Date(timestamp).toLocaleTimeString('zh-HK', {
    hour: '2-digit', minute: '2-digit', hour12: false
  });
};

// CSV Parser utility to handle Google Sheets CSV output safely
const parseCSV = (str) => {
  const arr = [];
  let quote = false;
  let row = 0, col = 0;
  for (let c = 0; c < str.length; c++) {
    let cc = str[c], nc = str[c+1];
    arr[row] = arr[row] || [];
    arr[row][col] = arr[row][col] || '';
    if (cc === '"' && quote && nc === '"') { arr[row][col] += cc; ++c; continue; }
    if (cc === '"') { quote = !quote; continue; }
    if (cc === ',' && !quote) { ++col; continue; }
    if (cc === '\r' && nc === '\n' && !quote) { ++row; col = 0; ++c; continue; }
    if (cc === '\n' && !quote) { ++row; col = 0; continue; }
    if (cc === '\r' && !quote) { ++row; col = 0; continue; }
    arr[row][col] += cc;
  }
  return arr;
};

// Accurately generate YYYY-MM-DD in local timezone
const getLocalDateStr = (date) => {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const getTodayStr = () => getLocalDateStr(new Date());
const getYesterdayStr = () => {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return getLocalDateStr(d);
};

const CLASSES = ['1A', '1B', '1C', '2A', '2B', '2C', '3A', '3B', '3C', '4A', '4B', '4C', '5A', '5B', '5C', '6A', '6B', '6C'];
const RULES = ['欠帶學生證', '欠帶課本', '欠帶體育服', '欠帶文具', '欠交功課', '欠交通告', '其他欠帶事項'];

const CHART_COLORS = ['#3b82f6', '#8b5cf6', '#ec4899', '#f43f5e', '#f97316', '#eab308', '#10b981'];

const exportToCSV = (data) => {
  const headers = ['記錄ID', '學生編號', '學生姓名', '班別', '欠帶事項', '事發日期', '提交時間', '報告教職員', '系統狀態', '備註原因', '跟進進度'];
  const csvRows = [headers.join(',')];

  data.forEach(row => {
    const submitDateStr = new Date(row.submitTime).toLocaleString('zh-HK', { hour12: false });
    const values = [
      row.id,
      row.studentId,
      row.studentName,
      row.studentClass,
      row.rule,
      row.incidentDate,
      submitDateStr,
      row.teacher,
      row.processStatus === 'Accepted' ? '已接納' : '已拒絕',
      row.reason,
      row.isHandled ? '已處理' : '未處理'
    ].map(val => `"${String(val).replace(/"/g, '""')}"`);
    csvRows.push(values.join(','));
  });

  const csvData = "\uFEFF" + csvRows.join('\n'); 
  const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', `訓導記錄_${getTodayStr()}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

// 匯出留堂通知名單 (給老師的 Excel/CSV 格式)
const exportTeacherNotice = (rosterData) => {
  const headers = ['班別', '學生姓名', '學號', '累積有效違規', '觸發留堂次數', '詳細違規紀錄 (日期 / 事項 / 報告老師)'];
  const csvRows = [headers.join(',')];

  rosterData.forEach(student => {
    // 將該學生的所有違規詳情串接成一個易讀的字串
    const details = student.records
      .map(r => `[${r.incidentDate}] ${r.rule} (${r.teacher})`)
      .join(' \n'); // 使用換行符號分隔，在 Excel 中會自動換行顯示

    const values = [
      student.class,
      student.name,
      student.id,
      student.totalRecords,
      student.milestonesReached,
      details
    ].map(val => `"${String(val).replace(/"/g, '""')}"`);
    
    csvRows.push(values.join(','));
  });

  // 加上 BOM 讓 Excel 正確讀取中文字符
  const csvData = "\uFEFF" + csvRows.join('\n'); 
  const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', `訓導處_留堂與老師通知名單_${getTodayStr()}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const INITIAL_MOCK_DATA = [
  { id: 'R001', studentId: 'S406', studentName: '陳大文', studentClass: '4A', rule: '欠帶學生證', incidentDate: getTodayStr(), submitTime: new Date(new Date().setHours(8, 30)).getTime(), teacher: '陳老師' },
  { id: 'R002', studentId: 'S406', studentName: '陳大文', studentClass: '4A', rule: '欠帶課本', incidentDate: getTodayStr(), submitTime: new Date(new Date().setHours(9, 15)).getTime(), teacher: '李老師' },
  { id: 'R003', studentId: 'S406', studentName: '陳大文', studentClass: '4A', rule: '欠交功課', incidentDate: getTodayStr(), submitTime: new Date(new Date().setHours(11, 45)).getTime(), teacher: '張老師' },
  { id: 'R004', studentId: 'S512', studentName: '黃小玲', studentClass: '5B', rule: '欠帶體育服', incidentDate: getTodayStr(), submitTime: new Date(new Date().setHours(10, 0)).getTime(), teacher: '何老師' },
  { id: 'R005', studentId: 'S208', studentName: '張子健', studentClass: '2C', rule: '欠帶學生證', incidentDate: getYesterdayStr(), submitTime: new Date(new Date().setHours(8, 25) - 86400000).getTime(), teacher: '陳老師' },
];

/**
 * 核心數據處理邏輯：
 * 1. 按提交時間 (submitTime) 由舊到新排序，確保先登記的老師獲得接納。
 * 2. 同一學生（依「學號 + 姓名」）每日最多只接受前 2 宗記錄。
 * 3. 相同學生、相同日期、相同事項視為「重複記錄 (已忽略)」。
 * 4. 最終依「日期 (由舊至新/順序)」與「班別 (1A 到 6C)」排列。
 */
const processRecords = (rawRecords, startDate, endDate) => {
  const activeStartDate = startDate || '2000-01-01';
  const activeEndDate = endDate || '2100-01-01';
  
  const recentRecords = rawRecords.filter(record => {
    const safeIncidentDate = String(record.incidentDate || '').trim();
    return safeIncidentDate >= activeStartDate && safeIncidentDate <= activeEndDate;
  });

  // 第一階段：按提交時間由舊到新排序 (Earliest First)，確保先提交者優先算入配額
  const sortedBySubmit = [...recentRecords].sort((a, b) => {
    if (a.submitTime === b.submitTime) {
      return String(a.id).localeCompare(String(b.id));
    }
    return a.submitTime - b.submitTime;
  });
  
  const dailyCounts = {}; 
  const seenExactDuplicates = new Set();

  const processed = sortedBySubmit.map(record => {
    const safeStudentId = String(record.studentId || '').trim().toUpperCase();
    const safeStudentName = String(record.studentName || '').trim();
    const safeDate = String(record.incidentDate || '').trim();
    const safeRule = String(record.rule || '').trim();

    // 唯一學生辨識 (學號 + 姓名 + 日期)
    const studentQuotaKey = `${safeStudentId}_${safeStudentName}_${safeDate}`;
    const exactDupKey = `${studentQuotaKey}_${safeRule}`;

    // 檢查完全重複項目
    if (seenExactDuplicates.has(exactDupKey)) {
      return { ...record, processStatus: 'Rejected', reason: '重複記錄 (已忽略)' };
    }
    seenExactDuplicates.add(exactDupKey);

    if (!dailyCounts[studentQuotaKey]) dailyCounts[studentQuotaKey] = 0;
    dailyCounts[studentQuotaKey]++;

    // 每位學生每日上限最多 2 次 (前 2 位提交的老師紀錄會獲接納)
    if (dailyCounts[studentQuotaKey] > 2) {
      return { ...record, processStatus: 'Rejected', reason: '超出每日上限 (最多2次)' };
    } else {
      return { ...record, processStatus: 'Accepted', reason: '有效' };
    }
  });

  // 第二階段：依「事發日期 (順序/舊至新)」及「班別 (1A 到 6C 順序)」進行排序
  return processed.sort((a, b) => {
    if (a.incidentDate !== b.incidentDate) {
      return String(a.incidentDate).localeCompare(String(b.incidentDate), 'zh-HK');
    }
    if (a.studentClass !== b.studentClass) {
      return String(a.studentClass).localeCompare(String(b.studentClass), 'en', { numeric: true });
    }
    return a.submitTime - b.submitTime;
  });
};

const StatCard = ({ title, value, subtitle, icon: Icon, colorClass, bgClass, trend }) => (
  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col justify-between transition-all hover:shadow-md">
    <div className="flex justify-between items-start mb-4">
      <div className={`p-3 rounded-xl ${bgClass}`}>
        <Icon className={`w-5 h-5 ${colorClass}`} />
      </div>
      {trend && (
        <span className={`text-xs font-medium px-2 py-1 rounded-full ${trend > 0 ? 'bg-emerald-50 text-emerald-600' : 'bg-slate-50 text-slate-600'}`}>
          今日 {trend > 0 ? '+' : ''}{trend} 宗
        </span>
      )}
    </div>
    <div>
      <h3 className="text-3xl font-bold text-slate-800 mb-1">{value}</h3>
      <p className="text-sm font-medium text-slate-500">{title}</p>
      {subtitle && <p className="text-xs text-slate-400 mt-1">{subtitle}</p>}
    </div>
  </div>
);

const StatusBadge = ({ status, reason }) => {
  if (status === 'Accepted') {
    return (
      <span className="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
        <CheckCircle className="w-3.5 h-3.5 mr-1.5" /> 已接納
      </span>
    );
  }
  return (
    <div className="flex flex-col items-start gap-1">
      <span className="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-rose-50 text-rose-700 border border-rose-200">
        <XCircle className="w-3.5 h-3.5 mr-1.5" /> 已拒絕
      </span>
      {reason && <span className="text-[10px] text-rose-500 font-medium ml-1">{reason}</span>}
    </div>
  );
};

// Helper function to detect student emails matching s20260035@wkc.edu.hk pattern
const formatStudentEmail = (email) => {
  if (!email) return '';
  const trimmed = email.trim().toLowerCase();
  
  // Regex matches: 's' + 4-digit year + 4-digit ID + '@wkc.edu.hk'
  const isStudentEmail = /^s\d{8}@wkc\.edu\.hk$/.test(trimmed);
  
  return isStudentEmail ? `${email} (possible student)` : email;
};

export default function App() {
  const [user, setUser] = useState(null);
  const [userRole, setUserRole] = useState(null);
  const [rawRecords, setRawRecords] = useState(INITIAL_MOCK_DATA);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterStatus, setFilterStatus] = useState('All');
  const [followUpFilter, setFollowUpFilter] = useState('All');
  
  const [selectedRecordIds, setSelectedRecordIds] = useState(new Set());
  const [showBulkDeleteConfirm, setShowBulkDeleteConfirm] = useState(false);

  // Real-time Cloud Synced States
  const [handledRecords, setHandledRecords] = useState(new Set());
  const [deletedRecordIds, setDeletedRecordIds] = useState(new Set());

  const [recordToDelete, setRecordToDelete] = useState(null);
  const [selectedClass, setSelectedClass] = useState('All');
  const [dateRange, setDateRange] = useState({ start: '', end: '' });
  
  const [activeTab, setActiveTab] = useState('dashboard');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [noticeEmails, setNoticeEmails] = useState('');
  const [isLiveMode, setIsLiveMode] = useState(false);
  const [isFetchingLive, setIsFetchingLive] = useState(false);
  const [sheetUrl, setSheetUrl] = useState('');
  const [importStartDate, setImportStartDate] = useState('');
  const [importEndDate, setImportEndDate] = useState('');
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncError, setSyncError] = useState('');
  const [lastSync, setLastSync] = useState(null);
  const [autoSync, setAutoSync] = useState(true);
  const [showSheetUrl, setShowSheetUrl] = useState(false);
  
  // Cloud Reset state
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const [isResetting, setIsResetting] = useState(false);
  const [isUploading, setIsUploading] = useState(false);

  const handleLoginSubmit = (e) => {
    e.preventDefault();
    const username = e.target.username.value;
    const password = e.target.password.value;
    
    if (username === 'admin' && password === 'wkc2027') {
      setIsLoggedIn(true);
    } else {
      alert('帳號或密碼錯誤，請重新輸入！');
    }
  };

  const handleLogout = () => {
  setUserRole(null);
  setUser(null);
};

  const activeRawRecords = useMemo(() => {
    return rawRecords.filter(r => !deletedRecordIds.has(r.id));
  }, [rawRecords, deletedRecordIds]);

  const processedData = useMemo(() => {
    return processRecords(activeRawRecords, importStartDate, importEndDate);
  }, [activeRawRecords, importStartDate, importEndDate]);

  // 提取重複及被拒絕記錄的詳細清單
  const rejectedDetails = useMemo(() => {
    const duplicates = processedData.filter(r => r.reason === '重複記錄 (已忽略)');
    const quotaExceeded = processedData.filter(r => r.reason === '超出每日上限 (最多2次)');
    return { duplicates, quotaExceeded };
  }, [processedData]);

  const toggleHandledStatus = async (recordId) => {
    if (!user) return;
    
    const docId = encodeURIComponent(recordId);
    const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'handled_records', docId);
    const isCurrentlyHandled = handledRecords.has(recordId);

    setHandledRecords(prev => {
      const next = new Set(prev);
      if (isCurrentlyHandled) next.delete(recordId);
      else next.add(recordId);
      return next;
    });

    try {
      if (isCurrentlyHandled) {
        await deleteDoc(docRef);
      } else {
        await setDoc(docRef, { recordId: recordId, handledAt: Date.now() });
      }
    } catch (err) {
      console.error("更新處理狀態失敗:", err);
    }
  };

  // ==========================================
  // INSERTED: Detention Execution logic
  // ==========================================
  const handleServeDetention = async (studentId, studentName, recordsToProcess) => {
    if (!user || recordsToProcess.length < 5) return;
    
    // Process the oldest 5 valid records
    const targets = recordsToProcess.sort((a,b) => a.submitTime - b.submitTime).slice(0, 5);
    
    // Optimistic Update
    setHandledRecords(prev => {
      const next = new Set(prev);
      targets.forEach(r => next.add(r.id));
      return next;
    });

    // Cloud Update
    try {
      for (const record of targets) {
        const docId = encodeURIComponent(record.id);
        const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'handled_records', docId);
        await setDoc(docRef, { recordId: record.id, handledAt: Date.now(), type: 'Detention_Served' });
      }
      alert(`系統通知已更新：已成功為學生 ${studentName} (${studentId}) 安排留堂！對應的 5 項違規記錄已自動轉為「已處理 (Processed)」。`);
    } catch(err) {
      console.error(err);
    }
  };
  // ==========================================

  useEffect(() => {
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (err) {
        console.error("Authentication error:", err);
      }
    };
    initAuth();
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!user) return;

    const handledColl = collection(db, 'artifacts', appId, 'public', 'data', 'handled_records');
    const unsubHandled = onSnapshot(handledColl, (snapshot) => {
      const ids = new Set();
      snapshot.forEach(docSnap => {
        const data = docSnap.data();
        if (data.recordId) ids.add(data.recordId);
      });
      setHandledRecords(ids);
    }, (err) => console.error("Cloud sync error (handled):", err));

    const deletedColl = collection(db, 'artifacts', appId, 'public', 'data', 'deleted_records');
    const unsubDeleted = onSnapshot(deletedColl, (snapshot) => {
      const ids = new Set();
      snapshot.forEach(docSnap => {
        const data = docSnap.data();
        if (data.recordId) ids.add(data.recordId);
      });
      setDeletedRecordIds(ids);
    }, (err) => console.error("Cloud sync error (deleted):", err));

    const configDocRef = doc(db, 'artifacts', appId, 'public', 'data', 'settings', 'config');
    const unsubConfig = onSnapshot(configDocRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        if (data.sheetUrl) setSheetUrl(data.sheetUrl);
        if (data.importStartDate) setImportStartDate(data.importStartDate);
        if (data.importEndDate) setImportEndDate(data.importEndDate);
      }
    }, (err) => console.error("Cloud sync error (config):", err));

    return () => {
      unsubHandled();
      unsubDeleted();
      unsubConfig();
    };
  }, [user]);

  useEffect(() => {
    if (sheetUrl && sheetUrl.includes('output=csv')) {
      fetchLiveSheet(null, true);
    }
  }, [sheetUrl, importStartDate, importEndDate]);

// AFTER (Syncs every 10 seconds in the background):
useEffect(() => {
  let interval;
  if (autoSync && isLiveMode && sheetUrl) {
    interval = setInterval(() => {
      fetchLiveSheet(null, false); 
    }, 10000); // Changed from 60000 to 10000
  }
  return () => clearInterval(interval);
}, [autoSync, isLiveMode, sheetUrl, importStartDate, importEndDate]);

useEffect(() => {
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible' && isLiveMode && sheetUrl) {
        fetchLiveSheet(null, false);
      }
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);
    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [isLiveMode, sheetUrl, importStartDate, importEndDate]);

  const saveConfigToCloud = async (url, startDate, endDate) => {
    if (!user) return;
    try {
      const configDocRef = doc(db, 'artifacts', appId, 'public', 'data', 'settings', 'config');
      await setDoc(configDocRef, {
        sheetUrl: url,
        importStartDate: startDate || '',
        importEndDate: endDate || '',
        updatedAt: Date.now()
      }, { merge: true });
    } catch (e) {
      console.error("Failed to save config to cloud:", e);
    }
  };
  
const fetchLiveSheet = async (e, isManual = true) => {
    if (e) e.preventDefault();
    if (!sheetUrl) return;

    if (isManual) setIsSyncing(true);
    if (isManual) setIsFetchingLive(true);
    if (isManual) setSyncError('');

    try {
      if (!sheetUrl.includes('output=csv')) {
        throw new Error('網址格式無效。請確保該連結已發布為 CSV 格式 (檢查網址是否包含 "output=csv")。');
      }

      if (isManual) {
        await saveConfigToCloud(sheetUrl, importStartDate, importEndDate);
      }

      const fetchUrl = `${sheetUrl}&_cache_bust=${new Date().getTime()}`;
      const response = await fetch(fetchUrl);
      
      if (!response.ok) throw new Error('無法獲取數據。請確保 Google Sheet 已「發布到網絡」。');
      
      const csvText = await response.text();
      const parsed = parseCSV(csvText);
      
      if (parsed.length < 2) throw new Error('未找到數據列。請確保您的表格已有提交記錄。');

      const rawHeaders = parsed[0].map(h => String(h || '').replace(/['"]+/g, '').trim());
      const lowerHeaders = rawHeaders.map(h => h.toLowerCase());
      
      const getIdx = (exactMatches, partialMatches = []) => {
        for (let kw of exactMatches) {
          const idx = lowerHeaders.findIndex(h => h === kw);
          if (idx !== -1) return idx;
        }
        for (let kw of partialMatches) {
          const idx = lowerHeaders.findIndex(h => h.includes(kw));
          if (idx !== -1) return idx;
        }
        return -1;
      };

      const idxTimestamp = getIdx(['時間戳記', 'timestamp'], ['time', '時間']);
      const idxDate = getIdx(['日期 date', '日期', 'date', 'incident date'], ['date', '日期']);
      
      // Specifically target "email part" or general emails for QR code detection
      const idxEmail = getIdx(
        ['email part', 'email address', '電子郵件地址', 'email', '電郵', 'username', '使用者名稱'], 
        ['email part', 'email', '電郵', 'username']
      );

      const idxTeacher = getIdx(
        ["老師姓名 teacher's name", '老師姓名', 'teacher name', "teacher's name", 'teacher', '老師', '教職員', '報告老師', 'reporter'], 
        ['teacher', '老師', 'staff', '教職', '負責人']
      );
      
      const idxClass = getIdx(['班別', 'class', 'form'], ['class', 'room', '班']);

      const studentColumnIndices = lowerHeaders
        .map((h, idx) => {
          if (idx === idxTeacher || idx === idxEmail) return -1;
          const isTeacherHeader = h.includes('老師') || h.includes('teacher') || h.includes('教職員') || h.includes('reporter');
          if (isTeacherHeader) return -1;

          const isItemOrSubject = h.includes('欠帶') || h.includes('用品') || h.includes('科目') || h.includes('subject') || h.includes('item') || h.includes('forgotten') || h.includes('which');
          if (isItemOrSubject) return -1;

          const isStudentNameCol = h.includes('姓名') || h.includes('名字') || (h.includes('student') && h.includes('name')) || h.includes('學生');
          return isStudentNameCol ? idx : -1;
        })
        .filter(idx => idx !== -1);

      if (studentColumnIndices.length === 0) {
        throw new Error('缺少必要欄位：在表格中找不到學生姓名欄位。');
      }

      const parseHKTime = (timeStr) => {
        if (!timeStr) return NaN;
        let d = new Date(timeStr);
        if (!isNaN(d.getTime())) return d.getTime();
        
        let normalized = timeStr.replace('上午', 'AM').replace('下午', 'PM');
        let normalizedDate = new Date(normalized);
        if (!isNaN(normalizedDate.getTime())) return normalizedDate.getTime();
        
        const regex = /(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s+(上午|下午)\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/;
        const match = timeStr.match(regex);
        if (match) {
          let [_, y, m, day, ampm, h, min, s] = match;
          h = parseInt(h, 10);
          if (ampm === '下午' && h < 12) h += 12;
          if (ampm === '上午' && h === 12) h = 0;
          return new Date(y, m - 1, day, h, min, s).getTime();
        }
        return NaN;
      };

      const cleanCell = (val) => String(val || '').replace(/['"]+/g, '').trim();
      const newRecords = [];

      for (let i = 1; i < parsed.length; i++) {
        const row = parsed[i];
        if (!row || row.length === 0 || row.join('').trim() === '') continue;

        let submitTime = Date.now();
        let hasValidTimestamp = false;
        if (idxTimestamp !== -1 && row[idxTimestamp]) {
          let t = parseHKTime(cleanCell(row[idxTimestamp]));
          if (!isNaN(t)) {
             submitTime = t;
             hasValidTimestamp = true;
          }
        }

        let incidentDate = hasValidTimestamp ? getLocalDateStr(submitTime) : getTodayStr();
        if (idxDate !== -1 && row[idxDate]) {
          const rawDate = cleanCell(row[idxDate]);
          let parsedDateObj = new Date(rawDate);
          if (!isNaN(parsedDateObj.getTime())) {
             incidentDate = getLocalDateStr(parsedDateObj);
          } else {
             const parts = rawDate.split(/[\/\-]/);
             if (parts.length >= 3) {
                if (parts[2].length === 4) { 
                   incidentDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                } else { 
                   incidentDate = `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].split(' ')[0].padStart(2, '0')}`;
                }
             }
          }
        }

        if (importStartDate && incidentDate < importStartDate) continue;
        if (importEndDate && incidentDate > importEndDate) continue;

        let globalClass = idxClass !== -1 ? cleanCell(row[idxClass]) : '';
        
        // Handling Email Part & Identifying Potential QR/Student Submissions
        let emailStr = idxEmail !== -1 ? cleanCell(row[idxEmail]) : '';
        // Pattern matches: s + year + 4 digits @wkc.edu.hk
        const isStudentEmail = /^s\d{8}(@wkc\.edu\.hk)?$/i.test(emailStr);
        let displayEmail = isStudentEmail ? `${emailStr} (possible student)` : emailStr;

        let teacher = idxTeacher !== -1 ? cleanCell(row[idxTeacher]) : '';
        if (!teacher && emailStr) teacher = emailStr.split('@')[0];
        else if (!teacher) teacher = '系統';

        for (let colIdx of studentColumnIndices) {
          const rawStudentVal = cleanCell(row[colIdx]);
          if (!rawStudentVal) continue;

          let recordClass = globalClass;
          if (!recordClass || recordClass === '未指定') {
            const headerText = rawHeaders[colIdx] || '';
            const classMatch = headerText.match(/([1-6][A-C])/i);
            recordClass = classMatch ? classMatch[1].toUpperCase() : '未指定';
          }

          let subject = '';
          for (let k = colIdx - 1; k >= Math.max(0, colIdx - 4); k--) {
             const h = lowerHeaders[k] || '';
             if (h.includes('科目') || h.includes('subject')) {
                if (row[k] && cleanCell(row[k]) !== '') {
                   subject = cleanCell(row[k]);
                   break;
                }
             }
          }

          let item = '';
          for (let k = colIdx + 1; k <= Math.min(rawHeaders.length - 1, colIdx + 4); k++) {
             const h = lowerHeaders[k] || '';
             if (h.includes('欠帶') || h.includes('用品') || h.includes('資料') || h.includes('items') || h.includes('forgotten')) {
                if (row[k] && cleanCell(row[k]) !== '') {
                   item = cleanCell(row[k]);
                   break;
                }
             }
          }

          let studentId = "未提供";
          let studentName = rawStudentVal;
          const nameMatch = rawStudentVal.match(/^(\d+)\s+(.+)$/);
          if (nameMatch) {
            studentId = nameMatch[1];
            studentName = nameMatch[2];
          }

          const ruleFormatted = subject ? `[${subject}] ${item || '未填寫'}` : (item || '未填寫');
          const recordId = `LIVE-${submitTime}-${studentId}-${subject}-${item}`.replace(/\s+/g, '')

          newRecords.push({
  id: recordId,
  studentId: studentId,
  studentName: studentName,
  studentClass: recordClass,
  rule: ruleFormatted,
  incidentDate: incidentDate,
  submitTime: submitTime,
  teacher: teacher,
  reporterEmail: formatStudentEmail(displayEmail) // <-- Adds "(possible student)" when matched
});
        }
      }
      
      // 防護機制：若新抓到的資料為空，但原本有資料，則跳過本次更新
      if (newRecords.length === 0 && rawRecords.length > 0) {
        console.warn("偵測到背景抓取空資料，已自動忽略本次更新以保護畫面。");
        return;
      }

      setRawRecords((prevRecords) => {
        if (JSON.stringify(prevRecords) === JSON.stringify(newRecords)) {
          return prevRecords;
        }
        return newRecords;
      });

      setLastSync(new Date());
      } catch (err) {
      if (isManual) {
        setSyncError(err.message);
        setAutoSync(false);
      }
      console.error("Fetch Live Sheet Error:", err);
    } finally {
      if (isManual) setIsSyncing(false);
      if (isManual) setIsFetchingLive(false);
    }
  };
const handleDeleteRecord = async (record) => {
    if (userRole !== 'admin' || !record) return;
    const docId = encodeURIComponent(record.id);

    setDeletedRecordIds(prev => new Set(prev).add(record.id));
    setRecordToDelete(null);

    if (user) {
      try {
        const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'deleted_records', docId);
        await setDoc(docRef, { recordId: record.id, deletedAt: Date.now() });
      } catch (err) {
        console.error("雲端刪除失敗:", err);
      }
    }
  };

  const handleBulkDelete = async () => {
    if (userRole !== 'admin') return;
    const idsToDelete = Array.from(selectedRecordIds);

    setDeletedRecordIds(prev => {
      const next = new Set(prev);
      idsToDelete.forEach(id => next.add(id));
      return next;
    });
    setSelectedRecordIds(new Set());
    setShowBulkDeleteConfirm(false);

    if (user) {
      for (const id of idsToDelete) {
        try {
          const docId = encodeURIComponent(id);
          const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'deleted_records', docId);
          await setDoc(docRef, { recordId: id, deletedAt: Date.now() });
        } catch (err) {
          console.error("雲端批量刪除失敗:", err);
        }
      }
    }
  };
const analyticsChartRef = useRef(null);

const exportAnalyticsToExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();

      // =============================================================
      // Sheet 1: 違規項目統計
      // =============================================================
      const wsRules = workbook.addWorksheet('違規項目統計');
      wsRules.columns = [
        { header: '違規/欠帶項目', key: 'category', width: 48 },
        { header: '累計次數', key: 'count', width: 15 },
        { header: '佔比 (%)', key: 'ratio', width: 15 },
      ];

      const ruleStatsMap = {};
      let totalCount = 0;

      displayData.forEach(item => {
        const rule = item.rule || '其他';
        ruleStatsMap[rule] = (ruleStatsMap[rule] || 0) + 1;
        totalCount += 1;
      });

      const ruleEntries = Object.entries(ruleStatsMap).sort((a, b) => b[1] - a[1]);
      ruleEntries.forEach(([rule, count]) => {
        const ratio = totalCount > 0 ? ((count / totalCount) * 100).toFixed(1) + '%' : '0%';
        wsRules.addRow({ category: rule, count, ratio });
      });

      // 美化樣式
      wsRules.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
      wsRules.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4F46E5' } };
      wsRules.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

      // =============================================================
      // Sheet 2: 日期趨勢分析
      // =============================================================
      const wsTrend = workbook.addWorksheet('日期趨勢分析');
      wsTrend.columns = [
        { header: '日期', key: 'date', width: 20 },
        { header: '當日違規人次', key: 'count', width: 18 },
      ];

      const dateStatsMap = {};
      displayData.forEach(item => {
        const date = item.incidentDate || '未知日期';
        dateStatsMap[date] = (dateStatsMap[date] || 0) + 1;
      });

      const dateEntries = Object.entries(dateStatsMap).sort((a, b) => a[0].localeCompare(b[0]));
      dateEntries.forEach(([date, count]) => {
        wsTrend.addRow({ date, count });
      });

      wsTrend.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
      wsTrend.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '059669' } };
      wsTrend.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

      // =============================================================
      // Sheet 3: 班級與學生排行
      // =============================================================
      const wsStudents = workbook.addWorksheet('班級與學生排行');
      wsStudents.columns = [
        { header: '班別', key: 'studentClass', width: 15 },
        { header: '學生姓名', key: 'studentName', width: 22 },
        { header: '學號', key: 'studentId', width: 15 },
        { header: '違規總次數', key: 'count', width: 15 },
      ];

      const studentMap = {};
      displayData.forEach(item => {
        const key = `${item.studentClass}-${item.studentId}`;
        if (!studentMap[key]) {
          studentMap[key] = {
            studentClass: item.studentClass,
            studentName: item.studentName,
            studentId: item.studentId,
            count: 0
          };
        }
        studentMap[key].count += 1;
      });

      const studentEntries = Object.values(studentMap).sort((a, b) => b.count - a.count);
      studentEntries.forEach(s => {
        wsStudents.addRow(s);
      });

      wsStudents.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
      wsStudents.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DC2626' } };
      wsStudents.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

      // =============================================================
      // 在 Excel 內部自動注入圖表數據範圍參照定義 (Drawing XML Ref)
      // =============================================================
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `統計及趨勢分析報告_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("匯出失敗:", error);
      alert("匯出失敗：" + error.message);
    }
  };

  const saveAcceptedRecordsToCloud = async () => {
    if (!user) return;
    const accepted = processedData.filter(r => r.processStatus === 'Accepted');
    if (accepted.length === 0) return;
    
    setIsUploading(true);
    try {
      const chunkSize = 100;
      for (let i = 0; i < accepted.length; i += chunkSize) {
        const chunk = accepted.slice(i, i + chunkSize);
        await Promise.all(chunk.map(record => {
          const docId = encodeURIComponent(record.id);
          const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'accepted_records', docId);
          return setDoc(docRef, { ...record, backedUpAt: Date.now() }, { merge: true });
        }));
      }
    } catch (err) {
      console.error("Failed to backup accepted records:", err);
    } finally {
      setIsUploading(false);
    }
  };

  const displayData = useMemo(() => {
    return processedData.map(record => ({
      ...record,
      isHandled: handledRecords.has(record.id)
    })).filter(record => {
      const matchesSearch = 
        record.studentId.toLowerCase().includes(searchTerm.toLowerCase()) ||
        record.studentName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        record.teacher.toLowerCase().includes(searchTerm.toLowerCase());
      
      const filterMap = { 'All': 'All', 'Accepted': 'Accepted', 'Rejected': 'Rejected' };
      const matchesStatus = filterStatus === 'All' ? true : record.processStatus === filterMap[filterStatus];
      
      const matchesClass = selectedClass === 'All' ? true : record.studentClass === selectedClass;
      
      const matchesFollowUp = followUpFilter === 'All' ? true : 
                              followUpFilter === 'Handled' ? record.isHandled : 
                              !record.isHandled;

      let matchesDateRange = true;
      if (dateRange.start) {
        matchesDateRange = matchesDateRange && record.incidentDate >= dateRange.start;
      }
      if (dateRange.end) {
        matchesDateRange = matchesDateRange && record.incidentDate <= dateRange.end;
      }
      
      return matchesSearch && matchesStatus && matchesClass && matchesDateRange && matchesFollowUp;
    });
  }, [processedData, searchTerm, filterStatus, selectedClass, dateRange, followUpFilter, handledRecords]);

  // Calculate Detentions Candidates (5 entries = 1 detention)
  const detentionCandidates = useMemo(() => {
    const activeAccepted = processedData.filter(r => r.processStatus === 'Accepted' && !handledRecords.has(r.id));
    const counts = {};
    activeAccepted.forEach(r => {
      if (!counts[r.studentId]) counts[r.studentId] = { name: r.studentName, class: r.studentClass, records: [] };
      counts[r.studentId].records.push(r);
    });
    
    const candidates = [];
    for (const [id, data] of Object.entries(counts)) {
      if (data.records.length >= 5) {
        candidates.push({
          id,
          name: data.name,
          studentClass: data.class,
          pendingCount: data.records.length,
          detentionsNeeded: Math.floor(data.records.length / 5),
          records: data.records
        });
      }
    }
    return candidates.sort((a,b) => b.pendingCount - a.pendingCount);
  }, [processedData, handledRecords]);

  const toggleSelectAll = () => {
    if (selectedRecordIds.size === displayData.length && displayData.length > 0) {
      setSelectedRecordIds(new Set());
    } else {
      setSelectedRecordIds(new Set(displayData.map(r => r.id)));
    }
  };

  const toggleSelectRecord = (id) => {
    setSelectedRecordIds(prev => {
      const newSet = new Set(prev);
      if (newSet.has(id)) newSet.delete(id);
      else newSet.add(id);
      return newSet;
    });
  };

  const handleBulkExport = () => {
    const dataToExport = displayData.filter(r => selectedRecordIds.has(r.id));
    exportToCSV(dataToExport);
  };

  const stats = useMemo(() => {
    const today = getTodayStr();
    const todayRecords = processedData.filter(r => r.incidentDate === today);
    return {
      total: processedData.length,
      todayTotal: todayRecords.length,
      todayAccepted: todayRecords.filter(r => r.processStatus === 'Accepted').length,
      todayRejected: todayRecords.filter(r => r.processStatus === 'Rejected').length,
    };
  }, [processedData]);

// 登入關卡阻擋未授權存取
  if (!userRole) {
    return <Login onLogin={(role) => setUserRole(role)} />;
  }

  const renderDashboard = () => (
    <div className="space-y-6">
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
        <StatCard 
          title="總處理記錄" 
          subtitle={`指定範圍內有效數據`}
          value={stats.total} 
          icon={Database} 
          colorClass="text-indigo-600" 
          bgClass="bg-indigo-50" 
          trend={stats.todayTotal}
        />
        <StatCard 
          title="今日事發記錄" 
          subtitle="今天發生的欠帶總數"
          value={stats.todayTotal} 
          icon={Calendar} 
          colorClass="text-blue-600" 
          bgClass="bg-blue-50" 
        />
        <StatCard 
          title="已接納記錄" 
          subtitle="符合規則的記錄 (今日)"
          value={stats.todayAccepted} 
          icon={CheckCircle} 
          colorClass="text-emerald-600" 
          bgClass="bg-emerald-50" 
        />
        <StatCard 
          title="已攔截記錄" 
          subtitle="重複或超額記錄 (今日)"
          value={stats.todayRejected} 
          icon={AlertTriangle} 
          colorClass="text-rose-600" 
          bgClass="bg-rose-50" 
        />
      </div>

      {/* 重複記錄與拒絕明細面板 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6">
          <div className="flex items-center gap-2 mb-4">
            <XCircle className="w-5 h-5 text-rose-500" />
            <h3 className="text-lg font-bold text-slate-800">重複記錄 (已忽略) 詳情列表</h3>
          </div>
          {rejectedDetails.duplicates.length === 0 ? (
            <p className="text-sm text-slate-400 py-4">目前沒有完全重複的記錄。</p>
          ) : (
            <div className="overflow-y-auto max-h-60">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-500 border-b border-slate-100">
                  <tr>
                    <th className="p-2">日期</th>
                    <th className="p-2">班別</th>
                    <th className="p-2">學生</th>
                    <th className="p-2">欠帶事項</th>
                    <th className="p-2">報告老師</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {rejectedDetails.duplicates.map(r => (
                    <tr key={r.id} className="hover:bg-slate-50/50">
                      <td className="p-2 whitespace-nowrap">{r.incidentDate}</td>
                      <td className="p-2 font-bold">{r.studentClass}</td>
                      <td className="p-2">{r.studentName} ({r.studentId})</td>
                      <td className="p-2 text-rose-600 font-medium">{r.rule}</td>
                      <td className="p-2">{r.teacher}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6">
          <div className="flex items-center gap-2 mb-4">
            <AlertTriangle className="w-5 h-5 text-amber-500" />
            <h3 className="text-lg font-bold text-slate-800">超出每日上限 / 已拒絕記錄 詳情列表</h3>
          </div>
          {rejectedDetails.quotaExceeded.length === 0 ? (
            <p className="text-sm text-slate-400 py-4">目前沒有超出每日上限的記錄。</p>
          ) : (
            <div className="overflow-y-auto max-h-60">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-500 border-b border-slate-100">
                  <tr>
                    <th className="p-2">日期</th>
                    <th className="p-2">班別</th>
                    <th className="p-2">學生</th>
                    <th className="p-2">欠帶事項</th>
                    <th className="p-2">報告老師</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {rejectedDetails.quotaExceeded.map(r => (
                    <tr key={r.id} className="hover:bg-slate-50/50">
                      <td className="p-2 whitespace-nowrap">{r.incidentDate}</td>
                      <td className="p-2 font-bold">{r.studentClass}</td>
                      <td className="p-2">{r.studentName} ({r.studentId})</td>
                      <td className="p-2 text-amber-600 font-medium">{r.rule}</td>
                      <td className="p-2">{r.teacher}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-6">
        <div className="bg-slate-900 rounded-2xl shadow-sm p-6 lg:p-8 text-slate-300 flex flex-col">
          <div className="flex items-center gap-3 mb-6">
            <div className="p-2.5 bg-blue-500/20 rounded-xl">
              <Filter className="w-6 h-6 text-blue-400" />
            </div>
            <div>
              <h3 className="text-xl font-bold text-white">系統處理規則說明</h3>
              <p className="text-sm text-slate-400 mt-1">此數據處理器作為訓導記錄的自動過濾閘道。</p>
            </div>
          </div>
          <div className="space-y-5 flex-1 text-sm md:text-base">
            <ul className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <li className="flex gap-4">
                <div className="min-w-[28px] h-7 rounded-full bg-slate-800 flex items-center justify-center text-sm font-bold text-white shrink-0">1</div>
                <span className="leading-relaxed">數據嚴格按照 Google Sheet 的<strong>提交時間 (Submit Time)</strong> 先後次序進行計算與配額發放。</span>
              </li>
              <li className="flex gap-4">
                <div className="min-w-[28px] h-7 rounded-full bg-slate-800 flex items-center justify-center text-sm font-bold text-white shrink-0">2</div>
                <span className="leading-relaxed">如發現<strong>完全重複</strong>的記錄 (相同學生、日期及欠帶事項)，後續記錄將被自動標記為 `重複記錄 (已忽略)`。</span>
              </li>
              <li className="flex gap-4">
                <div className="min-w-[28px] h-7 rounded-full bg-slate-800 flex items-center justify-center text-sm font-bold text-white shrink-0">3</div>
                <span className="leading-relaxed">同一位學生（依學號與姓名識別）每日<strong>最先提交的前 2 宗</strong>紀錄將獲 <span className="text-emerald-400 font-medium">接納</span>。</span>
              </li>
              <li className="flex gap-4">
                <div className="min-w-[28px] h-7 rounded-full bg-slate-800 flex items-center justify-center text-sm font-bold text-white shrink-0">4</div>
                <span className="leading-relaxed">同一天內即使由不同老師登記，第三宗（含）以上的任何紀錄將即時標記為 <span className="text-rose-400 font-medium">拒絕（超出每日上限）</span>。</span>
              </li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );

  const renderRecords = () => (
    <div className="bg-white rounded-2xl shadow-sm border border-slate-100 flex flex-col h-[calc(100vh-140px)]">
      <div className="p-5 border-b border-slate-100 flex flex-col gap-4">
        {(!isLiveMode && !sheetUrl) && (
          <div className="bg-amber-50 border border-amber-200 p-4 rounded-xl flex items-start gap-3 animate-in fade-in">
            <AlertTriangle className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" />
            <div>
              <p className="text-sm font-bold text-amber-800">目前顯示為系統內建的「示範數據 (Demo Data)」</p>
              <p className="text-xs text-amber-700 mt-1">您尚未連結 Google Sheet 數據庫，重新整理網頁後數據將會重置。請前往「實時數據連接」綁定您的試算表。</p>
            </div>
          </div>
        )}
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center bg-indigo-50 p-4 rounded-xl mb-2 gap-4 border border-indigo-100/50">
          <div className="flex items-center gap-3 text-indigo-800 text-sm">
            <div className="bg-indigo-100/80 p-2 rounded-lg shrink-0">
              <Cloud className="w-5 h-5 text-indigo-600" />
            </div>
            <p className="leading-relaxed">
              <strong className="font-semibold mr-1">雲端多使用者同步：</strong>
              所有教職員標記的「已處理」或「已刪除」狀態已同步至雲端，所有裝置皆可即時檢視。
            </p>
          </div>
          <div className="flex items-center gap-3 w-full sm:w-auto shrink-0">
            {selectedRecordIds.size > 0 ? (
              <div className="flex flex-col sm:flex-row items-center gap-3 w-full sm:w-auto animate-in fade-in zoom-in-95 duration-200">
                <span className="text-sm font-semibold text-indigo-700 bg-white/50 px-3 py-1.5 rounded-lg border border-indigo-200/50 w-full sm:w-auto text-center">
                  已選取 {selectedRecordIds.size} 項
                </span>
                <div className="flex items-center gap-2 w-full sm:w-auto">      
            <button 
                    onClick={handleBulkExport}
                    className="flex-1 sm:flex-none flex items-center justify-center gap-2 bg-indigo-600 border border-transparent text-white hover:bg-indigo-700 px-4 py-2 rounded-lg text-sm font-bold transition-all shadow-sm active:scale-95"
                  >
                    <Download className="w-4 h-4" />
                    匯出
                  </button>
                </div>
              </div>
            ) : (
              <button 
                onClick={() => exportToCSV(displayData)}
                className="w-full sm:w-auto flex items-center justify-center gap-2 bg-white border border-indigo-200 text-indigo-700 hover:bg-indigo-50 px-4 py-2 rounded-lg text-sm font-bold transition-all shadow-sm active:scale-95 hover:border-indigo-300"
              >
                <Download className="w-4 h-4" />
                匯出全部 (CSV)
              </button>
            )}
          </div>
        </div>
        
        <div className="flex flex-col xl:flex-row justify-between items-start xl:items-center gap-4">
          <div className="flex flex-wrap gap-3 w-full xl:w-auto">
            <div className="flex bg-slate-100 p-1 rounded-xl w-full sm:w-auto overflow-x-auto shrink-0">
              {[
                { id: 'All', label: '全部' }, 
                { id: 'Accepted', label: '已接納' }, 
                { id: 'Rejected', label: '已拒絕' }
              ].map(status => (
                <button
                  key={status.id}
                  onClick={() => setFilterStatus(status.id)}
                  className={`flex-1 sm:flex-none px-4 py-2 rounded-lg text-sm font-semibold transition-all ${
                    filterStatus === status.id 
                      ? 'bg-white text-slate-800 shadow-sm' 
                      : 'text-slate-500 hover:text-slate-700'
                  }`}
                >
                  {status.label}
                </button>
              ))}
            </div>

            <div className="flex bg-slate-100 p-1 rounded-xl w-full sm:w-auto overflow-x-auto shrink-0">
              {[
                { id: 'All', label: '所有進度' }, 
                { id: 'Pending', label: '未處理' }, 
                { id: 'Handled', label: '已處理' }
              ].map(status => (
                <button
                  key={status.id}
                  onClick={() => setFollowUpFilter(status.id)}
                  className={`flex-1 sm:flex-none px-4 py-2 rounded-lg text-sm font-semibold transition-all ${
                    followUpFilter === status.id 
                      ? 'bg-white text-slate-800 shadow-sm' 
                      : 'text-slate-500 hover:text-slate-700'
                  }`}
                >
                  {status.label}
                </button>
              ))}
            </div>
          </div>
          
          <div className="flex flex-col sm:flex-row items-center gap-3 w-full xl:w-auto">
            <select
              value={selectedClass}
              onChange={(e) => setSelectedClass(e.target.value)}
              className="w-full sm:w-36 text-sm border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 p-2.5 border bg-slate-50 transition-all outline-none shrink-0"
            >
              <option value="All">所有班別</option>
              {CLASSES.map(cls => <option key={cls} value={cls}>{cls}</option>)}
            </select>

            <div className="flex items-center gap-2 w-full sm:w-auto shrink-0">
              <input
                type="date"
                value={dateRange.start}
                onChange={(e) => setDateRange(prev => ({ ...prev, start: e.target.value }))}
                className="w-full sm:w-auto text-sm border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 p-2.5 border bg-slate-50 transition-all outline-none"
              />
              <span className="text-slate-400 text-sm font-medium">至</span>
              <input
                type="date"
                value={dateRange.end}
                onChange={(e) => setDateRange(prev => ({ ...prev, end: e.target.value }))}
                className="w-full sm:w-auto text-sm border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 p-2.5 border bg-slate-50 transition-all outline-none"
              />
            </div>

            <div className="relative w-full sm:w-64 shrink-0">
              <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                <Search className="h-4 w-4 text-slate-400" />
              </div>
              <input
                type="text"
                placeholder="搜尋學生或老師姓名..."
                className="pl-10 w-full text-sm border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 p-2.5 border bg-slate-50 transition-all outline-none"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>
        </div>
      </div>

      <div className="flex-1 overflow-auto">
        <table className="w-full text-left border-collapse">
          <thead className="sticky top-0 z-10 bg-white/95 backdrop-blur-sm border-b border-slate-100">
            <tr className="text-xs uppercase tracking-wider text-slate-400 font-bold">
              <th className="p-5 whitespace-nowrap w-12 text-center">
                <button 
                  onClick={toggleSelectAll}
                  className="text-slate-400 hover:text-indigo-500 transition-colors focus:outline-none"
                >
                  {displayData.length > 0 && selectedRecordIds.size === displayData.length ? (
                    <CheckSquare className="w-5 h-5 text-indigo-500" />
                  ) : (
                    <SquareCheck className="w-5 h-5" />
                  )}
                </button>
              </th>
              <th className="p-5 whitespace-nowrap">提交時間</th>
              <th className="p-5 whitespace-nowrap">學生資料</th>
              <th className="p-5 whitespace-nowrap">班別</th>
              <th className="p-5 whitespace-nowrap">事發日期</th>
              <th className="p-5 whitespace-nowrap">欠帶事項</th>
              <th className="p-5 whitespace-nowrap">報告教職員</th>
              <th className="p-5 whitespace-nowrap">系統狀態</th>
              <th className="p-5 whitespace-nowrap">操作</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50 text-sm">
            {isFetchingLive ? (
              <tr>
                <td colSpan="9" className="p-12 text-center">
                  <div className="flex flex-col items-center justify-center text-slate-400">
                    <RefreshCw className="w-8 h-8 mb-3 text-indigo-400 animate-spin" />
                    <p className="font-medium text-slate-500">正在同步雲端即時數據...</p>
                    <p className="text-xs mt-1">請稍候，正在從 Google Sheet 載入最新記錄</p>
                  </div>
                </td>
              </tr>
            ) : displayData.length === 0 ? (
              <tr>
                <td colSpan="9" className="p-12 text-center">
                  <div className="flex flex-col items-center justify-center text-slate-400">
                    <FileSpreadsheet className="w-12 h-12 mb-3 text-slate-200" />
                    <p className="font-medium text-slate-500">找不到任何記錄</p>
                    <p className="text-xs mt-1">請嘗試調整過濾器或搜尋關鍵字。</p>
                  </div>
                </td>
              </tr>
            ) : (
              displayData.map((record) => (
                <tr key={record.id} className={`hover:bg-slate-50/80 transition-colors group ${selectedRecordIds.has(record.id) ? 'bg-indigo-50/30' : ''}`}>
                  <td className="p-5 whitespace-nowrap text-center">
                    <button 
                      onClick={() => toggleSelectRecord(record.id)}
                      className="text-slate-300 hover:text-indigo-400 transition-colors focus:outline-none"
                    >
                      {selectedRecordIds.has(record.id) ? (
                        <CheckSquare className="w-5 h-5 text-indigo-500" />
                      ) : (
                        <SquareCheck className="w-5 h-5" />
                      )}
                    </button>
                  </td>
                  <td className="p-5 text-slate-500 whitespace-nowrap">
                    <div className="flex items-center gap-2">
                      <Clock className="w-4 h-4 text-slate-300 group-hover:text-indigo-400 transition-colors" />
                      {formatTime(record.submitTime)}
                    </div>
                  </td>
                  <td className="p-5 whitespace-nowrap">
                    <div className="font-semibold text-slate-800">{record.studentName}</div>
                    <div className="text-xs text-slate-500 font-mono mt-0.5">{record.studentId}</div>
                  </td>
                  <td className="p-5 whitespace-nowrap">
                    <span className="font-medium text-slate-700 bg-slate-100 px-2 py-1 rounded-md text-xs">{record.studentClass}</span>
                  </td>
                  <td className="p-5 whitespace-nowrap">
                    <div className="flex items-center gap-2 text-slate-600">
                      <Calendar className="w-4 h-4 text-slate-300" />
                      {formatDate(record.incidentDate)}
                    </div>
                  </td>
                  <td className="p-5 whitespace-nowrap">
                    <span className="font-medium text-slate-700 bg-slate-100 px-2.5 py-1 rounded-md text-xs">{record.rule}</span>
                  </td>
<td className="p-5 whitespace-nowrap">
  <div className="flex items-center gap-2 text-slate-500">
    <div className="w-6 h-6 rounded-full bg-indigo-100 flex items-center justify-center text-indigo-700 font-bold text-[10px] shrink-0">
      {record.teacher ? record.teacher.charAt(0) : ''}
    </div>
    <div className="flex flex-col">
      <span>{record.teacher}</span>
      {record.reporterEmail && (
        <span className={`text-[11px] mt-0.5 ${record.reporterEmail.includes('(possible student)') ? 'text-amber-600 font-bold' : 'text-slate-400'}`}>
          {record.reporterEmail}
        </span>
      )}
    </div>
  </div>
</td>
      <td className="p-5 whitespace-nowrap">
        <StatusBadge status={record.processStatus} reason={record.reason} />
      </td>
                  <td className="p-5 whitespace-nowrap">
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => toggleHandledStatus(record.id)}
                        className={`inline-flex items-center px-3 py-1.5 rounded-lg text-xs font-bold transition-all border ${
                          record.isHandled 
                            ? 'bg-indigo-50 text-indigo-700 border-indigo-200 hover:bg-indigo-100' 
                            : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-50 hover:text-slate-700 shadow-sm'
                        }`}
                      >
                        {record.isHandled ? (
                          <><CheckCircle className="w-3.5 h-3.5 mr-1.5 text-indigo-500" /> 已處理</>
                        ) : (
                          <><Clock className="w-3.5 h-3.5 mr-1.5 text-slate-400" /> 未處理</>
                        )}
                      </button>
                     {userRole === 'admin' && (
  <button
    onClick={() => setRecordToDelete(record)}
    className="p-1.5 text-slate-400 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-all border border-transparent hover:border-rose-100"
    title="刪除紀錄"
  >
    <Trash2 className="w-4.5 h-4.5" />
  </button>
)}
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );

  const renderIntegration = () => (
    <div className="max-w-3xl mx-auto">
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="p-8 border-b border-slate-100 bg-gradient-to-br from-emerald-500 to-teal-600 text-white">
          <div className="flex items-center gap-3 mb-3">
            <div className="bg-white/20 p-2 rounded-xl backdrop-blur-sm">
              <Database className="w-6 h-6 text-white" />
            </div>
            <h2 className="text-2xl font-bold">實時數據整合 (雲端多使用者)</h2>
          </div>
          <p className="text-emerald-50 opacity-90">
            全量讀取 Google Sheet 試算表內所有數據，並透過 Firebase 雲端資料庫實時同步跨裝置操作狀態（包含「已處理」標記與「刪除記錄」）。
          </p>
        </div>
        
        <div className="p-8">
          <form onSubmit={(e) => fetchLiveSheet(e, false)} className="space-y-6">
            
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-5 mb-2 relative">
              <label className="block text-sm font-bold text-slate-800 mb-2 flex items-center gap-2">
                <Calendar className="w-4 h-4 text-emerald-600" />
                自訂數據過濾範圍 (可留空以載入全部)
              </label>
              <p className="text-xs text-slate-500 mb-3 font-medium">若留空，系統將會直接載入 Google Sheet 中的所有歷史違規記錄。</p>
              <div className="flex items-center gap-4">
                <input
                  type="date"
                  value={importStartDate}
                  onChange={(e) => setImportStartDate(e.target.value)}
                  className="w-full sm:w-64 border-slate-200 rounded-xl shadow-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 p-2.5 border bg-white transition-all outline-none font-mono text-sm"
                />
                <span className="text-slate-400 font-medium">至</span>
                <input
                  type="date"
                  value={importEndDate}
                  onChange={(e) => setImportEndDate(e.target.value)}
                  className="w-full sm:w-64 border-slate-200 rounded-xl shadow-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 p-2.5 border bg-white transition-all outline-none font-mono text-sm"
                />
              </div>
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-800 mb-2 flex items-center gap-2">
                <Lock className="w-4 h-4 text-emerald-600" />
                Google Sheet CSV 連結
              </label>
              <div className="flex flex-col sm:flex-row gap-3">
                <div className="relative flex-1">
                  <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                    <LinkIcon className="h-5 w-5 text-slate-400" />
                  </div>
                  <input
                    type={showSheetUrl ? "text" : "password"}
                    required
                    placeholder="https://docs.google.com/spreadsheets/d/e/.../pub?output=csv"
                    className="pl-11 pr-12 w-full border-slate-200 rounded-xl shadow-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 p-3 border bg-slate-50 transition-all outline-none font-mono text-sm"
                    value={sheetUrl}
                    onChange={(e) => setSheetUrl(e.target.value)}
                  />
                  <button 
                    type="button"
                    onClick={() => setShowSheetUrl(!showSheetUrl)}
                    className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-slate-400 hover:text-emerald-600 transition-colors focus:outline-none"
                    title={showSheetUrl ? "隱藏連結" : "顯示連結"}
                  >
                    {showSheetUrl ? <EyeOff className="h-5 w-5" /> : <Eye className="h-5 w-5" />}
                  </button>
                </div>
                <button 
                  type="submit" 
                  disabled={isSyncing}
                  className="bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-400 text-white font-semibold py-3 px-6 rounded-xl shadow-sm shadow-emerald-600/20 flex items-center justify-center transition-all whitespace-nowrap active:scale-[0.98]"
                >
                  <RefreshCw className={`w-5 h-5 mr-2 ${isSyncing ? 'animate-spin' : ''}`} />
                  {isSyncing ? '同步中...' : '儲存設定並同步全量數據'}
                </button>
              </div>
            </div>

            <div className="border-t border-slate-100 pt-6 mt-6">
              <div className="flex items-center justify-between bg-indigo-50 border border-indigo-100 rounded-xl p-4">
                <div>
                  <h4 className="text-sm font-bold text-indigo-800 flex items-center gap-2">
                    <Database className="w-4 h-4 text-indigo-600" />
                    備份「已接納」記錄至雲端
                  </h4>
                  <p className="text-xs text-indigo-600 mt-1">
                    將當前表格中所有標記為「已接納」的記錄正式儲存至 Firebase 資料庫中。
                  </p>
                </div>
                <button
                  type="button"
                  onClick={saveAcceptedRecordsToCloud}
                  disabled={isUploading || processedData.length === 0}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:bg-indigo-400 text-white rounded-lg text-xs font-bold transition-all shadow-sm shrink-0 flex items-center gap-2"
                >
                  <Cloud className="w-4 h-4" />
                  {isUploading ? '上傳中...' : '開始備份'}
                </button>
              </div>
            </div>

            <div className="border-t border-slate-100 pt-6 mt-6">
              <div className="flex items-center justify-between bg-rose-50 border border-rose-100 rounded-xl p-4">
                <div>
                  <h4 className="text-sm font-bold text-rose-800 flex items-center gap-2">
                    <Trash2 className="w-4 h-4 text-rose-600" />
                    重置/清空雲端同步數據
                  </h4>
                  <p className="text-xs text-rose-600 mt-1">
                    清除已於 Firebase 雲端註記的「已刪除」與「已處理」記錄狀態，回復初始顯示。
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setShowResetConfirm(true)}
                  className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-xs font-bold transition-all shadow-sm shrink-0"
                >
                  清空雲端狀態
                </button>
              </div>
            </div>
            
            {syncError && (
              <div className="flex items-start gap-3 text-sm text-rose-700 bg-rose-50 p-4 rounded-xl border border-rose-100">
                <AlertTriangle className="w-5 h-5 shrink-0 mt-0.5" />
                <div>
                  <p className="font-bold mb-1">連接錯誤</p>
                  <p>{syncError}</p>
                </div>
              </div>
            )}
            
            {lastSync && !syncError && (
              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 text-sm text-emerald-800 bg-emerald-50/50 p-4 rounded-xl border border-emerald-100">
                <div className="flex items-center gap-3">
                  <div className="relative flex h-3 w-3">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-3 w-3 bg-emerald-500"></span>
                  </div>
                  <div>
                    <p className="font-bold text-emerald-900">雲端全量數據已同步</p>
                    <p className="text-emerald-700">正在追蹤全量 {rawRecords.length} 條記錄（已同步至 Firebase）。</p>
                  </div>
                </div>
                
                <div className="flex items-center gap-3 border-t sm:border-t-0 sm:border-l border-emerald-200 pt-3 sm:pt-0 sm:pl-4 w-full sm:w-auto">
                  <span className="font-semibold text-emerald-900">自動同步資料</span>
                  <button
                    type="button"
                    onClick={() => setAutoSync(!autoSync)}
                    className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:ring-offset-2 ${autoSync ? 'bg-emerald-500' : 'bg-slate-300'}`}
                  >
                    <span className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${autoSync ? 'translate-x-6' : 'translate-x-1'}`} />
                  </button>
                </div>
              </div>
            )}
          </form>

          <div className="mt-10">
            <h4 className="font-bold text-slate-800 mb-4">系統連接教學</h4>
            <div className="space-y-4 relative before:absolute before:inset-0 before:ml-5 before:-translate-x-px md:before:mx-auto md:before:translate-x-0 before:h-full before:w-0.5 before:bg-gradient-to-b before:from-transparent before:via-slate-200 before:to-transparent">
              {[
                { step: 1, title: '建立表單', desc: '確保 Google Form 包含所需欄位：學號、姓名、班別、欠帶事項。' },
                { step: 2, title: '連結至試算表', desc: '在表單的「回覆」分頁中，點擊「連結至 Google 試算表」建立數據庫。' },
                { step: 3, title: '發布到網絡', desc: '打開試算表，前往 檔案 (File) > 共用 (Share) > 發布到網絡 (Publish to web)。' },
                { step: 4, title: '匯出為 CSV', desc: '選擇您的資料分頁 (通常是 Form Responses 1)，格式選擇「逗號分隔值 (.csv)」，然後點擊發布。' },
                { step: 5, title: '貼上連結', desc: '複製生成的連結，貼上至上方的輸入框即可開始同步。' },
              ].map((item) => (
                <div key={item.step} className="relative flex items-center justify-between md:justify-normal md:odd:flex-row-reverse group is-active">
                  <div className="flex items-center justify-center w-10 h-10 rounded-full border-4 border-white bg-slate-100 text-slate-500 font-bold shadow shrink-0 md:order-1 md:group-odd:-translate-x-1/2 md:group-even:translate-x-1/2">
                    {item.step}
                  </div>
                  <div className="w-[calc(100%-4rem)] md:w-[calc(50%-2.5rem)] p-4 rounded-xl border border-slate-100 bg-slate-50 shadow-sm">
                    <h5 className="font-bold text-slate-800 mb-1">{item.title}</h5>
                    <p className="text-sm text-slate-500">{item.desc}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );

const handleSendEmailNotice = async (rosterData) => {
  if (!noticeEmails.trim()) {
    alert('請先輸入收件人 Email（多個信箱請用逗號分隔）');
    return;
  }

  // 替換為您專屬的 WKC Google Apps Script 網址
  const GAS_URL = 'https://script.google.com/a/macros/wkc.edu.hk/s/AKfycbyuyYYCRXqdMODdL9uAtwBG1BCkuE_XgKkdyYFljn6EcOf5w2Sga_YHOq2Oe1OG98jX/exec';
  const teacherEmails = noticeEmails; // 使用系統面板上輸入的信箱
    
    // 組合信件內容
    let emailBody = '各位老師好，\n\n以下為系統自動產生的最新【留堂及通知名單】（累積滿 5 次違規）：\n\n';
    
    rosterData.forEach(student => {
      emailBody += `-----------------------------------\n`;
      emailBody += `學生：${student.class} ${student.name} (${student.id})\n`;
      emailBody += `累積違規：${student.totalRecords} 次\n`;
      emailBody += `最新三次違規紀錄：\n`;
      // 只列出最新的 3 筆避免信件過長
      student.records.slice(0, 3).forEach(r => {
        emailBody += ` - [${r.incidentDate}] ${r.rule} (報告: ${r.teacher})\n`;
      });
      emailBody += `\n`;
    });
    
    emailBody += '\n本信件由系統自動發送，請勿直接回覆。';

    try {
      const response = await fetch(GAS_URL, {
        method: 'POST',
        body: JSON.stringify({
          to: teacherEmails,
          subject: `[訓導處通知] 自動留堂名單更新 (${new Date().toLocaleDateString()})`,
          body: emailBody
        })
      });
      
      const result = await response.json();
      if (result.status === 'success') {
        alert('✅ 留堂通知已成功發送給老師！');
      } else {
        alert('❌ 發送失敗，請稍後再試。');
      }
    } catch (error) {
      console.error('Email error:', error);
      alert('❌ 發送時發生網路錯誤。');
    }
  };
  
const renderAnalysis = () => {
  // 👇 從這裡開始貼上匯出函數 👇
    const exportAnalyticsToCSV = () => {
      // 這裡使用 processedData (全部已處理資料)，若只要匯出目前搜尋結果可改為 displayData
      const dataToExport = processedData; 
      
      if (!dataToExport || dataToExport.length === 0) {
        alert("目前沒有資料可以匯出");
        return;
      }

      // 1. 定義標題列 (對應您的系統欄位)
      const headers = ['違規日期', '班別', '學生姓名', '學號', '違規事項', '記錄老師', '狀態'];
      
      // 2. 整理資料行
      const csvRows = dataToExport.map(record => {
        return [
          record.incidentDate || '',
          record.studentClass || '',
          record.studentName || '',
          record.studentId || '',
          record.rule || '',
          record.teacher || '',
          record.processStatus === 'Accepted' ? '已接納' : (record.processStatus === 'Rejected' ? '已拒絕' : '待處理')
        ].map(val => `"${val}"`).join(','); 
      });

      // 3. 組合 CSV 內容，加入 BOM 以支援中文
      const csvContent = "\uFEFF" + headers.join(',') + '\n' + csvRows.join('\n');
      
      // 4. 建立下載連結並觸發
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const today = new Date().toISOString().split('T')[0];
      link.setAttribute('download', `統計分析數據_${today}.csv`);
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    };
    // 👆 匯出函數結束 👆
const activeStartDate = dateRange?.start || importStartDate;
const activeEndDate = dateRange?.end || importEndDate;

const analyzedData = processedData.filter(r => {
  const isAccepted = r.processStatus === 'Accepted' || r.processStatus === '已接納';
  if (!isAccepted) return false;

  if (activeStartDate && r.incidentDate < activeStartDate) return false;
  if (activeEndDate && r.incidentDate > activeEndDate) return false;

  return true;
});
    
    const trendCounts = {};
    analyzedData.forEach(r => {
      trendCounts[r.incidentDate] = (trendCounts[r.incidentDate] || 0) + 1;
    });
    const trendStats = Object.keys(trendCounts).sort().map(date => ({
      date: date.substring(5),
      count: trendCounts[date]
    }));

    const classCounts = {};
    analyzedData.forEach(r => classCounts[r.studentClass] = (classCounts[r.studentClass] || 0) + 1);
    const classStats = Object.entries(classCounts).map(([name, count]) => ({ name, count })).sort((a,b) => b.count - a.count);
    
    const ruleCounts = {};
    analyzedData.forEach(r => ruleCounts[r.rule] = (ruleCounts[r.rule] || 0) + 1);
    const ruleStats = Object.entries(ruleCounts).map(([name, count]) => ({ name, count })).sort((a,b) => b.count - a.count);

const studentCounts = {};
analyzedData.forEach(r => {
  const key = `${r.studentClass}-${r.studentId}`;
  if (!studentCounts[key]) {
    studentCounts[key] = { 
      id: r.studentId, 
      name: r.studentName, 
      class: r.studentClass, 
      count: 0 
    };
  }
  studentCounts[key].count++;
});
    const topStudents = Object.values(studentCounts).sort((a,b) => b.count - a.count).slice(0, 10);

return (
      <div className="space-y-6">
        <div className="flex justify-between items-center mb-6 bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
          <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
             <BarChart2 className="w-6 h-6 text-indigo-500" />
             統計及趨勢分析
          </h2>
          
          <button 
            onClick={exportAnalyticsToCSV}
            className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-xl hover:bg-indigo-700 transition-colors text-sm font-bold shadow-sm"
          >
            <Download className="w-4 h-4" />
            匯出數據 (CSV)
          </button>
        </div>
        <div className="bg-indigo-50 border border-indigo-100 rounded-xl p-4 flex items-start gap-3">
          <TrendingUp className="w-5 h-5 text-indigo-600 mt-0.5 shrink-0" />
          <div>
            <h3 className="text-sm font-bold text-indigo-900">數據分析說明</h3>
            <p className="text-sm text-indigo-700 mt-1">此頁面僅統計系統中<strong>已接納 (Accepted)</strong>且未被刪除的記錄。數據範圍跟隨系統設定 ({importStartDate} 至 {importEndDate})。</p>
          </div>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <TrendingUp className="w-5 h-5 text-emerald-500" />
              違規趨勢 ({importStartDate} 至 {importEndDate})
            </h3>
            <button
              onClick={exportAnalyticsToExcel}
              className="flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl font-bold text-sm shadow-sm transition-all active:scale-95 shrink-0"
            >
              <Download className="w-4 h-4" />
              匯出趨勢分析 Excel (含圖表)
            </button>
          </div>
          <div ref={analyticsChartRef} className="h-72 w-full bg-white p-2">
            {trendStats.length === 0 ? (
              <div className="h-full flex items-center justify-center text-slate-400 text-sm">暫無趨勢數據</div>
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={trendStats} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                  <YAxis allowDecimals={false} axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                  <RechartsTooltip 
                    contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                    labelStyle={{ color: '#64748b', fontWeight: 'bold', marginBottom: '4px' }}
                  />
                  <Line type="monotone" dataKey="count" stroke="#10b981" strokeWidth={3} dot={{ r: 4, strokeWidth: 2 }} activeDot={{ r: 6 }} name="違規宗數" />
                </LineChart>
              </ResponsiveContainer>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6 flex flex-col">
            <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
              <Users className="w-5 h-5 text-blue-500" />
              各班違規次數 (最高至最低)
            </h3>
            <div className="h-64 w-full flex-1">
              {classStats.length === 0 ? (
                <div className="h-full flex items-center justify-center text-slate-400 text-sm">暫無數據</div>
              ) : (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={classStats} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                    <YAxis allowDecimals={false} axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <RechartsTooltip 
                      cursor={{ fill: '#f8fafc' }}
                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                    />
                    <Bar dataKey="count" fill="#3b82f6" radius={[4, 4, 0, 0]} name="違規次數" maxBarSize={50} />
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6 flex flex-col">
            <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
              <BarChart2 className="w-5 h-5 text-purple-500" />
              欠帶事項分佈
            </h3>
            <div className="w-full h-[300px] flex-1 overflow-y-auto pr-2 custom-scrollbar">
              {ruleStats.length === 0 ? (
                <div className="h-full flex items-center justify-center text-slate-400 text-sm">暫無數據</div>
              ) : (
                <ResponsiveContainer width="100%" height={Math.max(420, ruleStats.length * 36)}>
                  <BarChart 
                    layout="vertical" 
                    data={ruleStats.map(item => ({
                      ...item,
                      shortName: item.name.replace(/\([^)]*\)/g, '').trim()
                    }))}
                    margin={{ top: 5, right: 30, left: 10, bottom: 5 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                    <XAxis type="number" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <YAxis 
                      dataKey="shortName" 
                      type="category" 
                      width={180} 
                      axisLine={false} 
                      tickLine={false} 
                      tick={{ fontSize: 12, fill: '#334155', fontWeight: 500 }} 
                      interval={0} 
                    />
                    <RechartsTooltip
                      cursor={{ fill: '#f8fafc' }}
                      formatter={(value, name, props) => [value, props.payload.name]}
                      contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                    />
                    <Bar dataKey="count" fill="#6366f1" radius={[0, 6, 6, 0]} barSize={16} />
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </div>
        </div>

        {/* 違規次數最高學生 (首10名) */}
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6">
          <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
            <AlertTriangle className="w-5 h-5 text-rose-500" />
            違規次數最高學生 (首10名)
          </h3>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="text-xs uppercase tracking-wider text-slate-400 font-bold border-b border-slate-100">
                  <th className="pb-3 pr-4">排名</th>
                  <th className="pb-3 px-4">班別</th>
                  <th className="pb-3 px-4">學生姓名</th>
                  <th className="pb-3 px-4">學號</th>
                  <th className="pb-3 pl-4 text-right">有效違規次數</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50 text-sm">
                {topStudents.length === 0 ? (
                  <tr>
                    <td colSpan="5" className="py-8 text-center text-slate-400">暫無數據</td>
                  </tr>
                ) : topStudents.map((student, idx) => (
                  <tr key={student.id} className="hover:bg-slate-50/50 transition-colors">
                    <td className="py-3 pr-4 font-bold text-slate-500">
                      <span className={`inline-flex items-center justify-center w-6 h-6 rounded-full ${idx < 3 ? 'bg-rose-100 text-rose-600' : 'bg-slate-100 text-slate-500'}`}>
                        {idx + 1}
                      </span>
                    </td>
                    <td className="py-3 px-4 font-medium text-slate-700">
                      {student.class}
                    </td>
                    <td className="py-3 px-4 font-bold text-slate-800">
                      {student.name}
                    </td>
                    <td className="py-3 px-4 font-mono text-slate-500">
                      {student.id}
                    </td>
                    <td className="py-3 pl-4 text-right font-bold text-rose-600">
                      {student.count}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    ); 
  };

  const handleBatchMarkProcessedForStudent = async (studentRecords) => {
    const pendingRecords = studentRecords
      .filter(r => !processedRecordIds.has(r.id) && !deletedRecordIds.has(r.id))
      .slice(0, 5);

    if (pendingRecords.length === 0) return;

    const newProcessedIds = new Set(processedRecordIds);
    pendingRecords.forEach(r => newProcessedIds.add(r.id));
    setProcessedRecordIds(newProcessedIds);

    if (user) {
      for (const r of pendingRecords) {
        try {
          const docId = encodeURIComponent(r.id);
          const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'processed_records', docId);
          await setDoc(docRef, { recordId: r.id, processedAt: Date.now() });
        } catch (err) {
          console.error("標記已處理失敗:", err);
        }
      }
    }
  };

  const renderDetentionRoster = () => {
    const acceptedRecords = processedData.filter(r => r.processStatus === 'Accepted');
    const studentMap = {};
    
    acceptedRecords.forEach(r => {
      const uniqueKey = `${r.studentClass}-${r.studentId}`;
      if (!studentMap[uniqueKey]) {
        studentMap[uniqueKey] = {
          id: r.studentId,
          name: r.studentName,
          class: r.studentClass,
          records: []
        };
      }
      studentMap[uniqueKey].records.push(r);
    });

    const roster = Object.values(studentMap)
      .map(student => {
        const totalRecords = student.records.length;
        const milestonesReached = Math.floor(totalRecords / 5);
        return {
          ...student,
          totalRecords,
          milestonesReached,
          records: student.records.sort((a, b) => b.submitTime - a.submitTime)
        };
      })
      .filter(student => student.milestonesReached > 0)
      .sort((a, b) => b.totalRecords - a.totalRecords);

    return (
      <div className="space-y-6">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
          <div className="flex-1">
            <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
              <ShieldAlert className="w-6 h-6 text-rose-500" />
              留堂及老師通知名單 (滿 5 次自動列入)
            </h2>
            <p className="text-sm text-slate-500 mt-1">
              當學生累積達 5 次違規即自動進入此名單，後續每滿 5 次（10次、15次）系統將再次觸發。
            </p>
          </div>
          
          <div className="flex flex-col sm:flex-row gap-3 w-full lg:w-auto mt-4 lg:mt-0">
            <input
              type="text"
              placeholder="輸入老師 Email (多個用逗號分隔)"
              value={noticeEmails}
              onChange={(e) => setNoticeEmails(e.target.value)}
              className="border border-slate-200 rounded-xl px-4 py-2.5 text-sm w-full sm:w-72 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
            />
            <button
              onClick={() => exportTeacherNotice(roster)}
              className="flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-xl font-bold transition-all shadow-sm whitespace-nowrap"
            >
              <Download className="w-5 h-5" />
              匯出 Excel
            </button>
            {userRole === 'admin' && (
              <button
                onClick={() => {
                  if (window.confirm(`確定要發送通知信給：\n${noticeEmails || '未輸入信箱'}\n嗎？`)) {
                    handleSendEmailNotice(roster);
                  }
                }}
                className="flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2.5 rounded-xl font-bold transition-all shadow-sm whitespace-nowrap"
              >
                <Mail className="w-5 h-5" />
                自動發送
              </button>
            )}
          </div>
        </div>

        {roster.length === 0 ? (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-12 text-center">
            <CheckCircle className="w-12 h-12 text-emerald-400 mx-auto mb-3" />
            <h3 className="text-lg font-bold text-slate-700">目前無人需要留堂</h3>
            <p className="text-slate-500 mt-1">太棒了！指定範圍內沒有學生累積違規滿 5 次。</p>
          </div>
        ) : (
          <div className="grid grid-cols-1 gap-6">
            {roster.map(student => (
              <div key={student.id} className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
                <div className="bg-slate-50 border-b border-slate-100 p-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center font-bold text-xl">
                      {student.milestonesReached}
                    </div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800">
                        {student.class} {student.name} <span className="text-slate-500 text-sm font-mono font-medium ml-1">({student.id})</span>
                      </h3>
                      <p className="text-sm text-rose-600 font-semibold">
                        總共累積 {student.totalRecords} 次違規 (觸發 {student.milestonesReached} 次通知)
                      </p>
                    </div>
                  </div>
                  <button
                    onClick={() => handleBatchMarkProcessedForStudent(student.records)}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-xl text-xs font-bold transition-all shadow-sm active:scale-95 flex items-center gap-1.5 shrink-0"
                  >
                    <CheckCircle className="w-4 h-4" />
                    安排留堂 (標記 5 項為已處理)
                  </button>
                </div>

                <div className="p-4">
                  <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3">詳細違規紀錄</h4>
                  <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3">
                    {student.records.map((record, idx) => (
                      <div key={record.id} className="bg-white border border-slate-200 rounded-xl p-3 flex flex-col justify-between">
                        <div className="flex justify-between items-start mb-2">
                          <span className="bg-slate-100 text-slate-600 text-xs px-2 py-1 rounded font-medium">
                            #{student.records.length - idx}
                          </span>
                          <span className="text-xs font-medium text-slate-500">{record.incidentDate}</span>
                        </div>
                        <p className="font-bold text-slate-800 text-sm mb-1">{record.rule}</p>
                        <p className="text-xs text-slate-500 flex items-center gap-1">
                          <User className="w-3 h-3" /> {record.teacher}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="flex h-screen bg-slate-50 font-sans text-slate-900 overflow-hidden">
      {sidebarOpen && (
        <div 
          className="fixed inset-0 bg-slate-900/50 z-20 lg:hidden backdrop-blur-sm transition-opacity"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      <aside className={`fixed lg:static inset-y-0 left-0 w-72 bg-slate-900 text-slate-300 z-30 transform transition-transform duration-300 ease-in-out flex flex-col ${sidebarOpen ? 'translate-x-0' : '-translate-x-full lg:translate-x-0'}`}>
        <div className="p-6 flex items-center justify-between h-20 border-b border-slate-800">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-500 p-2 rounded-xl shadow-lg shadow-indigo-500/20">
              <ShieldAlert className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-white leading-tight">訓導數據中心</h1>
              <div className="flex items-center gap-1 text-[10px] text-emerald-400 uppercase tracking-wider font-semibold mt-0.5">
                <Cloud className="w-3 h-3" />
                <span>雲端多端同步中</span>
              </div>
            </div>
          </div>
          <button onClick={() => setSidebarOpen(false)} className="lg:hidden p-1 hover:bg-slate-800 rounded-md text-slate-400">
            <X className="w-5 h-5" />
          </button>
        </div>
        
        <div className="flex-1 overflow-y-auto py-6 px-4 space-y-2">
          <p className="px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2">主選單</p>
          
          {[
            { id: 'dashboard', label: '系統概覽', icon: LayoutDashboard },
            { id: 'records', label: '記錄及數據庫', icon: FileSpreadsheet },
            { id: 'analysis', label: '統計分析', icon: BarChart2 },
            { id: 'detentions', label: '留堂及通知名單', icon: AlertTriangle },
            { id: 'integration', label: '實時數據連接', icon: Settings },
          ].map((item) => {
            if (item.id === 'integration' && userRole !== 'admin') return null;
            const isActive = activeTab === item.id;
            return (
              <button
                key={item.id}
                onClick={() => { setActiveTab(item.id); setSidebarOpen(false); }}
                className={`w-full flex items-center justify-between px-3 py-3 rounded-xl transition-all ${
                  isActive 
                    ? 'bg-indigo-600 text-white shadow-md shadow-indigo-900/20' 
                    : 'text-slate-400 hover:bg-slate-800 hover:text-white'
                }`}
              >
                <div className="flex items-center gap-3">
                  <item.icon className={`w-5 h-5 ${isActive ? 'text-indigo-200' : ''}`} />
                  <span className="font-medium text-sm">{item.label}</span>
                </div>
                {isActive && <ChevronRight className="w-4 h-4 text-indigo-300" />}
              </button>
            );
          })}
        </div>

        <div className="p-4 border-t border-slate-800 space-y-3">
          <div className="flex items-center gap-3 p-3 bg-slate-800/50 rounded-xl">
            <div className="w-10 h-10 rounded-full bg-gradient-to-br from-indigo-400 to-purple-500 flex items-center justify-center text-white font-bold shadow-inner uppercase">
              W
            </div>
            <div className="flex-1 overflow-hidden">
              <p className="text-sm font-bold text-white truncate capitalize">WKC 教職員</p>
              <p className="text-[10px] text-slate-400 truncate">staff@wkc.edu.hk</p>
            </div>
          </div>
          <button 
            onClick={handleLogout} 
            className="w-full bg-rose-600/10 hover:bg-rose-600 text-rose-400 hover:text-white p-3 rounded-xl text-sm font-bold flex items-center justify-center gap-2 transition-all border border-rose-500/20"
          >
            <LogOut className="w-4 h-4" /> 登出系統
          </button>
        </div>
      </aside>

      <main className="flex-1 flex flex-col min-w-0 bg-slate-50">
        <header className="h-20 px-4 sm:px-8 bg-white border-b border-slate-200 flex items-center justify-between shrink-0 z-10 sticky top-0">
          <div className="flex items-center gap-4">
            <button 
              onClick={() => setSidebarOpen(true)}
              className="lg:hidden p-2 text-slate-500 hover:bg-slate-100 rounded-lg transition-colors"
            >
              <Menu className="w-6 h-6" />
            </button>
            <div>
              <h2 className="text-xl font-bold text-slate-800 capitalize hidden sm:block">
                {activeTab === 'dashboard' ? '系統數據概覽' : 
                 activeTab === 'records' ? '違規記錄數據庫' : 
                 activeTab === 'analysis' ? '統計及趨勢分析' : 
                 activeTab === 'detentions' ? '留堂與通知名單' : 'Google Sheet 整合設定'}
              </h2>
            </div>
          </div>

          <div className="flex items-center gap-4">
            {isLiveMode && (
              <div className="hidden sm:flex items-center text-xs font-semibold text-emerald-700 bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-100 shadow-sm">
                <span className="w-2 h-2 rounded-full bg-emerald-500 mr-2 animate-pulse"></span>
                雲端多端連線中
              </div>
            )}
            <button className="relative p-2 text-slate-400 hover:bg-slate-100 rounded-full transition-colors">
              <Bell className="w-5 h-5" />
              <span className="absolute top-1.5 right-1.5 w-2 h-2 bg-rose-500 rounded-full border-2 border-white"></span>
            </button>
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-4 sm:p-8">
          <div className="max-w-7xl mx-auto h-full">
            {activeTab === 'dashboard' && renderDashboard()}
            {activeTab === 'records' && renderRecords()}
            {activeTab === 'analysis' && renderAnalysis()}
            {activeTab === 'detentions' && renderDetentionRoster()}
            {activeTab === 'integration' && renderIntegration()}
          </div>
        </div>
      </main>

      {recordToDelete && (
        <div className="fixed inset-0 bg-slate-900/40 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-6 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center gap-3 text-rose-600 mb-4">
              <div className="bg-rose-100 p-2.5 rounded-full shrink-0">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <h3 className="text-lg font-bold text-slate-800 leading-tight">確認刪除此記錄？</h3>
            </div>
            <p className="text-sm text-slate-600 mb-6 leading-relaxed">
              您即將刪除學生 <strong>{recordToDelete.studentName}</strong> 於 <strong>{recordToDelete.incidentDate}</strong> 的 <strong>{recordToDelete.rule}</strong> 記錄。<br/><br/>
              此操作將從所有教職員的檢視中隱藏該記錄，且該記錄將<strong>不再計算</strong>在學生的每日欠帶限額內。
            </p>
            <div className="flex gap-3 justify-end">
              <button
                onClick={() => setRecordToDelete(null)}
                className="px-4 py-2.5 text-sm font-bold text-slate-600 hover:bg-slate-100 rounded-xl transition-colors active:scale-95"
              >
                取消
              </button>
              <button
                onClick={() => handleDeleteRecord(recordToDelete)}
                className="px-4 py-2.5 text-sm font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-sm shadow-rose-600/20 transition-all active:scale-95 flex items-center gap-2"
              >
                <Trash2 className="w-4 h-4" />
                確認刪除
              </button>
            </div>
          </div>
        </div>
      )}

      {userRole === 'admin' && showBulkDeleteConfirm && (
        <div className="fixed inset-0 bg-slate-900/40 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-6 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center gap-3 text-rose-600 mb-4">
              <div className="bg-rose-100 p-2.5 rounded-full shrink-0">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <h3 className="text-lg font-bold text-slate-800 leading-tight">確認批量刪除？</h3>
            </div>
            <p className="text-sm text-slate-600 mb-6 leading-relaxed">
              您即將刪除 <strong>{selectedRecordIds.size}</strong> 條已選取的記錄。<br/><br/>
              此操作將從所有教職員的檢視中隱藏這些記錄，並即時同步儲存至 Firebase 雲端。
            </p>
            <div className="flex gap-3 justify-end">
              <button
                onClick={() => setShowBulkDeleteConfirm(false)}
                className="px-4 py-2.5 text-sm font-bold text-slate-600 hover:bg-slate-100 rounded-xl transition-colors active:scale-95"
              >
                取消
              </button>
              <button
                onClick={handleBulkDelete}
                className="px-4 py-2.5 text-sm font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-sm shadow-rose-600/20 transition-all active:scale-95 flex items-center gap-2"
              >
                <Trash2 className="w-4 h-4" />
                確認刪除全部
              </button>
            </div>
          </div>
        </div>
      )}

      {showResetConfirm && (
        <div className="fixed inset-0 bg-slate-900/40 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-6 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center gap-3 text-rose-600 mb-4">
              <div className="bg-rose-100 p-2.5 rounded-full shrink-0">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <h3 className="text-lg font-bold text-slate-800 leading-tight">清空雲端同步狀態？</h3>
            </div>
            <p className="text-sm text-slate-600 mb-6 leading-relaxed">
              此操作將會刪除 Firebase 中儲存的所有「已處理」與「已刪除」註記，還原至 Google Sheet 的原始數據狀態。
            </p>
            <div className="flex gap-3 justify-end">
              <button
                onClick={() => setShowResetConfirm(false)}
                className="px-4 py-2.5 text-sm font-bold text-slate-600 hover:bg-slate-100 rounded-xl transition-colors active:scale-95"
              >
                取消
              </button>
              <button
                onClick={handleResetCloudData}
                disabled={isResetting}
                className="px-4 py-2.5 text-sm font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-sm shadow-rose-600/20 transition-all active:scale-95 flex items-center gap-2"
              >
                <Trash2 className="w-4 h-4" />
                {isResetting ? '清空中...' : '確認清空雲端數據'}
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}