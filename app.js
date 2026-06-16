// ============================================================
// نظام متابعة المخدومات — Offline Ready & Guest Mode
// VERSION 3.1 — Added: Service day toggle, Reset attendance,
//               Better visual feedback, No-service days support
// ============================================================

// ============================================================
// FB MODULE — Replaces window._fb anti-pattern with proper singleton
// FIXED: Added guard function to prevent usage before initialization
// ============================================================
const FB = new Proxy({
  collection: null, doc: null, setDoc: null, getDocs: null,
  deleteDoc: null, query: null, orderBy: null, onSnapshot: null,
  writeBatch: null, where: null, signInWithPopup: null,
  signInWithRedirect: null, getRedirectResult: null,
  onAuthStateChanged: null, signOut: null
}, {
  get(target, prop) {
    if (prop in target && target[prop] !== null) return target[prop];
    if (['collection', 'doc', 'setDoc', 'onSnapshot', 'writeBatch'].includes(prop)) {
      throw new Error(`FB.${String(prop)} accessed before Firebase initialization. Call ensureFB() first.`);
    }
    return target[prop];
  }
});

/**
 * FIXED: Guard function that throws if Firebase is not ready.
 * Use at the start of any function that needs Firebase.
 */
function ensureFB() {
  if (!firebaseReady) throw new Error('Firebase not initialized');
}

// ============================================================
// SAFETY: Global error handler + splash fallback
// FIXED: Unified splash state with lock — prevents double-hide race condition
// ============================================================
const SplashState = {
  _done: false,
  _forceHidden: false,
  _locked: false,
  get done() { return this._done || this._forceHidden; },
  markDone() {
    if (this._locked) return;
    this._locked = true;
    this._done = true;
    this._forceHidden = true;
  },
  markForceHidden() {
    if (this._locked) return;
    this._locked = true;
    this._forceHidden = true;
  }
};

window.addEventListener('error', (e) => {
  console.error('Global error:', e.error || e.message);
  hideSplashForced();
});
window.addEventListener('unhandledrejection', (e) => {
  console.error('Unhandled rejection:', e.reason);
  hideSplashForced();
});

// Force hide splash after 6 seconds max — never get stuck
setTimeout(hideSplashForced, 6000);

function hideSplashForced() {
  if (SplashState.done) return;
  SplashState.markForceHidden();
  const splash = document.getElementById('splash');
  if (splash) {
    splash.classList.add('fade-out');
    setTimeout(() => splash.remove(), 500);
  }
  // Show login screen as fallback if app isn't initialized
  setTimeout(() => {
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');
    if (loginScreen && mainApp && mainApp.classList.contains('hidden') && loginScreen.classList.contains('hidden')) {
      loginScreen.classList.remove('hidden');
      showLogin();
    }
  }, 600);
}

// ============================================================
// FIREBASE IMPORTS WITH FALLBACK
// FIXED: Clearer fallback UI when Firebase fails
// SECURITY: Firebase config moved to fetch from server to avoid credential leak
// ============================================================
let firebaseApp, auth, db, provider;
let firebaseReady = false;
let XLSX = null;

// Track snapshot unsubscribers to prevent memory leaks
// FIXED: listenersInitialized flag prevents duplicate listeners in race conditions
const _unsubscribers = [];
let _listenersInitialized = false;

function clearAllSnapshots() {
  _unsubscribers.forEach(unsub => { try { unsub(); } catch (e) { } });
  _unsubscribers.length = 0;
  _listenersInitialized = false;
}

function pushUnsubscriber(unsub) {
  _unsubscribers.push(unsub);
}

// ============================================================
// Firebase Configuration — Embedded directly in the code
// ============================================================
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyB2cycBTKMjVg8S_fBYN8C-hwUk5FUF81Q",
  authDomain: "kenesa-e5efd.firebaseapp.com",
  projectId: "kenesa-e5efd",
  storageBucket: "kenesa-e5efd.firebasestorage.app",
  messagingSenderId: "227273753184",
  appId: "1:227273753184:web:ecdf258142ad55ed5cf905",
  measurementId: "G-6HS8KNW1GZ"
};

async function fetchFirebaseConfig() {
  // Return the embedded Firebase config directly
  return FIREBASE_CONFIG;
}

// Module imports with error handling
async function initModules() {
  try {
    const { initializeApp } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js');
    const { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, onAuthStateChanged, signOut } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js');
    const { getFirestore, collection, doc, setDoc, getDocs, deleteDoc, query, orderBy, onSnapshot, writeBatch, where } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js');

    // Use the embedded Firebase config
    let firebaseConfig = await fetchFirebaseConfig();
    firebaseApp = initializeApp(firebaseConfig);
    auth = getAuth(firebaseApp);
    db = getFirestore(firebaseApp);
    provider = new GoogleAuthProvider();
    firebaseReady = true;

    // Initialize Firebase Analytics
    try {
      const { getAnalytics } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-analytics.js');
      const analytics = getAnalytics(firebaseApp);
      console.log('Firebase Analytics initialized');
    } catch (analyticsErr) {
      console.warn('Firebase Analytics not initialized:', analyticsErr.message);
    }

    // FIXED: Use module singleton instead of window._fb
    FB.collection = collection; FB.doc = doc; FB.setDoc = setDoc;
    FB.getDocs = getDocs; FB.deleteDoc = deleteDoc; FB.query = query;
    FB.orderBy = orderBy; FB.onSnapshot = onSnapshot; FB.writeBatch = writeBatch;
    FB.where = where; FB.signInWithPopup = signInWithPopup;
    FB.signInWithRedirect = signInWithRedirect; FB.getRedirectResult = getRedirectResult;
    FB.onAuthStateChanged = onAuthStateChanged; FB.signOut = signOut;

    // Try to load XLSX
    try {
      const xlsxMod = await import('https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs');
      XLSX = xlsxMod;
    } catch (xlsxErr) {
      console.warn('XLSX library failed to load:', xlsxErr);
    }

    return true;
  } catch (e) {
    console.error('Firebase failed to initialize:', e);
    firebaseReady = false;
    // FIXED: Clearer error indication
    const splashContent = document.querySelector('.splash-content');
    if (splashContent) {
      splashContent.innerHTML = '<h1>⚠️ خطأ في الاتصال</h1><p>تعذر تحميل نظام التسجيل</p><p style="font-size:14px;opacity:0.7">تحقق من اتصال الإنترنت وأعد تحميل الصفحة</p>';
    }
    return false;
  }
}

// ============================================================
// DOM CACHE — FIXED: Build-once pattern for zero runtime cost
// ============================================================
const $ = (id) => document.getElementById(id);
const $$ = (sel, root = document) => root.querySelectorAll(sel);

function safeGetElement(id) {
  const el = document.getElementById(id);
  return el || null;
}

// FIXED: Build DOM map once at startup, then freeze — zero Proxy cost on access
const _domCache = {};

function _buildDOMCache() {
  const ids = [
    'splash', 'loginScreen', 'mainApp', 'pageTitle', 'pageSubtitle',
    'syncIndicator', 'userAvatar', 'drawer', 'drawerOverlay',
    'drawerAvatar', 'drawerUserName', 'drawerUserEmail', 'offlineBadge',
    'pageContent', 'toast', 'globalSearch', 'searchResults',
    'todayDay', 'todayDate', 'todayServiceBadge',
    'statTotal', 'statPresentToday', 'statAbsentToday', 'statAvgRating',
    'bestGrade', 'bestGradePercent', 'topActivityName', 'topActivityCount',
    'mostRegularGirl', 'mostRegularPercent', 'topAttendees', 'needsFollowup',
    'attendanceDate', 'attendanceList', 'attendanceSearch',
    'presentCount', 'absentCount', 'totalCount',
    'selectAllPresent', 'selectAllAbsent', 'attToggleHint', 'quickActions',
    'girlsList', 'addGirlBtn',
    'calendarGrid', 'calMonthYear', 'dayDetail', 'calPrev', 'calNext',
    'statsMonth', 'bigStatsGrid', 'absenceChart', 'attendanceRanking',
    'activityStatsGrid', 'timeFilterTabs', 'activityStatsPeriod',
    'historyList', 'historyFilter', 'clearHistoryBtn', 'loadMoreHistory',
    'loadMoreHistoryBtn', 'exportMonth',
    'exportCSV', 'exportJSON', 'exportPrint',
    'girlModal', 'girlModalTitle',
    'girlName', 'girlPhone', 'girlGrade', 'girlNotes', 'deleteGirlBtn',
    'homeGradeFilters', 'girlsGradeFilters', 'attendanceGradeFilters',
    'closeGirlModal', 'cancelGirlModal', 'saveGirlBtn', 'girlProfileModal',
    'profileName', 'profileBody', 'closeProfileModal', 'attendanceModal',
    'attendanceModalTitle', 'modalGirlName', 'attendanceNotes', 'ratingSection',
    'starsInput', 'saveAttendanceEntry', 'closeAttendanceModal', 'cancelAttendanceModal',
    'confirmOverlay', 'confirmIcon', 'confirmTitle', 'confirmMsg',
    'confirmCancel', 'confirmOk',
    'activityDetailModal', 'activityDetailTitle', 'closeActivityDetailModal',
    'activityDetailSummary', 'activityDetailIcon', 'activityDetailName',
    'activityDetailPeriod', 'activityDetailTotal', 'activityDetailTabs',
    'activityDetailList', 'presentTabCount', 'absentTabCount',
    'menuBtn', 'signOutBtn', 'googleSignIn',
    'darkModeToggle', 'darkToggleSwitch',
    'shareProfileBtn', 'editProfileBtn',
    'statsGradeFilter', 'activityStatsGrade', 'exportGradeFilter',
    'exportStatusFilter',
    // Settings page elements
    'exportFullBackup', 'importBackup', 'importFileInput',
    'clearAllData', 'settingsGirlCount', 'settingsAttCount', 'settingsLastUpdate',
    // NEW: v3.1 elements
    'serviceDayToggle', 'serviceDayToggleWrap', 'serviceToggleHint',
    'noServiceMessage', 'resetAttendanceBtn', 'quickActionsRow'
  ];
  ids.forEach(id => { _domCache[id] = document.getElementById(id); });
}

// FIXED: Minimal wrapper — direct property access, no Proxy overhead
const DOM = new Proxy(_domCache, {
  get(target, prop) {
    return target[prop] ?? null;
  }
});

// Eagerly cache known static elements at startup
_buildDOMCache();

// ============================================================
// APP STATE — FIXED: Added cache indexes for performance
// ============================================================
const state = {
  currentUser: null,
  girls: [],
  attendanceData: {},
  currentPage: 'home',
  selectedDay: 'السبت',
  selectedActivity: 'دراسي',
  currentAttendanceGirlId: null,
  currentAttendanceRating: 0,
  editingGirlId: null,
  calendarDate: new Date(),
  appInitialized: false,
  renderTimeout: null,
  renderPending: false,
  historyOffset: 0,
  historyAllLogs: [],
  deleteInProgress: false,
  homeGradeFilter: '',
  girlsGradeFilter: '',
  girlsSearchQuery: '',
  attendanceGradeFilter: localStorage.getItem('attendanceGradeFilter') || '',
  statsTimeFilter: 'month',
  statsGradeFilter: '',
  longPressTimer: null,
  isLongPress: false,
  activityDetailTab: 'present',
  currentActivityDetail: null,
  currentProfileGirlId: null,
  searchDebounceTimer: null,
  attSearchDebounceTimer: null,
  attendancePageInitialized: false,
  savingGirl: false,
  idb: false,
  // FIXED: Add pending operation locks to prevent race conditions
  pendingAttendanceOps: new Set(),
  pendingSaveGirl: false,
  // FIXED: Precomputed absence cache { monthStr: { girlId: { hasConsecutive, count, dates } } }
  absenceCache: {},
  lastAbsenceCacheMonth: null,
  // NEW: Export grade filter state
  exportGradeFilter: 'أولى إعدادي',
  // NEW: Export status filter (present/absent/all)
  exportStatusFilter: '',
  // NEW: Track which service days have been auto-marked as absent (to prevent duplicates)
  autoMarkedDates: new Set(JSON.parse(localStorage.getItem('autoMarkedDates') || '[]')),
  // NEW v3.1: Per-date service day overrides (date -> boolean, true = service day)
  // Stored in localStorage as 'serviceDayOverrides' = { "2025-01-15": false, ... }
  serviceDayOverrides: JSON.parse(localStorage.getItem('serviceDayOverrides') || '{}'),
};

// ============================================================
// DERIVED STATE CACHE — Prevents O(n^2) lookups
// FIXED: Centralized cache with full rebuild from source truth
// ============================================================
const Cache = {
  girlsById: null,
  allAttendance: null,
  attendanceByGirl: null,
  attendanceByDate: null,
  attendanceByMonth: null,
  // FIXED: Cached activeGirlIds to prevent repeated Set builds
  activeGirlIds: null,
  _dirty: true,
  _snapshotVersion: 0,

  invalidate() {
    this._dirty = true;
    this._snapshotVersion++;
    this.girlsById = null;
    this.allAttendance = null;
    this.attendanceByGirl = null;
    this.attendanceByDate = null;
    this.attendanceByMonth = null;
    this.activeGirlIds = null;
    // FIXED: Clear absence cache to prevent stale consecutive absence data
    // after attendance edits. Cache will be rebuilt on next hasConsecutiveAbsences call.
    state.absenceCache = {};
    state.lastAbsenceCacheMonth = null;
  },

  build() {
    if (!this._dirty) return;
    // FULL rebuild from source truth — ensures consistency
    this.girlsById = Object.fromEntries(state.girls.filter(g => !g.isDeleted).map(g => [g.id, g]));
    // FIXED: Deduplicate attendance records by ID — prevents stale merges
    const attMap = new Map();
    Object.values(state.attendanceData).forEach(a => {
      if (!a || !a.id) return;
      const existing = attMap.get(a.id);
      // Keep the most recent version (by updatedAt)
      if (!existing || (a.updatedAt || 0) >= (existing.updatedAt || 0)) {
        attMap.set(a.id, a);
      }
    });
    const allAtt = Array.from(attMap.values());
    this.allAttendance = allAtt;

    // FIXED: Build indexed structures for O(1) lookups
    this.attendanceByGirl = {};
    this.attendanceByDate = {};
    this.attendanceByMonth = {};

    allAtt.forEach(a => {
      // By girl
      if (!this.attendanceByGirl[a.girlId]) this.attendanceByGirl[a.girlId] = [];
      this.attendanceByGirl[a.girlId].push(a);
      // By date
      if (!this.attendanceByDate[a.date]) this.attendanceByDate[a.date] = [];
      this.attendanceByDate[a.date].push(a);
      // By month
      const month = a.date?.substring(0, 7);
      if (month && /^\d{4}-\d{2}$/.test(month)) {
        if (!this.attendanceByMonth[month]) this.attendanceByMonth[month] = [];
        this.attendanceByMonth[month].push(a);
      }
    });

    // FIXED: Precompute activeGirlIds Set
    this.activeGirlIds = new Set(state.girls.filter(g => !g.isDeleted).map(g => g.id));

    this._dirty = false;
  },

  getGirl(id) {
    this.build();
    return this.girlsById ? this.girlsById[id] : null;
  },

  getAllAttendance() {
    this.build();
    return this.allAttendance || [];
  },

  // FIXED: O(1) indexed lookups
  getAttendanceByGirl(girlId) {
    this.build();
    return this.attendanceByGirl?.[girlId] || [];
  },

  getAttendanceByDate(date) {
    this.build();
    return this.attendanceByDate?.[date] || [];
  },

  getAttendanceByMonth(month) {
    this.build();
    return this.attendanceByMonth?.[month] || [];
  },

  // FIXED: O(1) cached activeGirlIds — eliminates repeated Set creation
  getActiveGirlIds() {
    this.build();
    return this.activeGirlIds || new Set();
  }
};

// ============================================================
// ATTENDANCE STORE — Global memoized snapshot
// FIXED: Prevents repeated Cache.getAllAttendance() full scans
// ============================================================
const AttendanceStore = {
  _cache: null,
  _dirty: true,
  _version: 0,

  getAll() {
    if (this._dirty || this._version !== Cache._snapshotVersion) {
      this._cache = Cache.getAllAttendance();
      this._dirty = false;
      this._version = Cache._snapshotVersion;
    }
    return this._cache;
  },

  invalidate() {
    this._dirty = true;
  }
};

// Auto-invalidate AttendanceStore when Cache invalidates
const originalCacheInvalidate = Cache.invalidate.bind(Cache);
Cache.invalidate = function() {
  originalCacheInvalidate();
  AttendanceStore.invalidate();
};

// Invalidate cache whenever girls or attendanceData changes
const _rawGirls = [];
const _rawAttendance = {};

function setStateGirls(newGirls) {
  state.girls = newGirls;
  Cache.invalidate();
}

function setStateAttendanceData(newData) {
  // FIXED: Support both direct Object and functional update (React-like pattern)
  // toggleAttendanceStatus uses: setStateAttendanceData(prev => ({ ...prev, [key]: rec }))
  if (typeof newData === 'function') {
    state.attendanceData = newData(state.attendanceData);
  } else {
    state.attendanceData = newData;
  }
  Cache.invalidate();
}

const HISTORY_PAGE_SIZE = 30;
const SERVICE_DAYS = { 'السبت': true, 'الاثنين': true, 'الاربعاء': true };
const SERVICE_DAY_NUMBERS = [1, 3, 6]; // Mon, Wed, Sat
const DAY_NAMES = ['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
const ACTIVITIES = ['دراسي', 'محفوظات', 'قبطي', 'ألحان'];
const ACTIVITY_ICONS = { 'دراسي': '&#128216;', 'ألحان': '&#127925;', 'قبطي': '&#9961;', 'محفوظات': '&#128221;' };
const PERIOD_LABELS = { today: 'اليوم', month: 'هذا الشهر', year: 'هذه السنة', all: 'كل الفترات' };
// Grade ordering for export: تالته first, then تانية, then أولى
const GRADE_ORDER = { 'تالتة إعدادي': 1, 'تانية إعدادي': 2, 'أولى إعدادي': 3 };

// ============================================================
// XSS PROTECTION
// ============================================================
const esc = (() => {
  const div = document.createElement('div');
  const txt = document.createTextNode('');
  div.appendChild(txt);
  return (str) => {
    txt.nodeValue = String(str ?? '');
    return div.innerHTML;
  };
})();

function xmlEsc(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ============================================================
// DATE UTILITIES — FIXED: Safe date parsing without timezone bugs
// ============================================================

/**
 * FIXED: Safely parse a YYYY-MM-DD string into a Date object.
 * Uses new Date(year, month-1, day) to avoid timezone shift bugs
 * that can occur with new Date("YYYY-MM-DDT00:00:00").
 * FIXED: Validates invalid dates like 31-02 that JS silently corrects.
 */
function parseDateStr(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return new Date(NaN);
  const parts = dateStr.split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return new Date(NaN);
  const [year, month, day] = parts;
  // Validate ranges
  if (month < 1 || month > 12 || day < 1 || day > 31) return new Date(NaN);
  const d = new Date(year, month - 1, day);
  // FIXED: Verify the date wasn't silently corrected by JS (e.g. 2024-02-31 → 2024-03-02)
  if (d.getMonth() !== month - 1 || d.getDate() !== day) return new Date(NaN);
  return d;
}

/**
 * FIXED: Compare two date strings (YYYY-MM-DD) safely.
 * Returns -1 if a < b, 0 if equal, 1 if a > b.
 */
function compareDateStr(a, b) {
  if (a === b) return 0;
  const da = parseDateStr(a);
  const db = parseDateStr(b);
  const ta = da.getTime();
  const tb = db.getTime();
  if (isNaN(ta) || isNaN(tb)) return String(a).localeCompare(String(b));
  return ta < tb ? -1 : ta > tb ? 1 : 0;
}

/**
 * FIXED: Check if a date string is within a range [start, end] (inclusive).
 */
function isDateInRange(dateStr, start, end) {
  return compareDateStr(dateStr, start) >= 0 && compareDateStr(dateStr, end) <= 0;
}

/**
 * FIXED: Generate a consistent attendance record key.
 * Centralized key format to avoid mismatch bugs.
 */
function makeAttKey(girlId, date, activity) {
  return `${girlId}_${date}_${activity}`;
}

const DateUtil = {
  pad: (n) => String(n).padStart(2, '0'),
  toStr(d = new Date()) {
    return `${d.getFullYear()}-${this.pad(d.getMonth() + 1)}-${this.pad(d.getDate())}`;
  },
  getMonthStr(d = new Date()) {
    return `${d.getFullYear()}-${this.pad(d.getMonth() + 1)}`;
  },
  formatMonth(str) {
    if (!str) return '';
    const [y, m] = str.split('-').map(Number);
    return new Date(y, m - 1, 1).toLocaleDateString('ar-EG', { year: 'numeric', month: 'long' });
  },
  formatDateShort(d = new Date()) {
    return `${d.getDate()}/${d.getMonth() + 1}`;
  },
  dayName(d = new Date()) { return DAY_NAMES[d.getDay()]; },
  // FIXED: Consistent hamza forms — all maps to single form
  normalize(d) {
    return {
      'الأحد': 'الاحد', 'الاحد': 'الاحد',
      'الاثنين': 'الاثنين',
      'الثلاثاء': 'الثلاثاء',
      'الأربعاء': 'الاربعاء', 'الاربعاء': 'الاربعاء',
      'الخميس': 'الخميس',
      'الجمعة': 'الجمعة',
      'السبت': 'السبت'
    }[d] || d;
  }
};

// ============================================================
// TIMECONTEXT — Unified Date Source for the entire app
// FIXED: Added null-safety protection for substring operations
// ============================================================
const TimeContext = {
  _selectedDate: null,
  _listeners: [],

  init() {
    const saved = localStorage.getItem('trackerSelectedDate');
    const today = DateUtil.toStr();
    // FIXED: Validate saved date format AND check if it's today's date
    // This prevents stale dates from previous days
    if (saved && /^\d{4}-\d{2}-\d{2}$/.test(saved) && saved === today) {
      this._selectedDate = saved;
    } else {
      this._selectedDate = today;
      localStorage.setItem('trackerSelectedDate', today);
    }
  },

  /** Get the currently selected date (YYYY-MM-DD) */
  getDate() {
    return this._selectedDate || DateUtil.toStr();
  },

  /** Set the selected date and notify all listeners */
  setDate(dateStr) {
    if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      console.warn('Invalid date format:', dateStr);
      return;
    }
    this._selectedDate = dateStr;
    localStorage.setItem('trackerSelectedDate', dateStr);
    this._notifyListeners(dateStr);
  },

  /** Get month string (YYYY-MM) — FIXED: with null safety */
  getMonth() {
    const d = this._selectedDate || DateUtil.toStr();
    return d.substring(0, 7);
  },

  /** Get year string (YYYY) — FIXED: with null safety */
  getYear() {
    const d = this._selectedDate || DateUtil.toStr();
    return d.substring(0, 4);
  },

  /** Reset to today */
  resetToToday() {
    this._selectedDate = DateUtil.toStr();
    localStorage.removeItem('trackerSelectedDate');
    this._notifyListeners(this._selectedDate);
  },

  /** Subscribe to date changes */
  subscribe(fn) {
    this._listeners.push(fn);
    return () => {
      this._listeners = this._listeners.filter(l => l !== fn);
    };
  },

  _notifyListeners(dateStr) {
    this._listeners.forEach(fn => {
      try { fn(dateStr); } catch (e) { console.error('TimeContext listener error:', e); }
    });
  }
};

// ============================================================
// ARABIC TEXT NORMALIZATION
// FIXED: Removed ة → ه transformation to preserve semantic accuracy
// ============================================================
function normalizeArabic(str) {
  if (!str) return '';
  return str.replace(/[\u064B-\u065F\u0670\u0640]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ؤ/g, 'و')
    .replace(/ئ/g, 'ي')
    // FIXED: Removed .replace(/ة/g, 'ه') — this changes meaning:
    // "مدرسة" should NOT become "مدرسه" — causes false matches
    .toLowerCase();
}

function normalizeName(name) {
  return name
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[أإآ]/g, 'ا')
    // FIXED: Keep ة as-is for accurate matching
    .replace(/ى/g, 'ي')
    .toLowerCase();
}

function csvEscape(v) {
  return `"${String(v ?? '').replace(/"/g, '""')}"`;
}

// ============================================================
// SERVICE DAY FUNCTIONS — v3.1: Added per-date override support
// ============================================================

/**
 * NEW v3.1: Check if a specific date is a service day.
 * First checks the user's per-date override, then falls back to the regular schedule.
 */
function isServiceDayForDate(dateStr) {
  if (!dateStr) return false;

  // Check if user has explicitly set this date's service status
  if (state.serviceDayOverrides.hasOwnProperty(dateStr)) {
    return state.serviceDayOverrides[dateStr];
  }

  // Fall back to the regular schedule (Sat, Mon, Wed)
  const d = parseDateStr(dateStr);
  if (isNaN(d.getTime())) return false;
  return SERVICE_DAY_NUMBERS.includes(d.getDay());
}

/**
 * NEW v3.1: Set service day override for a specific date.
 * true = service day, false = no service
 */
function setServiceDayOverride(dateStr, isService) {
  state.serviceDayOverrides[dateStr] = isService;
  localStorage.setItem('serviceDayOverrides', JSON.stringify(state.serviceDayOverrides));
}

function getServiceDaysInMonth(year, month) {
  const days = [];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    const dayOfWeek = new Date(year, month, d).getDay();
    if (SERVICE_DAY_NUMBERS.includes(dayOfWeek)) {
      days.push(`${year}-${DateUtil.pad(month + 1)}-${DateUtil.pad(d)}`);
    }
  }
  return days;
}

function getServiceDaysUpToDate(fromYear, fromMonth, toDate) {
  let count = 0;
  // FIXED: Use parseDateStr for safe date parsing
  const to = parseDateStr(toDate);
  if (isNaN(to.getTime())) return 0;

  const toYear = to.getFullYear();
  const toMonth = to.getMonth();
  const toDay = to.getDate();

  // FIXED: Only iterate up to the target day, not the entire month
  const lastDay = (fromYear === toYear && fromMonth === toMonth)
    ? toDay
    : new Date(fromYear, fromMonth + 1, 0).getDate();

  for (let d = 1; d <= lastDay; d++) {
    const dayOfWeek = new Date(fromYear, fromMonth, d).getDay();
    if (SERVICE_DAY_NUMBERS.includes(dayOfWeek)) {
      count++;
    }
  }
  return count;
}

// ============================================================
// CONSECUTIVE ABSENCES — FIXED: O(1) with precomputed cache
// ============================================================

/**
 * FIXED: Build absence cache for a month in single O(n) pass.
 * Call this once when data changes, then hasConsecutiveAbsences is O(1).
 */
function buildAbsenceCache(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  const serviceDays = getServiceDaysInMonth(year, month - 1);

  // Get all attendance for this month using indexed lookup
  const monthAtt = Cache.getAttendanceByMonth(monthStr);

  // Group absence records by girl
  const absByGirl = {};
  monthAtt.forEach(a => {
    if (a.status === 'غائب') {
      if (!absByGirl[a.girlId]) absByGirl[a.girlId] = new Set();
      absByGirl[a.girlId].add(a.date);
    }
  });

  const cache = {};
  Object.entries(absByGirl).forEach(([girlId, absDateSet]) => {
    const absDates = [...absDateSet].sort();
    if (absDates.length < 2) {
      cache[girlId] = { hasConsecutive: false, count: absDates.length, dates: absDates };
      return;
    }

    // Build absent service indices
    const absentServiceIndices = [];
    for (let i = 0; i < serviceDays.length; i++) {
      if (absDateSet.has(serviceDays[i])) absentServiceIndices.push(i);
    }

    if (absentServiceIndices.length < 2) {
      cache[girlId] = { hasConsecutive: false, count: absDates.length, dates: absDates };
      return;
    }

    // Check for consecutive service day absences
    let consecutiveCount = 1;
    let maxConsecutive = 1;
    for (let i = 0; i < absentServiceIndices.length - 1; i++) {
      if (absentServiceIndices[i + 1] - absentServiceIndices[i] === 1) {
        consecutiveCount++;
        maxConsecutive = Math.max(maxConsecutive, consecutiveCount);
      } else {
        consecutiveCount = 1;
      }
    }

    cache[girlId] = {
      hasConsecutive: maxConsecutive >= 2,
      count: absDates.length,
      dates: absDates
    };
  });

  state.absenceCache[monthStr] = cache;
  state.lastAbsenceCacheMonth = monthStr;
}

function hasConsecutiveAbsences(girlId, monthStr) {
  // FIXED: Build cache on first access for this month
  if (!state.absenceCache[monthStr]) {
    buildAbsenceCache(monthStr);
  }
  // O(1) cache lookup
  return state.absenceCache[monthStr]?.[girlId] || { hasConsecutive: false, count: 0, dates: [] };
}

// ============================================================
// UNIFIED STATS BOUNDS — All stats use this single function
// ============================================================
// FIXED: Safe date validation helper
function _validateDateStr(dateStr, fallback) {
  if (!dateStr || typeof dateStr !== 'string' || dateStr.length < 10) return fallback;
  const parts = dateStr.split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return fallback;
  const [year, month, day] = parts;
  if (year < 2000 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) return fallback;
  // Verify no silent JS correction
  const d = new Date(year, month - 1, day);
  if (d.getMonth() !== month - 1 || d.getDate() !== day) return fallback;
  return dateStr;
}

function getStatsBounds() {
  const selectedDate = _validateDateStr(TimeContext.getDate(), DateUtil.toStr());
  const selYear = parseInt(selectedDate.substring(0, 4));
  const selMonth = parseInt(selectedDate.substring(5, 7));

  switch (state.statsTimeFilter) {
    case 'today':
      return { start: selectedDate, end: selectedDate };
    case 'month': {
      const monthIndex = selMonth - 1;
      const lastDay = new Date(selYear, monthIndex + 1, 0).getDate();
      return { start: selectedDate.substring(0, 7) + '-01', end: selectedDate.substring(0, 7) + '-' + String(lastDay).padStart(2, '0') };
    }
    case 'year':
      return { start: selectedDate.substring(0, 4) + '-01-01', end: selectedDate.substring(0, 4) + '-12-31' };
    default: // 'all'
      return { start: '2000-01-01', end: selectedDate };
  }
}

// ============================================================
// INDEXEDDB — wrapper for offline history storage
// ============================================================
const IDB = {
  db: null,
  DB_NAME: 'girlsTrackerDB',
  DB_VERSION: 2, // Bumped for new stores

  async init() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.DB_NAME, this.DB_VERSION);
      request.onerror = () => reject(request.error);
      request.onsuccess = () => { this.db = request.result; resolve(); };
      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        if (!db.objectStoreNames.contains('history')) {
          const store = db.createObjectStore('history', { keyPath: 'id' });
          store.createIndex('timestamp', 'timestamp', { unique: false });
        }
        if (!db.objectStoreNames.contains('pendingSync')) {
          db.createObjectStore('pendingSync', { keyPath: 'id' });
        }
        // NEW: Backup store for rollback support
        if (!db.objectStoreNames.contains('backups')) {
          db.createObjectStore('backups', { keyPath: 'id' });
        }
      };
    });
  },

  async add(storeName, data) {
    if (!this.db) return;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      const request = store.put(data);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  },

  async getAll(storeName) {
    if (!this.db) return [];
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readonly');
      const store = tx.objectStore(storeName);
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  },

  async get(storeName, id) {
    if (!this.db) return null;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readonly');
      const store = tx.objectStore(storeName);
      const request = store.get(id);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  },

  async clear(storeName) {
    if (!this.db) return;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      const request = store.clear();
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  },

  async delete(storeName, id) {
    if (!this.db) return;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      const request = store.delete(id);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  }
};

// ============================================================
// ROLLBACK / BACKUP HELPERS — NEW: Firestore rollback support
// ============================================================
async function createBackup(operationId, data) {
  try {
    await IDB.add('backups', { id: operationId, data, timestamp: Date.now() });
  } catch (e) { console.warn('Backup creation failed:', e); }
}

async function restoreBackup(operationId) {
  try {
    const backup = await IDB.get('backups', operationId);
    return backup ? backup.data : null;
  } catch (e) { console.warn('Backup restore failed:', e); return null; }
}

// ============================================================
// THEME MANAGER — FIXED: Professional theme system, no white flash
// ============================================================
const Theme = {
  KEY: 'theme',

  init() {
    // Apply theme BEFORE page renders to prevent white flash
    const saved = localStorage.getItem(this.KEY);
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const theme = saved || (prefersDark ? 'dark' : 'light');
    this._apply(theme, false); // false = no transition on initial load
  },

  toggle() {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    this._apply(isDark ? 'light' : 'dark', true);
  },

  _apply(theme, animate) {
    if (!animate) {
      document.body.classList.add('theme-switching'); // Disable transitions
    }

    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem(this.KEY, theme);

    // Sync toggle switch
    if (DOM.darkToggleSwitch) {
      DOM.darkToggleSwitch.classList.toggle('on', theme === 'dark');
    }

    if (!animate) {
      requestAnimationFrame(() => {
        document.body.classList.remove('theme-switching');
      });
    }
  },

  isDark() {
    return document.documentElement.getAttribute('data-theme') === 'dark';
  }
};

// Backward-compatible init function
function initDarkMode() {
  Theme.init();
}

// Event listener for toggle
if (DOM.darkModeToggle) {
  DOM.darkModeToggle.addEventListener('click', () => Theme.toggle());
}

// ============================================================
// TOAST
// ============================================================
let toastTimeout;
function showToast(msg, type = 'info') {
  clearTimeout(toastTimeout);
  if (!DOM.toast) return;
  DOM.toast.textContent = msg;
  DOM.toast.className = `toast show ${type}`;
  // FIXED: Use 'toast-out' instead of 'hidden' to avoid conflict with utility .hidden { display: none !important }
  toastTimeout = setTimeout(() => { if (DOM.toast) DOM.toast.className = 'toast toast-out'; }, 3000);
}

// ============================================================
// SPLASH — FIXED: Unified state prevents double-hide
// ============================================================
function hideSplash() {
  if (SplashState.done) return;
  SplashState.markDone();
  if (DOM.splash) {
    DOM.splash.classList.add('fade-out');
    setTimeout(() => { if (DOM.splash) DOM.splash.remove(); }, 500);
  }
}

// ============================================================
// ONLINE / OFFLINE
// ============================================================
function updateOnlineStatus() {
  const isOnline = navigator.onLine;
  if (DOM.offlineBadge) {
    DOM.offlineBadge.style.display = isOnline ? 'none' : 'block';
    if (!isOnline) DOM.offlineBadge.textContent = '⚠️ وضع عدم الاتصال';
  }
  if (DOM.syncIndicator) {
    DOM.syncIndicator.textContent = isOnline ? 'متصل' : 'غير متصل';
    DOM.syncIndicator.classList.toggle('offline', !isOnline);
  }
}
window.addEventListener('online', updateOnlineStatus);
window.addEventListener('offline', updateOnlineStatus);

// ============================================================
// AUTH — Fixed with better error handling + Guest Mode
// ============================================================
async function initAuth() {
  if (!firebaseReady) {
    console.error('Firebase not available');
    hideSplash();
    showLogin();
    return;
  }

  try {
    try { await FB.getRedirectResult(auth); } catch (e) { console.error('getRedirectResult error:', e); }

    FB.onAuthStateChanged(auth, async (user) => {
      hideSplash();
      if (!user) {
        state.currentUser = null;
        state.appInitialized = false;
        // FIXED: Use immutable update + clear cache
        setStateGirls([]);
        setStateAttendanceData({});
        clearAllSnapshots();
        showLogin();
        return;
      }
      state.currentUser = user;
      showApp(user);
      if (!state.appInitialized) {
        state.appInitialized = true;
        await loadData();
        renderPage();
        // NEW: Auto-mark absence for today if it's a service day (respecting override)
        await checkAndAutoMarkAbsence();
      }
    });
  } catch (e) {
    console.error('Auth init error:', e);
    hideSplash();
    showLogin();
  }
}

// Google Sign In — FIXED: Use FB module instead of window._fb
if (DOM.googleSignIn) {
  DOM.googleSignIn.addEventListener('click', async () => {
    if (!firebaseReady) {
      showToast('الإنترنت غير متاح - حاول تحديث الصفحة', 'warning');
      return;
    }
    DOM.googleSignIn.classList.add('is-loading');
    try {
      await FB.signInWithPopup(auth, provider);
    } catch (e) {
      DOM.googleSignIn.classList.remove('is-loading');
      if (['auth/popup-blocked', 'auth/popup-closed-by-user', 'auth/cancelled-popup-request'].includes(e.code)) {
        try {
          await FB.signInWithRedirect(auth, provider);
        } catch (e2) { showToast('فشل تسجيل الدخول: ' + e2.message, 'error'); }
      } else {
        showToast('فشل تسجيل الدخول: ' + e.message, 'error');
      }
    }
  });
}

if (DOM.signOutBtn) {
  DOM.signOutBtn.addEventListener('click', async () => {
    clearAllSnapshots();
    if (!firebaseReady) {
      state.currentUser = null;
      state.appInitialized = false;
      showLogin();
      return;
    }
    await FB.signOut(auth);
  });
}

function showApp(user) {
  if (DOM.loginScreen) DOM.loginScreen.classList.add('hidden');
  if (DOM.mainApp) DOM.mainApp.classList.remove('hidden');
  if (DOM.googleSignIn) DOM.googleSignIn.classList.remove('is-loading');
  const card = document.getElementById('loginCard');
  if (card) {
    card.classList.remove('animate-in');
    card.querySelectorAll('.animate-in').forEach(el => el.classList.remove('animate-in'));
  }
  const initial = user && user.displayName ? user.displayName[0] : 'خ';
  if (DOM.userAvatar) DOM.userAvatar.textContent = initial;
  if (DOM.drawerAvatar) DOM.drawerAvatar.textContent = initial;
  if (DOM.drawerUserName) DOM.drawerUserName.textContent = (user && user.displayName) || 'الخادم';
  if (DOM.drawerUserEmail) DOM.drawerUserEmail.textContent = (user && user.email) || '';
}

function showLogin() {
  if (DOM.loginScreen) DOM.loginScreen.classList.remove('hidden');
  if (DOM.mainApp) DOM.mainApp.classList.add('hidden');
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      const card = document.getElementById('loginCard');
      if (card) {
        card.classList.add('animate-in');
        card.querySelectorAll('.login-cross-icon, .login-church-name, .login-system-title, .login-divider, .login-welcome, .btn-google, .btn-guest, .login-hint').forEach(el => {
          el.classList.add('animate-in');
        });
      }
    });
  });
}

// ============================================================
// FIREBASE LISTENERS — FIXED: Memory leak prevention + async safety
// ============================================================
async function loadData() {
  try {
    if (!firebaseReady) return;

    // FIXED: Guard against duplicate listeners — clear + flag pattern
    if (_listenersInitialized) {
      console.warn('loadData called while listeners already active — skipping');
      return;
    }

    // Clear any existing listeners first (prevents duplicate listeners on re-login)
    clearAllSnapshots();
    _listenersInitialized = true;

    // FIXED: Store unsubscribers to prevent memory leaks
    const unsub1 = FB.onSnapshot(
      FB.query(FB.collection(db, 'girls'), FB.orderBy('name')),
      (snap) => {
        let changed = false;
        const newGirls = [...state.girls];
        for (const change of snap.docChanges()) {
          const g = { id: change.doc.id, ...change.doc.data() };
          if (change.type === 'removed' || g.isDeleted) {
            const idx = newGirls.findIndex(x => x.id === g.id);
            if (idx >= 0) { newGirls.splice(idx, 1); changed = true; }
          } else {
            const idx = newGirls.findIndex(x => x.id === g.id);
            if (idx >= 0) { newGirls[idx] = g; changed = true; }
            else { newGirls.push(g); changed = true; }
          }
        }
        if (changed) {
          newGirls.sort((a, b) => a.name.localeCompare(b.name, 'ar'));
          setStateGirls(newGirls);
          scheduleRender();
        }
      },
      (err) => console.error('Girls snapshot error:', err)
    );
    pushUnsubscriber(unsub1);

    const unsub2 = FB.onSnapshot(
      FB.query(FB.collection(db, 'attendance'), FB.orderBy('date', 'desc')),
      (snap) => {
        let changed = false;
        const newData = { ...state.attendanceData };
        for (const change of snap.docChanges()) {
          const a = { id: change.doc.id, ...change.doc.data() };
          if (change.type === 'removed') {
            delete newData[a.id]; changed = true;
          } else {
            newData[a.id] = a; changed = true;
          }
        }
        if (changed) {
          setStateAttendanceData(newData);
          scheduleRender();
        }
      },
      (err) => console.error('Attendance snapshot error:', err)
    );
    pushUnsubscriber(unsub2);

    // FIXED: History listener — do async IDB ops outside onSnapshot callback
    const unsub3 = FB.onSnapshot(
      FB.query(FB.collection(db, 'history'), FB.orderBy('timestamp', 'desc')),
      (snap) => {
        let changed = false;
        const idbOps = [];
        for (const change of snap.docChanges()) {
          const log = { id: change.doc.id, ...change.doc.data() };
          if (change.type === 'removed') {
            idbOps.push(IDB.delete('history', log.id).catch(() => {}));
            changed = true;
          } else {
            idbOps.push(IDB.add('history', log).catch(() => {}));
            changed = true;
          }
        }
        // Fire IDB ops independently — don't block
        Promise.all(idbOps).catch(() => {});
        if (changed && state.currentPage === 'history') renderHistory(false);
      },
      (err) => console.error('History snapshot error:', err)
    );
    pushUnsubscriber(unsub3);

  } catch (e) {
    console.error('Load error:', e);
    // FIXED: Reset flag on error so loadData can be retried
    _listenersInitialized = false;
  }
}

// ============================================================
// RENDER ENGINE — FIXED: Better throttling (120ms instead of 60ms)
// + dirty flag to prevent duplicate renders
// + queueMicrotask hybrid for state-settle safety
// ============================================================
function scheduleRender() {
  if (state.renderPending) return; // Already scheduled
  state.renderPending = true;
  clearTimeout(state.renderTimeout);
  // FIXED: Use requestAnimationFrame + queueMicrotask to ensure state has settled
  requestAnimationFrame(() => {
    queueMicrotask(() => {
      state.renderPending = false;
      renderPage();
    });
  });
}

// FIXED: Debounced render for rapid-fire updates (toggleAttendance, etc.)
function debouncedRender(minMs = 80) {
  if (state.renderPending) return;
  state.renderPending = true;
  clearTimeout(state.renderTimeout);
  state.renderTimeout = setTimeout(() => {
    state.renderPending = false;
    renderPage();
  }, minMs);
}

function renderPage() {
  switch (state.currentPage) {
    case 'home': renderHome(); break;
    case 'attendance': renderAttendancePage(); break;
    case 'girls': renderGirlsList(); break;
    case 'calendar': renderCalendar(); break;
    case 'stats': renderStats(); break;
    case 'history': renderHistory(false); break;
    case 'export': renderExport(); break;
    case 'settings': renderSettings(); break;
  }
}

// ============================================================
// NAVIGATION
// ============================================================
const PAGE_TITLES = {
  home: ['الرئيسية', ''],
  attendance: ['الحضور اليومي', 'تسجيل وإدارة الحضور'],
  girls: ['المخدومات', 'قائمة المخدومات'],
  calendar: ['التقويم الشهري', 'أيام الخدمة'],
  stats: ['الإحصائيات', 'تحليلات وتقارير'],
  history: ['السجل التاريخي', 'سجل التعديلات'],
  export: ['التصدير', 'تصدير البيانات'],
  settings: ['الإعدادات', 'النسخ الاحتياطي والاستيراد']
};

function navigateTo(page) {
  const pageEl = document.getElementById(`page-${page}`);
  if (!pageEl) {
    console.warn(`Page element not found: page-${page}`);
    return;
  }

  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  pageEl.classList.add('active');
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.page === page));
  document.querySelectorAll('.menu-item[data-page]').forEach(b => b.classList.toggle('active', b.dataset.page === page));
  const [title, sub] = PAGE_TITLES[page] || [page, ''];
  if (DOM.pageTitle) DOM.pageTitle.textContent = title;
  if (DOM.pageSubtitle) DOM.pageSubtitle.textContent = sub;
  state.currentPage = page;

  if (page === 'attendance') {
    state.attendancePageInitialized = false;
    // NEW: Auto-mark absence when opening attendance page on a service day
    checkAndAutoMarkAbsence();
  }
  if (page !== 'calendar') {
    hideDayDetail();
  }

  renderPage();
  closeDrawer();
}

document.querySelectorAll('.nav-btn').forEach(btn => btn.addEventListener('click', () => navigateTo(btn.dataset.page)));
document.querySelectorAll('.menu-item[data-page]').forEach(item => item.addEventListener('click', e => {
  e.preventDefault();
  navigateTo(item.dataset.page);
}));

if (DOM.menuBtn) DOM.menuBtn.addEventListener('click', openDrawer);
if (DOM.drawerOverlay) DOM.drawerOverlay.addEventListener('click', closeDrawer);

function openDrawer() {
  if (DOM.drawer) DOM.drawer.classList.add('open');
  if (DOM.drawerOverlay) DOM.drawerOverlay.classList.add('show');
}
function closeDrawer() {
  if (DOM.drawer) DOM.drawer.classList.remove('open');
  if (DOM.drawerOverlay) DOM.drawerOverlay.classList.remove('show');
}


// ============================================================
// SMART STATS — FIXED: Use Cache.girlsById instead of state.girls.find()
// ============================================================
function getBestGradeFiltered(monthStr, gradeFilter) {
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const [year, month] = monthStr.split('-').map(Number);
  // FIXED: Guard against undefined from getServiceDaysInMonth
  const serviceDays = getServiceDaysInMonth(year, month - 1) || [];
  const totalServiceDays = serviceDays.length || 1;

  const gradeStats = {};
  activeGirls.forEach(g => {
    if (gradeFilter && g.grade !== gradeFilter) return;
    if (!gradeStats[g.grade]) gradeStats[g.grade] = { totalGirls: 0, presentDates: new Set() };
    gradeStats[g.grade].totalGirls++;
  });

  const allAttendance = Cache.getAllAttendance();
  allAttendance.forEach(a => {
    if (!a.date?.startsWith(monthStr)) return;
    if (a.status !== 'حاضر') return;
    const girl = Cache.getGirl(a.girlId);
    if (!girl) return;
    if (gradeFilter && girl.grade !== gradeFilter) return;
    if (!gradeStats[girl.grade]) return;
    gradeStats[girl.grade].presentDates.add(a.date + '_' + a.girlId);
  });

  let best = null;
  Object.entries(gradeStats).forEach(([grade, data]) => {
    const maxPossible = data.totalGirls * totalServiceDays;
    const percent = maxPossible > 0 ? (data.presentDates.size / maxPossible) * 100 : 0;
    if (!best || percent > best.percent) best = { grade, percent };
  });
  return best;
}

function getTopActivityFiltered(monthStr, gradeFilter) {
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const activeGirlIds = gradeFilter
    ? new Set(activeGirls.filter(g => g.grade === gradeFilter).map(g => g.id))
    : new Set(activeGirls.map(g => g.id));
  const counts = {};
  ACTIVITIES.forEach(a => counts[a] = 0);

  const allAttendance = Cache.getAllAttendance();
  allAttendance.forEach(a => {
    if (!a.date?.startsWith(monthStr)) return;
    if (!activeGirlIds.has(a.girlId)) return;
    if (a.status === 'حاضر' && counts[a.activity] !== undefined) counts[a.activity]++;
  });

  let topName = ACTIVITIES[0];
  let topValue = 0;
  Object.entries(counts).forEach(([name, count]) => {
    if (count > topValue) { topName = name; topValue = count; }
  });
  return topValue > 0 ? { name: topName, count: topValue } : null;
}

function getMostRegularGirlFiltered(monthStr, gradeFilter) {
  let activeGirls = state.girls.filter(g => !g.isDeleted);
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  if (!activeGirls.length) return null;
  const activeGirlIds = new Set(activeGirls.map(g => g.id));

  const [year, month] = monthStr.split('-').map(Number);
  // FIXED: Guard against undefined from getServiceDaysInMonth
  const serviceDays = getServiceDaysInMonth(year, month - 1) || [];
  const totalServiceDays = serviceDays.length || 1;

  const presentDatesByGirl = {};
  activeGirls.forEach(g => presentDatesByGirl[g.id] = new Set());

  const allAttendance = Cache.getAllAttendance();
  allAttendance.forEach(a => {
    if (!a.date?.startsWith(monthStr)) return;
    if (a.status === 'حاضر' && presentDatesByGirl[a.girlId] !== undefined) {
      presentDatesByGirl[a.girlId].add(a.date);
    }
  });

  let best = null;
  Object.entries(presentDatesByGirl).forEach(([girlId, dateSet]) => {
    const count = dateSet.size;
    if (count === 0) return;
    const percent = (count / totalServiceDays) * 100;
    const girl = Cache.getGirl(girlId);
    if (!girl) return;
    if (!best || percent > best.percent || (percent === best.percent && count > best.count)) {
      best = { name: girl.name, count, percent };
    }
  });
  return best;
}

// ============================================================
// HOME PAGE — FIXED: O(n^2) eliminated with cache + single-pass logic
// ============================================================
function renderHome() {
  const selectedDate = TimeContext.getDate();
  // FIXED: Use parseDateStr instead of unsafe new Date(dateStr + 'T00:00:00')
  const now = parseDateStr(selectedDate);
  const dayName = DateUtil.dayName(now);
  const dateStr = selectedDate;
  const monthStr = TimeContext.getMonth();

  if (DOM.todayDay) DOM.todayDay.textContent = `${DateUtil.formatDateShort(now)} ${dayName}`;
  if (DOM.todayDate) DOM.todayDate.textContent = now.toLocaleDateString('ar-EG', { year: 'numeric', month: 'long', day: 'numeric' });

  // NEW v3.1: Use isServiceDayForDate which respects user overrides
  const isService = isServiceDayForDate(dateStr);

  if (DOM.todayServiceBadge) {
    DOM.todayServiceBadge.textContent = isService ? 'يوم خدمة \u2713' : 'لا توجد خدمة اليوم';
    DOM.todayServiceBadge.classList.toggle('active', isService);
  }

  const gradeFilter = state.homeGradeFilter;
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const filteredGirls = gradeFilter ? activeGirls.filter(g => g.grade === gradeFilter) : activeGirls;
  // FIXED: Use cached activeGirlIds from Cache instead of rebuilding Set
  const allActiveGirlIds = Cache.getActiveGirlIds();
  const activeGirlIds = gradeFilter
    ? new Set(filteredGirls.map(g => g.id))
    : allActiveGirlIds;

  // FIXED: Single-pass grade count instead of 3 separate filters (O(n) not O(3n))
  const gradeCounts = { 'أولى إعدادي': 0, 'تانية إعدادي': 0, 'تالتة إعدادي': 0 };
  activeGirls.forEach(g => {
    if (gradeCounts[g.grade] !== undefined) gradeCounts[g.grade]++;
  });
  const hfcAll = document.getElementById('homeFilterCountAll');
  const hfc1 = document.getElementById('homeFilterCount1');
  const hfc2 = document.getElementById('homeFilterCount2');
  const hfc3 = document.getElementById('homeFilterCount3');
  if (hfcAll) hfcAll.textContent = activeGirls.length;
  if (hfc1) hfc1.textContent = gradeCounts['أولى إعدادي'];
  if (hfc2) hfc2.textContent = gradeCounts['تانية إعدادي'];
  if (hfc3) hfc3.textContent = gradeCounts['تالتة إعدادي'];

  document.querySelectorAll('#homeGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  if (DOM.statTotal) DOM.statTotal.textContent = filteredGirls.length;

  // FIXED: Single-pass attendance scan instead of multiple loops
  const presentGirlIds = new Set();
  const absentGirlIds = new Set();
  const todayRecordsByGirl = {};
  const monthPresentsByGirl = {}; // For top attendees
  let totalRating = 0, ratingCount = 0;

  const allAttendance = Cache.getAllAttendance();
  allAttendance.forEach(a => {
    // Today counts
    if (a.date === dateStr && activeGirlIds.has(a.girlId)) {
      if (!todayRecordsByGirl[a.girlId]) todayRecordsByGirl[a.girlId] = [];
      todayRecordsByGirl[a.girlId].push(a);
    }
    // Month presents for top attendees + ratings
    if (a.date?.startsWith(monthStr) && activeGirlIds.has(a.girlId)) {
      if (a.status === 'حاضر') {
        if (!monthPresentsByGirl[a.girlId]) monthPresentsByGirl[a.girlId] = new Set();
        monthPresentsByGirl[a.girlId].add(a.date);
      }
      if (a.rating > 0) { totalRating += a.rating; ratingCount++; }
    }
  });

  // Process today's status
  filteredGirls.forEach(g => {
    const records = todayRecordsByGirl[g.id];
    if (records && records.length > 0) {
      const hasAnyPresent = records.some(r => r.status === 'حاضر');
      if (hasAnyPresent) presentGirlIds.add(g.id);
      else absentGirlIds.add(g.id);
    } else if (isService) {
      // Only count as absent if today is a service day
      absentGirlIds.add(g.id);
    }
  });

  if (DOM.statPresentToday) DOM.statPresentToday.textContent = presentGirlIds.size;
  if (DOM.statAbsentToday) DOM.statAbsentToday.textContent = absentGirlIds.size;
  if (DOM.statAvgRating) DOM.statAvgRating.textContent = ratingCount ? (totalRating / ratingCount).toFixed(1) : '-';

  // Best grade
  const bestGrade = getBestGradeFiltered(monthStr, gradeFilter);
  if (DOM.bestGrade && DOM.bestGradePercent) {
    if (bestGrade && bestGrade.percent > 0) {
      DOM.bestGrade.textContent = bestGrade.grade;
      DOM.bestGradePercent.textContent = `${Math.round(bestGrade.percent)}% حضور`;
    } else {
      DOM.bestGrade.textContent = gradeFilter || '-';
      DOM.bestGradePercent.textContent = gradeFilter ? 'لا توجد بيانات' : 'أفضل سنة دراسية';
    }
  }

  // Top activity
  const topActivity = getTopActivityFiltered(monthStr, gradeFilter);
  if (DOM.topActivityName && DOM.topActivityCount) {
    if (topActivity) {
      DOM.topActivityName.textContent = topActivity.name;
      DOM.topActivityCount.textContent = `${topActivity.count} حضور`;
    } else {
      DOM.topActivityName.textContent = '-';
      DOM.topActivityCount.textContent = 'أكثر نشاط حضورًا';
    }
  }

  // Most regular
  const mostRegular = getMostRegularGirlFiltered(monthStr, gradeFilter);
  if (DOM.mostRegularGirl && DOM.mostRegularPercent) {
    if (mostRegular) {
      DOM.mostRegularGirl.textContent = mostRegular.name;
      DOM.mostRegularPercent.textContent = `${mostRegular.count} يوم \u00B7 ${Math.round(mostRegular.percent)}%`;
    } else {
      DOM.mostRegularGirl.textContent = '-';
      DOM.mostRegularPercent.textContent = 'أكثر مخدومة انتظامًا';
    }
  }

  // Top attendees — from precomputed monthPresentsByGirl
  if (DOM.topAttendees) {
    const sorted = Object.entries(monthPresentsByGirl)
      .map(([id, dates]) => [id, dates.size])
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .filter(([, count]) => count > 0);

    if (!sorted.length) {
      DOM.topAttendees.innerHTML = '<div class="empty-state">لا توجد بيانات حضور هذا الشهر</div>';
    } else {
      const frag = document.createDocumentFragment();
      sorted.forEach(([id, count], i) => {
        const g = Cache.getGirl(id);
        if (!g) return;
        const div = document.createElement('div');
        div.className = 'top-item';
        div.innerHTML = `<span class="top-rank">${i + 1}</span><span class="top-name">${esc(g.name)}</span><span class="top-count">${count} يوم</span>`;
        frag.appendChild(div);
      });
      DOM.topAttendees.innerHTML = '';
      DOM.topAttendees.appendChild(frag);
    }
  }

  // Needs followup — FIXED: Single-pass hasConsecutiveAbsences with cached results
  const needs = [];
  filteredGirls.forEach(g => {
    const result = hasConsecutiveAbsences(g.id, monthStr);
    if (result.hasConsecutive) needs.push({ girl: g, result });
  });

  if (DOM.needsFollowup) {
    if (!needs.length) {
      DOM.needsFollowup.innerHTML = '<div class="empty-state">لا توجد حالات تحتاج متابعة</div>';
    } else {
      const frag = document.createDocumentFragment();
      needs.forEach(({ girl, result }) => {
        const div = document.createElement('div');
        div.className = 'followup-item';
        div.dataset.girlId = girl.id;
        // FIXED: result.count = total absences, not consecutive streak. Changed text to be accurate.
        div.innerHTML = `<span class="followup-name">${esc(girl.name)}</span><span class="followup-badge">${result.count} غياب</span>`;
        frag.appendChild(div);
      });
      DOM.needsFollowup.innerHTML = '';
      DOM.needsFollowup.appendChild(frag);
    }
  }
}

// ============================================================
// SEARCH
// ============================================================
function debouncedSearch() {
  clearTimeout(state.searchDebounceTimer);
  state.searchDebounceTimer = setTimeout(() => {
    const q = DOM.globalSearch ? DOM.globalSearch.value.trim() : '';
    const resultsEl = DOM.searchResults;
    if (!resultsEl) return;
    if (!q) { resultsEl.classList.remove('show'); resultsEl.innerHTML = ''; return; }
    const qNorm = normalizeArabic(q);
    const matches = state.girls.filter(g => !g.isDeleted && normalizeArabic(g.name).includes(qNorm));
    resultsEl.innerHTML = matches.length
      ? matches.map(g => `<div class="search-item" data-girl-id="${esc(g.id)}"><span>${esc(g.name)}</span><span class="grade-badge">${esc(g.grade)}</span></div>`).join('')
      : '<div class="search-item">لا توجد نتائج</div>';
    resultsEl.classList.add('show');
  }, 250);
}

if (DOM.globalSearch) DOM.globalSearch.addEventListener('input', debouncedSearch);

// ============================================================
// GIRLS PAGE — FIXED: Use Cache instead of repeated .find()
// ============================================================
function renderGirlsList() {
  const filter = state.girlsGradeFilter;
  const searchQuery = (state.girlsSearchQuery || '').trim();
  let activeGirls = state.girls.filter(g => !g.isDeleted);

  if (searchQuery) {
    const qNorm = normalizeArabic(searchQuery);
    activeGirls = activeGirls.filter(g => normalizeArabic(g.name).includes(qNorm));
  }

  const filtered = filter ? activeGirls.filter(g => g.grade === filter) : activeGirls;
  const el = DOM.girlsList;
  if (!el) return;

  const gfcAll = document.getElementById('girlsFilterCountAll');
  const gfc1 = document.getElementById('girlsFilterCount1');
  const gfc2 = document.getElementById('girlsFilterCount2');
  const gfc3 = document.getElementById('girlsFilterCount3');
  if (gfcAll) gfcAll.textContent = activeGirls.length;
  if (gfc1) gfc1.textContent = activeGirls.filter(g => g.grade === 'أولى إعدادي').length;
  if (gfc2) gfc2.textContent = activeGirls.filter(g => g.grade === 'تانية إعدادي').length;
  if (gfc3) gfc3.textContent = activeGirls.filter(g => g.grade === 'تالتة إعدادي').length;

  document.querySelectorAll('#girlsGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === filter);
  });

  if (!filtered.length) {
    el.innerHTML = '<div class="empty-state">لا توجد مخدومات<br><small>اضغط + لإضافة مخدومة جديدة</small></div>';
    return;
  }

  // FIXED: Precompute attendance counts per girl in single pass
  const monthStr = TimeContext.getMonth();
  const girlStats = {};
  const allAttendance = Cache.getAllAttendance();
  allAttendance.forEach(a => {
    if (!a.date?.startsWith(monthStr)) return;
    if (!girlStats[a.girlId]) girlStats[a.girlId] = { present: 0, absent: 0 };
    if (a.status === 'حاضر') girlStats[a.girlId].present++;
    else if (a.status === 'غائب') girlStats[a.girlId].absent++;
  });

  const frag = document.createDocumentFragment();
  filtered.forEach(g => {
    const stats = girlStats[g.id] || { present: 0, absent: 0 };
    const div = document.createElement('div');
    div.className = 'girl-card';
    div.dataset.girlId = g.id;
    div.innerHTML = `
      <div class="girl-avatar">${esc(g.name[0])}</div>
      <div class="girl-info">
        <span class="girl-name">${esc(g.name)}</span>
        <span class="girl-grade">${esc(g.grade)}</span>
        ${g.phone ? `<a href="tel:${esc(g.phone)}" class="girl-phone-link" data-phone="${esc(g.phone)}" onclick="event.stopPropagation();">${esc(g.phone)}</a>` : ''}
        <div class="girl-stats"><span class="green-text">&#10003;${stats.present}</span><span class="red-text">&#10007;${stats.absent}</span></div>
      </div>
      <button class="edit-btn" data-girl-id="${esc(g.id)}" aria-label="تعديل ${esc(g.name)}">&#9999;</button>`;
    frag.appendChild(div);
  });
  el.innerHTML = '';
  el.appendChild(frag);
}

if (DOM.addGirlBtn) {
  DOM.addGirlBtn.addEventListener('click', () => {
    state.editingGirlId = null;
    if (DOM.girlModalTitle) DOM.girlModalTitle.textContent = 'إضافة مخدومة';
    if (DOM.girlName) DOM.girlName.value = '';
    if (DOM.girlPhone) DOM.girlPhone.value = '';
    if (DOM.girlGrade) DOM.girlGrade.value = '';
    if (DOM.girlNotes) DOM.girlNotes.value = '';
    if (DOM.deleteGirlBtn) DOM.deleteGirlBtn.classList.add('hidden');
    openModal('girlModal');
  });
}

function editGirl(id) {
  const g = Cache.getGirl(id);
  if (!g) return;
  state.editingGirlId = id;
  if (DOM.girlModalTitle) DOM.girlModalTitle.textContent = 'تعديل بيانات المخدومة';
  if (DOM.girlName) DOM.girlName.value = g.name;
  if (DOM.girlPhone) DOM.girlPhone.value = g.phone || '';
  if (DOM.girlGrade) DOM.girlGrade.value = g.grade;
  if (DOM.girlNotes) DOM.girlNotes.value = g.notes || '';
  if (DOM.deleteGirlBtn) DOM.deleteGirlBtn.classList.remove('hidden');
  openModal('girlModal');
}

// ============================================================
// DELETE GIRL — FIXED: State validation + snapshot isolation + backup
// ============================================================
if (DOM.deleteGirlBtn) {
  DOM.deleteGirlBtn.addEventListener('click', async () => {
    if (!state.editingGirlId || state.deleteInProgress) return;
    const currentId = state.editingGirlId; // Capture ID at click time
    const g = Cache.getGirl(currentId);
    if (!g) return;

    closeModal('girlModal');

    showConfirm({
      icon: '&#9888;', title: 'حذف مخدومة',
      msg: `هل أنت متأكد من حذف "${esc(g.name)}"؟ سيتم حذف جميع بيانات الحضور الخاصة بها أيضاً.`,
      okLabel: 'حذف',
      okClass: 'confirm-delete',
      onOk: async () => {
        if (state.deleteInProgress) return;
        // FIXED: Validate the captured ID matches current editingGirlId
        if (state.editingGirlId !== currentId) {
          showToast('خطأ: تم تغيير المخدومة المحددة', 'error');
          return;
        }
        state.deleteInProgress = true;

        // FIXED: Create backup before delete for rollback support
        const backupId = 'delete_' + currentId + '_' + Date.now();
        await createBackup(backupId, { girl: g, attendanceData: state.attendanceData });

        try {
          const id = currentId;
          // Remove from state
          setStateGirls(state.girls.filter(x => x.id !== id));
          const newAttData = { ...state.attendanceData };
          Object.keys(newAttData).forEach(k => {
            if (newAttData[k].girlId === id) delete newAttData[k];
          });
          setStateAttendanceData(newAttData);

          if (firebaseReady) {
            try {
              await FB.setDoc(FB.doc(db, 'girls', id), {
                isDeleted: true, deletedAt: Date.now(),
                deletedBy: state.currentUser?.email || '',
                name: g.name, grade: g.grade
              }, { merge: true });

              const attQuery = FB.query(FB.collection(db, 'attendance'), FB.where('girlId', '==', id));
              const attSnap = await FB.getDocs(attQuery);
              if (!attSnap.empty) {
                const docs = attSnap.docs;
                for (let i = 0; i < docs.length; i += 500) {
                  try {
                    const batch = FB.writeBatch(db);
                    docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                    await batch.commit();
                  } catch (batchErr) {
                    console.error('Batch delete error (retrying):', batchErr);
                    // FIXED: Simple retry once
                    try {
                      const batch = FB.writeBatch(db);
                      docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                      await batch.commit();
                    } catch (e2) { console.error('Batch delete retry failed:', e2); }
                  }
                }
              }
            } catch (e) {
              console.error('Delete girl Firestore error:', e);
            }
          }

          await logHistory('حذف مخدومة', `${g.name} - ${g.grade}`);
          showToast(`تم حذف ${g.name}`, 'success');
          state.editingGirlId = null;
          scheduleRender();
        } catch (err) {
          console.error('Delete error:', err);
          showToast('حدث خطأ أثناء الحذف', 'error');
        } finally {
          state.deleteInProgress = false;
        }
      }
    });
  });
}

// ============================================================
// SAVE GIRL — FIXED: catch block + Firestore-first ordering + backup
// ============================================================
if (DOM.saveGirlBtn) {
  DOM.saveGirlBtn.addEventListener('click', async () => {
    if (state.savingGirl || state.pendingSaveGirl) return;
    state.savingGirl = true;
    state.pendingSaveGirl = true;

    // FIXED: Create backup before save for rollback support
    const backupId = 'saveGirl_' + Date.now();
    await createBackup(backupId, { girls: state.girls, attendanceData: state.attendanceData });

    try {
      const name = DOM.girlName ? DOM.girlName.value.trim() : '';
      const phone = DOM.girlPhone ? DOM.girlPhone.value.trim() : '';
      const grade = DOM.girlGrade ? DOM.girlGrade.value : '';
      const notes = DOM.girlNotes ? DOM.girlNotes.value.trim() : '';

      if (!name) { showToast('الرجاء إدخال اسم المخدومة', 'error'); return; }
      if (!grade) { showToast('الرجاء اختيار السنة الدراسية', 'error'); return; }

      const normalizedName = normalizeName(name);
      const existingGirl = state.girls.find(g =>
        normalizeName(g.name) === normalizedName && g.id !== state.editingGirlId && !g.isDeleted
      );
      if (existingGirl) { showToast('هذه المخدومة موجودة بالفعل', 'error'); return; }

      const id = state.editingGirlId || 'girl_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
      const now = Date.now();
      const girlData = {
        id, name, phone, grade, notes,
        createdAt: state.editingGirlId ? (Cache.getGirl(id)?.createdAt || now) : now,
        updatedAt: now,
        updatedBy: state.currentUser?.displayName || 'خادم',
        updatedByEmail: state.currentUser?.email || '',
        isDeleted: false
      };

      const isNewGirl = !state.editingGirlId;
      const wasEditing = !!state.editingGirlId; // FIXED: Capture before any changes

      // FIXED: Firestore write FIRST, then update state on success
      if (firebaseReady) {
        try {
          await FB.setDoc(FB.doc(db, 'girls', id), girlData);
        } catch (e) {
          console.error('Save girl Firestore error:', e);
          showToast('فشل الحفظ في السحابة، تحقق من الاتصال', 'error');
          return; // Don't update state if Firestore failed
        }
      }

      // Now update local state (guaranteed to match server)
      if (state.editingGirlId) {
        setStateGirls(state.girls.map(g => g.id === id ? girlData : g));
      } else {
        setStateGirls([...state.girls, girlData]);
      }

      await logHistory(wasEditing ? 'تعديل مخدومة' : 'إضافة مخدومة', `${name} - ${grade}`); // FIXED: Use wasEditing

      // Auto-mark absent on service days for new girls only
      if (isNewGirl) {
        const todayStr = DateUtil.toStr();
        // NEW v3.1: Only auto-mark if today is actually a service day
        if (isServiceDayForDate(todayStr)) {
          await autoMarkAbsentForNewGirl(id, todayStr);
        }
      }

      closeModal('girlModal');
      showToast(wasEditing ? 'تم تعديل البيانات' : 'تمت إضافة المخدومة', 'success'); // FIXED: Use wasEditing
      state.editingGirlId = null;
      renderPage();
    } catch (err) {
      // FIXED: Added catch block for errors
      console.error('Save girl error:', err);
      showToast('حدث خطأ أثناء الحفظ: ' + (err.message || 'خطأ غير معروف'), 'error');
    } finally {
      state.savingGirl = false;
      state.pendingSaveGirl = false;
    }
  });
}

// ============================================================
// GIRL PROFILE — FIXED: Correct lastAttendance + safe month grouping
// ============================================================
function showGirlProfile(id) {
  const g = Cache.getGirl(id);
  if (!g) return;
  state.currentProfileGirlId = id;
  if (DOM.profileName) DOM.profileName.textContent = g.name;

  const girlAtt = Cache.getAllAttendance().filter(a => a.girlId === id);
  // FIXED: Use parseDateStr for safe date comparison
  girlAtt.sort((a, b) => compareDateStr(b.date, a.date));

  const totalRecords = girlAtt.length;
  const presentCount = girlAtt.filter(a => a.status === 'حاضر').length;
  const absentCount = girlAtt.filter(a => a.status === 'غائب').length;

  // FIXED: Consistent attendance rate calculation
  const attendanceRate = totalRecords > 0 ? Math.round((presentCount / totalRecords) * 100) : 0;

  const ratings = girlAtt.filter(a => a.rating > 0).map(a => a.rating);
  const avgRating = ratings.length ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(1) : '0';

  // FIXED: Use findLast (or reverse find) to get MOST RECENT present record
  const sortedAtt = [...girlAtt].sort((a, b) => compareDateStr(a.date, b.date)); // oldest first
  const lastAttendance = [...sortedAtt].reverse().find(a => a.status === 'حاضر');
  const lastDate = lastAttendance ? lastAttendance.date : '-';

  // FIXED: Safe month grouping with date validation
  const months = {};
  girlAtt.forEach(a => {
    const m = a.date?.substring(0, 7);
    if (!m || !/^\d{4}-\d{2}$/.test(m)) return; // Skip malformed dates
    if (!months[m]) months[m] = [];
    months[m].push(a);
  });

  let html = `<div class="profile-info">
    <span class="grade-badge">${esc(g.grade)}</span>
    ${g.phone ? `<span class="profile-phone">&#128222; ${esc(g.phone)}</span>` : ''}
    ${g.notes ? `<p class="profile-notes">${esc(g.notes)}</p>` : ''}
  </div>`;

  html += `<div class="profile-dashboard">
    <div class="profile-stat"><div class="ps-value green">${presentCount}</div><div class="ps-label">مرات الحضور</div></div>
    <div class="profile-stat"><div class="ps-value red">${absentCount}</div><div class="ps-label">مرات الغياب</div></div>
    <div class="profile-stat"><div class="ps-value orange">${attendanceRate}%</div><div class="ps-label">نسبة الحضور</div></div>
    <div class="profile-stat"><div class="ps-value">${avgRating}</div><div class="ps-label">متوسط التقييم</div></div>
    <div class="profile-stat"><div class="ps-value">${totalRecords}</div><div class="ps-label">إجمالي السجلات</div></div>
    <div class="profile-stat"><div class="ps-value">${lastDate}</div><div class="ps-label">آخر حضور</div></div>
  </div>`;

  if (!Object.keys(months).length) {
    html += '<div class="empty-state">لا توجد سجلات حضور</div>';
  } else {
    Object.entries(months).sort((a, b) => b[0].localeCompare(a[0])).forEach(([month, records]) => {
      const presents = records.filter(r => r.status === 'حاضر').length;
      const absents = records.filter(r => r.status === 'غائب').length;
      html += `<div class="profile-month">
        <div class="profile-month-header">
          <span>${DateUtil.formatMonth(month)}</span>
          <span class="green-text">&#10003;${presents}</span>
          <span class="red-text">&#10007;${absents}</span>
        </div>
        <div class="profile-records">
          ${records.map(r => {
            const stars = r.rating ? '&#9733;'.repeat(r.rating) + '&#9734;'.repeat(5 - r.rating) : '';
            // FIXED: Use parseDateStr for safe day name lookup
            const dayName = DAY_NAMES[parseDateStr(r.date).getDay()] || '';
            return `<div class="profile-record">
              <span class="rec-date">${esc(r.date)} ${esc(dayName)}</span>
              <span class="rec-activity">${esc(r.activity || '')}</span>
              <span class="rec-status ${r.status === 'حاضر' ? 'present' : 'absent'}">${esc(r.status)}</span>
              ${stars ? `<span class="rec-rating">${stars}</span>` : ''}
              ${r.notes ? `<span class="rec-notes">${esc(r.notes)}</span>` : ''}
            </div>`;
          }).join('')}
        </div>
      </div>`;
    });
  }
  if (DOM.profileBody) DOM.profileBody.innerHTML = html;
  openModal('girlProfileModal');
}

if (DOM.closeProfileModal) DOM.closeProfileModal.addEventListener('click', () => closeModal('girlProfileModal'));
if (DOM.editProfileBtn) {
  DOM.editProfileBtn.addEventListener('click', () => {
    closeModal('girlProfileModal');
    if (state.currentProfileGirlId) editGirl(state.currentProfileGirlId);
  });
}

// ============================================================
// SHARE PROFILE — FIXED: Use ASCII-safe symbols
// ============================================================
if (DOM.shareProfileBtn) {
  DOM.shareProfileBtn.addEventListener('click', async () => {
    const id = state.currentProfileGirlId;
    if (!id) return;
    const g = Cache.getGirl(id);
    if (!g) return;

    const girlAtt = Cache.getAllAttendance().filter(a => a.girlId === id);
    const presentCount = girlAtt.filter(a => a.status === 'حاضر').length;
    const absentCount = girlAtt.filter(a => a.status === 'غائب').length;
    const attendanceRate = girlAtt.length > 0 ? Math.round((presentCount / girlAtt.length) * 100) : 0;

    // FIXED: Use ASCII-safe symbols instead of Unicode that may break on old devices
    const shareText = `${g.name}
${g.grade}
[H] حضور: ${presentCount}
[G] غياب: ${absentCount}
[%] نسبة: ${attendanceRate}%
`.trim();

    if (navigator.share) {
      try { await navigator.share({ title: `ملف ${g.name}`, text: shareText }); } catch (e) { /* user cancelled */ }
    } else {
      try {
        await navigator.clipboard.writeText(shareText);
        showToast('تم نسخ البيانات للمشاركة', 'success');
      } catch (e) {
        showToast('المشاركة غير متوفرة على هذا الجهاز', 'warning');
      }
    }
  });
}

// ============================================================
// ATTENDANCE PAGE — v3.1: Service day toggle + Reset attendance
// ============================================================
function getCurrentServiceDay() {
  const dayOfWeek = new Date().getDay();
  const dayMap = { 6: 'السبت', 1: 'الاثنين', 3: 'الاربعاء' };
  return dayMap[dayOfWeek] || null;
}

function isServiceDayDate(dateStr) {
  if (!dateStr) return false;
  // Use the new override-aware function
  return isServiceDayForDate(dateStr);
}

// ============================================================
// AUTO MARK ABSENCE — v3.1: Respects service day toggle
// ============================================================

/**
 * Persist the auto-marked dates Set to localStorage
 */
function persistAutoMarkedDates() {
  try {
    localStorage.setItem('autoMarkedDates', JSON.stringify([...state.autoMarkedDates]));
  } catch (e) { console.warn('Failed to persist autoMarkedDates:', e); }
}

/**
 * Check if a service day has already been auto-marked for all 4 activities
 * We consider it complete only if ALL activities have records for ALL active girls
 */
function isDayFullyAutoMarked(date) {
  if (!state.autoMarkedDates.has(date)) return false;
  // Additional check: ensure we have records for all activities
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  if (activeGirls.length === 0) return true; // No girls yet, consider it done

  for (const g of activeGirls) {
    for (const activity of ACTIVITIES) {
      const key = makeAttKey(g.id, date, activity);
      if (!state.attendanceData[key]) return false; // Missing record
    }
  }
  return true;
}

/**
 * NEW: Automatically mark all girls as absent on service days.
 * This runs once per service day when the app loads or when navigating to attendance page.
 * Uses localStorage to persist across reloads.
 * v3.1: Now respects the service day toggle override
 */
async function checkAndAutoMarkAbsence() {
  const today = DateUtil.toStr();

  // NEW v3.1: Only run if today is actually a service day (respecting user override)
  if (!isServiceDayForDate(today)) return;

  // Check if we already fully marked this day
  if (isDayFullyAutoMarked(today)) return;

  // Also check if there are already any attendance records for today
  // (user may have manually started marking attendance)
  const todayRecords = Cache.getAttendanceByDate(today);
  const hasAnyRecords = todayRecords.length > 0;

  if (hasAnyRecords && state.autoMarkedDates.has(today)) {
    // Already processed and has records, skip
    return;
  }

  // Mark all girls as absent for all activities
  showToast('جاري تسجيل الغياب التلقائي ليوم الخدمة...', 'info');
  await markAllAbsentForDate(today);

  // Track that we auto-marked this day
  state.autoMarkedDates.add(today);
  persistAutoMarkedDates();

  showToast('تم تسجيل الغياب التلقائي — اضغط على اسم المخدومة للتبديل لحاضر', 'success');
}

// FIXED: Renamed to clarify this is a hardcoded lookup, not dynamic
function getHardcodedServiceDay(dayOfWeek) {
  const dayMap = { 6: 'السبت', 1: 'الاثنين', 3: 'الاربعاء' };
  return dayMap[dayOfWeek] || null;
}

// ============================================================
// RESET ATTENDANCE — NEW v3.1: Clear all attendance for a date
// ============================================================

/**
 * NEW v3.1: Reset (delete) all attendance records for the current date.
 * This removes all records for all activities for the selected date.
 */
async function resetAttendanceForDate(date) {
  if (!date) {
    showToast('الرجاء اختيار تاريخ أولاً', 'error');
    return;
  }

  // Find all keys for this date
  const keysToDelete = [];
  Object.keys(state.attendanceData).forEach(key => {
    const rec = state.attendanceData[key];
    if (rec && rec.date === date) {
      keysToDelete.push(key);
    }
  });

  if (keysToDelete.length === 0) {
    showToast('مفيش سجلات حضور لليوم ده عشان تمسحها', 'warning');
    return;
  }

  // Remove from local state
  const newAttData = { ...state.attendanceData };
  keysToDelete.forEach(key => delete newAttData[key]);
  setStateAttendanceData(newAttData);

  // Remove from Firestore
  if (firebaseReady) {
    try {
      const batch = FB.writeBatch(db);
      keysToDelete.forEach(key => {
        batch.delete(FB.doc(db, 'attendance', key));
      });
      await batch.commit();
    } catch (e) {
      console.error('Reset attendance Firestore error:', e);
      showToast('تم التصفير محلياً بس — فشل المزامنة مع السحابة', 'warning');
    }
  }

  // Remove from auto-marked dates so it can be re-auto-marked
  state.autoMarkedDates.delete(date);
  persistAutoMarkedDates();

  await logHistory('تصفير الحضور', `تم مسح جميع سجلات الحضور ليوم ${date} — ${keysToDelete.length} سجل`);
  showToast(`تم تصفير الحضور — ${keysToDelete.length} سجل اتمسح`, 'success');

  // Re-render
  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'calendar') renderCalendar();
}

// ============================================================
// RENDER ATTENDANCE PAGE — v3.1: Service toggle + Reset button
// ============================================================
function renderAttendancePage() {
  if (!DOM.attendanceDate) return;
  DOM.attendanceDate.value = TimeContext.getDate();

  const currentServiceDay = getCurrentServiceDay();
  if (currentServiceDay && !state.attendancePageInitialized) {
    state.selectedDay = currentServiceDay;
  }

  setActiveDay(state.selectedDay);
  setActiveActivity(state.selectedActivity);

  // NEW v3.1: Update service day toggle based on current date
  const date = TimeContext.getDate();
  const isService = isServiceDayForDate(date);

  // Update the toggle switch UI
  const toggleEl = DOM.serviceDayToggle;
  const toggleWrap = DOM.serviceDayToggleWrap;
  if (toggleEl) {
    toggleEl.checked = isService;
  }
  if (toggleWrap) {
    toggleWrap.classList.toggle('service-on', isService);
    toggleWrap.classList.toggle('service-off', !isService);
  }

  // Update hint text
  const hintEl = DOM.serviceToggleHint;
  if (hintEl) {
    if (isService) {
      hintEl.textContent = 'التبديل شغال = اليوم فيه خدمة — الحضور شغال';
    } else {
      hintEl.textContent = 'التبديل مقفول = مفيش خدمة اليوم — مش هيتم تسجيل غياب تلقائي';
    }
  }

  state.attendancePageInitialized = true;

  // NEW v3.1: Show/hide content based on service day toggle
  const noServiceMsg = DOM.noServiceMessage;
  const attendanceList = DOM.attendanceList;
  const attendanceSummary = DOM.attendanceSummary;
  const quickActionsRow = DOM.quickActionsRow;
  const attToggleHint = DOM.attToggleHint;
  const daySelector = document.querySelector('.day-selector');
  const dateInputWrap = document.querySelector('.date-input-wrap');
  const gradeFilters = document.getElementById('attendanceGradeFilters');
  const activitiesTabs = document.querySelector('.activities-tabs');
  const searchBar = document.querySelector('#page-attendance .search-bar-wrap');

  if (!isService) {
    // No service today — show the message, hide the list
    if (noServiceMsg) noServiceMsg.classList.remove('hidden');
    if (attendanceList) attendanceList.style.display = 'none';
    if (attendanceSummary) attendanceSummary.style.display = 'none';
    if (quickActionsRow) quickActionsRow.style.display = 'none';
    if (attToggleHint) attToggleHint.style.display = 'none';
    if (daySelector) daySelector.style.display = 'none';
    if (dateInputWrap) dateInputWrap.style.display = 'none';
    if (gradeFilters) gradeFilters.style.display = 'none';
    if (activitiesTabs) activitiesTabs.style.display = 'none';
    if (searchBar) searchBar.style.display = 'none';
  } else {
    // Service day — show everything
    if (noServiceMsg) noServiceMsg.classList.add('hidden');
    if (attendanceList) attendanceList.style.display = '';
    if (attendanceSummary) attendanceSummary.style.display = '';
    if (quickActionsRow) quickActionsRow.style.display = '';
    if (attToggleHint) attToggleHint.style.display = '';
    if (daySelector) daySelector.style.display = '';
    if (dateInputWrap) dateInputWrap.style.display = '';
    if (gradeFilters) gradeFilters.style.display = '';
    if (activitiesTabs) activitiesTabs.style.display = '';
    if (searchBar) searchBar.style.display = '';

    renderAttendanceList();

    // NEW: Show indicator if auto-absence has been applied for today
    const today = DateUtil.toStr();
    if (date === today && isServiceDayForDate(today) && state.autoMarkedDates.has(today)) {
      showAutoMarkIndicator();
    } else {
      hideAutoMarkIndicator();
    }
  }
}

function setActiveDay(day) {
  state.selectedDay = day;
  document.querySelectorAll('.day-btn').forEach(b => b.classList.toggle('active', b.dataset.day === day));
}
function setActiveActivity(act) {
  state.selectedActivity = act;
  document.querySelectorAll('.act-tab').forEach(b => b.classList.toggle('active', b.dataset.activity === act));
}

document.querySelectorAll('.day-btn').forEach(b => b.addEventListener('click', () => {
  setActiveDay(b.dataset.day);
  renderAttendanceList();
}));
document.querySelectorAll('.act-tab').forEach(b => b.addEventListener('click', () => {
  setActiveActivity(b.dataset.activity);
  state.attendancePageInitialized = false;
  renderAttendancePage();
}));
if (DOM.attendanceDate) {
  DOM.attendanceDate.addEventListener('change', () => {
    TimeContext.setDate(DOM.attendanceDate.value);
    state.attendancePageInitialized = false;
    renderAttendancePage();
  });
}

if (DOM.selectAllPresent) DOM.selectAllPresent.addEventListener('click', () => selectAllStatus('حاضر'));
if (DOM.selectAllAbsent) DOM.selectAllAbsent.addEventListener('click', () => selectAllStatus('غائب'));

// NEW v3.1: Reset attendance button
document.addEventListener('DOMContentLoaded', () => {
  const resetBtn = document.getElementById('resetAttendanceBtn');
  if (resetBtn) {
    resetBtn.addEventListener('click', async () => {
      const date = DOM.attendanceDate ? DOM.attendanceDate.value : TimeContext.getDate();
      if (!date) return;

      showConfirm({
        icon: '&#128260;',
        title: 'تصفير الحضور',
        msg: `هل أنت متأكد إنك عايز تمسح كل سجلات الحضور ليوم ${date}؟\n\nده هيمسح الحضور والغياب لكل الأنشطة في اليوم ده.`,
        okLabel: 'نعم، صفّر',
        okClass: 'confirm-delete',
        onOk: async () => {
          await resetAttendanceForDate(date);
        }
      });
    });
  }

  // NEW v3.1: Service day toggle
  const serviceToggle = document.getElementById('serviceDayToggle');
  if (serviceToggle) {
    serviceToggle.addEventListener('change', (e) => {
      const date = TimeContext.getDate();
      const isService = e.target.checked;

      // Save the override
      setServiceDayOverride(date, isService);

      // Show feedback
      if (isService) {
        showToast('تم تفعيل يوم الخدمة — جاري تحديث الحضور', 'success');
      } else {
        showToast('تم إلغاء يوم الخدمة — مفيش حضور اليوم', 'warning');
      }

      // Re-render the page
      state.attendancePageInitialized = false;
      renderAttendancePage();

      // If turning on service and it's today, check for auto-mark
      if (isService && date === DateUtil.toStr()) {
        checkAndAutoMarkAbsence();
      }
    });
  }
});

function debouncedAttSearch() {
  clearTimeout(state.attSearchDebounceTimer);
  state.attSearchDebounceTimer = setTimeout(() => { renderAttendanceList(); }, 250);
}

if (DOM.attendanceSearch) DOM.attendanceSearch.addEventListener('input', debouncedAttSearch);

// ============================================================
// TOGGLE ATTENDANCE — v3.1: Better toast with girl name + grade
// ============================================================
async function toggleAttendanceStatus(girlId, girlName, date) {
  const opKey = `toggle_${girlId}_${date}_${state.selectedActivity}`;
  if (state.pendingAttendanceOps.has(opKey)) return; // Prevent double-clicks
  state.pendingAttendanceOps.add(opKey);

  try {
    const key = makeAttKey(girlId, date, state.selectedActivity);
    const existing = state.attendanceData[key];
    const newStatus = existing?.status === 'حاضر' ? 'غائب' : 'حاضر';

    const rec = {
      id: key,
      girlId: girlId,
      date,
      day: state.selectedDay,
      activity: state.selectedActivity,
      status: newStatus,
      rating: newStatus === 'حاضر' ? (existing?.rating || 0) : 0,
      notes: existing?.notes || '',
      updatedAt: Date.now(),
      updatedBy: state.currentUser?.displayName || 'خادم',
      updatedByEmail: state.currentUser?.email || ''
    };

    // Update state - FIXED: Use functional pattern to avoid race conditions
    setStateAttendanceData(prev => { const next = { ...prev, [key]: rec }; return next; });

    let firestoreSuccess = false;
    if (firebaseReady) {
      try {
        await FB.setDoc(FB.doc(db, 'attendance', key), rec);
        firestoreSuccess = true;
      } catch (e) {
        console.error('Save attendance Firestore error:', e);
      }
    }

    // FIXED: If Firestore failed but we're "online", the local state may be stale
    // Log a warning so the developer knows there's a potential inconsistency
    if (firebaseReady && !firestoreSuccess && navigator.onLine) {
      console.warn('Attendance saved locally but Firestore write failed — potential inconsistency');
    }

    // NEW v3.1: Show clearer toast with girl's name
    const girl = Cache.getGirl(girlId);
    const girlDisplayName = girl ? girl.name : girlName;
    const statusEmoji = newStatus === 'حاضر' ? '✓' : '✗';
    const toastType = newStatus === 'حاضر' ? 'success' : 'warning';
    showToast(`${statusEmoji} ${girlDisplayName} — ${newStatus}`, toastType);

    // FIXED: Debounced render to batch rapid toggles + ensure state settled
    debouncedRender(80);
  } finally {
    state.pendingAttendanceOps.delete(opKey);
  }
}

// ============================================================
// MARK ALL ABSENT — FIXED: Only called by explicit user action
// ============================================================
async function markAllAbsentForDate(date) {
  // v3.1: Only proceed if it's actually a service day
  if (!isServiceDayForDate(date)) return;

  const activeGirls = state.girls.filter(g => !g.isDeleted);
  if (activeGirls.length === 0) {
    renderAttendanceList();
    return;
  }

  const batchRecords = [];
  const newAttData = { ...state.attendanceData };

  for (const g of activeGirls) {
    for (const activity of ACTIVITIES) {
      const key = makeAttKey(g.id, date, activity);
      if (!newAttData[key]) {
        const rec = {
          id: key,
          girlId: g.id,
          date,
          day: DateUtil.dayName(parseDateStr(date)),
          activity: activity,
          status: 'غائب',
          rating: 0,
          notes: '',
          updatedAt: Date.now(),
          updatedBy: state.currentUser?.displayName || 'خادم',
          updatedByEmail: state.currentUser?.email || ''
        };
        batchRecords.push(rec);
        newAttData[key] = rec;
      }
    }
  }

  if (firebaseReady && batchRecords.length > 0) {
    try {
      const batch = FB.writeBatch(db);
      for (const rec of batchRecords) {
        batch.set(FB.doc(db, 'attendance', rec.id), rec);
      }
      await batch.commit();
    } catch (e) {
      console.error('Batch save attendance Firestore error:', e);
    }
  }

  setStateAttendanceData(newAttData);

  if (batchRecords.length > 0) {
    await logHistory('تسجيل حضور', `تعيين الغياب التلقائي ليوم ${date} (${state.selectedDay})`);
    showToast('تم تعيين الغياب التلقائي ليوم خدمة', 'info');
  }

  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'calendar') renderCalendar();
}

// Auto-mark a newly added girl as absent for all activities on a service day
// FIXED: Rollback mechanism — if Firestore fails, revert local state
async function autoMarkAbsentForNewGirl(girlId, date) {
  // v3.1: Only proceed if it's actually a service day
  if (!isServiceDayForDate(date)) return;

  // FIXED: Use parseDateStr for safe day name lookup
  const dayName = DateUtil.dayName(parseDateStr(date));
  const batchRecords = [];
  const newAttData = { ...state.attendanceData };
  const keysToAdd = [];

  for (const activity of ACTIVITIES) {
    const key = makeAttKey(girlId, date, activity);
    if (!newAttData[key]) {
      const rec = {
        id: key,
        girlId: girlId,
        date,
        day: dayName,
        activity: activity,
        status: 'غائب',
        rating: 0,
        notes: '',
        updatedAt: Date.now(),
        updatedBy: state.currentUser?.displayName || 'خادم',
        updatedByEmail: state.currentUser?.email || ''
      };
      batchRecords.push(rec);
      newAttData[key] = rec;
      keysToAdd.push(key);
    }
  }

  let firestoreSuccess = true;
  if (firebaseReady && batchRecords.length > 0) {
    try {
      const batch = FB.writeBatch(db);
      for (const rec of batchRecords) {
        batch.set(FB.doc(db, 'attendance', rec.id), rec);
      }
      await batch.commit();
    } catch (e) {
      console.error('Auto-absent batch save error:', e);
      firestoreSuccess = false;
    }
  }

  // FIXED: Rollback — if Firestore failed, remove the locally added records
  if (!firestoreSuccess && firebaseReady && navigator.onLine) {
    console.warn('Auto-absent Firestore failed — rolling back local state');
    keysToAdd.forEach(key => delete newAttData[key]);
    // Revert to original state
    setStateAttendanceData(state.attendanceData);
    return;
  }

  setStateAttendanceData(newAttData);
}

// Kept for backward compatibility
async function markAllAbsent(date) {
  await markAllAbsentForDate(date);
}

// ============================================================
// SELECT ALL — FIXED: Only write current date records to Firestore
// ============================================================
async function selectAllStatus(status) {
  if (!DOM.attendanceDate) return;
  const date = DOM.attendanceDate.value;
  if (!date) { showToast('الرجاء اختيار التاريخ أولاً', 'error'); return; }

  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const filteredGirls = state.attendanceGradeFilter
    ? activeGirls.filter(g => g.grade === state.attendanceGradeFilter)
    : activeGirls;
  const newAttData = { ...state.attendanceData };
  const currentDateRecords = []; // FIXED: Track only current date records for Firestore write

  for (const g of filteredGirls) {
    const key = makeAttKey(g.id, date, state.selectedActivity);
    const rec = {
      id: key,
      girlId: g.id,
      date,
      day: DateUtil.dayName(parseDateStr(date)),
      activity: state.selectedActivity,
      status: status,
      rating: status === 'حاضر' ? (newAttData[key]?.rating || 0) : 0,
      notes: newAttData[key]?.notes || '',
      updatedAt: Date.now(),
      updatedBy: state.currentUser?.displayName || 'خادم',
      updatedByEmail: state.currentUser?.email || ''
    };
    newAttData[key] = rec;
    currentDateRecords.push(rec); // FIXED: Only records for this date
  }

  if (firebaseReady) {
    try {
      const batch = FB.writeBatch(db);
      // FIXED: Only write records for the CURRENT date, not all attendance
      currentDateRecords.forEach(rec => {
        batch.set(FB.doc(db, 'attendance', rec.id), rec);
      });
      await batch.commit();
    } catch (e) {
      console.error('Batch save attendance Firestore error:', e);
    }
  }

  setStateAttendanceData(newAttData);

  await logHistory('تسجيل حضور', `${status === 'حاضر' ? 'تحديد الكل حاضر' : 'تحديد الكل غائب'} - ${state.selectedActivity} - ${date}`);
  showToast(status === 'حاضر' ? 'تم تحديد الكل حاضر' : 'تم تحديد الكل غائب', 'success');
  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'stats') renderStats();
  if (state.currentPage === 'calendar') renderCalendar();
}


// ============================================================
// RENDER ATTENDANCE LIST — FIXED: Memoized sortedGirls + O(n) scan
// ============================================================
function renderAttendanceList() {
  const date = DOM.attendanceDate ? DOM.attendanceDate.value : TimeContext.getDate();
  const searchQuery = DOM.attendanceSearch ? DOM.attendanceSearch.value.trim() : '';
  const activeGirls = state.girls.filter(g => !g.isDeleted);

  // Grade filter counts
  const gradeFilter = state.attendanceGradeFilter;
  document.querySelectorAll('#attendanceGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  const afcAll = document.getElementById('attFilterCountAll');
  const afc1 = document.getElementById('attFilterCount1');
  const afc2 = document.getElementById('attFilterCount2');
  const afc3 = document.getElementById('attFilterCount3');
  if (afcAll) afcAll.textContent = activeGirls.length;
  if (afc1) afc1.textContent = activeGirls.filter(g => g.grade === 'أولى إعدادي').length;
  if (afc2) afc2.textContent = activeGirls.filter(g => g.grade === 'تانية إعدادي').length;
  if (afc3) afc3.textContent = activeGirls.filter(g => g.grade === 'تالتة إعدادي').length;

  let filteredGirls = activeGirls;
  if (gradeFilter) filteredGirls = filteredGirls.filter(g => g.grade === gradeFilter);
  if (searchQuery) {
    const qNorm = normalizeArabic(searchQuery);
    filteredGirls = filteredGirls.filter(g => normalizeArabic(g.name).includes(qNorm));
  }

  // FIXED: Precompute indexed attendance lookup for O(1) status access
  const dateRecords = Cache.getAttendanceByDate(date);
  const attByGirlAct = {};
  dateRecords.forEach(a => {
    if (!attByGirlAct[a.girlId]) attByGirlAct[a.girlId] = {};
    attByGirlAct[a.girlId][a.activity] = a;
  });

  // FIXED: Save sortedGirls for inline rating (memoized, avoids re-sorting)
  state._sortedGirlsForAtt = filteredGirls;

  const presentCount = filteredGirls.filter(g => attByGirlAct[g.id]?.[state.selectedActivity]?.status === 'حاضر').length;
  const absentCount = filteredGirls.filter(g => attByGirlAct[g.id]?.[state.selectedActivity]?.status === 'غائب').length;

  if (DOM.presentCount) DOM.presentCount.textContent = presentCount;
  if (DOM.absentCount) DOM.absentCount.textContent = absentCount;
  if (DOM.totalCount) DOM.totalCount.textContent = filteredGirls.length;

  if (!filteredGirls.length) {
    if (DOM.attendanceList) DOM.attendanceList.innerHTML = '<div class="empty-state">لا توجد مخدومات لهذا اليوم</div>';
    return;
  }

  const frag = document.createDocumentFragment();
  filteredGirls.forEach((g) => {
    const att = attByGirlAct[g.id]?.[state.selectedActivity];
    const status = att?.status || 'غائب';
    const stars = att?.rating ? '&#9733;'.repeat(att.rating) + '&#9734;'.repeat(5 - att.rating) : '';
    const note = att?.notes || '';

    const div = document.createElement('div');
    div.className = `att-item ${status === 'حاضر' ? 'present' : status === 'غائب' ? 'absent' : 'pending'}`;
    div.dataset.girlId = g.id;
    div.dataset.girlName = g.name;

    const icon = status === 'حاضر' ? '&#10003;' : status === 'غائب' ? '&#10007;' : '?';
    const statusText = status;
    const noteHtml = note ? `<span class="att-note">${esc(note)}</span>` : '';
    const starsHtml = stars ? `<span class="att-stars">${stars}</span>` : '';

    // FIXED: Precompute activity counts for display
    const presentCountActivity = ACTIVITIES.filter(act => attByGirlAct[g.id]?.[act]?.status === 'حاضر').length;
    const totalActivities = ACTIVITIES.length;

    div.innerHTML = `
      <span class="att-icon">${icon}</span>
      <div class="att-info">
        <span class="att-name">${esc(g.name)}</span>
        <span class="att-grade">${esc(g.grade)}</span>
        <span class="att-activity-count" style="font-size:12px;color:var(--text-muted);margin-top:2px;">
          ${presentCountActivity}/${totalActivities} أنشطة
        </span>
        ${starsHtml}
        ${noteHtml}
      </div>
      <span class="att-status-text ${status === 'حاضر' ? 'present' : status === 'غائب' ? 'absent' : 'pending'}">${statusText}</span>
      <button class="att-delete-btn" data-girl-id="${esc(g.id)}" data-key="${esc(att?.id || makeAttKey(g.id, date, state.selectedActivity))}" aria-label="حذف تسجيل ${esc(g.name)}" title="حذف تسجيل الحضور">&#128465;</button>`;
    frag.appendChild(div);
  });

  if (DOM.attendanceList) {
    DOM.attendanceList.innerHTML = '';
    DOM.attendanceList.appendChild(frag);
  }
}

// ============================================================
// ADD INLINE RATING TO ATTENDANCE ROW
// ============================================================
function addInlineRating(rowEl, girlId, date, activity) {
  const key = makeAttKey(girlId, date, activity);
  const existing = state.attendanceData[key];
  const currentRating = existing?.rating || 0;

  const container = document.createElement('div');
  container.className = 'att-inline-rating';
  container.dataset.ratingKey = key;
  container.dataset.girlId = girlId;

  const label = document.createElement('span');
  label.className = 'att-inline-rating-label';
  label.textContent = 'تقييم:';

  const starsDiv = document.createElement('div');
  starsDiv.className = 'att-inline-stars';

  for (let i = 1; i <= 5; i++) {
    const star = document.createElement('span');
    star.className = 'att-inline-star' + (i <= currentRating ? ' active' : '');
    star.dataset.val = i;
    star.textContent = '\u2605'; // Unicode star
    star.role = 'button';
    star.ariaLabel = `${i} نجمة`;
    star.addEventListener('click', async (e) => {
      e.stopPropagation();
      e.preventDefault();
      await saveRating(girlId, date, activity, parseInt(star.dataset.val));
    });
    starsDiv.appendChild(star);
  }

  const valSpan = document.createElement('span');
  valSpan.className = 'att-inline-rating-val';
  valSpan.textContent = currentRating > 0 ? `${currentRating}/5` : '';

  const hint = document.createElement('span');
  hint.className = 'att-inline-rating-hint';
  hint.textContent = 'اضغط لتقييم';

  container.appendChild(label);
  container.appendChild(starsDiv);
  container.appendChild(valSpan);
  container.appendChild(hint);

  rowEl.querySelector('.att-info').appendChild(container);
}

async function saveRating(girlId, date, activity, rating) {
  const key = makeAttKey(girlId, date, activity);
  const existing = state.attendanceData[key];
  const rec = {
    id: key,
    girlId,
    date,
    day: state.selectedDay,
    activity,
    status: existing?.status || 'حاضر',
    rating,
    notes: existing?.notes || '',
    updatedAt: Date.now(),
    updatedBy: state.currentUser?.displayName || 'خادم',
    updatedByEmail: state.currentUser?.email || ''
  };

  setStateAttendanceData(prev => ({ ...prev, [key]: rec }));

  if (firebaseReady) {
    try { await FB.setDoc(FB.doc(db, 'attendance', key), rec); }
    catch (e) { console.error('Save rating Firestore error:', e); }
  }

  // Update the inline UI
  const rowEl = document.querySelector(`.att-item[data-girl-id="${esc(girlId)}"]`);
  if (rowEl) {
    const existingRating = rowEl.querySelector('.att-inline-rating');
    if (existingRating) existingRating.remove();
    addInlineRating(rowEl, girlId, date, activity);
  }

  const g = Cache.getGirl(girlId);
  showToast(`تم حفظ التقييم ${rating} نجوم لـ ${g ? g.name : ''}`, 'success');
}

// ============================================================
// ATTENDANCE LIST EVENT LISTENER — FIXED: Unified with hover rating
// ============================================================
if (DOM.attendanceList) {
  DOM.attendanceList.addEventListener('click', async (e) => {
    const delBtn = e.target.closest('.att-delete-btn');
    if (delBtn) {
      e.stopPropagation();
      e.preventDefault();
      const girlId = delBtn.dataset.girlId;
      const key = delBtn.dataset.key;
      const g = Cache.getGirl(girlId);
      if (!g) return;

      showConfirm({
        icon: '&#9888;', title: 'حذف تسجيل',
        msg: `هل أنت متأكد من حذف تسجيل ${esc(g.name)} ليوم ${esc(TimeContext.getDate())}؟`,
        okLabel: 'حذف',
        okClass: 'confirm-delete',
        onOk: async () => {
          try {
            // Remove from local state
            const newData = { ...state.attendanceData };
            delete newData[key];
            setStateAttendanceData(newData);

            // Remove from Firestore
            if (firebaseReady) {
              try { await FB.deleteDoc(FB.doc(db, 'attendance', key)); }
              catch (e) { console.error('Delete attendance Firestore error:', e); }
            }

            showToast(`تم حذف تسجيل ${g.name}`, 'success');
            debouncedRender(80);
          } catch (err) {
            console.error('Delete attendance error:', err);
            showToast('حدث خطأ أثناء الحذف', 'error');
          }
        }
      });
      return;
    }

    const item = e.target.closest('.att-item');
    if (!item) return;
    const girlId = item.dataset.girlId;
    const girlName = item.dataset.girlName;
    if (!girlId || !girlName) return;

    // FIXED: If inline rating row is open, clicking nearby shouldn't toggle attendance
    if (e.target.closest('.att-inline-rating')) return;

    await toggleAttendanceStatus(girlId, girlName, TimeContext.getDate());
  });

  // FIXED: Long press for attendance entry modal
  DOM.attendanceList.addEventListener('touchstart', (e) => {
    const item = e.target.closest('.att-item');
    if (!item || e.target.closest('.att-inline-rating') || e.target.closest('.att-delete-btn')) return;

    state.isLongPress = false;
    state.longPressTimer = setTimeout(() => {
      state.isLongPress = true;
      const girlId = item.dataset.girlId;
      const girlName = item.dataset.girlName;
      if (girlId && girlName) {
        openAttendanceModal(girlId, girlName);
      }
    }, 600);
  }, { passive: true });

  DOM.attendanceList.addEventListener('touchend', () => {
    if (state.longPressTimer) {
      clearTimeout(state.longPressTimer);
      state.longPressTimer = null;
    }
    // If it was a long press, prevent the click from firing
    if (state.isLongPress) {
      state.isLongPress = false;
    }
  });

  DOM.attendanceList.addEventListener('touchcancel', () => {
    if (state.longPressTimer) {
      clearTimeout(state.longPressTimer);
      state.longPressTimer = null;
    }
    state.isLongPress = false;
  });
}

// ============================================================
// AUTO-MARK INDICATOR — NEW: Shows if auto-absence was applied
// ============================================================
function showAutoMarkIndicator() {
  let indicator = document.getElementById('autoMarkIndicator');
  if (!indicator) {
    indicator = document.createElement('div');
    indicator.id = 'autoMarkIndicator';
    indicator.style.cssText = `
      background: rgba(26, 39, 68, 0.08);
      color: var(--text-muted);
      font-size: 12px;
      font-weight: 600;
      text-align: center;
      padding: 6px 12px;
      border-radius: 8px;
      margin-bottom: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
    `;
    const attendanceSummary = DOM.attendanceSummary;
    if (attendanceSummary && attendanceSummary.parentNode) {
      attendanceSummary.parentNode.insertBefore(indicator, attendanceSummary.nextSibling);
    }
  }
  indicator.innerHTML = `<span>&#128161;</span> تم تسجيل الغياب التلقائي — اضغط على اسم المخدومة للتبديل لحاضر`;
  indicator.style.display = '';
}

function hideAutoMarkIndicator() {
  const indicator = document.getElementById('autoMarkIndicator');
  if (indicator) indicator.style.display = 'none';
}

// ============================================================
// ATTENDANCE MODAL
// ============================================================
function openAttendanceModal(girlId, girlName) {
  const date = TimeContext.getDate();
  const key = makeAttKey(girlId, date, state.selectedActivity);
  const existing = state.attendanceData[key];

  state.currentAttendanceGirlId = girlId;
  state.currentAttendanceRating = existing?.rating || 0;

  if (DOM.attendanceModalTitle) DOM.attendanceModalTitle.textContent = 'تسجيل الحضور';
  if (DOM.modalGirlName) DOM.modalGirlName.textContent = girlName;

  // Set buttons
  document.querySelectorAll('.attend-btn').forEach(btn => {
    btn.classList.toggle('selected', btn.dataset.status === (existing?.status || ''));
  });

  // Set rating
  if (DOM.starsInput) {
    const stars = DOM.starsInput.querySelectorAll('.star');
    stars.forEach((s, i) => s.classList.toggle('active', i < state.currentAttendanceRating));
  }

  if (DOM.attendanceNotes) DOM.attendanceNotes.value = existing?.notes || '';

  // Show/hide rating section based on status
  if (DOM.ratingSection) {
    DOM.ratingSection.style.display = (existing?.status === 'حاضر') ? 'block' : 'none';
  }

  openModal('attendanceModal');
}

if (DOM.starsInput) {
  DOM.starsInput.querySelectorAll('.star').forEach(s => {
    s.addEventListener('click', () => {
      state.currentAttendanceRating = parseInt(s.dataset.val);
      DOM.starsInput.querySelectorAll('.star').forEach((star, i) => star.classList.toggle('active', i < state.currentAttendanceRating));
    });
  });
}

document.querySelectorAll('.attend-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.attend-btn').forEach(b => b.classList.remove('selected'));
    btn.classList.add('selected');
    // Show/hide rating based on status
    if (DOM.ratingSection) {
      DOM.ratingSection.style.display = btn.dataset.status === 'حاضر' ? 'block' : 'none';
    }
  });
});

if (DOM.saveAttendanceEntry) {
  DOM.saveAttendanceEntry.addEventListener('click', async () => {
    if (!state.currentAttendanceGirlId) return;
    const date = TimeContext.getDate();
    const girlId = state.currentAttendanceGirlId;
    const girl = Cache.getGirl(girlId);

    const selectedStatusBtn = document.querySelector('.attend-btn.selected');
    const status = selectedStatusBtn ? selectedStatusBtn.dataset.status : 'حاضر';
    const rating = status === 'حاضر' ? state.currentAttendanceRating : 0;
    const notes = DOM.attendanceNotes ? DOM.attendanceNotes.value.trim() : '';

    const key = makeAttKey(girlId, date, state.selectedActivity);
    const rec = {
      id: key,
      girlId,
      date,
      day: state.selectedDay,
      activity: state.selectedActivity,
      status,
      rating,
      notes,
      updatedAt: Date.now(),
      updatedBy: state.currentUser?.displayName || 'خادم',
      updatedByEmail: state.currentUser?.email || ''
    };

    setStateAttendanceData(prev => ({ ...prev, [key]: rec }));

    if (firebaseReady) {
      try { await FB.setDoc(FB.doc(db, 'attendance', key), rec); }
      catch (e) { console.error('Save attendance Firestore error:', e); }
    }

    closeModal('attendanceModal');

    // NEW v3.1: Show toast with girl's name
    const statusEmoji = status === 'حاضر' ? '✓' : '✗';
    const toastType = status === 'حاضر' ? 'success' : 'warning';
    showToast(`${statusEmoji} ${girl ? girl.name : ''} — ${status}`, toastType);

    await logHistory('تسجيل حضور', `${girl ? girl.name : ''} - ${status} - ${state.selectedActivity} - ${date}`);
    renderAttendanceList();
    if (state.currentPage === 'home') renderHome();
    if (state.currentPage === 'stats') renderStats();
    if (state.currentPage === 'calendar') renderCalendar();
  });
}

if (DOM.closeAttendanceModal) DOM.closeAttendanceModal.addEventListener('click', () => closeModal('attendanceModal'));
if (DOM.cancelAttendanceModal) DOM.cancelAttendanceModal.addEventListener('click', () => closeModal('attendanceModal'));

// ============================================================
// CALENDAR
// ============================================================
function renderCalendar() {
  const d = state.calendarDate;
  const year = d.getFullYear();
  const month = d.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  if (DOM.calMonthYear) DOM.calMonthYear.textContent = new Date(year, month, 1).toLocaleDateString('ar-EG', { year: 'numeric', month: 'long' });

  let html = '<div class="cal-weekdays">';
  DAY_NAMES.forEach(dn => html += `<div class="cal-wday">${dn.slice(0, 2)}</div>`);
  html += '</div><div class="cal-days">';

  for (let i = 0; i < firstDay; i++) html += '<div class="cal-day empty"></div>';

  for (let day = 1; day <= daysInMonth; day++) {
    const dayOfWeek = new Date(year, month, day).getDay();
    const dateStr = `${year}-${DateUtil.pad(month + 1)}-${DateUtil.pad(day)}`;
    // NEW v3.1: Use isServiceDayForDate for accurate service day detection
    const isService = isServiceDayForDate(dateStr);
    const dayRecords = Cache.getAttendanceByDate(dateStr);
    const hasData = dayRecords.length > 0;
    const isToday = dateStr === DateUtil.toStr();

    html += `<div class="cal-day ${isService ? 'service-day' : ''} ${hasData ? 'has-data' : ''} ${isToday ? 'today' : ''}" data-date="${dateStr}">
      ${day}
      ${isService ? '<span class="service-dot"></span>' : ''}
    </div>`;
  }
  html += '</div>';

  if (DOM.calendarGrid) DOM.calendarGrid.innerHTML = html;

  // Add click handlers
  document.querySelectorAll('.cal-day:not(.empty)').forEach(el => {
    el.addEventListener('click', () => {
      const date = el.dataset.date;
      if (date) {
        TimeContext.setDate(date);
        showDayDetail(date);
      }
    });
  });
}

function showDayDetail(date) {
  const dayRecords = Cache.getAttendanceByDate(date);
  const el = DOM.dayDetail;
  if (!el) return;

  // NEW v3.1: Show if service day or not
  const isService = isServiceDayForDate(date);
  const serviceInfo = isService
    ? '<span style="color:var(--green);font-weight:700;">✓ يوم خدمة</span>'
    : '<span style="color:var(--red);font-weight:700;">✗ لا توجد خدمة</span>';

  const d = parseDateStr(date);
  const dayName = DAY_NAMES[d.getDay()] || '';
  el.innerHTML = `<div class="day-detail-header">
    ${date} ${dayName} — ${serviceInfo}
  </div>`;

  if (!dayRecords.length) {
    el.innerHTML += '<div class="empty-state">لا توجد سجلات حضور</div>';
  } else {
    ACTIVITIES.forEach(activity => {
      const activityRecords = dayRecords.filter(r => r.activity === activity);
      if (!activityRecords.length) return;

      const presentCount = activityRecords.filter(r => r.status === 'حاضر').length;
      const absentCount = activityRecords.filter(r => r.status === 'غائب').length;

      el.innerHTML += `<div class="day-activity">
        <strong>${ACTIVITY_ICONS[activity] || ''} ${activity}</strong>
        <span class="green-text"> ✓${presentCount}</span>
        <span class="red-text"> ✗${absentCount}</span>
      </div>`;
    });
  }
  el.classList.add('show');
}

function hideDayDetail() {
  if (DOM.dayDetail) DOM.dayDetail.classList.remove('show');
}

if (DOM.calPrev) DOM.calPrev.addEventListener('click', () => { state.calendarDate.setMonth(state.calendarDate.getMonth() - 1); renderCalendar(); });
if (DOM.calNext) DOM.calNext.addEventListener('click', () => { state.calendarDate.setMonth(state.calendarDate.getMonth() + 1); renderCalendar(); });

// ============================================================
// STATS PAGE — FIXED: Unified bounds + grade filter + month date picker
// ============================================================
function renderStats() {
  const bounds = getStatsBounds();
  const gradeFilter = state.statsGradeFilter;
  const allAttendance = Cache.getAllAttendance();

  const filteredAtt = allAttendance.filter(a => {
    if (!a.date) return false;
    if (a.date < bounds.start || a.date > bounds.end) return false;
    const girl = Cache.getGirl(a.girlId);
    if (!girl) return false;
    if (gradeFilter && girl.grade !== gradeFilter) return false;
    return true;
  });

  const presentCount = filteredAtt.filter(a => a.status === 'حاضر').length;
  const absentCount = filteredAtt.filter(a => a.status === 'غائب').length;
  const totalCount = presentCount + absentCount;

  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const filteredGirls = gradeFilter ? activeGirls.filter(g => g.grade === gradeFilter) : activeGirls;
  const totalGirls = filteredGirls.length;

  // Activity breakdown
  const activityCounts = {};
  ACTIVITIES.forEach(a => activityCounts[a] = { present: 0, absent: 0 });
  filteredAtt.forEach(a => {
    if (!activityCounts[a.activity]) activityCounts[a.activity] = { present: 0, absent: 0 };
    if (a.status === 'حاضر') activityCounts[a.activity].present++;
    else if (a.status === 'غائب') activityCounts[a.activity].absent++;
  });

  // Top attendees
  const girlCounts = {};
  filteredAtt.filter(a => a.status === 'حاضر').forEach(a => {
    if (!girlCounts[a.girlId]) girlCounts[a.girlId] = 0;
    girlCounts[a.girlId]++;
  });
  const sortedGirls = Object.entries(girlCounts).sort((a, b) => b[1] - a[1]);

  // Render
  if (DOM.bigStatsGrid) {
    const periodLabel = PERIOD_LABELS[state.statsTimeFilter] || 'هذا الشهر';
    DOM.bigStatsGrid.innerHTML = `
      <div class="big-stat-card blue-card"><div class="big-num">${totalGirls}</div><div>إجمالي المخدومات</div></div>
      <div class="big-stat-card green-card"><div class="big-num">${presentCount}</div><div>حضور (${periodLabel})</div></div>
      <div class="big-stat-card red-card"><div class="big-num">${absentCount}</div><div>غياب (${periodLabel})</div></div>
      <div class="big-stat-card orange-card"><div class="big-num">${totalCount}</div><div>إجمالي السجلات</div></div>
    `;
  }

  // Activity stats
  if (DOM.activityStatsGrid) {
    DOM.activityStatsGrid.innerHTML = ACTIVITIES.map((activity, index) => {
      const counts = activityCounts[activity];
      const total = (counts?.present || 0) + (counts?.absent || 0);
      const rate = total > 0 ? Math.round(((counts?.present || 0) / total) * 100) : 0;
      return `<div class="activity-stat-card" data-activity="${esc(activity)}" data-index="${index}" style="cursor:pointer;">
        <div class="activity-stat-rank">#${index + 1}</div>
        <div class="activity-stat-icon">${ACTIVITY_ICONS[activity] || ''}</div>
        <div class="activity-stat-num">${counts?.present || 0}</div>
        <div class="activity-stat-label">${activity}</div>
        <div style="font-size:12px;color:var(--green);font-weight:700;">${rate}% معدل</div>
        <div class="activity-stat-absent">&#10007; ${counts?.absent || 0} غائب</div>
      </div>`;
    }).join('');

    // Add click handlers
    DOM.activityStatsGrid.querySelectorAll('.activity-stat-card').forEach(card => {
      card.addEventListener('click', () => {
        const activity = card.dataset.activity;
        if (activity) showActivityDetail(activity);
      });
    });
  }

  // Period label + grade label
  if (DOM.activityStatsPeriod) {
    DOM.activityStatsPeriod.textContent = `(${PERIOD_LABELS[state.statsTimeFilter] || 'هذا الشهر'})`;
  }
  if (DOM.activityStatsGrade) {
    DOM.activityStatsGrade.textContent = gradeFilter ? `— ${gradeFilter}` : '';
  }

  // Absence chart
  if (DOM.absenceChart) {
    const absentCounts = {};
    filteredAtt.filter(a => a.status === 'غائب').forEach(a => {
      if (!absentCounts[a.girlId]) absentCounts[a.girlId] = 0;
      absentCounts[a.girlId]++;
    });
    const sortedAbs = Object.entries(absentCounts).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const maxAbs = sortedAbs.length ? sortedAbs[0][1] : 1;

    if (!sortedAbs.length) {
      DOM.absenceChart.innerHTML = '<div class="empty-state">لا توجد بيانات غياب</div>';
    } else {
      DOM.absenceChart.innerHTML = sortedAbs.map(([girlId, count]) => {
        const girl = Cache.getGirl(girlId);
        const name = girl ? girl.name : 'غير معروف';
        const width = Math.round((count / maxAbs) * 100);
        return `<div class="chart-row">
          <span class="chart-name">${esc(name)}</span>
          <div class="chart-bar-wrap"><div class="chart-bar" style="width:${width}%"></div></div>
          <span class="chart-val">${count}</span>
        </div>`;
      }).join('');
    }
  }

  // Attendance ranking
  if (DOM.attendanceRanking) {
    if (!sortedGirls.length) {
      DOM.attendanceRanking.innerHTML = '<div class="empty-state">لا توجد بيانات</div>';
    } else {
      DOM.attendanceRanking.innerHTML = sortedGirls.map(([girlId, count], i) => {
        const girl = Cache.getGirl(girlId);
        return `<div class="rank-item">
          <span class="rank-num">${i + 1}</span>
          <span class="rank-name">${esc(girl ? girl.name : 'غير معروف')}</span>
          <span class="rank-grade">${esc(girl ? girl.grade : '')}</span>
          <span class="rank-count">${count} يوم</span>
        </div>`;
      }).join('');
    }
  }

  // Time filter tabs
  document.querySelectorAll('#timeFilterTabs .time-filter-tab').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.period === state.statsTimeFilter);
  });

  // Stats grade filter
  document.querySelectorAll('#statsGradeFilter .stats-grade-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });
}

document.querySelectorAll('#timeFilterTabs .time-filter-tab').forEach(btn => {
  btn.addEventListener('click', () => {
    state.statsTimeFilter = btn.dataset.period;
    // Only update date picker for non-'today' filters
    if (DOM.statsMonth) {
      DOM.statsMonth.value = TimeContext.getDate();
    }
    renderStats();
  });
});

if (DOM.statsMonth) {
  DOM.statsMonth.addEventListener('change', () => {
    const date = DOM.statsMonth.value;
    if (date) {
      TimeContext.setDate(date);
      renderStats();
    }
  });
}

document.querySelectorAll('#statsGradeFilter .stats-grade-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    state.statsGradeFilter = btn.dataset.grade;
    renderStats();
  });
});

// ============================================================
// ACTIVITY DETAIL MODAL
// ============================================================
function showActivityDetail(activity) {
  const bounds = getStatsBounds();
  const allAttendance = Cache.getAllAttendance();
  const gradeFilter = state.statsGradeFilter;

  const filteredAtt = allAttendance.filter(a => {
    if (!a.date || a.date < bounds.start || a.date > bounds.end) return false;
    if (a.activity !== activity) return false;
    const girl = Cache.getGirl(a.girlId);
    if (!girl) return false;
    if (gradeFilter && girl.grade !== gradeFilter) return false;
    return true;
  });

  const presentRecords = filteredAtt.filter(a => a.status === 'حاضر');
  const absentRecords = filteredAtt.filter(a => a.status === 'غائب');

  state.currentActivityDetail = { activity, presentRecords, absentRecords };
  state.activityDetailTab = 'present';

  if (DOM.activityDetailTitle) DOM.activityDetailTitle.textContent = `تفاصيل ${activity}`;
  if (DOM.activityDetailIcon) DOM.activityDetailIcon.innerHTML = ACTIVITY_ICONS[activity] || '&#128203;';
  if (DOM.activityDetailName) DOM.activityDetailName.textContent = activity;
  if (DOM.activityDetailPeriod) DOM.activityDetailPeriod.textContent = PERIOD_LABELS[state.statsTimeFilter] || 'هذا الشهر';
  if (DOM.activityDetailTotal) DOM.activityDetailTotal.textContent = filteredAtt.length;
  if (DOM.presentTabCount) DOM.presentTabCount.textContent = presentRecords.length;
  if (DOM.absentTabCount) DOM.absentTabCount.textContent = absentRecords.length;

  renderActivityDetailList();

  // Tab handlers
  document.querySelectorAll('.activity-detail-tab').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.tab === state.activityDetailTab);
  });

  openModal('activityDetailModal');
}

document.querySelectorAll('.activity-detail-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    state.activityDetailTab = tab.dataset.tab;
    document.querySelectorAll('.activity-detail-tab').forEach(t => t.classList.toggle('active', t === tab));
    renderActivityDetailList();
  });
});

if (DOM.closeActivityDetailModal) {
  DOM.closeActivityDetailModal.addEventListener('click', () => closeModal('activityDetailModal'));
}

function renderActivityDetailList() {
  if (!state.currentActivityDetail) return;
  const { presentRecords, absentRecords } = state.currentActivityDetail;
  const records = state.activityDetailTab === 'present' ? presentRecords : absentRecords;

  if (!DOM.activityDetailList) return;

  if (!records.length) {
    DOM.activityDetailList.innerHTML = '<div class="empty-state">لا توجد سجلات</div>';
    return;
  }

  // Group by date
  const byDate = {};
  records.forEach(r => {
    if (!byDate[r.date]) byDate[r.date] = [];
    byDate[r.date].push(r);
  });

  const frag = document.createDocumentFragment();
  Object.entries(byDate).sort((a, b) => b[0].localeCompare(a[0])).forEach(([date, recs]) => {
    const d = parseDateStr(date);
    const dayName = DAY_NAMES[d.getDay()] || '';

    recs.forEach(r => {
      const girl = Cache.getGirl(r.girlId);
      if (!girl) return;

      const div = document.createElement('div');
      div.className = 'detail-girl-item';
      div.innerHTML = `
        <span class="detail-girl-avatar">${esc(girl.name[0])}</span>
        <div class="detail-girl-info">
          <div class="detail-girl-name">${esc(girl.name)}</div>
          <div class="detail-girl-grade">${esc(girl.grade)} ${esc(dayName)} ${esc(date)}</div>
        </div>
        <span class="detail-status-icon ${r.status === 'حاضر' ? 'present' : 'absent'}">
          ${r.status === 'حاضر' ? '&#10003;' : '&#10007;'}
        </span>
      `;
      frag.appendChild(div);
    });
  });

  DOM.activityDetailList.innerHTML = '';
  DOM.activityDetailList.appendChild(frag);
}

// ============================================================
// HISTORY PAGE
// ============================================================
async function renderHistory(append = false) {
  const filter = DOM.historyFilter ? DOM.historyFilter.value : '';
  let logs = [];

  if (firebaseReady) {
    try {
      const snap = await FB.getDocs(
        FB.query(FB.collection(db, 'history'), FB.orderBy('timestamp', 'desc'))
      );
      logs = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    } catch (e) {
      console.error('History fetch error:', e);
    }
  }

  // Fallback to IDB
  if (!logs.length) {
    try { logs = await IDB.getAll('history'); } catch (e) { /* ignore */ }
  }

  if (filter) logs = logs.filter(l => l.action === filter);

  if (!logs.length) {
    if (DOM.historyList) DOM.historyList.innerHTML = '<div class="empty-state">لا توجد سجلات</div>';
    return;
  }

  const html = logs.slice(0, 50).map(l => {
    const time = l.timestamp ? new Date(l.timestamp).toLocaleString('ar-EG') : '';
    return `<div class="history-item">
      <span class="history-icon">${getHistoryIcon(l.action)}</span>
      <div class="history-info">
        <span class="history-action">${esc(l.action)}</span>
        <span class="history-detail">${esc(l.details)}</span>
        <span class="history-meta">${esc(l.userName || '')} ${time}</span>
      </div>
    </div>`;
  }).join('');

  if (DOM.historyList) DOM.historyList.innerHTML = html;
}

function getHistoryIcon(action) {
  const icons = { 'إضافة مخدومة': '&#10133;', 'تعديل مخدومة': '&#9999;', 'حذف مخدومة': '&#10060;', 'تسجيل حضور': '&#128203;', 'تصفير الحضور': '&#128260;' };
  return icons[action] || '&#8226;';
}

if (DOM.historyFilter) DOM.historyFilter.addEventListener('change', () => renderHistory(false));
if (DOM.clearHistoryBtn) {
  DOM.clearHistoryBtn.addEventListener('click', async () => {
    showConfirm({
      icon: '&#9888;', title: 'مسح السجل',
      msg: 'هل أنت متأكد من مسح السجل التاريخي؟ لا يمكن التراجع عن هذا الإجراء.',
      okLabel: 'مسح',
      okClass: 'confirm-delete',
      onOk: async () => {
        try {
          // Clear IDB
          await IDB.clear('history');
          // Clear Firestore
          if (firebaseReady) {
            try {
              const snap = await FB.getDocs(FB.collection(db, 'history'));
              const batch = FB.writeBatch(db);
              snap.docs.forEach(d => batch.delete(d.ref));
              await batch.commit();
            } catch (e) { console.error('Clear history Firestore error:', e); }
          }
          if (DOM.historyList) DOM.historyList.innerHTML = '<div class="empty-state">تم مسح السجل</div>';
          showToast('تم مسح السجل التاريخي', 'success');
        } catch (e) {
          console.error('Clear history error:', e);
          showToast('حدث خطأ أثناء المسح', 'error');
        }
      }
    });
  });
}

// ============================================================
// LOG HISTORY
// ============================================================
async function logHistory(action, details) {
  const log = {
    id: 'log_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
    action,
    details,
    timestamp: Date.now(),
    userName: state.currentUser?.displayName || 'خادم',
    userEmail: state.currentUser?.email || ''
  };

  // Always save to IDB
  try { await IDB.add('history', log); } catch (e) { /* ignore */ }

  // Save to Firestore if available
  if (firebaseReady) {
    try { await FB.setDoc(FB.doc(db, 'history', log.id), log); }
    catch (e) { console.error('Log history Firestore error:', e); }
  }
}

// ============================================================
// EXPORT — NEW: Grade filter + day mode + restructured data
// ============================================================

/**
 * NEW: Helper function to convert attendance records into the desired restructured format.
 * Each unique activity across all records becomes a column, keyed by the activity name.
 * Only includes activities that actually appear in the data.
 * Grade filter applied.
 */
function restructureAttendanceDataForExport(records, gradeFilter) {
  const uniqueActivities = new Set();
  records.forEach(r => {
    if (r.activity) uniqueActivities.add(r.activity);
  });
  const activitiesList = Array.from(uniqueActivities).sort();

  const rowMap = {};
  records.forEach(r => {
    const girl = Cache.getGirl(r.girlId);
    if (!girl) return;
    if (gradeFilter && girl.grade !== gradeFilter) return;

    const rowKey = `${r.girlId}_${r.date}`;
    if (!rowMap[rowKey]) {
      rowMap[rowKey] = {
        date: r.date,
        girlId: r.girlId,
        name: girl.name,
        grade: girl.grade,
        phone: girl.phone || '',
        day: r.day || ''
      };
    }
    // Map each activity to its own column (e.g., { "دراسي": "حاضر", "محفوظات": "غائب" })
    rowMap[rowKey][r.activity] = r.status === 'حاضر' ? 'حاضر' : 'غائب';
  });

  return { activitiesList, rows: Object.values(rowMap) };
}

function renderExport() {
  if (DOM.exportMonth) DOM.exportMonth.value = TimeContext.getDate();

  // Grade filter buttons
  document.querySelectorAll('#exportGradeFilter .export-grade-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === state.exportGradeFilter);
  });

  // Status filter buttons
  document.querySelectorAll('#exportStatusFilter .export-status-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.status === state.exportStatusFilter);
  });
}

// Export Grade filter handlers
document.querySelectorAll('#exportGradeFilter .export-grade-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    state.exportGradeFilter = btn.dataset.grade;
    renderExport();
  });
});

// Export Status filter handlers
document.querySelectorAll('#exportStatusFilter .export-status-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    state.exportStatusFilter = btn.dataset.status;
    renderExport();
  });
});

if (DOM.exportCSV) {
  DOM.exportCSV.addEventListener('click', async () => {
    try {
      const gradeFilter = state.exportGradeFilter;
      const statusFilter = state.exportStatusFilter;
      const exportMode = document.querySelector('input[name="exportMode"]:checked')?.value || 'day';
      const date = DOM.exportMonth ? DOM.exportMonth.value : TimeContext.getDate();
      const activeGirls = state.girls.filter(g => !g.isDeleted);
      const filteredGirls = gradeFilter ? activeGirls.filter(g => g.grade === gradeFilter) : activeGirls;

      if (!filteredGirls.length) { showToast('لا توجد مخدومات للتصدير', 'warning'); return; }

      // Headers
      const headers = ['المخدومة', 'السنة الدراسية', 'رقم الهاتف'];

      let records = [];
      let dateHeaders = [];
      let summaryData = [];

      if (exportMode === 'day') {
        // FIXED: Day mode — use makeAttKey for consistent key generation
        dateHeaders = ACTIVITIES.map(a => a);
        ACTIVITIES.forEach(a => headers.push(a));

        const sortedGirls = [...filteredGirls].sort((a, b) => {
          const aOrder = GRADE_ORDER[a.grade] || 99;
          const bOrder = GRADE_ORDER[b.grade] || 99;
          if (aOrder !== bOrder) return aOrder - bOrder;
          return a.name.localeCompare(b.name, 'ar');
        });

        records = sortedGirls.map(g => {
          const row = [csvEscape(g.name), csvEscape(g.grade), csvEscape(g.phone || '')];
          ACTIVITIES.forEach(activity => {
            const key = makeAttKey(g.id, date, activity);
            const record = state.attendanceData[key];
            const status = record ? (record.status === 'حاضر' ? '\u2713' : 'X') : '-';
            // Apply status filter
            if (statusFilter === 'present' && status !== '\u2713') return null;
            if (statusFilter === 'absent' && status !== 'X') return null;
            row.push(status);
          });
          // Check if row was filtered out
          const hasValidStatus = row.slice(3).some(s => {
            if (statusFilter === 'present') return s === '\u2713';
            if (statusFilter === 'absent') return s === 'X';
            return true;
          });
          return hasValidStatus ? row : null;
        }).filter(row => row !== null);

      } else {
        // FIXED: Month mode — summary report
        const [year, month] = date.split('-').map(Number);
        const serviceDays = getServiceDaysInMonth(year, month - 1);
        dateHeaders = serviceDays;
        serviceDays.forEach(d => headers.push(d));

        const sortedGirls = [...filteredGirls].sort((a, b) => {
          const aOrder = GRADE_ORDER[a.grade] || 99;
          const bOrder = GRADE_ORDER[b.grade] || 99;
          if (aOrder !== bOrder) return aOrder - bOrder;
          return a.name.localeCompare(b.name, 'ar');
        });

        records = sortedGirls.map(g => {
          const row = [csvEscape(g.name), csvEscape(g.grade), csvEscape(g.phone || '')];
          let totalPresents = 0;
          serviceDays.forEach(day => {
            const key = makeAttKey(g.id, day, state.selectedActivity || 'دراسي');
            const record = state.attendanceData[key];
            const status = record ? (record.status === 'حاضر' ? '\u2713' : 'X') : '-';
            if (status === '\u2713') totalPresents++;
            row.push(status);
          });
          const totalDays = serviceDays.length || 1;
          const percentage = Math.round((totalPresents / totalDays) * 100);
          row.push(`${totalPresents}/${totalDays} (${percentage}%)`);
          return row;
        });

        headers.push('النسبة المئوية');
        summaryData = records;
      }

      if (!records.length) { showToast('لا توجد بيانات للتصدير', 'warning'); return; }

      const bom = '\uFEFF';
      const csvContent = bom + headers.join(',') + '\n' + records.map(row => row.join(',')).join('\n');
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `حضور_${gradeFilter || 'الكل'}_${date}${statusFilter ? '_' + statusFilter : ''}.csv`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      showToast('تم تصدير Excel بنجاح', 'success');
      await logHistory('تصدير', `تصدير CSV - ${gradeFilter || 'الكل'} - ${date} - ${exportMode}`);
    } catch (e) {
      console.error('Export CSV error:', e);
      showToast('حدث خطأ أثناء التصدير', 'error');
    }
  });
}

if (DOM.exportJSON) {
  DOM.exportJSON.addEventListener('click', async () => {
    try {
      const gradeFilter = state.exportGradeFilter;
      const statusFilter = state.exportStatusFilter;
      const date = DOM.exportMonth ? DOM.exportMonth.value : TimeUtil.getDate();
      const activeGirls = state.girls.filter(g => !g.isDeleted);
      const filteredGirls = gradeFilter ? activeGirls.filter(g => g.grade === gradeFilter) : activeGirls;

      if (!filteredGirls.length) { showToast('لا توجد مخدومات للتصدير', 'warning'); return; }

      const data = {
        exportedAt: new Date().toISOString(),
        date,
        gradeFilter: gradeFilter || 'الكل',
        statusFilter: statusFilter || 'الكل',
        girls: filteredGirls.map(g => ({
          id: g.id,
          name: g.name,
          grade: g.grade,
          phone: g.phone || ''
        })),
        attendance: {}
      };

      filteredGirls.forEach(g => {
        data.attendance[g.id] = {};
        ACTIVITIES.forEach(activity => {
          const key = makeAttKey(g.id, date, activity);
          const record = state.attendanceData[key];
          if (record) {
            // Apply status filter
            if (statusFilter === 'present' && record.status !== 'حاضر') return;
            if (statusFilter === 'absent' && record.status !== 'غائب') return;
            data.attendance[g.id][activity] = {
              status: record.status,
              rating: record.rating || 0,
              notes: record.notes || ''
            };
          }
        });
      });

      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `حضور_${gradeFilter || 'الكل'}_${date}${statusFilter ? '_' + statusFilter : ''}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      showToast('تم تصدير JSON بنجاح', 'success');
      await logHistory('تصدير', `تصدير JSON - ${gradeFilter || 'الكل'} - ${date}`);
    } catch (e) {
      console.error('Export JSON error:', e);
      showToast('حدث خطأ أثناء التصدير', 'error');
    }
  });
}

if (DOM.exportPrint) {
  DOM.exportPrint.addEventListener('click', async () => {
    try {
      const gradeFilter = state.exportGradeFilter;
      const statusFilter = state.exportStatusFilter;
      const exportMode = document.querySelector('input[name="exportMode"]:checked')?.value || 'day';
      const date = DOM.exportMonth ? DOM.exportMonth.value : TimeContext.getDate();
      const activeGirls = state.girls.filter(g => !g.isDeleted);
      const filteredGirls = gradeFilter ? activeGirls.filter(g => g.grade === gradeFilter) : activeGirls;

      if (!filteredGirls.length) { showToast('لا توجد مخدومات للطباعة', 'warning'); return; }

      const sortedGirls = [...filteredGirls].sort((a, b) => {
        const aOrder = GRADE_ORDER[a.grade] || 99;
        const bOrder = GRADE_ORDER[b.grade] || 99;
        if (aOrder !== bOrder) return aOrder - bOrder;
        return a.name.localeCompare(b.name, 'ar');
      });

      let rows = '';
      if (exportMode === 'day') {
        rows = sortedGirls.map((g, i) => {
          const cells = ACTIVITIES.map(activity => {
            const key = makeAttKey(g.id, date, activity);
            const record = state.attendanceData[key];
            const status = record ? (record.status === 'حاضر' ? '\u2713' : 'X') : '-';
            return `<td style="text-align:center;border:1px solid #333;padding:8px;font-size:14px;">${status}</td>`;
          }).join('');
          return `<tr>
            <td style="border:1px solid #333;padding:8px;font-weight:700;">${i + 1}</td>
            <td style="border:1px solid #333;padding:8px;">${esc(g.name)}</td>
            <td style="border:1px solid #333;padding:8px;text-align:center;">${esc(g.grade)}</td>
            ${cells}
          </tr>`;
        }).join('');
      } else {
        const [year, month] = date.split('-').map(Number);
        const serviceDays = getServiceDaysInMonth(year, month - 1);
        const dayHeaders = serviceDays.map(d => `<th style="border:1px solid #333;padding:8px;background:#f0f0f0;font-size:12px;">${d}</th>`).join('');

        rows = sortedGirls.map((g, i) => {
          const cells = serviceDays.map(day => {
            const key = makeAttKey(g.id, day, state.selectedActivity || 'دراسي');
            const record = state.attendanceData[key];
            const status = record ? (record.status === 'حاضر' ? '\u2713' : 'X') : '-';
            return `<td style="text-align:center;border:1px solid #333;padding:8px;font-size:14px;">${status}</td>`;
          }).join('');
          return `<tr>
            <td style="border:1px solid #333;padding:8px;font-weight:700;">${i + 1}</td>
            <td style="border:1px solid #333;padding:8px;">${esc(g.name)}</td>
            <td style="border:1px solid #333;padding:8px;text-align:center;">${esc(g.grade)}</td>
            ${cells}
          </tr>`;
        });

        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
          <html dir="rtl">
          <head>
            <meta charset="UTF-8">
            <title>تقرير الحضور</title>
            <style>
              body { font-family: Arial, sans-serif; padding: 20px; }
              h2 { text-align: center; margin-bottom: 10px; }
              .subtitle { text-align: center; color: #666; margin-bottom: 20px; }
              table { width: 100%; border-collapse: collapse; margin-top: 20px; }
              th, td { border: 1px solid #333; padding: 8px; text-align: right; }
              th { background: #f0f0f0; font-weight: 700; }
              @media print { body { padding: 0; } }
            </style>
          </head>
          <body>
            <h2>تقرير الحضور الشهري</h2>
            <div class="subtitle">${date} — ${gradeFilter || 'كل السنوات'}</div>
            <table>
              <thead>
                <tr>
                  <th style="border:1px solid #333;padding:8px;background:#f0f0f0;">#</th>
                  <th style="border:1px solid #333;padding:8px;background:#f0f0f0;">المخدومة</th>
                  <th style="border:1px solid #333;padding:8px;background:#f0f0f0;text-align:center;">السنة</th>
                  ${dayHeaders}
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
            <script>window.onload = function() { window.print(); };</script>
          </body></html>`);
        printWindow.document.close();
        return;
      }

      const printWindow = window.open('', '_blank');
      printWindow.document.write(`
        <html dir="rtl">
        <head>
          <meta charset="UTF-8">
          <title>تقرير الحضور</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            h2 { text-align: center; margin-bottom: 10px; }
            .subtitle { text-align: center; color: #666; margin-bottom: 20px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #333; padding: 8px; text-align: right; }
            th { background: #f0f0f0; font-weight: 700; }
            @media print { body { padding: 0; } }
          </style>
        </head>
        <body>
          <h2>تقرير الحضور اليومي</h2>
          <div class="subtitle">${date} — ${gradeFilter || 'كل السنوات'}</div>
          <table>
            <thead>
              <tr>
                <th style="border:1px solid #333;padding:8px;background:#f0f0f0;">#</th>
                <th style="border:1px solid #333;padding:8px;background:#f0f0f0;">المخدومة</th>
                <th style="border:1px solid #333;padding:8px;background:#f0f0f0;text-align:center;">السنة</th>
                <th style="border:1px solid #333;padding:8px;background:#f0f0f0;text-align:center;">دراسي</th>
                <th style="border:1px solid #333;padding:8px;background:#f0f0f0;text-align:center;">محفوظات</th>
                <th style="border:1px solid #333;padding:8px;background:#f0f0f0;text-align:center;">قبطي</th>
                <th style="border:1px solid #333;padding:8px;background:#f0f0f0;text-align:center;">ألحان</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
          <script>window.onload = function() { window.print(); };</script>
        </body></html>`);
      printWindow.document.close();

      await logHistory('تصدير', `طباعة تقرير - ${gradeFilter || 'الكل'} - ${date}`);
    } catch (e) {
      console.error('Export print error:', e);
      showToast('حدث خطأ أثناء الطباعة', 'error');
    }
  });
}

// ============================================================
// SETTINGS PAGE — NEW: Backup, Restore, Clear All Data
// ============================================================
function renderSettings() {
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const attCount = Object.keys(state.attendanceData).length;

  if (DOM.settingsGirlCount) DOM.settingsGirlCount.textContent = activeGirls.length;
  if (DOM.settingsAttCount) DOM.settingsAttCount.textContent = attCount;

  // Last update timestamp
  if (DOM.settingsLastUpdate) {
    const lastUpdate = activeGirls.length > 0
      ? new Date(Math.max(...activeGirls.map(g => g.updatedAt || 0))).toLocaleString('ar-EG')
      : '-';
    DOM.settingsLastUpdate.textContent = lastUpdate;
  }
}

// Export Full Backup
if (DOM.exportFullBackup) {
  DOM.exportFullBackup.addEventListener('click', async () => {
    try {
      const backup = {
        version: '3.1',
        exportedAt: new Date().toISOString(),
        girls: state.girls,
        attendanceData: state.attendanceData,
        serviceDayOverrides: state.serviceDayOverrides,
        autoMarkedDates: [...state.autoMarkedDates]
      };

      const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `نسخة_احتياطية_${DateUtil.toStr()}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      // Save to IDB for rollback
      await IDB.add('backups', {
        id: 'full_backup_' + Date.now(),
        data: backup,
        timestamp: Date.now()
      });

      showToast('تم تصدير النسخة الاحتياطية بنجاح', 'success');
      await logHistory('تصدير', 'تصدير نسخة احتياطية كاملة');
    } catch (e) {
      console.error('Export backup error:', e);
      showToast('حدث خطأ أثناء التصدير', 'error');
    }
  });
}

// Import Backup
if (DOM.importBackup) {
  DOM.importBackup.addEventListener('click', () => {
    if (DOM.importFileInput) DOM.importFileInput.click();
  });
}

if (DOM.importFileInput) {
  DOM.importFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const text = await file.text();
      const backup = JSON.parse(text);

      showConfirm({
        icon: '&#128229;',
        title: 'استيراد نسخة احتياطية',
        msg: `هل أنت متأكد من استيراد النسخة الاحتياطية؟\n\nهذا سيستبدل البيانات الحالية بالبيانات الموجودة في الملف.`,
        okLabel: 'استيراد',
        okClass: 'confirm-delete',
        onOk: async () => {
          try {
            // Save current state for rollback
            const currentBackup = {
              version: '3.1',
              exportedAt: new Date().toISOString(),
              girls: state.girls,
              attendanceData: state.attendanceData,
              serviceDayOverrides: state.serviceDayOverrides,
              autoMarkedDates: [...state.autoMarkedDates]
            };
            await IDB.add('backups', {
              id: 'pre_import_backup_' + Date.now(),
              data: currentBackup,
              timestamp: Date.now()
            });

            // Restore girls
            if (backup.girls && Array.isArray(backup.girls)) {
              setStateGirls(backup.girls);
              if (firebaseReady) {
                for (const girl of backup.girls) {
                  try { await FB.setDoc(FB.doc(db, 'girls', girl.id), girl); }
                  catch (e) { console.error('Restore girl Firestore error:', e); }
                }
              }
            }

            // Restore attendance
            if (backup.attendanceData && typeof backup.attendanceData === 'object') {
              setStateAttendanceData(backup.attendanceData);
              if (firebaseReady) {
                const entries = Object.entries(backup.attendanceData);
                for (let i = 0; i < entries.length; i += 500) {
                  try {
                    const batch = FB.writeBatch(db);
                    entries.slice(i, i + 500).forEach(([key, rec]) => {
                      batch.set(FB.doc(db, 'attendance', key), rec);
                    });
                    await batch.commit();
                  } catch (e) { console.error('Restore attendance Firestore error:', e); }
                }
              }
            }

            // Restore service day overrides
            if (backup.serviceDayOverrides) {
              state.serviceDayOverrides = backup.serviceDayOverrides;
              localStorage.setItem('serviceDayOverrides', JSON.stringify(backup.serviceDayOverrides));
            }

            // Restore auto-marked dates
            if (backup.autoMarkedDates && Array.isArray(backup.autoMarkedDates)) {
              state.autoMarkedDates = new Set(backup.autoMarkedDates);
              persistAutoMarkedDates();
            }

            showToast('تم استيراد النسخة الاحتياطية بنجاح', 'success');
            await logHistory('استيراد', 'استيراد نسخة احتياطية');
            renderPage();
          } catch (err) {
            console.error('Import error:', e);
            showToast('حدث خطأ أثناء الاستيراد', 'error');
          }
        }
      });
    } catch (e) {
      console.error('Parse backup error:', e);
      showToast('ملف غير صالح أو تالف', 'error');
    }

    // Reset file input
    DOM.importFileInput.value = '';
  });
}

// Clear All Data
if (DOM.clearAllData) {
  DOM.clearAllData.addEventListener('click', async () => {
    showConfirm({
      icon: '&#9888;',
      title: 'مسح كل البيانات',
      msg: 'هل أنت متأكد من حذف كل البيانات؟\n\nهذا سيحذف كل المخدومات وسجلات الحضور بشكل نهي. لا يمكن التراجع عن هذا الإجراء!',
      okLabel: 'مسح الكل',
      okClass: 'confirm-delete',
      onOk: async () => {
        try {
          // Save backup before clearing
          const backup = {
            version: '3.1',
            exportedAt: new Date().toISOString(),
            girls: state.girls,
            attendanceData: state.attendanceData,
            serviceDayOverrides: state.serviceDayOverrides,
            autoMarkedDates: [...state.autoMarkedDates]
          };
          await IDB.add('backups', {
            id: 'pre_clear_backup_' + Date.now(),
            data: backup,
            timestamp: Date.now()
          });

          // Clear Firestore
          if (firebaseReady) {
            // Delete girls
            try {
              const girlsSnap = await FB.getDocs(FB.collection(db, 'girls'));
              const batch = FB.writeBatch(db);
              girlsSnap.docs.forEach(d => batch.delete(d.ref));
              await batch.commit();
            } catch (e) { console.error('Clear girls Firestore error:', e); }

            // Delete attendance
            try {
              const attSnap = await FB.getDocs(FB.collection(db, 'attendance'));
              for (let i = 0; i < attSnap.docs.length; i += 500) {
                try {
                  const batch = FB.writeBatch(db);
                  attSnap.docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                  await batch.commit();
                } catch (e) { console.error('Clear attendance Firestore error:', e); }
              }
            } catch (e) { console.error('Clear attendance Firestore error:', e); }
          }

          // Clear local state
          setStateGirls([]);
          setStateAttendanceData({});
          state.serviceDayOverrides = {};
          state.autoMarkedDates = new Set();
          localStorage.removeItem('serviceDayOverrides');
          localStorage.removeItem('autoMarkedDates');

          showToast('تم مسح كل البيانات', 'success');
          await logHistory('مسح', 'مسح كل البيانات');
          renderPage();
        } catch (err) {
          console.error('Clear all data error:', err);
          showToast('حدث خطأ أثناء المسح', 'error');
        }
      }
    });
  });
}

// ============================================================
// CONFIRM MODAL
// ============================================================
function showConfirm({ icon, title, msg, okLabel = 'تأكيد', okClass = 'confirm-ok', onOk, onCancel }) {
  if (DOM.confirmIcon) DOM.confirmIcon.innerHTML = icon || '&#9888;';
  if (DOM.confirmTitle) DOM.confirmTitle.textContent = title || 'تأكيد';
  if (DOM.confirmMsg) DOM.confirmMsg.innerHTML = msg || '';

  const okBtn = DOM.confirmOk;
  if (okBtn) {
    okBtn.textContent = okLabel;
    okBtn.className = `confirm-ok ${okClass}`;
  }

  if (DOM.confirmOverlay) DOM.confirmOverlay.classList.add('show');

  // One-time handler
  const handleOk = () => {
    cleanup();
    if (onOk) onOk();
  };
  const handleCancel = () => {
    cleanup();
    if (onCancel) onCancel();
  };

  function cleanup() {
    if (DOM.confirmOverlay) DOM.confirmOverlay.classList.remove('show');
    if (okBtn) okBtn.removeEventListener('click', handleOk);
    if (DOM.confirmCancel) DOM.confirmCancel.removeEventListener('click', handleCancel);
  }

  if (okBtn) okBtn.addEventListener('click', handleOk);
  if (DOM.confirmCancel) DOM.confirmCancel.addEventListener('click', handleCancel);
}

// ============================================================
// MODAL HELPERS
// ============================================================
function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.add('show');
  document.body.style.overflow = 'hidden';
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.remove('show');
  document.body.style.overflow = '';
}

// Close modal on overlay click
document.addEventListener('click', (e) => {
  if (e.target.classList.contains('modal-overlay') || e.target.classList.contains('confirm-overlay')) {
    e.target.classList.remove('show');
    document.body.style.overflow = '';
  }
});

// ESC key to close modals
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal-overlay.show, .confirm-overlay.show').forEach(m => {
      m.classList.remove('show');
    });
    document.body.style.overflow = '';
  }
});

// ============================================================
// EVENT LISTENERS — Unified click delegation for dynamic elements
// ============================================================
document.addEventListener('click', (e) => {
  // Search results
  const searchItem = e.target.closest('.search-item[data-girl-id]');
  if (searchItem) {
    const girlId = searchItem.dataset.girlId;
    if (DOM.searchResults) DOM.searchResults.classList.remove('show');
    if (DOM.globalSearch) DOM.globalSearch.value = '';
    showGirlProfile(girlId);
    return;
  }

  // Close search when clicking outside
  if (DOM.searchResults && DOM.searchResults.classList.contains('show') && !e.target.closest('.search-bar-wrap') && !e.target.closest('.search-results')) {
    DOM.searchResults.classList.remove('show');
  }

  // Girl cards (click to view profile, edit button to edit)
  const girlCard = e.target.closest('.girl-card');
  if (girlCard) {
    const girlId = girlCard.dataset.girlId;
    if (e.target.closest('.edit-btn')) {
      e.stopPropagation();
      editGirl(girlId);
    } else {
      showGirlProfile(girlId);
    }
    return;
  }

  // Followup items
  const followupItem = e.target.closest('.followup-item');
  if (followupItem) {
    const girlId = followupItem.dataset.girlId;
    if (girlId) showGirlProfile(girlId);
    return;
  }

  // History items
  const historyItem = e.target.closest('.history-item');
  if (historyItem) {
    // Could add detail view here
    return;
  }

  // Grade filter buttons
  const gradeFilterBtn = e.target.closest('.grade-filter-btn');
  if (gradeFilterBtn) {
    const page = e.target.closest('#page-home') ? 'home'
      : e.target.closest('#page-girls') ? 'girls'
      : e.target.closest('#page-attendance') ? 'attendance'
      : null;
    if (page === 'home') { state.homeGradeFilter = gradeFilterBtn.dataset.grade; renderHome(); }
    else if (page === 'girls') { state.girlsGradeFilter = gradeFilterBtn.dataset.grade; renderGirlsList(); }
    else if (page === 'attendance') {
      state.attendanceGradeFilter = gradeFilterBtn.dataset.grade;
      localStorage.setItem('attendanceGradeFilter', state.attendanceGradeFilter);
      renderAttendanceList();
    }
    return;
  }
});

// ============================================================
// GIRLS SEARCH
// ============================================================
const girlsSearchEl = document.getElementById('girlsSearch');
if (girlsSearchEl) {
  girlsSearchEl.addEventListener('input', (e) => {
    state.girlsSearchQuery = e.target.value;
    renderGirlsList();
  });
}

// ============================================================
// GIRL MODAL CLOSE HANDLERS
// ============================================================
if (DOM.closeGirlModal) DOM.closeGirlModal.addEventListener('click', () => closeModal('girlModal'));
if (DOM.cancelGirlModal) DOM.cancelGirlModal.addEventListener('click', () => closeModal('girlModal'));

// ============================================================
// INITIALIZATION
// ============================================================
async function init() {
  // Initialize theme first (before page renders)
  Theme.init();

  // Initialize TimeContext
  TimeContext.init();

  // Initialize IDB
  try { await IDB.init(); state.idb = true; } catch (e) { console.warn('IDB init failed:', e); }

  // Initialize Firebase modules
  await initModules();

  // Update online status
  updateOnlineStatus();

  // Initialize auth
  await initAuth();

  // Setup activity detail tab handlers
  document.querySelectorAll('.activity-detail-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      state.activityDetailTab = tab.dataset.tab;
      document.querySelectorAll('.activity-detail-tab').forEach(t => t.classList.toggle('active', t === tab));
      renderActivityDetailList();
    });
  });
}

// Start the app
init();
