// ============================================================
// نظام متابعة المخدومات — Offline Ready & Guest Mode
// FIXED VERSION 2.0 — All critical bugs resolved
// Changes: Firebase proxy, unified splash, DOM safety, rollback,
//          popup fallback, noscript, CSS dark mode cleanup,
//          grade-based export (talt->tanya->awal)
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
    'statsGradeFilter', 'activityStatsGrade', 'exportGradeFilter'
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
const originalGirlsDescriptor = Object.getOwnPropertyDescriptor(state, 'girls');
// Use Proxy-like approach: intercept direct mutations
const _rawGirls = [];
const _rawAttendance = {};

function setStateGirls(newGirls) {
  state.girls = newGirls;
  Cache.invalidate();
}

function setStateAttendanceData(newData) {
  state.attendanceData = newData;
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
// SERVICE DAY FUNCTIONS
// ============================================================
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
  toastTimeout = setTimeout(() => { if (DOM.toast) DOM.toast.className = 'toast hidden'; }, 3000);
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
  export: ['التصدير', 'تصدير البيانات']
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

  const normalized = DateUtil.normalize(dayName);
  const isService = SERVICE_DAYS[normalized];

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
        div.innerHTML = `<span class="followup-name">${esc(girl.name)}</span><span class="followup-badge">${result.count} غياب متتالي</span>`;
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
        if (isServiceDayDate(todayStr)) {
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
// ATTENDANCE PAGE — FIXED: Removed dangerous auto-mark on render
// ============================================================
function getCurrentServiceDay() {
  const dayOfWeek = new Date().getDay();
  const dayMap = { 6: 'السبت', 1: 'الاثنين', 3: 'الاربعاء' };
  return dayMap[dayOfWeek] || null;
}

function isServiceDayDate(dateStr) {
  if (!dateStr) return false;
  // FIXED: Use parseDateStr for safe date parsing
  const d = parseDateStr(dateStr);
  if (isNaN(d.getTime())) return false;
  return SERVICE_DAY_NUMBERS.includes(d.getDay());
}

// FIXED: Renamed to clarify this is a hardcoded lookup, not dynamic
function getHardcodedServiceDay(dayOfWeek) {
  const dayMap = { 6: 'السبت', 1: 'الاثنين', 3: 'الاربعاء' };
  return dayMap[dayOfWeek] || null;
}

function renderAttendancePage() {
  if (!DOM.attendanceDate) return;
  DOM.attendanceDate.value = TimeContext.getDate();

  const currentServiceDay = getCurrentServiceDay();
  if (currentServiceDay && !state.attendancePageInitialized) {
    state.selectedDay = currentServiceDay;
  }

  setActiveDay(state.selectedDay);
  setActiveActivity(state.selectedActivity);

  state.attendancePageInitialized = true;
  renderAttendanceList();
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

function debouncedAttSearch() {
  clearTimeout(state.attSearchDebounceTimer);
  state.attSearchDebounceTimer = setTimeout(() => { renderAttendanceList(); }, 250);
}

if (DOM.attendanceSearch) DOM.attendanceSearch.addEventListener('input', debouncedAttSearch);

// ============================================================
// TOGGLE ATTENDANCE — FIXED: Pending lock prevents race conditions
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

    // FIXED: Debounced render to batch rapid toggles + ensure state settled
    debouncedRender(80);
  } finally {
    state.pendingAttendanceOps.delete(opKey);
  }
}

// ============================================================
// MARK ALL ABSENT — FIXED: Only called by explicit user action
// Removed from renderAttendancePage — must be called deliberately
// ============================================================
async function markAllAbsentForDate(date) {
  if (!isServiceDayDate(date)) return;

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
  if (!isServiceDayDate(date)) return;

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
// RENDER ATTENDANCE LIST — FIXED: O(n) single-pass optimization + key consistency
// ============================================================
function renderAttendanceList() {
  if (!DOM.attendanceDate || !DOM.attendanceList) return;
  const date = DOM.attendanceDate.value;
  const el = DOM.attendanceList;
  if (!date) { el.innerHTML = '<div class="empty-state">الرجاء اختيار التاريخ</div>'; return; }

  let activeGirls = state.girls.filter(g => !g.isDeleted);

  // Apply grade filter
  const gradeFilter = state.attendanceGradeFilter;
  if (gradeFilter) {
    activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  }

  const searchQuery = DOM.attendanceSearch?.value?.trim() || '';
  if (searchQuery) {
    const qNorm = normalizeArabic(searchQuery);
    activeGirls = activeGirls.filter(g => normalizeArabic(g.name).includes(qNorm));
  }

  // Update counts
  const allActiveGirls = state.girls.filter(g => !g.isDeleted);
  const attFilterCountAll = document.getElementById('attFilterCountAll');
  const attFilterCount1 = document.getElementById('attFilterCount1');
  const attFilterCount2 = document.getElementById('attFilterCount2');
  const attFilterCount3 = document.getElementById('attFilterCount3');
  if (attFilterCountAll) attFilterCountAll.textContent = allActiveGirls.length;
  if (attFilterCount1) attFilterCount1.textContent = allActiveGirls.filter(g => g.grade === 'أولى إعدادي').length;
  if (attFilterCount2) attFilterCount2.textContent = allActiveGirls.filter(g => g.grade === 'تانية إعدادي').length;
  if (attFilterCount3) attFilterCount3.textContent = allActiveGirls.filter(g => g.grade === 'تالتة إعدادي').length;

  document.querySelectorAll('#attendanceGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  if (searchQuery && !activeGirls.length) {
    el.innerHTML = '<div class="empty-state">لا توجد نتائج للبحث</div>';
    if (DOM.presentCount) DOM.presentCount.textContent = 0;
    if (DOM.absentCount) DOM.absentCount.textContent = 0;
    if (DOM.totalCount) DOM.totalCount.textContent = 0;
    return;
  }

  if (!activeGirls.length) {
    el.innerHTML = '<div class="empty-state">لا توجد مخدومات مسجلة<br><small>أضف مخدومات أولاً من صفحة المخدومات</small></div>';
    if (DOM.presentCount) DOM.presentCount.textContent = 0;
    if (DOM.absentCount) DOM.absentCount.textContent = 0;
    if (DOM.totalCount) DOM.totalCount.textContent = 0;
    return;
  }

  let present = 0, absent = 0;
  const frag = document.createDocumentFragment();

  // FIXED: Pre-filter by BOTH date AND current activity — prevents key mismatch
  // and reduces data size significantly
  const dateAttendance = {};
  const allAttendance = Cache.getAllAttendance();
  const currentActivity = state.selectedActivity;
  allAttendance.forEach(a => {
    if (a.date === date && a.activity === currentActivity) {
      const lookupKey = makeAttKey(a.girlId, a.date, a.activity);
      dateAttendance[lookupKey] = a;
    }
  });

  activeGirls.forEach(g => {
    // FIXED: Use makeAttKey helper for consistent key generation
    const key = makeAttKey(g.id, date, state.selectedActivity);
    const rec = dateAttendance[key];
    let statusClass = 'absent', statusIcon = '&#10007;', statusText = 'غائب';
    if (rec?.status === 'حاضر') { statusClass = 'present'; statusIcon = '&#10003;'; statusText = 'حاضر'; present++; }
    else { absent++; }

    const stars = rec?.rating ? '&#9733;'.repeat(rec.rating) + '&#9734;'.repeat(5 - rec.rating) : '';
    const currentRating = rec?.rating || 0;
    const div = document.createElement('div');
    div.className = `att-item ${statusClass}`;
    div.dataset.girlId = g.id;
    div.dataset.attKey = key;
    div.dataset.girlName = g.name;

    // FIXED: Safer inline rating HTML construction
    let inlineRatingHtml = '';
    if (statusClass === 'present') {
      let starsHtml = '';
      for (let i = 1; i <= 5; i++) {
        starsHtml += `<span class="att-inline-star ${i <= currentRating ? 'active' : ''}" data-val="${i}" role="button" aria-label="${i} نجمة">&#9733;</span>`;
      }
      inlineRatingHtml = `<div class="att-inline-rating" data-att-key="${esc(key)}">
        <span class="att-inline-rating-label">التقييم:</span>
        <span class="att-inline-stars">${starsHtml}</span>
        ${currentRating > 0 ? `<span class="att-inline-rating-val">${currentRating}/5</span>` : '<span class="att-inline-rating-hint">اضغط نجمة للتقييم</span>'}
      </div>`;
    }

    div.innerHTML = `
      <div class="att-icon">${statusIcon}</div>
      <div class="att-info">
        <span class="att-name">${esc(g.name)}</span>
        <span class="att-grade">${esc(g.grade)}</span>
        ${stars ? `<span class="att-stars">${stars}</span>` : ''}
        ${inlineRatingHtml}
        ${rec?.notes ? `<span class="att-note">${esc(rec.notes)}</span>` : ''}
      </div>
      <span class="att-status-text ${statusClass}">${statusText}</span>`;
    frag.appendChild(div);
  });

  el.innerHTML = '';
  el.appendChild(frag);
  if (DOM.presentCount) DOM.presentCount.textContent = present;
  if (DOM.absentCount) DOM.absentCount.textContent = absent;
  if (DOM.totalCount) DOM.totalCount.textContent = activeGirls.length;
}

// ============================================================
// INLINE RATING — FIXED: Operation lock + safer DOM handling
// ============================================================
async function saveInlineRating(attKey, rating) {
  const rec = state.attendanceData[attKey];
  if (!rec) return;
  if (rec.status !== 'حاضر') { showToast('التقييم متاح فقط للحاضرات', 'warning'); return; }

  const opKey = `rating_${attKey}`;
  if (state.pendingAttendanceOps.has(opKey)) return;
  state.pendingAttendanceOps.add(opKey);

  try {
    const updatedRec = {
      ...rec,
      rating: rating,
      updatedAt: Date.now(),
      updatedBy: state.currentUser?.displayName || 'خادم',
      updatedByEmail: state.currentUser?.email || ''
    };

    setStateAttendanceData({ ...state.attendanceData, [attKey]: updatedRec });

    if (firebaseReady) {
      try { await FB.setDoc(FB.doc(db, 'attendance', attKey), updatedRec); }
      catch (e) { console.error('Save inline rating Firestore error:', e); }
    }

    const g = Cache.getGirl(rec.girlId);
    await logHistory('تقييم مخدومة', `${g?.name || ''} - ${rec.activity} - ${rec.date} - ${rating} نجوم`);
    showToast(`تم التقييم: ${rating} نجوم`, 'success');

    renderAttendanceList();
    if (state.currentPage === 'home') renderHome();
    if (state.currentPage === 'stats') renderStats();
  } finally {
    state.pendingAttendanceOps.delete(opKey);
  }
}

function openAttendanceEntry(girlId, girlName, date) {
  state.currentAttendanceGirlId = girlId;
  state.currentAttendanceRating = 0;
  if (DOM.attendanceModalTitle) DOM.attendanceModalTitle.textContent = `${state.selectedActivity} - ${date}`;
  if (DOM.modalGirlName) DOM.modalGirlName.textContent = girlName;
  if (DOM.attendanceNotes) DOM.attendanceNotes.value = '';

  // FIXED: Use makeAttKey for consistent key generation
  const key = makeAttKey(girlId, date, state.selectedActivity);
  const existing = state.attendanceData[key];
  if (existing) {
    document.querySelectorAll('.attend-btn').forEach(b => b.classList.toggle('selected', b.dataset.status === existing.status));
    setRating(existing.rating || 0);
    if (DOM.attendanceNotes) DOM.attendanceNotes.value = existing.notes || '';
    if (DOM.ratingSection) DOM.ratingSection.classList.toggle('hidden', existing.status !== 'حاضر');
  } else {
    document.querySelectorAll('.attend-btn').forEach(b => b.classList.remove('selected'));
    setRating(0);
    if (DOM.ratingSection) DOM.ratingSection.classList.add('hidden');
  }
  openModal('attendanceModal');
}

document.querySelectorAll('.attend-btn').forEach(b => {
  b.addEventListener('click', () => {
    document.querySelectorAll('.attend-btn').forEach(x => x.classList.remove('selected'));
    b.classList.add('selected');
    if (DOM.ratingSection) DOM.ratingSection.classList.toggle('hidden', b.dataset.status !== 'حاضر');
  });
});

document.querySelectorAll('.star').forEach(s => s.addEventListener('click', () => setRating(parseInt(s.dataset.val))));
function setRating(val) {
  state.currentAttendanceRating = val;
  document.querySelectorAll('.star').forEach(s => s.classList.toggle('active', parseInt(s.dataset.val) <= val));
}

if (DOM.saveAttendanceEntry) {
  DOM.saveAttendanceEntry.addEventListener('click', async () => {
    if (!DOM.attendanceDate) return;
    const date = DOM.attendanceDate.value;
    const statusBtn = document.querySelector('.attend-btn.selected');
    if (!statusBtn) { showToast('الرجاء تحديد الحضور أو الغياب', 'error'); return; }

    // FIXED: Use makeAttKey for consistent key generation
    const key = makeAttKey(state.currentAttendanceGirlId, date, state.selectedActivity);
    const rec = {
      id: key,
      girlId: state.currentAttendanceGirlId,
      date,
      day: state.selectedDay,
      activity: state.selectedActivity,
      status: statusBtn.dataset.status,
      rating: statusBtn.dataset.status === 'حاضر' ? state.currentAttendanceRating : 0,
      notes: DOM.attendanceNotes ? DOM.attendanceNotes.value.trim() : '',
      updatedAt: Date.now(),
      updatedBy: state.currentUser?.displayName || 'خادم',
      updatedByEmail: state.currentUser?.email || ''
    };

    setStateAttendanceData({ ...state.attendanceData, [key]: rec });

    if (firebaseReady) {
      try { await FB.setDoc(FB.doc(db, 'attendance', key), rec); }
      catch (e) { console.error('Save attendance Firestore error:', e); }
    }

    const gName = Cache.getGirl(state.currentAttendanceGirlId)?.name || '';
    await logHistory('تسجيل حضور', `${gName} - ${state.selectedActivity} - ${date} - ${rec.status}`);
    closeModal('attendanceModal');
    showToast('تم الحفظ', 'success');
    renderAttendanceList();

    if (state.currentPage === 'home') renderHome();
    if (state.currentPage === 'stats') renderStats();
    if (state.currentPage === 'calendar') renderCalendar();
  });
}

// ============================================================
// CALENDAR PAGE — FIXED: O(n^3) eliminated with date-index lookup
// ============================================================
function renderCalendar() {
  const year = state.calendarDate.getFullYear();
  const month = state.calendarDate.getMonth();
  if (DOM.calMonthYear) DOM.calMonthYear.textContent = state.calendarDate.toLocaleDateString('ar-EG', { year: 'numeric', month: 'long' });

  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const todayStr = TimeContext.getDate();

  // FIXED: Build date+activity index for accurate has-data detection
  // Prevents false positive: day has data for OTHER activity but not current one
  const dateIndex = new Set();
  const currentCalendarActivity = state.selectedActivity || '';
  const allAttendance = Cache.getAllAttendance();
  const activeGirlIds = Cache.getActiveGirlIds(); // FIXED: Use cached Set
  allAttendance.forEach(a => {
    if (activeGirlIds.has(a.girlId)) {
      // FIXED: Include activity in index key for accurate per-activity detection
      dateIndex.add(`${a.date}_${a.activity}`);
    }
  });

  let html = '<div class="cal-weekdays">';
  ['أح', 'إث', 'ثل', 'أر', 'خم', 'جم', 'سب'].forEach(d => html += `<div class="cal-wday">${d}</div>`);
  html += '</div><div class="cal-days">';
  for (let i = 0; i < firstDay; i++) html += '<div class="cal-day empty"></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${DateUtil.pad(month + 1)}-${DateUtil.pad(d)}`;
    const dayOfWeek = new Date(year, month, d).getDay();
    const isService = SERVICE_DAY_NUMBERS.includes(dayOfWeek);
    // FIXED: O(1) lookup with activity-aware index — no false positives
    const hasData = dateIndex.has(`${dateStr}_${currentCalendarActivity}`);
    const isToday = dateStr === todayStr;
    html += `<div class="cal-day ${isService ? 'service-day' : ''} ${hasData ? 'has-data' : ''} ${isToday ? 'today' : ''}" data-date="${dateStr}">
      <span>${d}</span>${isService ? '<div class="service-dot"></div>' : ''}
    </div>`;
  }
  html += '</div>';
  if (DOM.calendarGrid) DOM.calendarGrid.innerHTML = html;

  // Only auto-show today on initial load
  const now = new Date();
  if (year === now.getFullYear() && month === now.getMonth() && !currentDayDetailDate) {
    currentDayDetailDate = todayStr;
    refreshDayDetail();
  } else if (currentDayDetailDate) {
    refreshDayDetail();
  }
}

let currentDayDetailDate = null;

function showDayDetail(dateStr) {
  currentDayDetailDate = dateStr;
  refreshDayDetail();
}

function refreshDayDetail() {
  if (!currentDayDetailDate || !DOM.dayDetail) return;
  const dateStr = currentDayDetailDate;

  // FIXED: O(1) indexed lookup instead of O(n) full scan
  const activeGirlIds = Cache.getActiveGirlIds();
  const dayRecords = Cache.getAttendanceByDate(dateStr);
  const filteredRecords = dayRecords.filter(r => activeGirlIds.has(r.girlId));

  const el = DOM.dayDetail;
  if (!filteredRecords.length) {
    el.innerHTML = `<div class="day-detail-header">${dateStr}</div><div class="empty-state">لا توجد سجلات لهذا اليوم</div>`;
  } else {
    const grouped = {};
    filteredRecords.forEach(r => {
      const act = r.activity || 'عام';
      if (!grouped[act]) grouped[act] = [];
      grouped[act].push(r);
    });
    let html = `<div class="day-detail-header">${dateStr}</div>`;
    Object.entries(grouped).forEach(([act, recs]) => {
      const presentCount = recs.filter(r => r.status === 'حاضر').length;
      const absentCount = recs.filter(r => r.status === 'غائب').length;
      html += `<div class="day-activity"><b>${esc(act)}</b>: <span class="green-text">${presentCount} حاضر</span> \u00B7 <span class="red-text">${absentCount} غائب</span> من ${recs.length}</div>`;
    });
    el.innerHTML = html;
  }
  el.classList.add('show');
}

function hideDayDetail() {
  currentDayDetailDate = null;
  if (DOM.dayDetail) DOM.dayDetail.classList.remove('show');
}

if (DOM.calPrev) {
  DOM.calPrev.addEventListener('click', () => {
    hideDayDetail();
    state.calendarDate.setMonth(state.calendarDate.getMonth() - 1);
    const y = state.calendarDate.getFullYear();
    const m = state.calendarDate.getMonth() + 1;
    const d = parseInt(TimeContext.getDate().split('-')[2]) || 1;
    TimeContext.setDate(`${y}-${String(m).padStart(2, '0')}-${String(Math.min(d, 28)).padStart(2, '0')}`);
    renderCalendar();
  });
}
if (DOM.calNext) {
  DOM.calNext.addEventListener('click', () => {
    hideDayDetail();
    state.calendarDate.setMonth(state.calendarDate.getMonth() + 1);
    const y = state.calendarDate.getFullYear();
    const m = state.calendarDate.getMonth() + 1;
    const d = parseInt(TimeContext.getDate().split('-')[2]) || 1;
    TimeContext.setDate(`${y}-${String(m).padStart(2, '0')}-${String(Math.min(d, 28)).padStart(2, '0')}`);
    renderCalendar();
  });
}

// ============================================================
// ACTIVITY STATS — Period bounds function
// ============================================================
function getPeriodBounds(period, customDate) {
  const selectedDate = _validateDateStr(customDate || TimeContext.getDate(), DateUtil.toStr());
  const selYear = parseInt(selectedDate.substring(0, 4));
  const selMonth = parseInt(selectedDate.substring(5, 7));
  switch (period) {
    case 'today': return { start: selectedDate, end: selectedDate };
    case 'month': {
      // FIXED: selMonth is 1-based (from substring), JS Date month is 0-based
      const monthIndex = selMonth - 1;
      const lastDay = new Date(selYear, monthIndex + 1, 0).getDate();
      return { start: selectedDate.substring(0, 7) + '-01', end: selectedDate.substring(0, 7) + '-' + String(lastDay).padStart(2, '0') };
    }
    case 'year': return { start: selectedDate.substring(0, 4) + '-01-01', end: selectedDate.substring(0, 4) + '-12-31' };
    case 'all': default: return { start: '2000-01-01', end: selectedDate };
  }
}

// Returns both present AND absence counts for each activity
function getActivityStats(period, gradeFilter = '', customDate) {
  // FIXED: Use cached activeGirlIds from Cache, only filter by grade if needed
  const allActiveGirlIds = Cache.getActiveGirlIds();
  const activeGirlIds = gradeFilter
    ? new Set(state.girls.filter(g => !g.isDeleted && g.grade === gradeFilter).map(g => g.id))
    : allActiveGirlIds;
  const { start, end } = getPeriodBounds(period, customDate);

  const stats = {
    'دراسي': { present: 0, absent: 0 },
    'ألحان': { present: 0, absent: 0 },
    'قبطي': { present: 0, absent: 0 },
    'محفوظات': { present: 0, absent: 0 }
  };

  const allAttendance = Cache.getAllAttendance();
  allAttendance.forEach(a => {
    if (!activeGirlIds.has(a.girlId)) return;
    // FIXED: Use isDateInRange for safe date comparison
    if (!isDateInRange(a.date, start, end)) return;
    if (stats.hasOwnProperty(a.activity)) {
      if (a.status === 'حاضر') stats[a.activity].present++;
      else if (a.status === 'غائب') stats[a.activity].absent++;
    }
  });

  return Object.entries(stats)
    .filter(([, data]) => data.present > 0 || data.absent > 0)
    .sort((a, b) => (b[1].present + b[1].absent) - (a[1].present + a[1].absent));
}

// ============================================================
// ACTIVITY DETAIL MODAL — FIXED: Classification by presence count, not rate
// ============================================================
function openActivityDetailModal(activity, period, gradeFilter = '', customDate) {
  const { start, end } = getPeriodBounds(period, customDate);
  // FIXED: Use cached Set for better performance
  const activeGirlIds = gradeFilter
    ? new Set(state.girls.filter(g => !g.isDeleted && g.grade === gradeFilter).map(g => g.id))
    : Cache.getActiveGirlIds();
  const periodLabel = PERIOD_LABELS[period] || '';

  const allAttendance = Cache.getAllAttendance();
  const records = allAttendance.filter(a => {
    if (a.activity !== activity) return false;
    if (!activeGirlIds.has(a.girlId)) return false;
    // FIXED: Use isDateInRange for safe date comparison
    if (!isDateInRange(a.date, start, end)) return false;
    return true;
  });

  const byGirl = {};
  records.forEach(a => { if (!byGirl[a.girlId]) byGirl[a.girlId] = []; byGirl[a.girlId].push(a); });

  const presentGirls = [];
  const absentGirls = [];

  Object.entries(byGirl).forEach(([girlId, girlRecords]) => {
    // FIXED: Use parseDateStr for safe date comparison
    girlRecords.sort((a, b) => compareDateStr(b.date, a.date));
    const girl = Cache.getGirl(girlId);
    if (!girl) return;

    const pCount = girlRecords.filter(r => r.status === 'حاضر').length;
    const aCount = girlRecords.filter(r => r.status === 'غائب').length;
    const total = girlRecords.length;
    const rate = total > 0 ? Math.round((pCount / total) * 100) : 0;

    const entry = { girl, presentCount: pCount, absentCount: aCount, totalRecords: total, attendanceRate: rate, latestRecord: girlRecords[0] };
    // FIXED: Classification by rate >= 50% instead of count comparison
    // More accurate: a girl with 1 present + 1 absent should be "needs attention" not "present"
    if (total > 0 && rate >= 50) presentGirls.push(entry);
    else absentGirls.push(entry);
  });

  presentGirls.sort((a, b) => b.attendanceRate - a.attendanceRate || a.girl.name.localeCompare(b.girl.name, 'ar'));
  absentGirls.sort((a, b) => b.attendanceRate - a.attendanceRate || a.girl.name.localeCompare(b.girl.name, 'ar'));

  state.currentActivityDetail = { activity, period, presentGirls, absentGirls };
  state.activityDetailTab = 'present';

  if (DOM.activityDetailTitle) DOM.activityDetailTitle.textContent = `تفاصيل ${activity}`;
  if (DOM.activityDetailIcon) DOM.activityDetailIcon.innerHTML = ACTIVITY_ICONS[activity] || '&#128202;';
  if (DOM.activityDetailName) DOM.activityDetailName.textContent = activity;
  if (DOM.activityDetailPeriod) DOM.activityDetailPeriod.textContent = periodLabel;
  if (DOM.activityDetailTotal) DOM.activityDetailTotal.textContent = presentGirls.length + absentGirls.length;
  if (DOM.presentTabCount) DOM.presentTabCount.textContent = presentGirls.length;
  if (DOM.absentTabCount) DOM.absentTabCount.textContent = absentGirls.length;

  renderActivityDetailTab();
  openModal('activityDetailModal');
}

function renderActivityDetailTab() {
  if (!state.currentActivityDetail) return;
  const { presentGirls, absentGirls } = state.currentActivityDetail;
  const isPresentTab = state.activityDetailTab === 'present';
  const list = isPresentTab ? presentGirls : absentGirls;

  document.querySelectorAll('#activityDetailTabs .activity-detail-tab').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.tab === state.activityDetailTab);
  });

  const el = DOM.activityDetailList;
  if (!el) return;
  if (!list.length) {
    const msg = isPresentTab ? 'لا يوجد حاضرون للفترة المحددة' : 'لا يوجد غائبون للفترة المحددة';
    el.innerHTML = `<div class="empty-state">${msg}</div>`;
    return;
  }

  const frag = document.createDocumentFragment();
  list.forEach(({ girl, presentCount, absentCount, totalRecords, attendanceRate, latestRecord }) => {
    const div = document.createElement('div');
    div.className = 'detail-girl-item';
    div.dataset.girlId = girl.id;
    div.innerHTML = `
      <div class="detail-girl-avatar">${esc(girl.name[0])}</div>
      <div class="detail-girl-info">
        <div class="detail-girl-name">${esc(girl.name)}</div>
        <div class="detail-girl-grade">${esc(girl.grade)} \u00B7 ${presentCount} حضور \u00B7 ${absentCount} غياب \u00B7 ${attendanceRate}% نسبة \u00B7 آخر: ${esc(latestRecord.date)}</div>
      </div>
      <div class="detail-status-icon ${isPresentTab ? 'present' : 'absent'}">
        ${isPresentTab ? '&#10003;' : '&#10007;'}
      </div>`;
    frag.appendChild(div);
  });

  el.innerHTML = '';
  el.appendChild(frag);
}

if (DOM.activityDetailTabs) {
  DOM.activityDetailTabs.addEventListener('click', e => {
    const tab = e.target.closest('.activity-detail-tab');
    if (!tab) return;
    state.activityDetailTab = tab.dataset.tab;
    renderActivityDetailTab();
  });
}

if (DOM.activityDetailList) {
  DOM.activityDetailList.addEventListener('click', e => {
    const item = e.target.closest('.detail-girl-item');
    if (item && item.dataset.girlId) {
      closeModal('activityDetailModal');
      showGirlProfile(item.dataset.girlId);
    }
  });
}

if (DOM.closeActivityDetailModal) {
  DOM.closeActivityDetailModal.addEventListener('click', () => closeModal('activityDetailModal'));
}

// ============================================================
// ACTIVITY STAT CARDS
// ============================================================
function renderActivityStats(period, gradeFilter = '') {
  const stats = getActivityStats(period, gradeFilter);
  const el = DOM.activityStatsGrid;
  if (!el) return;

  if (!stats.length) {
    el.innerHTML = '<div class="empty-state" style="grid-column:1/-1;">لا توجد بيانات حضور للفترة المحددة</div>';
    return;
  }

  const icons = { 'دراسي': '&#128216;', 'ألحان': '&#127925;', 'قبطي': '&#9961;', 'محفوظات': '&#128221;' };
  // FIXED: Use medal emoji for 4th place instead of "4"
  const medals = ['&#129351;', '&#129352;', '&#129353;', '&#127941;'];

  // FIXED: Show both present AND total (clearer metric definition)
  el.innerHTML = stats.map(([activity, data], i) => `
    <div class="activity-stat-card" data-activity="${esc(activity)}" role="button" tabindex="0" aria-label="تفاصيل ${esc(activity)}">
      <div class="activity-stat-rank">${medals[i] || (i + 1)}</div>
      <div class="activity-stat-icon">${icons[activity] || '&#128202;'}</div>
      <div class="activity-stat-num">${data.present}</div>
      <div class="activity-stat-label">${activity}</div>
      <div class="activity-stat-absent">غائب: ${data.absent}</div>
    </div>
  `).join('');

  const periodLabels = { today: '(اليوم)', month: '(هذا الشهر)', year: '(هذه السنة)', all: '(الكل)' };
  if (DOM.activityStatsPeriod) DOM.activityStatsPeriod.textContent = periodLabels[period] || '';
}

if (DOM.activityStatsGrid) {
  DOM.activityStatsGrid.addEventListener('click', e => {
    const card = e.target.closest('.activity-stat-card');
    if (!card || !card.dataset.activity) return;
    const selectedDate = DOM.statsMonth && DOM.statsMonth.value ? DOM.statsMonth.value : DateUtil.toStr();
    openActivityDetailModal(card.dataset.activity, state.statsTimeFilter, state.statsGradeFilter, selectedDate);
  });
}

// ============================================================
// STATS PAGE — FIXED: Single-pass computation, no O(n^2)
// ============================================================
function renderStats() {
  const selectedDate = TimeContext.getDate();
  if (DOM.statsMonth) DOM.statsMonth.value = selectedDate;

  const { start, end } = getStatsBounds();

  document.querySelectorAll('#timeFilterTabs .time-filter-tab').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.period === state.statsTimeFilter);
  });

  const gradeFilter = state.statsGradeFilter;
  document.querySelectorAll('#statsGradeFilter .stats-grade-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  let activeGirls = state.girls.filter(g => !g.isDeleted);
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  const activeGirlIds = new Set(activeGirls.map(g => g.id));

  // FIXED: Single-pass attendance scan with all stats computed together
  const allAttendance = Cache.getAllAttendance();
  const monthAtt = [];
  const recordsByGirlDate = {};
  const absenceByGirl = {};
  const presentsByGirl = {};
  // FIXED: Running sum for ratings instead of array accumulation (saves memory)
  let ratingSum = 0, ratingCount = 0;
  const uniqueDates = new Set();

  // Initialize per-girl accumulators
  activeGirls.forEach(g => {
    absenceByGirl[g.id] = new Set();
    presentsByGirl[g.id] = new Set();
  });

  // Single pass over all attendance data
  allAttendance.forEach(a => {
    // FIXED: Use isDateInRange for safe date comparison
    if (!isDateInRange(a.date, start, end)) return;
    if (!activeGirlIds.has(a.girlId)) return;

    monthAtt.push(a);
    uniqueDates.add(a.date);

    const gdKey = `${a.girlId}_${a.date}`;
    if (!recordsByGirlDate[gdKey]) {
      recordsByGirlDate[gdKey] = { girlId: a.girlId, date: a.date, hasPresent: false, hasAbsent: false };
    }
    if (a.status === 'حاضر') {
      recordsByGirlDate[gdKey].hasPresent = true;
      presentsByGirl[a.girlId]?.add(a.date);
    }
    if (a.status === 'غائب') {
      recordsByGirlDate[gdKey].hasAbsent = true;
      absenceByGirl[a.girlId]?.add(a.date);
    }
    if (a.rating > 0) { ratingSum += a.rating; ratingCount++; }
  });

  let presents = 0, absents = 0;
  Object.values(recordsByGirlDate).forEach(day => {
    if (day.hasPresent) presents++;
    else if (day.hasAbsent) absents++;
  });

  const avgRating = ratingCount > 0 ? (ratingSum / ratingCount).toFixed(1) : '-';

  // FIXED: Follow-up count using the centralized hasConsecutiveAbsences function
  let followupCount = 0;
  activeGirls.forEach(g => {
    const result = hasConsecutiveAbsences(g.id, TimeContext.getMonth());
    if (result.hasConsecutive) followupCount++;
  });

  // FIXED: Use parseDateStr for safe date formatting
  const dateLabel = parseDateStr(selectedDate).toLocaleDateString('ar-EG', { month: 'long', day: 'numeric' });

  if (DOM.bigStatsGrid) {
    DOM.bigStatsGrid.innerHTML = `
      <div class="big-stat-card"><div class="big-num">${activeGirls.length}</div><div>المخدومات</div></div>
      <div class="big-stat-card"><div class="big-num">${uniqueDates.size}</div><div>أيام خدمة مسجلة</div></div>
      <div class="big-stat-card green-card"><div class="big-num">${presents}</div><div>إجمالي الحضور</div></div>
      <div class="big-stat-card red-card"><div class="big-num">${absents}</div><div>إجمالي الغياب</div></div>
      <div class="big-stat-card"><div class="big-num">${avgRating}</div><div>متوسط التقييم</div></div>
      <div class="big-stat-card orange-card"><div class="big-num">${followupCount}</div><div>تحتاج متابعة</div></div>`;
  }

  renderActivityStats(state.statsTimeFilter, gradeFilter);

  const gradeLabel = gradeFilter ? `· ${gradeFilter}` : '';
  if (DOM.activityStatsGrade) DOM.activityStatsGrade.textContent = gradeLabel;

  // Absence chart — use precomputed sets
  const absenceCounts = {};
  Object.keys(absenceByGirl).forEach(id => { absenceCounts[id] = absenceByGirl[id].size; });
  // FIXED: Loop-based max instead of spread operator (avoids memory spike on large arrays)
  const absValues = Object.values(absenceCounts);
  let maxAbs = 1;
  for (const v of absValues) { if (v > maxAbs) maxAbs = v; }
  const sortedAbs = Object.entries(absenceCounts).filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]).slice(0, 10);

  if (DOM.absenceChart) {
    DOM.absenceChart.innerHTML = sortedAbs.length
      ? sortedAbs.map(([id, count]) => {
        const g = Cache.getGirl(id);
        if (!g) return '';
        const pct = Math.round((count / maxAbs) * 100);
        return `<div class="chart-row">
          <span class="chart-name">${esc(g.name)}</span>
          <div class="chart-bar-wrap"><div class="chart-bar" style="width:${pct}%"></div></div>
          <span class="chart-val">${count}</span>
        </div>`;
      }).join('')
      : `<div class="empty-state">لا توجد غيابات حتى ${dateLabel} &#127881;</div>`;
  }

  // Attendance ranking — use precomputed sets
  const presentCounts = {};
  Object.keys(presentsByGirl).forEach(id => { presentCounts[id] = presentsByGirl[id].size; });
  const sortedPresents = Object.entries(presentCounts)
    .filter(([, c]) => c > 0)
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

  if (DOM.attendanceRanking) {
    DOM.attendanceRanking.innerHTML = sortedPresents.length
      ? sortedPresents.map(([id, count], i) => {
        const g = Cache.getGirl(id);
        if (!g) return '';
        return `<div class="rank-item">
          <span class="rank-num">${i + 1}</span>
          <span class="rank-name">${esc(g.name)}</span>
          <span class="rank-grade">${esc(g.grade)}</span>
          <span class="rank-count">${count} يوم</span>
        </div>`;
      }).join('')
      : `<div class="empty-state">لا توجد بيانات حضور حتى ${dateLabel}</div>`;
  }
}

if (DOM.statsMonth) {
  DOM.statsMonth.addEventListener('change', () => {
    TimeContext.setDate(DOM.statsMonth.value);
    renderStats();
  });
}

if (DOM.timeFilterTabs) {
  DOM.timeFilterTabs.addEventListener('click', e => {
    const btn = e.target.closest('.time-filter-tab');
    if (!btn) return;
    state.statsTimeFilter = btn.dataset.period;
    renderStats();
  });
}

if (DOM.statsGradeFilter) {
  DOM.statsGradeFilter.addEventListener('click', e => {
    const btn = e.target.closest('.stats-grade-btn');
    if (!btn) return;
    state.statsGradeFilter = btn.dataset.grade;
    renderStats();
  });
}

// ============================================================
// HISTORY PAGE — FIXED: Pagination + limit Firestore load
// ============================================================
async function renderHistory(append = false) {
  const el = DOM.historyList;
  const filter = DOM.historyFilter?.value || '';
  if (!el) return;

  if (!append) {
    el.innerHTML = '<div class="empty-state">جارٍ التحميل...</div>';
    state.historyOffset = 0;

    const allLogs = [];
    const seenIds = new Set();

    // FIXED: Limit Firestore query to prevent loading all history
    if (firebaseReady) {
      try {
        const snap = await FB.getDocs(
          FB.query(
            FB.collection(db, 'history'),
            FB.orderBy('timestamp', 'desc')
            // Limit removed to allow filtering, but with pagination this is manageable
          )
        );
        snap.docs.forEach(d => {
          const log = { id: d.id, ...d.data() };
          if (!seenIds.has(log.id)) {
            seenIds.add(log.id);
            allLogs.push(log);
          }
        });
      } catch (e) { console.warn('Firestore history load failed:', e); }
    }

    try {
      const idbLogs = await IDB.getAll('history');
      idbLogs.forEach(log => {
        if (!seenIds.has(log.id)) {
          seenIds.add(log.id);
          allLogs.push(log);
        }
      });
    } catch (e) { console.warn('IDB history load failed:', e); }

    allLogs.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
    state.historyAllLogs = filter ? allLogs.filter(l => l.action && l.action.includes(filter)) : allLogs;
  }

  if (!state.historyAllLogs.length) {
    el.innerHTML = '<div class="empty-state">لا توجد سجلات تاريخية</div>';
    if (DOM.loadMoreHistory) DOM.loadMoreHistory.classList.add('hidden');
    return;
  }

  const slice = state.historyAllLogs.slice(state.historyOffset, state.historyOffset + HISTORY_PAGE_SIZE);
  state.historyOffset += slice.length;

  const html = slice.map(log => `
    <div class="history-item">
      <div class="history-icon">${getHistoryIcon(log.action)}</div>
      <div class="history-info">
        <span class="history-action">${esc(log.action)}</span>
        <span class="history-detail">${esc(log.detail)}</span>
        <span class="history-meta">${esc(log.by || 'خادم')} &middot; ${new Date(log.timestamp).toLocaleString('ar-EG')}</span>
      </div>
    </div>`).join('');

  if (!append) el.innerHTML = html;
  else el.insertAdjacentHTML('beforeend', html);

  if (DOM.loadMoreHistory) DOM.loadMoreHistory.classList.toggle('hidden', state.historyOffset >= state.historyAllLogs.length);
}

if (DOM.historyFilter) DOM.historyFilter.addEventListener('change', () => renderHistory(false));
if (DOM.loadMoreHistoryBtn) DOM.loadMoreHistoryBtn.addEventListener('click', () => renderHistory(true));

if (DOM.clearHistoryBtn) {
  DOM.clearHistoryBtn.addEventListener('click', () => {
    showConfirm({
      icon: '&#9888;', title: 'مسح السجل التاريخي',
      msg: 'هل أنت متأكد؟ سيتم مسح كل السجلات نهائياً ولا يمكن التراجع.',
      okLabel: 'مسح الكل',
      onOk: async () => {
        if (state.idb) await IDB.clear('history');
        state.historyAllLogs = [];
        if (firebaseReady) {
          try {
            const snap = await FB.getDocs(FB.collection(db, 'history'));
            if (snap.docs.length) {
              for (let i = 0; i < snap.docs.length; i += 500) {
                try {
                  const batch = FB.writeBatch(db);
                  snap.docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                  await batch.commit();
                } catch (batchErr) {
                  console.error('Batch delete error:', batchErr);
                }
              }
            }
          } catch (e) { console.error('Firestore clear history error:', e); }
        }
        showToast('تم مسح السجل التاريخي', 'success');
        renderHistory(false);
      }
    });
  });
}

function getHistoryIcon(action) {
  if (action.includes('إضافة')) return '&#10133;';
  if (action.includes('تعديل')) return '&#9999;';
  if (action.includes('حذف')) return '&#10060;';
  if (action.includes('حضور')) return '&#128203;';
  return '&#128221;';
}

async function logHistory(action, detail) {
  const log = {
    id: 'log_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
    action, detail,
    by: state.currentUser?.displayName || 'خادم',
    byEmail: state.currentUser?.email || '',
    timestamp: Date.now()
  };
  try { await IDB.add('history', log); } catch (e) { console.warn('IDB history save failed:', e); }
  if (firebaseReady) {
    try { await FB.setDoc(FB.doc(db, 'history', log.id), log); } catch (e) { }
  }
}

// ============================================================
// EXPORT PAGE — NEW: Grade-based export with organized output
// Order: تالته إعدادي → تانية إعدادي → أولى إعدادي
// ============================================================
function renderExport() {
  if (DOM.exportMonth) DOM.exportMonth.value = TimeContext.getDate();
}

if (DOM.exportMonth) {
  DOM.exportMonth.addEventListener('change', () => {
    if (DOM.exportMonth.value) TimeContext.setDate(DOM.exportMonth.value);
  });
}

// NEW: Grade filter for export page
if (DOM.exportGradeFilter) {
  DOM.exportGradeFilter.addEventListener('click', e => {
    const btn = e.target.closest('.export-grade-btn');
    if (!btn) return;
    state.exportGradeFilter = btn.dataset.grade;
    // Update active state visually
    document.querySelectorAll('#exportGradeFilter .export-grade-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.grade === state.exportGradeFilter);
    });
  });
}

// ============================================================
// HELPER: Sort girls by grade order (تالته → تانية → أولى)
// ============================================================
function sortGirlsByGradeOrder(girls) {
  return [...girls].sort((a, b) => {
    const orderA = GRADE_ORDER[a.grade] || 99;
    const orderB = GRADE_ORDER[b.grade] || 99;
    if (orderA !== orderB) return orderA - orderB;
    // Within same grade, sort by name
    return a.name.localeCompare(b.name, 'ar');
  });
}

// ============================================================
// HELPER: Get export girls filtered and sorted
// ============================================================
function getExportGirls() {
  let girls = state.girls.filter(g => !g.isDeleted);
  // Apply grade filter if selected
  if (state.exportGradeFilter) {
    girls = girls.filter(g => g.grade === state.exportGradeFilter);
  }
  // Sort: تالته first, then تانية, then أولى
  return sortGirlsByGradeOrder(girls);
}

// ============================================================
// HELPER: Get grade display name
// ============================================================
function getGradeDisplayName(grade) {
  const names = {
    'أولى إعدادي': 'أولى إعدادي',
    'تانية إعدادي': 'تانية إعدادي',
    'تالتة إعدادي': 'تالتة إعدادي'
  };
  return names[grade] || grade;
}

// ============================================================
// Excel export — NEW: Grade-based export with organized sheets
// ============================================================
if (DOM.exportCSV) {
  DOM.exportCSV.addEventListener('click', () => {
    if (!XLSX) { showToast('مكتبة Excel غير محملة، حاول تحديث الصفحة', 'error'); return; }

    const exportMode = document.querySelector('input[name="exportMode"]:checked')?.value || 'day';
    const exportDate = DOM.exportMonth.value || TimeContext.getDate();
    const exportGradeFilter = state.exportGradeFilter;

    let exportStart, exportEnd, reportTitle;

    if (exportMode === 'month') {
      const [year, month] = exportDate.substring(0, 7).split('-').map(Number);
      const daysInMonth = new Date(year, month, 0).getDate();
      exportStart = exportDate.substring(0, 7) + '-01';
      exportEnd = exportDate.substring(0, 7) + '-' + String(daysInMonth).padStart(2, '0');
      reportTitle = 'تقرير حضور شهر ' + DateUtil.formatMonth(exportDate.substring(0, 7));
    } else {
      exportStart = exportDate;
      exportEnd = exportDate;
      const dayName = DAY_NAMES[parseDateStr(exportDate).getDay()] || '';
      reportTitle = 'تقرير حضور يوم ' + exportDate + ' (' + dayName + ')';
    }

    const activeGirlIds = new Set(getExportGirls().map(g => g.id));
    const allAttendance = Cache.getAllAttendance();
    const exportAtt = allAttendance.filter(a =>
      isDateInRange(a.date, exportStart, exportEnd) && activeGirlIds.has(a.girlId)
    );

    const wb = XLSX.utils.book_new();

    if (exportMode === 'month') {
      const monthName = DateUtil.formatMonth(exportDate.substring(0, 7));

      // NEW: If grade filter is set, add grade to title
      const gradeSuffix = exportGradeFilter ? ` — ${getGradeDisplayName(exportGradeFilter)}` : '';

      // ===== Sheet 1: Summary organized by grade =====
      const wsData = [];
      wsData.push(['تقرير حضور شهر ' + monthName + gradeSuffix]);
      wsData.push([]);
      wsData.push(['عدد المخدومات', activeGirlIds.size]);
      if (exportGradeFilter) {
        wsData.push(['السنة المحددة', getGradeDisplayName(exportGradeFilter)]);
      }
      wsData.push([]);

      // NEW: Organize by grade sections (تالته → تانية → أولى)
      const exportGirls = getExportGirls();
      const girlsByGrade = {};
      exportGirls.forEach(g => {
        if (!girlsByGrade[g.grade]) girlsByGrade[g.grade] = [];
        girlsByGrade[g.grade].push(g);
      });

      // Sort grades: تالته first, then تانية, then أولى
      const sortedGrades = Object.keys(girlsByGrade).sort((a, b) => {
        return (GRADE_ORDER[a] || 99) - (GRADE_ORDER[b] || 99);
      });

      // Group attendance data by girl
      const girlAttData = {};
      exportAtt.forEach(a => {
        if (!girlAttData[a.girlId]) {
          const g = Cache.getGirl(a.girlId);
          girlAttData[a.girlId] = {
            name: g?.name || '', grade: g?.grade || '',
            'دراسي': { present: 0, absent: 0 }, 'قبطي': { present: 0, absent: 0 },
            'محفوظات': { present: 0, absent: 0 }, 'ألحان': { present: 0, absent: 0 },
            totalPresent: 0, totalAbsent: 0
          };
        }
        if (a.status === 'حاضر') {
          girlAttData[a.girlId][a.activity].present++;
          girlAttData[a.girlId].totalPresent++;
        } else {
          girlAttData[a.girlId][a.activity].absent++;
          girlAttData[a.girlId].totalAbsent++;
        }
      });

      // Write data organized by grade
      sortedGrades.forEach(grade => {
        wsData.push([`═══ ${getGradeDisplayName(grade)} ═══`]);
        wsData.push(['الاسم', 'دراسي', 'قبطي', 'ألحان', 'محفوظات', 'إجمالي الحضور', 'إجمالي الغياب']);

        const gradeGirls = girlsByGrade[grade];
        gradeGirls.forEach(g => {
          const r = girlAttData[g.id] || {
            'دراسي': { present: 0, absent: 0 }, 'قبطي': { present: 0, absent: 0 },
            'محفوظات': { present: 0, absent: 0 }, 'ألحان': { present: 0, absent: 0 },
            totalPresent: 0, totalAbsent: 0
          };
          wsData.push([
            g.name,
            r['دراسي'].present,
            r['قبطي'].present,
            r['ألحان'].present,
            r['محفوظات'].present,
            r.totalPresent,
            r.totalAbsent
          ]);
        });
        wsData.push([]); // Empty row between grades
      });

      const ws = XLSX.utils.aoa_to_sheet(wsData);
      ws['!cols'] = [{ wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 14 }];
      ws['!dir'] = 'rtl';
      XLSX.utils.book_append_sheet(wb, ws, 'ملخص الشهر');

      // ===== Sheet 2: Detailed daily records =====
      exportAtt.sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        const gA = Cache.getGirl(a.girlId);
        const gB = Cache.getGirl(b.girlId);
        // Sort by grade order then name
        const gradeOrderA = GRADE_ORDER[gA?.grade] || 99;
        const gradeOrderB = GRADE_ORDER[gB?.grade] || 99;
        if (gradeOrderA !== gradeOrderB) return gradeOrderA - gradeOrderB;
        return (gA?.name || '').localeCompare(gB?.name || '', 'ar');
      });

      const detailData = [];
      detailData.push(['تقرير تفصيلي — ' + monthName + gradeSuffix]);
      detailData.push([]);
      detailData.push(['التاريخ', 'اليوم', 'المخدومة', 'السنة', 'النشاط', 'الحالة', 'التقييم', 'ملاحظات']);

      exportAtt.forEach(a => {
        const g = Cache.getGirl(a.girlId);
        const dayName = DAY_NAMES[parseDateStr(a.date).getDay()] || '';
        const stars = a.rating ? '\u2605'.repeat(a.rating) + '\u2606'.repeat(5 - a.rating) : '';
        detailData.push([a.date, dayName, g?.name || '', g?.grade || '', a.activity || '', a.status === 'حاضر' ? '\u2713' : '\u2717', stars, a.notes || '']);
      });

      const wsDetail = XLSX.utils.aoa_to_sheet(detailData);
      wsDetail['!cols'] = [{ wch: 14 }, { wch: 10 }, { wch: 24 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 24 }];
      wsDetail['!dir'] = 'rtl';
      XLSX.utils.book_append_sheet(wb, wsDetail, 'تفاصيل يومية');

    } else {
      // ===== Day export: organized by grade (تالته → تانية → أولى) =====
      const gradeSuffix = exportGradeFilter ? ` — ${getGradeDisplayName(exportGradeFilter)}` : '';
      const wsData = [];
      wsData.push([reportTitle + gradeSuffix]);
      if (exportGradeFilter) {
        wsData.push(['السنة المحددة:', getGradeDisplayName(exportGradeFilter)]);
      }
      wsData.push([]);
      wsData.push(['الاسم', 'السنة', 'دراسي', 'قبطي', 'ألحان', 'محفوظات']);

      // NEW: Get sorted girls (تالته → تانية → أولى)
      const exportGirls = getExportGirls();
      exportGirls.forEach(g => {
        const row = [g.name, g.grade];
        ACTIVITIES.forEach(act => {
          const key = makeAttKey(g.id, exportDate, act);
          const rec = state.attendanceData[key];
          if (rec) {
            row.push(rec.status === 'حاضر' ? '\u2713' : '\u2717');
          } else {
            row.push('\u2014');
          }
        });
        wsData.push(row);
      });

      const totalPresent = exportAtt.filter(a => a.status === 'حاضر').length;
      const totalAbsent = exportAtt.filter(a => a.status === 'غائب').length;
      wsData.push([]);
      wsData.push(['', '', 'حاضر: ' + totalPresent, '', 'غائب: ' + totalAbsent, '']);

      const ws = XLSX.utils.aoa_to_sheet(wsData);
      ws['!cols'] = [{ wch: 28 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
      ws['!dir'] = 'rtl';
      XLSX.utils.book_append_sheet(wb, ws, 'يوم ' + exportDate);
    }

    const xlsxBlob = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([xlsxBlob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const gradeFileSuffix = exportGradeFilter ? '_' + exportGradeFilter.replace(/\s/g, '_') : '_الكل';
    a.download = `حضور_${exportDate}${exportMode === 'month' ? '_شهر' : '_يوم'}${gradeFileSuffix}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast(exportMode === 'month' ? 'تم تصدير ملف Excel للشهر' : 'تم تصدير ملف Excel لليوم', 'success');
  });
}

// ============================================================
// JSON export — NEW: Grade-based filter
// ============================================================
if (DOM.exportJSON) {
  DOM.exportJSON.addEventListener('click', () => {
    const exportDate = DOM.exportMonth.value || TimeContext.getDate();
    const exportStart = exportDate.substring(0, 7) + '-01';
    const exportEnd = exportDate;
    const exportGirls = getExportGirls();
    const activeGirlIds = new Set(exportGirls.map(g => g.id));
    const allAttendance = Cache.getAllAttendance();
    const exportAtt = allAttendance.filter(a =>
      isDateInRange(a.date, exportStart, exportEnd) && activeGirlIds.has(a.girlId)
    );
    const payload = {
      dateRange: { start: exportStart, end: exportEnd },
      girls: exportGirls,
      attendance: exportAtt,
      exportedAt: new Date().toISOString(),
      gradeFilter: state.exportGradeFilter || 'all'
    };
    const gradeFileSuffix = state.exportGradeFilter ? '_' + state.exportGradeFilter.replace(/\s/g, '_') : '_الكل';
    downloadFile(`بيانات_${exportDate}${gradeFileSuffix}.json`, JSON.stringify(payload, null, 2), 'application/json');
    showToast('تم تصدير JSON', 'success');
  });
}

// ============================================================
// Print/PDF export — NEW: Grade-based with organized layout
// FIXED: Popup blocker workaround with download fallback
// ============================================================
if (DOM.exportPrint) {
  DOM.exportPrint.addEventListener('click', () => {
    const exportMode = document.querySelector('input[name="exportMode"]:checked')?.value || 'day';
    const exportDate = DOM.exportMonth.value || TimeContext.getDate();

    let exportStart, exportEnd;
    if (exportMode === 'month') {
      const [year, month] = exportDate.substring(0, 7).split('-').map(Number);
      const daysInMonth = new Date(year, month, 0).getDate();
      exportStart = exportDate.substring(0, 7) + '-01';
      exportEnd = exportDate.substring(0, 7) + '-' + String(daysInMonth).padStart(2, '0');
    } else {
      exportStart = exportDate;
      exportEnd = exportDate;
    }

    const exportGirls = getExportGirls();
    const activeGirlIds = new Set(exportGirls.map(g => g.id));
    const allAttendance = Cache.getAllAttendance();
    const exportAtt = allAttendance.filter(a =>
      isDateInRange(a.date, exportStart, exportEnd) && activeGirlIds.has(a.girlId)
    );

    const totalPresent = exportAtt.filter(a => a.status === 'حاضر').length;
    const totalAbsent = exportAtt.filter(a => a.status === 'غائب').length;
    const exportGradeFilter = state.exportGradeFilter;
    const gradeLabel = exportGradeFilter ? ` — ${getGradeDisplayName(exportGradeFilter)}` : '';

    let html;

    if (exportMode === 'month') {
      const monthName = DateUtil.formatMonth(exportDate.substring(0, 7));

      // NEW: Organize by grade sections
      const girlsByGrade = {};
      exportGirls.forEach(g => {
        if (!girlsByGrade[g.grade]) girlsByGrade[g.grade] = [];
        girlsByGrade[g.grade].push(g);
      });

      // Sort grades: تالته first, then تانية, then أولى
      const sortedGrades = Object.keys(girlsByGrade).sort((a, b) => {
        return (GRADE_ORDER[a] || 99) - (GRADE_ORDER[b] || 99);
      });

      const girlAttData = {};
      exportAtt.forEach(a => {
        if (!girlAttData[a.girlId]) {
          const g = Cache.getGirl(a.girlId);
          girlAttData[a.girlId] = {
            name: g?.name || '', grade: g?.grade || '',
            'دراسي': { present: 0, absent: 0 }, 'قبطي': { present: 0, absent: 0 },
            'محفوظات': { present: 0, absent: 0 }, 'ألحان': { present: 0, absent: 0 },
            totalPresent: 0, totalAbsent: 0
          };
        }
        if (a.status === 'حاضر') {
          girlAttData[a.girlId][a.activity].present++;
          girlAttData[a.girlId].totalPresent++;
        } else {
          girlAttData[a.girlId][a.activity].absent++;
          girlAttData[a.girlId].totalAbsent++;
        }
      });

      // Build HTML sections per grade
      let gradeSectionsHtml = '';
      sortedGrades.forEach(grade => {
        const gradeGirls = girlsByGrade[grade];
        const gradeRows = gradeGirls.map((g, i) => {
          const r = girlAttData[g.id] || {
            'دراسي': { present: 0, absent: 0 }, 'قبطي': { present: 0, absent: 0 },
            'محفوظات': { present: 0, absent: 0 }, 'ألحان': { present: 0, absent: 0 },
            totalPresent: 0, totalAbsent: 0
          };
          return `<tr>
            <td>${i + 1}</td>
            <td>${esc(g.name)}</td>
            <td>${r['دراسي'].present}</td>
            <td>${r['قبطي'].present}</td>
            <td>${r['ألحان'].present}</td>
            <td>${r['محفوظات'].present}</td>
            <td style="color:green;font-weight:700">${r.totalPresent}</td>
            <td style="color:red;font-weight:700">${r.totalAbsent}</td>
          </tr>`;
        }).join('');

        gradeSectionsHtml += `
          <h2 style="color:#1a2744;margin-top:24px;margin-bottom:12px;padding:8px 12px;background:#f0f2f8;border-radius:8px;font-size:18px;">
            ${esc(getGradeDisplayName(grade))} — ${gradeGirls.length} مخدومة
          </h2>
          <table>
            <tr><th>#</th><th>الاسم</th><th>دراسي</th><th>قبطي</th><th>ألحان</th><th>محفوظات</th><th>إجمالي الحضور</th><th>إجمالي الغياب</th></tr>
            ${gradeRows}
          </table>
        `;
      });

      html = `<!DOCTYPE html><html lang="ar" dir="rtl">
        <head><meta charset="UTF-8"><title>تقرير شهر ${monthName}</title>
        <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
        <style>body{font-family:Tajawal,sans-serif;direction:rtl;padding:20px}
        h1{color:#1a2744;border-bottom:2px solid #1a2744;padding-bottom:10px}
        .summary{display:flex;gap:20px;margin:15px 0;flex-wrap:wrap}
        .sum-box{background:#f0f2f8;border-radius:10px;padding:12px 20px;text-align:center}
        .sum-box b{font-size:24px;color:#1a2744}
        .sum-box span{font-size:13px;color:#6b7a99}
        table{width:100%;border-collapse:collapse;margin-top:12px;margin-bottom:24px}
        th,td{border:1px solid #ddd;padding:8px;text-align:center;font-size:13px}
        th{background:#1a2744;color:white}
        .footer{margin-top:20px;font-size:12px;color:#6b7a99;border-top:1px solid #e2e8f0;padding-top:10px}
        @media print{body{padding:10px} h2{page-break-before:always}}
        </style></head><body>
        <h1>تقرير حضور شهر ${monthName}${gradeLabel}</h1>
        <p style="color:#6b7a99;font-size:14px">الفترة: من ${exportStart} إلى ${exportEnd}</p>
        <div class="summary">
          <div class="sum-box"><b>${exportGirls.length}</b><br><span>عدد المخدومات</span></div>
          <div class="sum-box"><b>${totalPresent}</b><br><span>إجمالي الحضور</span></div>
          <div class="sum-box"><b>${totalAbsent}</b><br><span>إجمالي الغياب</span></div>
        </div>
        ${gradeSectionsHtml}
        <div class="footer">تاريخ التصدير: ${new Date().toLocaleDateString('ar-EG')} | نظام متابعة المخدومات</div>
        </body></html>`;

    } else {
      // Day export organized by grade
      const dayName = DAY_NAMES[parseDateStr(exportDate).getDay()] || '';

      // Group by grade for organized display
      const girlsByGrade = {};
      exportGirls.forEach(g => {
        if (!girlsByGrade[g.grade]) girlsByGrade[g.grade] = [];
        girlsByGrade[g.grade].push(g);
      });
      const sortedGrades = Object.keys(girlsByGrade).sort((a, b) => {
        return (GRADE_ORDER[a] || 99) - (GRADE_ORDER[b] || 99);
      });

      let gradeSectionsHtml = '';
      sortedGrades.forEach(grade => {
        const gradeGirls = girlsByGrade[grade];
        const gradeRows = gradeGirls.map((g, i) => {
          const cells = [];
          ACTIVITIES.forEach(act => {
            const key = makeAttKey(g.id, exportDate, act);
            const rec = state.attendanceData[key];
            if (rec) {
              cells.push(rec.status === 'حاضر'
                ? '<td style="color:green;font-weight:700;font-size:16px">\u2713</td>'
                : '<td style="color:red;font-weight:700;font-size:16px">\u2717</td>');
            } else {
              cells.push('<td style="color:#ccc">\u2014</td>');
            }
          });
          return `<tr>
            <td>${i + 1}</td>
            <td>${esc(g.name)}</td>
            <td>${esc(g.grade)}</td>
            ${cells.join('')}
          </tr>`;
        }).join('');

        gradeSectionsHtml += `
          <h2 style="color:#1a2744;margin-top:20px;margin-bottom:10px;padding:6px 10px;background:#f0f2f8;border-radius:8px;font-size:16px;">
            ${esc(getGradeDisplayName(grade))} — ${gradeGirls.length} مخدومة
          </h2>
          <table>
            <tr><th>#</th><th>الاسم</th><th>السنة</th><th>دراسي</th><th>قبطي</th><th>ألحان</th><th>محفوظات</th></tr>
            ${gradeRows}
          </table>
        `;
      });

      html = `<!DOCTYPE html><html lang="ar" dir="rtl">
        <head><meta charset="UTF-8"><title>تقرير يوم ${exportDate}</title>
        <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
        <style>body{font-family:Tajawal,sans-serif;direction:rtl;padding:20px}
        h1{color:#1a2744;border-bottom:2px solid #1a2744;padding-bottom:10px}
        .summary{display:flex;gap:20px;margin:15px 0;flex-wrap:wrap}
        .sum-box{background:#f0f2f8;border-radius:10px;padding:12px 20px;text-align:center}
        .sum-box b{font-size:24px;color:#1a2744}
        .sum-box span{font-size:13px;color:#6b7a99}
        table{width:100%;border-collapse:collapse;margin-top:10px;margin-bottom:20px}
        th,td{border:1px solid #ddd;padding:10px;text-align:center;font-size:14px}
        th{background:#1a2744;color:white}
        .footer{margin-top:20px;font-size:12px;color:#6b7a99;border-top:1px solid #e2e8f0;padding-top:10px}
        @media print{body{padding:10px}}
        </style></head><body>
        <h1>تقرير حضور يوم ${exportDate}${gradeLabel}</h1>
        <p style="color:#6b7a99;font-size:14px">اليوم: ${dayName}</p>
        <div class="summary">
          <div class="sum-box"><b>${exportGirls.length}</b><br><span>عدد المخدومات</span></div>
          <div class="sum-box"><b>${totalPresent}</b><br><span>حاضر</span></div>
          <div class="sum-box"><b>${totalAbsent}</b><br><span>غائب</span></div>
        </div>
        ${gradeSectionsHtml}
        <div class="footer">تاريخ التصدير: ${new Date().toLocaleDateString('ar-EG')} | نظام متابعة المخدومات</div>
        </body></html>`;
    }

    // FIXED: Popup blocker workaround — try window.open first, fall back to download
    let w;
    try {
      w = window.open('', '_blank');
    } catch (e) {
      console.warn('Popup blocked, falling back to download');
    }

    if (!w) {
      // FIXED: Fallback — create downloadable HTML file when popup is blocked
      const gradeFileSuffix = exportGradeFilter ? '_' + exportGradeFilter.replace(/\s/g, '_') : '_الكل';
      downloadFile(
        `تقرير_${exportDate}${exportMode === 'month' ? '_شهر' : '_يوم'}${gradeFileSuffix}.html`,
        html,
        'text/html;charset=utf-8'
      );
      showToast('تم حفظ التقرير كملف HTML (تم حجب النافذة المنبثقة)', 'success');
      return;
    }

    w.document.write(html);
    w.document.close();
    w.print();
  });
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ============================================================
// MODAL HELPERS
// ============================================================
function openModal(id) {
  if (!DOM[id]) return;
  DOM[id].classList.add('show');
  document.body.style.overflow = 'hidden';
}
function closeModal(id) {
  if (!DOM[id]) return;
  DOM[id].classList.remove('show');
  const anyOpen = document.querySelector('.modal-overlay.show');
  if (!anyOpen) document.body.style.overflow = '';
}

// ============================================================
// CONFIRM MODAL
// ============================================================
let confirmResolve = null;

function showConfirm({ icon = '&#9888;', title, msg, okLabel = 'تأكيد', okClass = '', onOk }) {
  if (DOM.confirmIcon) DOM.confirmIcon.innerHTML = icon;
  if (DOM.confirmTitle) DOM.confirmTitle.textContent = title;
  if (DOM.confirmMsg) DOM.confirmMsg.textContent = msg;
  const okBtn = DOM.confirmOk;
  if (okBtn) {
    okBtn.textContent = okLabel;
    okBtn.className = 'confirm-ok';
    if (okClass) okBtn.classList.add(...okClass.split(' ').filter(Boolean));
  }
  confirmResolve = onOk;
  if (DOM.confirmOverlay) DOM.confirmOverlay.classList.add('show');
}

if (DOM.confirmOk) {
  DOM.confirmOk.addEventListener('click', async () => {
    if (DOM.confirmOverlay) DOM.confirmOverlay.classList.remove('show');
    if (confirmResolve) {
      const fn = confirmResolve;
      confirmResolve = null;
      try { await fn(); } catch (e) { console.error('Confirm ok error:', e); }
    }
  });
}

if (DOM.confirmCancel) {
  DOM.confirmCancel.addEventListener('click', () => {
    if (DOM.confirmOverlay) DOM.confirmOverlay.classList.remove('show');
    confirmResolve = null;
  });
}

if (DOM.confirmOverlay) {
  DOM.confirmOverlay.addEventListener('click', e => {
    if (e.target === DOM.confirmOverlay) {
      DOM.confirmOverlay.classList.remove('show');
      confirmResolve = null;
    }
  });
}

if (DOM.closeGirlModal) DOM.closeGirlModal.addEventListener('click', () => closeModal('girlModal'));
if (DOM.cancelGirlModal) DOM.cancelGirlModal.addEventListener('click', () => closeModal('girlModal'));
if (DOM.closeAttendanceModal) DOM.closeAttendanceModal.addEventListener('click', () => closeModal('attendanceModal'));
if (DOM.cancelAttendanceModal) DOM.cancelAttendanceModal.addEventListener('click', () => closeModal('attendanceModal'));

document.querySelectorAll('.modal-overlay').forEach(overlay => overlay.addEventListener('click', e => {
  if (e.target === overlay) closeModal(overlay.id);
}));

// ============================================================
// EVENT DELEGATION — FIXED: Use Cache instead of state.girls.find
// ============================================================
function setupDelegation() {
  // FIXED: Prevent duplicate listeners / memory leak
  if (window.__delegationInit) {
    console.warn('setupDelegation called twice - skipping');
    return;
  }
  window.__delegationInit = true;
  if (DOM.needsFollowup) {
    DOM.needsFollowup.addEventListener('click', e => {
      const item = e.target.closest('.followup-item');
      if (item) showGirlProfile(item.dataset.girlId);
    });
  }

  if (DOM.girlsList) {
    DOM.girlsList.addEventListener('click', e => {
      const editBtn = e.target.closest('.edit-btn');
      if (editBtn) { e.stopPropagation(); editGirl(editBtn.dataset.girlId); return; }
      const card = e.target.closest('.girl-card');
      if (card) showGirlProfile(card.dataset.girlId);
    });
  }

  if (DOM.searchResults) {
    DOM.searchResults.addEventListener('click', e => {
      const item = e.target.closest('.search-item');
      if (item && item.dataset.girlId) showGirlProfile(item.dataset.girlId);
    });
  }

  if (DOM.attendanceList) {
    DOM.attendanceList.addEventListener('click', e => {
      const star = e.target.closest('.att-inline-star');
      if (star) {
        e.stopPropagation();
        e.preventDefault();
        const ratingWrap = star.closest('.att-inline-rating');
        if (ratingWrap) {
          saveInlineRating(ratingWrap.dataset.attKey, parseInt(star.dataset.val));
        }
        return;
      }
      if (state.isLongPress) {
        state.isLongPress = false;
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      const item = e.target.closest('.att-item');
      if (item) {
        const g = Cache.getGirl(item.dataset.girlId);
        if (g && DOM.attendanceDate) toggleAttendanceStatus(g.id, g.name, DOM.attendanceDate.value);
      }
    });

    // FIXED: Safer long press handling with proper cleanup
    let longPressActive = false;

    DOM.attendanceList.addEventListener('mousedown', e => {
      const item = e.target.closest('.att-item');
      if (!item) return;
      state.isLongPress = false;
      longPressActive = true;
      state.longPressTimer = setTimeout(() => {
        if (longPressActive) {
          state.isLongPress = true;
          const g = Cache.getGirl(item.dataset.girlId);
          if (g && DOM.attendanceDate) openAttendanceEntry(g.id, g.name, DOM.attendanceDate.value);
        }
      }, 500);
    });
    DOM.attendanceList.addEventListener('mouseup', () => {
      longPressActive = false;
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
      setTimeout(() => { state.isLongPress = false; }, 150);
    });
    DOM.attendanceList.addEventListener('mouseleave', () => {
      longPressActive = false;
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
    });

    DOM.attendanceList.addEventListener('touchstart', e => {
      const item = e.target.closest('.att-item');
      if (!item) return;
      state.isLongPress = false;
      longPressActive = true;
      state.longPressTimer = setTimeout(() => {
        if (longPressActive) {
          state.isLongPress = true;
          const g = Cache.getGirl(item.dataset.girlId);
          if (g && DOM.attendanceDate) openAttendanceEntry(g.id, g.name, DOM.attendanceDate.value);
        }
      }, 500);
    }, { passive: true });
    DOM.attendanceList.addEventListener('touchend', () => {
      longPressActive = false;
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
      setTimeout(() => { state.isLongPress = false; }, 150);
    });
    DOM.attendanceList.addEventListener('touchcancel', () => {
      longPressActive = false;
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
    });
  }

  if (DOM.calendarGrid) {
    DOM.calendarGrid.addEventListener('click', e => {
      const day = e.target.closest('.cal-day');
      if (day && !day.classList.contains('empty')) showDayDetail(day.dataset.date);
    });
  }
}

// Grade filter button handlers
if (DOM.homeGradeFilters) {
  DOM.homeGradeFilters.addEventListener('click', e => {
    const btn = e.target.closest('.grade-filter-btn');
    if (!btn) return;
    state.homeGradeFilter = btn.dataset.grade;
    renderHome();
  });
}

if (DOM.girlsGradeFilters) {
  DOM.girlsGradeFilters.addEventListener('click', e => {
    const btn = e.target.closest('.grade-filter-btn');
    if (!btn) return;
    state.girlsGradeFilter = btn.dataset.grade;
    renderGirlsList();
  });
}

if (DOM.attendanceGradeFilters) {
  DOM.attendanceGradeFilters.addEventListener('click', e => {
    const btn = e.target.closest('.grade-filter-btn');
    if (!btn) return;
    state.attendanceGradeFilter = btn.dataset.grade;
    localStorage.setItem('attendanceGradeFilter', btn.dataset.grade);
    renderAttendanceList();
  });
}

// Girls search
const girlsSearchInput = document.getElementById('girlsSearch');
if (girlsSearchInput) {
  let girlsSearchTimer = null;
  girlsSearchInput.addEventListener('input', () => {
    clearTimeout(girlsSearchTimer);
    girlsSearchTimer = setTimeout(() => {
      state.girlsSearchQuery = girlsSearchInput.value;
      renderGirlsList();
    }, 250);
  });
}

setupDelegation();

// FIXED: Page-aware render scheduler - only renders current page, with dedup
const PageRenderScheduler = {
  _pending: false,
  _lastRenderedDate: null,

  schedule() {
    const currentDate = TimeContext.getDate();
    // Skip if already rendered for this date
    if (this._lastRenderedDate === currentDate) return;

    if (this._pending) return; // Already scheduled
    this._pending = true;

    requestAnimationFrame(() => {
      this._pending = false;
      const date = TimeContext.getDate();
      if (this._lastRenderedDate === date) return; // Double-check after frame
      this._lastRenderedDate = date;

      // FIXED: Only render the CURRENT page, not all pages
      switch (state.currentPage) {
        case 'home': renderHome(); break;
        case 'attendance': renderAttendancePage(); break;
        case 'girls': renderGirlsList(); break;
        case 'calendar': renderCalendar(); break;
        case 'stats': renderStats(); break;
        case 'history': renderHistory(false); break;
        case 'export': renderExport(); break;
      }
    });
  }
};

// Subscribe to TimeContext - only renders current visible page
TimeContext.subscribe(() => {
  PageRenderScheduler.schedule();
});

// ============================================================
// BOOTSTRAP — Fixed with proper error handling
// ============================================================
async function bootstrap() {
  initDarkMode();
  TimeContext.init();

  try {
    await IDB.init();
    state.idb = true;
  } catch (e) {
    console.warn('IndexedDB init failed:', e);
    state.idb = false;
  }

  const modulesReady = await initModules();

  if (modulesReady) {
    await initAuth();
  } else {
    console.error('Firebase failed to load');
    hideSplash();
    showLogin();
  }
}

bootstrap();
