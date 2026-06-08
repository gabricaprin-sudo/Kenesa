// ============================================================
// نظام متابعة المخدومات — Offline Ready & Guest Mode
// ============================================================

// ============================================================
// SAFETY: Global error handler + splash fallback
// ============================================================
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

let splashForceHidden = false;
function hideSplashForced() {
  if (splashForceHidden) return;
  splashForceHidden = true;
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
// ============================================================
let firebaseApp, auth, db, provider;
let firebaseReady = false;
let XLSX = null;

// Module imports with error handling
async function initModules() {
  try {
    const { initializeApp } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js');
    const { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, onAuthStateChanged, signOut } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js');
    const { getFirestore, collection, doc, setDoc, getDocs, deleteDoc, query, orderBy, onSnapshot, writeBatch, where } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js');

    const firebaseConfig = {
      apiKey: "AIzaSyB2cycBTKMjVg8S_fBYN8C-hwUk5FUF81Q",
      authDomain: "kenesa-e5efd.firebaseapp.com",
      projectId: "kenesa-e5efd",
      storageBucket: "kenesa-e5efd.firebasestorage.app",
      messagingSenderId: "227273753184",
      appId: "1:227273753184:web:ecdf258142ad55ed5cf905",
      measurementId: "G-6HS8KNW1GZ"
    };

    firebaseApp = initializeApp(firebaseConfig);
    auth = getAuth(firebaseApp);
    db = getFirestore(firebaseApp);
    provider = new GoogleAuthProvider();
    firebaseReady = true;

    // Attach Firebase functions to global scope for the app
    window._fb = { collection, doc, setDoc, getDocs, deleteDoc, query, orderBy, onSnapshot, writeBatch, where, signInWithPopup, signInWithRedirect, getRedirectResult, onAuthStateChanged, signOut };

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
    return false;
  }
}

// ============================================================
// DOM CACHE
// ============================================================
const $ = (id) => document.getElementById(id);
const $$ = (sel, root = document) => root.querySelectorAll(sel);

const DOM = {
  splash: $('splash'), loginScreen: $('loginScreen'), mainApp: $('mainApp'),
  pageTitle: $('pageTitle'), pageSubtitle: $('pageSubtitle'),
  syncIndicator: $('syncIndicator'), userAvatar: $('userAvatar'),
  drawer: $('drawer'), drawerOverlay: $('drawerOverlay'),
  drawerAvatar: $('drawerAvatar'), drawerUserName: $('drawerUserName'),
  drawerUserEmail: $('drawerUserEmail'), offlineBadge: $('offlineBadge'),
  pageContent: $('pageContent'), toast: $('toast'),
  globalSearch: $('globalSearch'), searchResults: $('searchResults'),
  todayDay: $('todayDay'), todayDate: $('todayDate'), todayServiceBadge: $('todayServiceBadge'),
  statTotal: $('statTotal'), statPresentToday: $('statPresentToday'),
  statAbsentToday: $('statAbsentToday'), statAvgRating: $('statAvgRating'),
  bestGrade: $('bestGrade'), bestGradePercent: $('bestGradePercent'),
  topActivityName: $('topActivityName'), topActivityCount: $('topActivityCount'),
  mostRegularGirl: $('mostRegularGirl'), mostRegularPercent: $('mostRegularPercent'),
  topAttendees: $('topAttendees'), needsFollowup: $('needsFollowup'),
  attendanceDate: $('attendanceDate'), attendanceList: $('attendanceList'),
  attendanceSearch: $('attendanceSearch'),
  presentCount: $('presentCount'), absentCount: $('absentCount'), totalCount: $('totalCount'),
  selectAllPresent: $('selectAllPresent'), selectAllAbsent: $('selectAllAbsent'),
  attToggleHint: $('attToggleHint'), quickActions: $('quickActions'),
  girlsList: $('girlsList'), addGirlBtn: $('addGirlBtn'),
  calendarGrid: $('calendarGrid'), calMonthYear: $('calMonthYear'),
  dayDetail: $('dayDetail'), calPrev: $('calPrev'), calNext: $('calNext'),
  statsMonth: $('statsMonth'), bigStatsGrid: $('bigStatsGrid'),
  absenceChart: $('absenceChart'), attendanceRanking: $('attendanceRanking'),
  activityStatsGrid: $('activityStatsGrid'), timeFilterTabs: $('timeFilterTabs'), activityStatsPeriod: $('activityStatsPeriod'),
  historyList: $('historyList'), historyFilter: $('historyFilter'),
  clearHistoryBtn: $('clearHistoryBtn'), loadMoreHistory: $('loadMoreHistory'),
  loadMoreHistoryBtn: $('loadMoreHistoryBtn'), exportMonth: $('exportMonth'),
  exportCSV: $('exportCSV'), exportJSON: $('exportJSON'), exportPrint: $('exportPrint'),
  girlModal: $('girlModal'), girlModalTitle: $('girlModalTitle'),
  girlName: $('girlName'), girlPhone: $('girlPhone'), girlGrade: $('girlGrade'),
  girlNotes: $('girlNotes'), deleteGirlBtn: $('deleteGirlBtn'),
  homeGradeFilters: $('homeGradeFilters'), girlsGradeFilters: $('girlsGradeFilters'),
  closeGirlModal: $('closeGirlModal'), cancelGirlModal: $('cancelGirlModal'),
  saveGirlBtn: $('saveGirlBtn'), girlProfileModal: $('girlProfileModal'),
  profileName: $('profileName'), profileBody: $('profileBody'),
  closeProfileModal: $('closeProfileModal'), attendanceModal: $('attendanceModal'),
  attendanceModalTitle: $('attendanceModalTitle'), modalGirlName: $('modalGirlName'),
  attendanceNotes: $('attendanceNotes'), ratingSection: $('ratingSection'),
  starsInput: $('starsInput'), saveAttendanceEntry: $('saveAttendanceEntry'),
  closeAttendanceModal: $('closeAttendanceModal'), cancelAttendanceModal: $('cancelAttendanceModal'),
  confirmOverlay: $('confirmOverlay'), confirmIcon: $('confirmIcon'),
  confirmTitle: $('confirmTitle'), confirmMsg: $('confirmMsg'),
  confirmCancel: $('confirmCancel'), confirmOk: $('confirmOk'),
  activityDetailModal: $('activityDetailModal'),
  activityDetailTitle: $('activityDetailTitle'),
  closeActivityDetailModal: $('closeActivityDetailModal'),
  activityDetailSummary: $('activityDetailSummary'),
  activityDetailIcon: $('activityDetailIcon'),
  activityDetailName: $('activityDetailName'),
  activityDetailPeriod: $('activityDetailPeriod'),
  activityDetailTotal: $('activityDetailTotal'),
  activityDetailTabs: $('activityDetailTabs'),
  activityDetailList: $('activityDetailList'),
  presentTabCount: $('presentTabCount'),
  absentTabCount: $('absentTabCount'),
  menuBtn: $('menuBtn'), signOutBtn: $('signOutBtn'), googleSignIn: $('googleSignIn'),
  darkModeToggle: $('darkModeToggle'), darkToggleSwitch: $('darkToggleSwitch'),
  shareProfileBtn: $('shareProfileBtn'), editProfileBtn: $('editProfileBtn'),
  statsGradeFilter: $('statsGradeFilter'),
  activityStatsGrade: $('activityStatsGrade')
};

// ============================================================
// APP STATE
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
  historyOffset: 0,
  historyAllLogs: [],
  deleteInProgress: false,
  homeGradeFilter: '',
  girlsGradeFilter: '',
  girlsSearchQuery: '',
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
};

const HISTORY_PAGE_SIZE = 30;
const SERVICE_DAYS = { 'السبت': true, 'الاثنين': true, 'الاربعاء': true };
const SERVICE_DAY_NUMBERS = [1, 3, 6]; // Mon, Wed, Sat
const DAY_NAMES = ['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
const ACTIVITIES = ['دراسي', 'محفوظات', 'قبطي', 'ألحان'];
const ACTIVITY_ICONS = { 'دراسي': '&#128216;', 'ألحان': '&#127925;', 'قبطي': '&#9961;', 'محفوظات': '&#128221;' };
const PERIOD_LABELS = { today: 'اليوم', month: 'هذا الشهر', year: 'هذه السنة', all: 'كل الفترات' };

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
// DATE UTILITIES
// ============================================================
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
  normalize(d) {
    return { 'الأحد': 'الاحد', 'الاثنين': 'الاثنين', 'الثلاثاء': 'الثلاثاء', 'الأربعاء': 'الاربعاء', 'الخميس': 'الخميس', 'الجمعة': 'الجمعة', 'السبت': 'السبت' }[d] || d;
  }
};

// ============================================================
// TIMECONTEXT — Unified Date Source for the entire app
// Every page uses this instead of computing dates independently
// ============================================================
const TimeContext = {
  _selectedDate: null,
  _listeners: [],

  init() {
    const saved = localStorage.getItem('trackerSelectedDate');
    this._selectedDate = saved || DateUtil.toStr();
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

  /** Get month string (YYYY-MM) */
  getMonth() {
    return this._selectedDate.substring(0, 7);
  },

  /** Get year string (YYYY) */
  getYear() {
    return this._selectedDate.substring(0, 4);
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
  const to = new Date(toDate + 'T00:00:00');
  const daysInMonth = new Date(fromYear, fromMonth + 1, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${fromYear}-${DateUtil.pad(fromMonth + 1)}-${DateUtil.pad(d)}`;
    const current = new Date(dateStr + 'T00:00:00');
    if (current > to) break;
    const dayOfWeek = current.getDay();
    if (SERVICE_DAY_NUMBERS.includes(dayOfWeek)) {
      count++;
    }
  }
  return count;
}

function hasConsecutiveAbsences(girlId, monthStr) {
  const absRecords = Object.values(state.attendanceData)
    .filter(a => a.girlId === girlId && a.date?.startsWith(monthStr) && a.status === 'غائب');

  if (absRecords.length < 2) return { hasConsecutive: false, count: absRecords.length, dates: [] };

  const absDates = [...new Set(absRecords.map(a => a.date))].sort();

  for (let i = 0; i < absDates.length - 1; i++) {
    const d1 = new Date(absDates[i] + 'T00:00:00');
    const d2 = new Date(absDates[i + 1] + 'T00:00:00');
    const diffDays = (d2 - d1) / (1000 * 60 * 60 * 24);
    if (diffDays <= 3) {
      return { hasConsecutive: true, count: absDates.length, dates: absDates };
    }
  }
  return { hasConsecutive: false, count: absDates.length, dates: absDates };
}


// ============================================================
// UNIFIED STATS BOUNDS — All stats use this single function
// ============================================================
function getStatsBounds() {
  // Unified: always use TimeContext as the single date source
  const selectedDate = TimeContext.getDate();
  const selYear = parseInt(selectedDate.substring(0, 4));
  const selMonth = parseInt(selectedDate.substring(5, 7));

  switch (state.statsTimeFilter) {
    case 'today':
      return { start: selectedDate, end: selectedDate };
    case 'month': {
      const lastDay = new Date(selYear, selMonth, 0).getDate();
      return { start: selectedDate.substring(0, 7) + '-01', end: selectedDate.substring(0, 7) + '-' + String(lastDay).padStart(2, '0') };
    }
    case 'year':
      return { start: selectedDate.substring(0, 4) + '-01-01', end: selectedDate.substring(0, 4) + '-12-31' };
    default: // 'all'
      return { start: '2000-01-01', end: selectedDate };
  }
}

// ============================================================
// ARABIC TEXT NORMALIZATION
// ============================================================
function normalizeArabic(str) {
  if (!str) return '';
  return str.replace(/[\u064B-\u065F\u0670\u0640]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ؤ/g, 'و')
    .replace(/ئ/g, 'ي')
    .replace(/ة/g, 'ه')
    .toLowerCase();
}

function normalizeName(name) {
  return name
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    .toLowerCase();
}

function csvEscape(v) {
  return `"${String(v ?? '').replace(/"/g, '""')}"`;
}

// ============================================================
// INDEXEDDB — wrapper for offline history storage
// ============================================================
const IDB = {
  db: null,
  DB_NAME: 'girlsTrackerDB',
  DB_VERSION: 1,

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
// DARK MODE
// ============================================================
function initDarkMode() {
  const saved = localStorage.getItem('darkMode');
  if (saved === 'true') {
    document.documentElement.setAttribute('data-theme', 'dark');
    if (DOM.darkToggleSwitch) DOM.darkToggleSwitch.classList.add('on');
  }
}

if (DOM.darkModeToggle) {
  DOM.darkModeToggle.addEventListener('click', () => {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    if (isDark) {
      document.documentElement.removeAttribute('data-theme');
      if (DOM.darkToggleSwitch) DOM.darkToggleSwitch.classList.remove('on');
      localStorage.setItem('darkMode', 'false');
    } else {
      document.documentElement.setAttribute('data-theme', 'dark');
      if (DOM.darkToggleSwitch) DOM.darkToggleSwitch.classList.add('on');
      localStorage.setItem('darkMode', 'true');
    }
  });
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
// SPLASH
// ============================================================
let splashDone = false;
function hideSplash() {
  if (splashDone) return;
  splashDone = true;
  splashForceHidden = true;
  if (DOM.splash) {
    DOM.splash.classList.add('fade-out');
    setTimeout(() => { if (DOM.splash) DOM.splash.remove(); }, 500);
  }
}

// ============================================================
// ONLINE / OFFLINE
// ============================================================


// ============================================================
// AUTH — Fixed with better error handling + Guest Mode
// ============================================================
async function initAuth() {
  if (!firebaseReady || !window._fb) {
    console.error('Firebase not available');
    hideSplash();
    showLogin();
    return;
  }

  try {
    const { getRedirectResult, onAuthStateChanged } = window._fb;
    try { await getRedirectResult(auth); } catch (e) { console.error('getRedirectResult error:', e); }

    onAuthStateChanged(auth, async (user) => {
      hideSplash();
      if (!user) {
        state.currentUser = null;
        state.appInitialized = false;
        state.girls = [];
        state.attendanceData = {};
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

// Guest Sign In


if (DOM.googleSignIn) {
  DOM.googleSignIn.addEventListener('click', async () => {
    if (!firebaseReady || !window._fb) {
      showToast('الإنترنت غير متاح - استخدم وضع عدم الاتصال', 'warning');
      return;
    }
    DOM.googleSignIn.classList.add('is-loading');
    try {
      const { signInWithPopup } = window._fb;
      await signInWithPopup(auth, provider);
    } catch (e) {
      DOM.googleSignIn.classList.remove('is-loading');
      if (['auth/popup-blocked', 'auth/popup-closed-by-user', 'auth/cancelled-popup-request'].includes(e.code)) {
        try {
          const { signInWithRedirect } = window._fb;
          await signInWithRedirect(auth, provider);
        } catch (e2) { showToast('فشل تسجيل الدخول: ' + e2.message, 'error'); }
      } else {
        showToast('فشل تسجيل الدخول: ' + e.message, 'error');
      }
    }
  });
}

if (DOM.signOutBtn) {
  DOM.signOutBtn.addEventListener('click', async () => {
    if (!firebaseReady || !window._fb) {
      state.currentUser = null;
      state.appInitialized = false;
      showLogin();
      return;
    }
    const { signOut } = window._fb;
    await signOut(auth);
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
// FIREBASE LISTENERS
// ============================================================
async function loadData() {
  try {
    if (!firebaseReady || !window._fb) return;

    // Destructure Firestore functions once
    const { onSnapshot: _onSnapshot, query: _query, collection: _collection, orderBy: _orderBy, getDocs: _getDocs, doc: _doc, setDoc: _setDoc, writeBatch: _writeBatch, deleteDoc: _deleteDoc } = window._fb;

    _onSnapshot(_query(_collection(db, 'girls'), _orderBy('name')), snap => {
      let changed = false;
      for (const change of snap.docChanges()) {
        const g = { id: change.doc.id, ...change.doc.data() };
        if (change.type === 'removed' || g.isDeleted) {
          state.girls = state.girls.filter(x => x.id !== g.id);
          changed = true;
        } else {
          const idx = state.girls.findIndex(x => x.id === g.id);
          idx >= 0 ? (state.girls[idx] = g) : state.girls.push(g);
          changed = true;
        }
      }
      state.girls.sort((a, b) => a.name.localeCompare(b.name, 'ar'));
      if (changed) scheduleRender();
    });

    _onSnapshot(_query(_collection(db, 'attendance'), _orderBy('date', 'desc')), snap => {
      let changed = false;
      for (const change of snap.docChanges()) {
        const a = { id: change.doc.id, ...change.doc.data() };
        if (change.type === 'removed') {
          delete state.attendanceData[a.id];
          changed = true;
        } else {
          state.attendanceData[a.id] = a;
          changed = true;
        }
      }
      if (changed) scheduleRender();
    });

    // Listen to history collection and sync to IndexedDB
    _onSnapshot(_query(_collection(db, 'history'), _orderBy('timestamp', 'desc')), async snap => {
      let changed = false;
      for (const change of snap.docChanges()) {
        const log = { id: change.doc.id, ...change.doc.data() };
        if (change.type === 'removed') {
          try { await IDB.delete('history', log.id); } catch (e) { }
          changed = true;
        } else {
          try { await IDB.add('history', log); } catch (e) { }
          changed = true;
        }
      }
      if (changed && state.currentPage === 'history') renderHistory(false);
    });
  } catch (e) { console.error('Load error:', e); }
}

// ============================================================
// RENDER ENGINE
// ============================================================
function scheduleRender() {
  clearTimeout(state.renderTimeout);
  state.renderTimeout = setTimeout(() => renderPage(), 60);
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
  const pageEl = $(`page-${page}`);
  if (!pageEl) {
    console.warn(`Page element not found: page-${page}`);
    return;
  }

  $$('.page').forEach(p => p.classList.remove('active'));
  pageEl.classList.add('active');
  $$('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.page === page));
  $$('.menu-item[data-page]').forEach(b => b.classList.toggle('active', b.dataset.page === page));
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

$$('.nav-btn').forEach(btn => btn.addEventListener('click', () => navigateTo(btn.dataset.page)));
$$('.menu-item[data-page]').forEach(item => item.addEventListener('click', e => {
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
// SMART STATS
// ============================================================
function getBestGradeFiltered(monthStr, gradeFilter) {
  let activeGirls = state.girls.filter(g => !g.isDeleted);
  const [year, month] = monthStr.split('-').map(Number);
  const totalServiceDays = getServiceDaysInMonth(year, month - 1).length || 1;

  const gradeStats = {};
  activeGirls.forEach(g => {
    if (gradeFilter && g.grade !== gradeFilter) return;
    if (!gradeStats[g.grade]) gradeStats[g.grade] = { totalGirls: 0, presentDates: new Set() };
    gradeStats[g.grade].totalGirls++;
  });

  Object.values(state.attendanceData).forEach(a => {
    if (!a.date?.startsWith(monthStr)) return;
    if (a.status !== 'حاضر') return;
    const girl = activeGirls.find(g => g.id === a.girlId);
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
  let activeGirls = state.girls.filter(g => !g.isDeleted);
  const activeGirlIds = gradeFilter ? new Set(activeGirls.filter(g => g.grade === gradeFilter).map(g => g.id)) : new Set(activeGirls.map(g => g.id));
  const counts = {};
  ACTIVITIES.forEach(a => counts[a] = 0);
  Object.values(state.attendanceData).forEach(a => {
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
  const totalServiceDays = getServiceDaysInMonth(year, month - 1).length || 1;

  const presentDatesByGirl = {};
  activeGirls.forEach(g => presentDatesByGirl[g.id] = new Set());

  Object.values(state.attendanceData).forEach(a => {
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
    const girl = activeGirls.find(g => g.id === girlId);
    if (!girl) return;
    if (!best || percent > best.percent || (percent === best.percent && count > best.count)) {
      best = { name: girl.name, count, percent };
    }
  });
  return best;
}

// ============================================================
// HOME PAGE — FIXED: Auto-count absences on service days
// ============================================================
function renderHome() {
  // Use TimeContext for unified date source
  const selectedDate = TimeContext.getDate();
  const now = new Date(selectedDate + 'T00:00:00');
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
  let activeGirls = state.girls.filter(g => !g.isDeleted);
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  const activeGirlIds = new Set(activeGirls.map(g => g.id));

  const allActive = state.girls.filter(g => !g.isDeleted);
  const hfcAll = $('homeFilterCountAll');
  const hfc1 = $('homeFilterCount1');
  const hfc2 = $('homeFilterCount2');
  const hfc3 = $('homeFilterCount3');
  if (hfcAll) hfcAll.textContent = allActive.length;
  if (hfc1) hfc1.textContent = allActive.filter(g => g.grade === 'أولى إعدادي').length;
  if (hfc2) hfc2.textContent = allActive.filter(g => g.grade === 'تانية إعدادي').length;
  if (hfc3) hfc3.textContent = allActive.filter(g => g.grade === 'تالتة إعدادي').length;

  $$('#homeGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  if (DOM.statTotal) DOM.statTotal.textContent = activeGirls.length;

  const presentGirlIds = new Set();
  const absentGirlIds = new Set();
  const todayRecordsByGirl = {};

  // Collect all attendance records for today
  Object.values(state.attendanceData).forEach(a => {
    if (a.date !== dateStr) return;
    if (!activeGirlIds.has(a.girlId)) return;
    if (!todayRecordsByGirl[a.girlId]) todayRecordsByGirl[a.girlId] = [];
    todayRecordsByGirl[a.girlId].push(a);
  });

  // Check each girl's status for today
  activeGirls.forEach(g => {
    const records = todayRecordsByGirl[g.id];
    if (records && records.length > 0) {
      // Girl has attendance records - check if any are present
      const hasAnyPresent = records.some(r => r.status === 'حاضر');
      if (hasAnyPresent) presentGirlIds.add(g.id);
      else absentGirlIds.add(g.id);
    } else if (isService) {
      // Service day with NO records = auto counted as absent
      absentGirlIds.add(g.id);
    }
  });

  if (DOM.statPresentToday) DOM.statPresentToday.textContent = presentGirlIds.size;
  if (DOM.statAbsentToday) DOM.statAbsentToday.textContent = absentGirlIds.size;

  let totalRating = 0, ratingCount = 0;
  Object.values(state.attendanceData).forEach(a => {
    if (a.date?.startsWith(monthStr) && a.rating > 0 && activeGirlIds.has(a.girlId)) {
      totalRating += a.rating; ratingCount++;
    }
  });
  if (DOM.statAvgRating) DOM.statAvgRating.textContent = ratingCount ? (totalRating / ratingCount).toFixed(1) : '-';

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

  const presentDatesByGirl = {};
  activeGirls.forEach(g => presentDatesByGirl[g.id] = new Set());
  Object.values(state.attendanceData).forEach(a => {
    if (a.date?.startsWith(monthStr) && a.status === 'حاضر' && presentDatesByGirl[a.girlId] !== undefined) {
      presentDatesByGirl[a.girlId].add(a.date);
    }
  });
  const counts = {};
  Object.entries(presentDatesByGirl).forEach(([id, dateSet]) => { counts[id] = dateSet.size; });
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 5);

  if (DOM.topAttendees) {
    if (!sorted.length || !sorted[0][1]) {
      DOM.topAttendees.innerHTML = '<div class="empty-state">لا توجد بيانات حضور هذا الشهر</div>';
    } else {
      const frag = document.createDocumentFragment();
      sorted.forEach(([id, count], i) => {
        if (!count) return;
        const g = state.girls.find(x => x.id === id);
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

  const needs = activeGirls.filter(g => {
    const result = hasConsecutiveAbsences(g.id, monthStr);
    return result.hasConsecutive;
  });

  if (DOM.needsFollowup) {
    if (!needs.length) {
      DOM.needsFollowup.innerHTML = '<div class="empty-state">لا توجد حالات تحتاج متابعة</div>';
    } else {
      const frag = document.createDocumentFragment();
      needs.forEach(g => {
        const result = hasConsecutiveAbsences(g.id, monthStr);
        const div = document.createElement('div');
        div.className = 'followup-item';
        div.dataset.girlId = g.id;
        div.innerHTML = `<span class="followup-name">${esc(g.name)}</span><span class="followup-badge">${result.count} غياب متتالي</span>`;
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
// GIRLS PAGE
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

  const gfcAll = $('girlsFilterCountAll');
  const gfc1 = $('girlsFilterCount1');
  const gfc2 = $('girlsFilterCount2');
  const gfc3 = $('girlsFilterCount3');
  if (gfcAll) gfcAll.textContent = activeGirls.length;
  if (gfc1) gfc1.textContent = activeGirls.filter(g => g.grade === 'أولى إعدادي').length;
  if (gfc2) gfc2.textContent = activeGirls.filter(g => g.grade === 'تانية إعدادي').length;
  if (gfc3) gfc3.textContent = activeGirls.filter(g => g.grade === 'تالتة إعدادي').length;

  $$('#girlsGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === filter);
  });

  if (!filtered.length) {
    el.innerHTML = '<div class="empty-state">لا توجد مخدومات<br><small>اضغط + لإضافة مخدومة جديدة</small></div>';
    return;
  }
  const monthStr = TimeContext.getMonth();
  const frag = document.createDocumentFragment();
  filtered.forEach(g => {
    const presents = Object.values(state.attendanceData).filter(a =>
      a.girlId === g.id && a.date?.startsWith(monthStr) && a.status === 'حاضر'
    ).length;
    const absents = Object.values(state.attendanceData).filter(a =>
      a.girlId === g.id && a.date?.startsWith(monthStr) && a.status === 'غائب'
    ).length;
    const div = document.createElement('div');
    div.className = 'girl-card';
    div.dataset.girlId = g.id;
    div.innerHTML = `
      <div class="girl-avatar">${esc(g.name[0])}</div>
      <div class="girl-info">
        <span class="girl-name">${esc(g.name)}</span>
        <span class="girl-grade">${esc(g.grade)}</span>
        ${g.phone ? `<a href="tel:${esc(g.phone)}" class="girl-phone-link" data-phone="${esc(g.phone)}" onclick="event.stopPropagation();">${esc(g.phone)}</a>` : ''}
        <div class="girl-stats"><span class="green-text">&#10003;${presents}</span><span class="red-text">&#10007;${absents}</span></div>
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
  const g = state.girls.find(x => x.id === id);
  if (!g || g.isDeleted) return;
  state.editingGirlId = id;
  if (DOM.girlModalTitle) DOM.girlModalTitle.textContent = 'تعديل بيانات المخدومة';
  if (DOM.girlName) DOM.girlName.value = g.name;
  if (DOM.girlPhone) DOM.girlPhone.value = g.phone || '';
  if (DOM.girlGrade) DOM.girlGrade.value = g.grade;
  if (DOM.girlNotes) DOM.girlNotes.value = g.notes || '';
  if (DOM.deleteGirlBtn) DOM.deleteGirlBtn.classList.remove('hidden');
  openModal('girlModal');
}

if (DOM.deleteGirlBtn) {
  DOM.deleteGirlBtn.addEventListener('click', async () => {
    if (!state.editingGirlId || state.deleteInProgress) return;
    const g = state.girls.find(x => x.id === state.editingGirlId);
    if (!g) return;

    closeModal('girlModal');

    showConfirm({
      icon: '&#9888;', title: 'حذف مخدومة',
      msg: `هل أنت متأكد من حذف "${esc(g.name)}"؟ سيتم حذف جميع بيانات الحضور الخاصة بها أيضاً.`,
      okLabel: 'حذف',
      okClass: 'confirm-delete',
      onOk: async () => {
        if (state.deleteInProgress) return;
        state.deleteInProgress = true;

        try {
          const id = state.editingGirlId;
          state.girls = state.girls.filter(x => x.id !== id);
          const attKeys = Object.keys(state.attendanceData).filter(k => state.attendanceData[k].girlId === id);
          attKeys.forEach(k => delete state.attendanceData[k]);

          try {
            const { setDoc, doc, collection, query, where, getDocs, writeBatch } = window._fb;
            await setDoc(doc(db, 'girls', id), {
              isDeleted: true, deletedAt: Date.now(),
              deletedBy: state.currentUser?.email || '',
              name: g.name, grade: g.grade
            }, { merge: true });

            const attQuery = query(collection(db, 'attendance'), where('girlId', '==', id));
            const attSnap = await getDocs(attQuery);
            if (!attSnap.empty) {
              const docs = attSnap.docs;
              for (let i = 0; i < docs.length; i += 500) {
                const batch = writeBatch(db);
                docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                await batch.commit();
              }
            }
          } catch (e) {
            console.error('Delete girl Firestore error:', e);
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
// SAVE GIRL
// ============================================================
if (DOM.saveGirlBtn) {
  DOM.saveGirlBtn.addEventListener('click', async () => {
    if (state.savingGirl) return;
    state.savingGirl = true;
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
        createdAt: state.editingGirlId ? (state.girls.find(g => g.id === state.editingGirlId)?.createdAt || now) : now,
        updatedAt: now,
        updatedBy: state.currentUser?.displayName || 'خادم',
        updatedByEmail: state.currentUser?.email || '',
        isDeleted: false
      };

      if (state.editingGirlId) {
        state.girls = state.girls.map(g => g.id === id ? girlData : g);
      } else {
        state.girls.push(girlData);
      }

      await logHistory(state.editingGirlId ? 'تعديل مخدومة' : 'إضافة مخدومة', `${name} - ${grade}`);

      if (firebaseReady && window._fb) {
        try { await window._fb.setDoc(window._fb.doc(db, 'girls', id), girlData); }
        catch (e) { console.error('Save girl Firestore error:', e); }
      }

      closeModal('girlModal');
      showToast(state.editingGirlId ? 'تم تعديل البيانات' : 'تمت إضافة المخدومة', 'success');
      state.editingGirlId = null;
      renderPage();
    } finally {
      state.savingGirl = false;
    }
  });
}

// ============================================================
// GIRL PROFILE
// ============================================================
function showGirlProfile(id) {
  const g = state.girls.find(x => x.id === id);
  if (!g || g.isDeleted) return;
  state.currentProfileGirlId = id;
  if (DOM.profileName) DOM.profileName.textContent = g.name;

  const girlAtt = Object.values(state.attendanceData).filter(a => a.girlId === id);
  girlAtt.sort((a, b) => new Date(b.date) - new Date(a.date));

  const totalRecords = girlAtt.length;
  const presentCount = girlAtt.filter(a => a.status === 'حاضر').length;
  const absentCount = girlAtt.filter(a => a.status === 'غائب').length;
  const attendanceRate = totalRecords > 0 ? Math.round((presentCount / totalRecords) * 100) : 0;
  const ratings = girlAtt.filter(a => a.rating > 0).map(a => a.rating);
  const avgRating = ratings.length ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(1) : '0';
  const lastAttendance = girlAtt.find(a => a.status === 'حاضر');
  const lastDate = lastAttendance ? lastAttendance.date : '-';

  const months = {};
  girlAtt.forEach(a => {
    const m = a.date?.substring(0, 7);
    if (!m) return;
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
            return `<div class="profile-record">
              <span class="rec-date">${esc(r.date)} ${esc(DAY_NAMES[new Date(r.date + 'T00:00:00').getDay()] || '')}</span>
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

if (DOM.shareProfileBtn) {
  DOM.shareProfileBtn.addEventListener('click', async () => {
    const id = state.currentProfileGirlId;
    if (!id) return;
    const g = state.girls.find(x => x.id === id);
    if (!g) return;

    const girlAtt = Object.values(state.attendanceData).filter(a => a.girlId === id);
    const presentCount = girlAtt.filter(a => a.status === 'حاضر').length;
    const absentCount = girlAtt.filter(a => a.status === 'غائب').length;
    const attendanceRate = girlAtt.length > 0 ? Math.round((presentCount / girlAtt.length) * 100) : 0;

      const shareText = `👧 ${g.name}
📚 ${g.grade}
\u2705 حضور: ${presentCount}
\u274C غياب: ${absentCount}
\uD83D\uDCCA نسبة: ${attendanceRate}%
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
// ATTENDANCE PAGE — FIXED: Reliable auto-absence on service days
// ============================================================
function getCurrentServiceDay() {
  const dayOfWeek = new Date().getDay();
  const dayMap = { 6: 'السبت', 1: 'الاثنين', 3: 'الاربعاء' };
  return dayMap[dayOfWeek] || null;
}

function isServiceDayDate(dateStr) {
  if (!dateStr) return false;
  const d = new Date(dateStr + 'T00:00:00');
  return SERVICE_DAY_NUMBERS.includes(d.getDay());
}

function renderAttendancePage() {
  if (!DOM.attendanceDate) return;
  // Use TimeContext and always update to the selected date
  DOM.attendanceDate.value = TimeContext.getDate();

  const currentServiceDay = getCurrentServiceDay();
  if (currentServiceDay && !state.attendancePageInitialized) {
    state.selectedDay = currentServiceDay;
  }

  setActiveDay(state.selectedDay);
  setActiveActivity(state.selectedActivity);

  const date = DOM.attendanceDate.value;
  const activeGirls = state.girls.filter(g => !g.isDeleted);

  // Check if any records exist for this date across ALL activities
  const hasAnyRecordsForDate = activeGirls.some(g => {
    return ACTIVITIES.some(act => {
      const key = `${g.id}_${date}_${act}`;
      return state.attendanceData[key];
    });
  });

  // Auto-mark absent on service days if no records exist yet for this date
  if (activeGirls.length > 0 && !hasAnyRecordsForDate && isServiceDayDate(date) && !state.attendancePageInitialized) {
    state.attendancePageInitialized = true;
    markAllAbsentForDate(date);
    return;
  }

  state.attendancePageInitialized = true;
  renderAttendanceList();
}

function setActiveDay(day) {
  state.selectedDay = day;
  $$('.day-btn').forEach(b => b.classList.toggle('active', b.dataset.day === day));
}
function setActiveActivity(act) {
  state.selectedActivity = act;
  $$('.act-tab').forEach(b => b.classList.toggle('active', b.dataset.activity === act));
}

$$('.day-btn').forEach(b => b.addEventListener('click', () => {
  setActiveDay(b.dataset.day);
  state.attendancePageInitialized = false;
  renderAttendancePage();
}));
$$('.act-tab').forEach(b => b.addEventListener('click', () => {
  setActiveActivity(b.dataset.activity);
  state.attendancePageInitialized = false;
  renderAttendancePage();
}));
if (DOM.attendanceDate) {
  DOM.attendanceDate.addEventListener('change', () => {
    // Sync selected date to TimeContext so all pages stay in sync
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

async function toggleAttendanceStatus(girlId, girlName, date) {
  const key = `${girlId}_${date}_${state.selectedActivity}`;
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

  state.attendanceData[key] = rec;

  if (firebaseReady && window._fb) {
    try { await window._fb.setDoc(window._fb.doc(db, 'attendance', key), rec); }
    catch (e) { console.error('Save attendance Firestore error:', e); }
  }

  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'stats') renderStats();
  if (state.currentPage === 'calendar') renderCalendar();
}

// FIXED: Mark all girls as absent for ALL activities on a service day
async function markAllAbsentForDate(date) {
  if (!isServiceDayDate(date)) return;

  const activeGirls = state.girls.filter(g => !g.isDeleted);
  if (activeGirls.length === 0) {
    renderAttendanceList();
    return;
  }

  const batchRecords = [];

  for (const g of activeGirls) {
    for (const activity of ACTIVITIES) {
      const key = `${g.id}_${date}_${activity}`;
      if (!state.attendanceData[key]) {
        const rec = {
          id: key,
          girlId: g.id,
          date,
          day: state.selectedDay,
          activity: activity,
          status: 'غائب',
          rating: 0,
          notes: '',
          updatedAt: Date.now(),
          updatedBy: state.currentUser?.displayName || 'خادم',
          updatedByEmail: state.currentUser?.email || ''
        };
        batchRecords.push(rec);
      }
    }
  }

  for (const rec of batchRecords) {
    state.attendanceData[rec.id] = rec;
  }

  if (firebaseReady && window._fb && batchRecords.length > 0) {
    try {
      const batch = window._fb.writeBatch(db);
      for (const rec of batchRecords) {
        batch.set(window._fb.doc(db, 'attendance', rec.id), rec);
      }
      await batch.commit();
    } catch (e) {
      console.error('Batch save attendance Firestore error:', e);
    }
  }

  if (batchRecords.length > 0) {
    await logHistory('تسجيل حضور', `تعيين الغياب التلقائي ليوم ${date} (${state.selectedDay})`);
    showToast('تم تعيين الغياب التلقائي ليوم خدمة', 'info');
  }

  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'calendar') renderCalendar();
}

// Kept for backward compatibility - delegates to the new function
async function markAllAbsent(date) {
  await markAllAbsentForDate(date);
}

async function selectAllStatus(status) {
  if (!DOM.attendanceDate) return;
  const date = DOM.attendanceDate.value;
  if (!date) { showToast('الرجاء اختيار التاريخ أولاً', 'error'); return; }

  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const batchRecords = [];

  for (const g of activeGirls) {
    const key = `${g.id}_${date}_${state.selectedActivity}`;
    const rec = {
      id: key,
      girlId: g.id,
      date,
      day: state.selectedDay,
      activity: state.selectedActivity,
      status: status,
      rating: status === 'حاضر' ? (state.attendanceData[key]?.rating || 0) : 0,
      notes: state.attendanceData[key]?.notes || '',
      updatedAt: Date.now(),
      updatedBy: state.currentUser?.displayName || 'خادم',
      updatedByEmail: state.currentUser?.email || ''
    };
    batchRecords.push(rec);
    state.attendanceData[key] = rec;
  }

  if (firebaseReady && window._fb && batchRecords.length > 0) {
    try {
      const batch = window._fb.writeBatch(db);
      for (const rec of batchRecords) {
        batch.set(window._fb.doc(db, 'attendance', rec.id), rec);
      }
      await batch.commit();
    } catch (e) {
      console.error('Batch save attendance Firestore error:', e);
    }
  }

  await logHistory('تسجيل حضور', `${status === 'حاضر' ? 'تحديد الكل حاضر' : 'تحديد الكل غائب'} - ${state.selectedActivity} - ${date}`);
  showToast(status === 'حاضر' ? 'تم تحديد الكل حاضر' : 'تم تحديد الكل غائب', 'success');
  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'stats') renderStats();
  if (state.currentPage === 'calendar') renderCalendar();
}

function renderAttendanceList() {
  if (!DOM.attendanceDate || !DOM.attendanceList) return;
  const date = DOM.attendanceDate.value;
  const el = DOM.attendanceList;
  if (!date) { el.innerHTML = '<div class="empty-state">الرجاء اختيار التاريخ</div>'; return; }

  let activeGirls = state.girls.filter(g => !g.isDeleted);
  const searchQuery = DOM.attendanceSearch?.value?.trim() || '';
  if (searchQuery) {
    const qNorm = normalizeArabic(searchQuery);
    activeGirls = activeGirls.filter(g => normalizeArabic(g.name).includes(qNorm));
  }

  let present = 0, absent = 0;
  const frag = document.createDocumentFragment();

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

  activeGirls.forEach(g => {
    const key = `${g.id}_${date}_${state.selectedActivity}`;
    const rec = state.attendanceData[key];
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

    // Interactive inline rating (only for present girls)
    let inlineRatingHtml = '';
    if (statusClass === 'present') {
      inlineRatingHtml = `<div class="att-inline-rating" data-att-key="${esc(key)}">
        <span class="att-inline-rating-label">التقييم:</span>
        <span class="att-inline-stars">
          ${[1,2,3,4,5].map(i => `<span class="att-inline-star ${i <= currentRating ? 'active' : ''}" data-val="${i}" role="button" aria-label="${i} نجمة">&#9733;</span>`).join('')}
        </span>
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
      <span class="att-status-text ${statusClass}">${statusText}</span>
      <button class="att-delete-btn" data-att-key="${esc(key)}" title="حذف السجل">&#10060;</button>`;
    frag.appendChild(div);
  });

  el.innerHTML = '';
  el.appendChild(frag);
  if (DOM.presentCount) DOM.presentCount.textContent = present;
  if (DOM.absentCount) DOM.absentCount.textContent = absent;
  if (DOM.totalCount) DOM.totalCount.textContent = activeGirls.length;
}

async function deleteAttendanceRecord(key) {
  const rec = state.attendanceData[key];
  if (!rec) return;

  const g = state.girls.find(x => x.id === rec.girlId);
  const gName = g ? g.name : 'مخدومة';

  showConfirm({
    icon: '&#9888;', title: 'حذف سجل الحضور',
    msg: `هل أنت متأكد من حذف سجل ${esc(gName)} ليوم ${esc(rec.date)}؟`,
    okLabel: 'حذف',
    onOk: async () => {
      try {
        delete state.attendanceData[key];
        if (firebaseReady && window._fb) {
          try { await window._fb.deleteDoc(window._fb.doc(db, 'attendance', key)); }
          catch (e) { console.error('Delete attendance Firestore error:', e); }
        }
        await logHistory('حذف سجل حضور', `${gName} - ${rec.date} - ${rec.activity} - ${rec.status}`);
        showToast('تم حذف سجل الحضور', 'success');
        renderAttendanceList();
        if (state.currentPage === 'stats') renderStats();
        if (state.currentPage === 'home') renderHome();
        if (state.currentPage === 'calendar') renderCalendar();
      } catch (err) {
        console.error('Delete attendance error:', err);
        showToast('حدث خطأ أثناء الحذف', 'error');
      }
    }
  });
}

// ============================================================
// INLINE RATING — Quick star rating directly in attendance list
// ============================================================
async function saveInlineRating(attKey, rating) {
  const rec = state.attendanceData[attKey];
  if (!rec) return;
  if (rec.status !== 'حاضر') { showToast('التقييم متاح فقط للحاضرات', 'warning'); return; }

  const updatedRec = {
    ...rec,
    rating: rating,
    updatedAt: Date.now(),
    updatedBy: state.currentUser?.displayName || 'خادم',
    updatedByEmail: state.currentUser?.email || ''
  };

  state.attendanceData[attKey] = updatedRec;

  if (firebaseReady && window._fb) {
    try { await window._fb.setDoc(window._fb.doc(db, 'attendance', attKey), updatedRec); }
    catch (e) { console.error('Save inline rating Firestore error:', e); }
  }

  const g = state.girls.find(x => x.id === rec.girlId);
  await logHistory('تقييم مخدومة', `${g?.name || ''} - ${rec.activity} - ${rec.date} - ${rating} نجوم`);
  showToast(`تم التقييم: ${rating} نجوم`, 'success');

  // Refresh the attendance list to show updated stars
  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'stats') renderStats();
}

function openAttendanceEntry(girlId, girlName, date) {
  state.currentAttendanceGirlId = girlId;
  state.currentAttendanceRating = 0;
  if (DOM.attendanceModalTitle) DOM.attendanceModalTitle.textContent = `${state.selectedActivity} - ${date}`;
  if (DOM.modalGirlName) DOM.modalGirlName.textContent = girlName;
  if (DOM.attendanceNotes) DOM.attendanceNotes.value = '';

  const key = `${girlId}_${date}_${state.selectedActivity}`;
  const existing = state.attendanceData[key];
  if (existing) {
    $$('.attend-btn').forEach(b => b.classList.toggle('selected', b.dataset.status === existing.status));
    setRating(existing.rating || 0);
    if (DOM.attendanceNotes) DOM.attendanceNotes.value = existing.notes || '';
    if (DOM.ratingSection) DOM.ratingSection.classList.toggle('hidden', existing.status !== 'حاضر');
  } else {
    $$('.attend-btn').forEach(b => b.classList.remove('selected'));
    setRating(0);
    if (DOM.ratingSection) DOM.ratingSection.classList.add('hidden');
  }
  openModal('attendanceModal');
}

$$('.attend-btn').forEach(b => {
  b.addEventListener('click', () => {
    $$('.attend-btn').forEach(x => x.classList.remove('selected'));
    b.classList.add('selected');
    if (DOM.ratingSection) DOM.ratingSection.classList.toggle('hidden', b.dataset.status !== 'حاضر');
  });
});

$$('.star').forEach(s => s.addEventListener('click', () => setRating(parseInt(s.dataset.val))));
function setRating(val) {
  state.currentAttendanceRating = val;
  $$('.star').forEach(s => s.classList.toggle('active', parseInt(s.dataset.val) <= val));
}

if (DOM.saveAttendanceEntry) {
  DOM.saveAttendanceEntry.addEventListener('click', async () => {
    if (!DOM.attendanceDate) return;
    const date = DOM.attendanceDate.value;
    const statusBtn = document.querySelector('.attend-btn.selected');
    if (!statusBtn) { showToast('الرجاء تحديد الحضور أو الغياب', 'error'); return; }

    const key = `${state.currentAttendanceGirlId}_${date}_${state.selectedActivity}`;
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

    state.attendanceData[key] = rec;

    if (firebaseReady && window._fb) {
      try { await window._fb.setDoc(window._fb.doc(db, 'attendance', key), rec); }
      catch (e) { console.error('Save attendance Firestore error:', e); }
    }

    const gName = state.girls.find(g => g.id === state.currentAttendanceGirlId)?.name || '';
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
// CALENDAR PAGE — Fixed duplicate todayStr
// ============================================================
function renderCalendar() {
  const year = state.calendarDate.getFullYear();
  const month = state.calendarDate.getMonth();
  if (DOM.calMonthYear) DOM.calMonthYear.textContent = state.calendarDate.toLocaleDateString('ar-EG', { year: 'numeric', month: 'long' });

  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const todayStr = TimeContext.getDate();

  let html = '<div class="cal-weekdays">';
  ['أح', 'إث', 'ثل', 'أر', 'خم', 'جم', 'سب'].forEach(d => html += `<div class="cal-wday">${d}</div>`);
  html += '</div><div class="cal-days">';
  for (let i = 0; i < firstDay; i++) html += '<div class="cal-day empty"></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${DateUtil.pad(month + 1)}-${DateUtil.pad(d)}`;
    const dayOfWeek = new Date(year, month, d).getDay();
    const isService = SERVICE_DAY_NUMBERS.includes(dayOfWeek);
    const activeGirlIds = new Set(state.girls.filter(g => !g.isDeleted).map(g => g.id));
    const hasData = Object.values(state.attendanceData).some(a => a.date === dateStr && activeGirlIds.has(a.girlId));
    const isToday = dateStr === todayStr;
    html += `<div class="cal-day ${isService ? 'service-day' : ''} ${hasData ? 'has-data' : ''} ${isToday ? 'today' : ''}" data-date="${dateStr}">
      <span>${d}</span>${isService ? '<div class="service-dot"></div>' : ''}
    </div>`;
  }
  html += '</div>';
  if (DOM.calendarGrid) DOM.calendarGrid.innerHTML = html;

  // Only auto-show today on initial load, not when user navigates
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
  const records = Object.values(state.attendanceData).filter(a => a.date === dateStr);
  const el = DOM.dayDetail;

  const activeGirlIds = new Set(state.girls.filter(g => !g.isDeleted).map(g => g.id));
  const filteredRecords = records.filter(r => activeGirlIds.has(r.girlId));

  if (!filteredRecords.length) {
    el.innerHTML = `<div class="day-detail-header">${dateStr}</div><div class="empty-state">لا توجد سجلات لهذا اليوم</div>`;
  } else {
    const grouped = {};
    filteredRecords.forEach(r => { if (!grouped[r.activity || 'عام']) grouped[r.activity || 'عام'] = []; grouped[r.activity || 'عام'].push(r); });
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
    // Sync calendar navigation to TimeContext
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
    // Sync calendar navigation to TimeContext
    const y = state.calendarDate.getFullYear();
    const m = state.calendarDate.getMonth() + 1;
    const d = parseInt(TimeContext.getDate().split('-')[2]) || 1;
    TimeContext.setDate(`${y}-${String(m).padStart(2, '0')}-${String(Math.min(d, 28)).padStart(2, '0')}`);
    renderCalendar();
  });
}

// ============================================================
// ACTIVITY STATS — FIXED: Show both present AND absence data
// ============================================================
function getPeriodBounds(period, customDate) {
  // Unified: use TimeContext, with optional override
  const selectedDate = customDate || TimeContext.getDate();
  const selYear = parseInt(selectedDate.substring(0, 4));
  const selMonth = parseInt(selectedDate.substring(5, 7));
  switch (period) {
    case 'today': return { start: selectedDate, end: selectedDate };
    case 'month': {
      const lastDay = new Date(selYear, selMonth, 0).getDate();
      return { start: selectedDate.substring(0, 7) + '-01', end: selectedDate.substring(0, 7) + '-' + String(lastDay).padStart(2, '0') };
    }
    case 'year': return { start: selectedDate.substring(0, 4) + '-01-01', end: selectedDate.substring(0, 4) + '-12-31' };
    case 'all': default: return { start: '2000-01-01', end: selectedDate };
  }
}

// FIXED: Returns both present AND absence counts for each activity
function getActivityStats(period, gradeFilter = '', customDate) {
  let activeGirls = state.girls.filter(g => !g.isDeleted);
  const activeGirlIds = gradeFilter
    ? new Set(activeGirls.filter(g => g.grade === gradeFilter).map(g => g.id))
    : new Set(activeGirls.map(g => g.id));
  const { start, end } = getPeriodBounds(period, customDate);

  const stats = {
    'دراسي': { present: 0, absent: 0 },
    'ألحان': { present: 0, absent: 0 },
    'قبطي': { present: 0, absent: 0 },
    'محفوظات': { present: 0, absent: 0 }
  };

  Object.values(state.attendanceData).forEach(a => {
    if (!activeGirlIds.has(a.girlId)) return;
    if (a.date < start || a.date > end) return;
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
// ACTIVITY DETAIL MODAL
// ============================================================
function openActivityDetailModal(activity, period, gradeFilter = '', customDate) {
  const { start, end } = getPeriodBounds(period, customDate);
  let activeGirls = state.girls.filter(g => !g.isDeleted);
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  const activeGirlIds = new Set(activeGirls.map(g => g.id));
  const periodLabel = PERIOD_LABELS[period] || '';

  const records = Object.values(state.attendanceData).filter(a => {
    if (a.activity !== activity) return false;
    if (!activeGirlIds.has(a.girlId)) return false;
    if (a.date < start || a.date > end) return false;
    return true;
  });

  const byGirl = {};
  records.forEach(a => { if (!byGirl[a.girlId]) byGirl[a.girlId] = []; byGirl[a.girlId].push(a); });

  const presentGirls = [];
  const absentGirls = [];

  Object.entries(byGirl).forEach(([girlId, girlRecords]) => {
    girlRecords.sort((a, b) => new Date(b.date) - new Date(a.date));
    const girl = activeGirls.find(g => g.id === girlId);
    if (!girl) return;

    const pCount = girlRecords.filter(r => r.status === 'حاضر').length;
    const aCount = girlRecords.filter(r => r.status === 'غائب').length;
    const total = girlRecords.length;
    const rate = total > 0 ? Math.round((pCount / total) * 100) : 0;

    const entry = { girl, presentCount: pCount, absentCount: aCount, totalRecords: total, attendanceRate: rate, latestRecord: girlRecords[0] };
    if (pCount >= aCount) presentGirls.push(entry);
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

  $$('#activityDetailTabs .activity-detail-tab').forEach(tab => {
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
// ACTIVITY STAT CARDS — FIXED: Show both present and absent
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
  const medals = ['&#129351;', '&#129352;', '&#129353;', '4'];

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
// STATS PAGE
// ============================================================
function renderStats() {
  // Use TimeContext as the single source of truth
  const selectedDate = TimeContext.getDate();
  if (DOM.statsMonth) DOM.statsMonth.value = selectedDate;

  // Unified date bounds from the three interconnected filters
  const { start, end } = getStatsBounds();

  $$('#timeFilterTabs .time-filter-tab').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.period === state.statsTimeFilter);
  });

  const gradeFilter = state.statsGradeFilter;
  $$('#statsGradeFilter .stats-grade-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  let activeGirls = state.girls.filter(g => !g.isDeleted);
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  const activeGirlIds = new Set(activeGirls.map(g => g.id));

  // Filter attendance by unified bounds + grade filter
  const monthAtt = Object.values(state.attendanceData).filter(a =>
    a.date >= start && a.date <= end && activeGirlIds.has(a.girlId)
  );

  const totalSessions = new Set(monthAtt.map(a => a.date)).size;

  // Group records by girlId+date to count per-person-per-day (like home page)
  // A girl is present on a day if ANY activity is حاضر, absent otherwise
  const recordsByGirlDate = {};
  monthAtt.forEach(a => {
    const key = `${a.girlId}_${a.date}`;
    if (!recordsByGirlDate[key]) recordsByGirlDate[key] = { girlId: a.girlId, date: a.date, hasPresent: false, hasAbsent: false };
    if (a.status === 'حاضر') recordsByGirlDate[key].hasPresent = true;
    if (a.status === 'غائب') recordsByGirlDate[key].hasAbsent = true;
  });

  let presents = 0;
  let absents = 0;
  Object.values(recordsByGirlDate).forEach(day => {
    if (day.hasPresent) presents++;
    else if (day.hasAbsent) absents++;
  });

  const ratings = monthAtt.filter(a => a.rating > 0).map(a => a.rating);
  const avgRating = ratings.length ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(1) : '-';

  // Follow-up uses the same unified bounds
  const followupCount = activeGirls.filter(g => {
    const absDates = [...new Set(
      Object.values(state.attendanceData)
        .filter(a => a.girlId === g.id && a.date >= start && a.date <= end && a.status === 'غائب')
        .map(a => a.date)
    )].sort();
    if (absDates.length < 2) return false;
    for (let i = 0; i < absDates.length - 1; i++) {
      const d1 = new Date(absDates[i] + 'T00:00:00');
      const d2 = new Date(absDates[i + 1] + 'T00:00:00');
      const diffDays = (d2 - d1) / (1000 * 60 * 60 * 24);
      if (diffDays <= 3) return true;
    }
    return false;
  }).length;

  const dateLabel = new Date(selectedDate + 'T00:00:00').toLocaleDateString('ar-EG', { month: 'long', day: 'numeric' });

  if (DOM.bigStatsGrid) {
    DOM.bigStatsGrid.innerHTML = `
      <div class="big-stat-card"><div class="big-num">${activeGirls.length}</div><div>المخدومات</div></div>
      <div class="big-stat-card"><div class="big-num">${totalSessions}</div><div>أيام خدمة مسجلة</div></div>
      <div class="big-stat-card green-card"><div class="big-num">${presents}</div><div>إجمالي الحضور</div></div>
      <div class="big-stat-card red-card"><div class="big-num">${absents}</div><div>إجمالي الغياب</div></div>
      <div class="big-stat-card"><div class="big-num">${avgRating}</div><div>متوسط التقييم</div></div>
      <div class="big-stat-card orange-card"><div class="big-num">${followupCount}</div><div>تحتاج متابعة</div></div>`;
  }

  renderActivityStats(state.statsTimeFilter, gradeFilter);

  const gradeLabel = gradeFilter ? `· ${gradeFilter}` : '';
  if (DOM.activityStatsGrade) DOM.activityStatsGrade.textContent = gradeLabel;

  // Absence chart — count unique dates per girl (not individual activity records)
  const absenceByGirl = {};
  activeGirls.forEach(g => absenceByGirl[g.id] = new Set());
  monthAtt.filter(a => a.status === 'غائب').forEach(a => {
    if (absenceByGirl[a.girlId] !== undefined) absenceByGirl[a.girlId].add(a.date);
  });
  // Convert Sets to counts
  Object.keys(absenceByGirl).forEach(id => { absenceByGirl[id] = absenceByGirl[id].size; });
  const maxAbs = Math.max(...Object.values(absenceByGirl), 1);
  const sortedAbs = Object.entries(absenceByGirl).filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]).slice(0, 10);

  if (DOM.absenceChart) {
    DOM.absenceChart.innerHTML = sortedAbs.length
      ? sortedAbs.map(([id, count]) => {
        const g = state.girls.find(x => x.id === id);
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

  // Attendance ranking — count unique dates per girl (not individual activity records)
  const presentsByGirl = {};
  activeGirls.forEach(g => presentsByGirl[g.id] = new Set());
  monthAtt.filter(a => a.status === 'حاضر').forEach(a => {
    if (presentsByGirl[a.girlId] !== undefined) presentsByGirl[a.girlId].add(a.date);
  });
  // Convert Sets to counts
  Object.keys(presentsByGirl).forEach(id => { presentsByGirl[id] = presentsByGirl[id].size; });

  const sortedPresents = Object.entries(presentsByGirl)
    .filter(([, c]) => c > 0)
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

  if (DOM.attendanceRanking) {
    DOM.attendanceRanking.innerHTML = sortedPresents.length
      ? sortedPresents.map(([id, count], i) => {
        const g = state.girls.find(x => x.id === id);
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
// HISTORY PAGE — FIXED: Loads from Firestore + IndexedDB
// ============================================================
async function renderHistory(append = false) {
  const el = DOM.historyList;
  const filter = DOM.historyFilter?.value || '';
  if (!el) return;

  if (!append) {
    el.innerHTML = '<div class="empty-state">جارٍ التحميل...</div>';
    state.historyOffset = 0;

    // Collect logs from all sources
    const allLogs = [];
    const seenIds = new Set();

    // 1. Try Firestore first (online) — get ALL history documents
    if (firebaseReady && window._fb) {
      try {
        const snap = await window._fb.getDocs(
          window._fb.query(window._fb.collection(db, 'history'), window._fb.orderBy('timestamp', 'desc'))
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

    // 2. Also load from IndexedDB (covers offline entries + duplicates)
    try {
      const idbLogs = await IDB.getAll('history');
      idbLogs.forEach(log => {
        if (!seenIds.has(log.id)) {
          seenIds.add(log.id);
          allLogs.push(log);
        }
      });
    } catch (e) { console.warn('IDB history load failed:', e); }

    // Sort newest first
    allLogs.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

    // Apply filter
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
        // Clear IndexedDB
        if (state.idb) await IDB.clear('history');
        state.historyAllLogs = [];
        // Clear Firestore
        if (firebaseReady && window._fb) {
          try {
            const snap = await window._fb.getDocs(window._fb.collection(db, 'history'));
            if (snap.docs.length) {
              for (let i = 0; i < snap.docs.length; i += 500) {
                const batch = window._fb.writeBatch(db);
                snap.docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                await batch.commit();
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
  // Save to IndexedDB first (always works, even offline)
  try { await IDB.add('history', log); } catch (e) { console.warn('IDB history save failed:', e); }
  // Save to Firestore (works when online)
  if (firebaseReady && window._fb) {
    try { await window._fb.setDoc(window._fb.doc(db, 'history', log.id), log); } catch (e) { }
  }
}

// ============================================================
// EXPORT PAGE — FIXED: Day/Month selection with ✓ and X symbols
// ============================================================
function renderExport() {
  if (DOM.exportMonth) DOM.exportMonth.value = TimeContext.getDate();
}

// Sync exportMonth changes to TimeContext
if (DOM.exportMonth) {
  DOM.exportMonth.addEventListener('change', () => {
    if (DOM.exportMonth.value) TimeContext.setDate(DOM.exportMonth.value);
  });
}

// Excel export — FIXED: Daily report = ✓/✗ per activity, Monthly report = numeric counts
if (DOM.exportCSV) {
  DOM.exportCSV.addEventListener('click', () => {
    if (!XLSX) { showToast('مكتبة Excel غير محملة، حاول تحديث الصفحة', 'error'); return; }

    const exportMode = document.querySelector('input[name="exportMode"]:checked')?.value || 'day';
    const exportDate = DOM.exportMonth.value || TimeContext.getDate();

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
      const dayName = DAY_NAMES[new Date(exportDate + 'T00:00:00').getDay()] || '';
      reportTitle = 'تقرير حضور يوم ' + exportDate + ' (' + dayName + ')';
    }

    const activeGirlIds = new Set(state.girls.filter(g => !g.isDeleted).map(g => g.id));
    const exportAtt = Object.values(state.attendanceData).filter(a =>
      a.date >= exportStart && a.date <= exportEnd && activeGirlIds.has(a.girlId)
    );

    const wb = XLSX.utils.book_new();

    if (exportMode === 'month') {
      // ============================================================
      // MONTHLY REPORT: NUMBERS per activity + totals
      // الاسم | السنة | دراسي | قبطي | ألحان | محفوظات | إجمالي الحضور | إجمالي الغياب
      // ============================================================
      const monthName = DateUtil.formatMonth(exportDate.substring(0, 7));
      const wsData = [];
      wsData.push(['تقرير حضور شهر ' + monthName]);
      wsData.push([]);
      wsData.push(['عدد المخدومات', activeGirlIds.size]);
      wsData.push([]);
      // Header row
      wsData.push(['الاسم', 'السنة', 'دراسي', 'قبطي', 'ألحان', 'محفوظات', 'إجمالي الحضور', 'إجمالي الغياب']);

      // Group by girl - aggregate counts per activity
      const grouped = {};
      exportAtt.forEach(a => {
        if (!grouped[a.girlId]) {
          const g = state.girls.find(x => x.id === a.girlId);
          grouped[a.girlId] = {
            name: g?.name || '', grade: g?.grade || '',
            'دراسي': { present: 0, absent: 0 }, 'قبطي': { present: 0, absent: 0 },
            'محفوظات': { present: 0, absent: 0 }, 'ألحان': { present: 0, absent: 0 },
            totalPresent: 0, totalAbsent: 0
          };
        }
        if (a.status === 'حاضر') {
          grouped[a.girlId][a.activity].present++;
          grouped[a.girlId].totalPresent++;
        } else {
          grouped[a.girlId][a.activity].absent++;
          grouped[a.girlId].totalAbsent++;
        }
      });

      const sortedGirls = Object.values(grouped).sort((a, b) => a.name.localeCompare(b.name, 'ar'));

      sortedGirls.forEach(r => {
        wsData.push([
          r.name, r.grade,
          r['دراسي'].present,   // numeric count (not ✓/X)
          r['قبطي'].present,    // numeric count
          r['ألحان'].present,   // numeric count
          r['محفوظات'].present,  // numeric count
          r.totalPresent,        // total present
          r.totalAbsent          // total absent
        ]);
      });

      const ws = XLSX.utils.aoa_to_sheet(wsData);
      ws['!cols'] = [{ wch: 28 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 14 }];
      ws['!dir'] = 'rtl';
      XLSX.utils.book_append_sheet(wb, ws, 'ملخص الشهر');

      // === Sheet 2: Detailed Daily Records ===
      exportAtt.sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        return (a.activity || '').localeCompare(b.activity || '', 'ar');
      });

      const detailData = [];
      detailData.push(['تقرير تفصيلي — ' + monthName]);
      detailData.push([]);
      detailData.push(['التاريخ', 'اليوم', 'المخدومة', 'السنة', 'النشاط', 'الحالة', 'التقييم', 'ملاحظات']);

      exportAtt.forEach(a => {
        const g = state.girls.find(x => x.id === a.girlId);
        const dayName = DAY_NAMES[new Date(a.date + 'T00:00:00').getDay()] || '';
        const stars = a.rating ? '★'.repeat(a.rating) + '☆'.repeat(5 - a.rating) : '';
        detailData.push([a.date, dayName, g?.name || '', g?.grade || '', a.activity || '', a.status === 'حاضر' ? '✓' : '✗', stars, a.notes || '']);
      });

      const wsDetail = XLSX.utils.aoa_to_sheet(detailData);
      wsDetail['!cols'] = [{ wch: 14 }, { wch: 10 }, { wch: 24 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 24 }];
      wsDetail['!dir'] = 'rtl';
      XLSX.utils.book_append_sheet(wb, wsDetail, 'تفاصيل يومية');

    } else {
      // ============================================================
      // DAILY REPORT: ✓/✗ per activity for each girl
      // الاسم | السنة | دراسي | قبطي | ألحان | محفوظات
      // ============================================================
      const wsData = [];
      wsData.push([reportTitle]);
      wsData.push([]);
      // Header: one row per girl, activities as columns
      wsData.push(['الاسم', 'السنة', 'دراسي', 'قبطي', 'ألحان', 'محفوظات']);

      // Build per-girl per-activity status map
      const activeGirls = state.girls.filter(g => !g.isDeleted).sort((a, b) => a.name.localeCompare(b.name, 'ar'));
      activeGirls.forEach(g => {
        const row = [g.name, g.grade];
        ACTIVITIES.forEach(act => {
          const key = `${g.id}_${exportDate}_${act}`;
          const rec = state.attendanceData[key];
          if (rec) {
            row.push(rec.status === 'حاضر' ? '✓' : '✗');
          } else {
            row.push('—');
          }
        });
        wsData.push(row);
      });

      // Summary
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
    a.download = `حضور_${exportDate}${exportMode === 'month' ? '_شهر' : '_يوم'}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast(exportMode === 'month' ? 'تم تصدير ملف Excel للشهر' : 'تم تصدير ملف Excel لليوم', 'success');
  });
}
if (DOM.exportJSON) {
  DOM.exportJSON.addEventListener('click', () => {
    const exportDate = DOM.exportMonth.value || TimeContext.getDate();
    const exportStart = exportDate.substring(0, 7) + '-01';
    const exportEnd = exportDate;
    const activeGirlIds = new Set(state.girls.filter(g => !g.isDeleted).map(g => g.id));
    const exportAtt = Object.values(state.attendanceData).filter(a =>
      a.date >= exportStart && a.date <= exportEnd && activeGirlIds.has(a.girlId)
    );
    const payload = {
      dateRange: { start: exportStart, end: exportEnd },
      girls: state.girls.filter(g => !g.isDeleted),
      attendance: exportAtt,
      exportedAt: new Date().toISOString()
    };
    downloadFile(`بيانات_${exportDate}.json`, JSON.stringify(payload, null, 2), 'application/json');
    showToast('تم تصدير JSON', 'success');
  });
}

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

    const activeGirlIds = new Set(state.girls.filter(g => !g.isDeleted).map(g => g.id));
    const exportAtt = Object.values(state.attendanceData).filter(a =>
      a.date >= exportStart && a.date <= exportEnd && activeGirlIds.has(a.girlId)
    );

    const activeGirls = state.girls.filter(g => !g.isDeleted);
    const totalPresent = exportAtt.filter(a => a.status === 'حاضر').length;
    const totalAbsent = exportAtt.filter(a => a.status === 'غائب').length;

    let html;

    if (exportMode === 'month') {
      // ============================================================
      // MONTHLY PRINT REPORT: Numbers per activity + totals
      // ============================================================
      const monthName = DateUtil.formatMonth(exportDate.substring(0, 7));

      // Group by girl - aggregate counts
      const grouped = {};
      exportAtt.forEach(a => {
        if (!grouped[a.girlId]) {
          const g = state.girls.find(x => x.id === a.girlId);
          grouped[a.girlId] = {
            name: g?.name || '', grade: g?.grade || '',
            'دراسي': { present: 0, absent: 0 }, 'قبطي': { present: 0, absent: 0 },
            'محفوظات': { present: 0, absent: 0 }, 'ألحان': { present: 0, absent: 0 },
            totalPresent: 0, totalAbsent: 0
          };
        }
        if (a.status === 'حاضر') {
          grouped[a.girlId][a.activity].present++;
          grouped[a.girlId].totalPresent++;
        } else {
          grouped[a.girlId][a.activity].absent++;
          grouped[a.girlId].totalAbsent++;
        }
      });

      const sortedGirls = Object.values(grouped).sort((a, b) => a.name.localeCompare(b.name, 'ar'));

      const rows = sortedGirls.map((r, i) => {
        return `<tr>
          <td>${i + 1}</td>
          <td>${esc(r.name)}</td>
          <td>${esc(r.grade)}</td>
          <td>${r['دراسي'].present}</td>
          <td>${r['قبطي'].present}</td>
          <td>${r['ألحان'].present}</td>
          <td>${r['محفوظات'].present}</td>
          <td style="color:green;font-weight:700">${r.totalPresent}</td>
          <td style="color:red;font-weight:700">${r.totalAbsent}</td>
        </tr>`;
      }).join('');

      html = `<!DOCTYPE html><html lang="ar" dir="rtl">
        <head><meta charset="UTF-8"><title>تقرير شهر ${monthName}</title>
        <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
        <style>body{font-family:Tajawal,sans-serif;direction:rtl;padding:20px}
        h1{color:#1a2744;border-bottom:2px solid #1a2744;padding-bottom:10px}
        .summary{display:flex;gap:20px;margin:15px 0;flex-wrap:wrap}
        .sum-box{background:#f0f2f8;border-radius:10px;padding:12px 20px;text-align:center}
        .sum-box b{font-size:24px;color:#1a2744}
        .sum-box span{font-size:13px;color:#6b7a99}
        table{width:100%;border-collapse:collapse;margin-top:20px}
        th,td{border:1px solid #ddd;padding:8px;text-align:center;font-size:13px}
        th{background:#1a2744;color:white}
        .footer{margin-top:20px;font-size:12px;color:#6b7a99;border-top:1px solid #e2e8f0;padding-top:10px}
        @media print{body{padding:10px}}
        </style></head><body>
        <h1>تقرير حضور شهر ${monthName}</h1>
        <p style="color:#6b7a99;font-size:14px">الفترة: من ${exportStart} إلى ${exportEnd}</p>
        <div class="summary">
          <div class="sum-box"><b>${activeGirls.length}</b><br><span>عدد المخدومات</span></div>
          <div class="sum-box"><b>${totalPresent}</b><br><span>إجمالي الحضور</span></div>
          <div class="sum-box"><b>${totalAbsent}</b><br><span>إجمالي الغياب</span></div>
          <div class="sum-box"><b>${sortedGirls.length}</b><br><span>مخدومات مشاركة</span></div>
        </div>
        <table>
          <tr><th>#</th><th>الاسم</th><th>السنة</th><th>دراسي</th><th>قبطي</th><th>ألحان</th><th>محفوظات</th><th>إجمالي الحضور</th><th>إجمالي الغياب</th></tr>
          ${rows}
        </table>
        <div class="footer">تاريخ التصدير: ${new Date().toLocaleDateString('ar-EG')} | نظام متابعة المخدومات</div>
        </body></html>`;

    } else {
      // ============================================================
      // DAILY PRINT REPORT: ✓/✗ per activity for each girl
      // ============================================================
      const dayName = DAY_NAMES[new Date(exportDate + 'T00:00:00').getDay()] || '';

      // Build per-girl per-activity status map
      const sortedGirls = state.girls.filter(g => !g.isDeleted).sort((a, b) => a.name.localeCompare(b.name, 'ar'));

      const rows = sortedGirls.map((g, i) => {
        const cells = [];
        ACTIVITIES.forEach(act => {
          const key = `${g.id}_${exportDate}_${act}`;
          const rec = state.attendanceData[key];
          if (rec) {
            cells.push(rec.status === 'حاضر'
              ? '<td style="color:green;font-weight:700;font-size:16px">✓</td>'
              : '<td style="color:red;font-weight:700;font-size:16px">✗</td>');
          } else {
            cells.push('<td style="color:#ccc">—</td>');
          }
        });
        return `<tr>
          <td>${i + 1}</td>
          <td>${esc(g.name)}</td>
          <td>${esc(g.grade)}</td>
          ${cells.join('')}
        </tr>`;
      }).join('');

      html = `<!DOCTYPE html><html lang="ar" dir="rtl">
        <head><meta charset="UTF-8"><title>تقرير يوم ${exportDate}</title>
        <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
        <style>body{font-family:Tajawal,sans-serif;direction:rtl;padding:20px}
        h1{color:#1a2744;border-bottom:2px solid #1a2744;padding-bottom:10px}
        .summary{display:flex;gap:20px;margin:15px 0;flex-wrap:wrap}
        .sum-box{background:#f0f2f8;border-radius:10px;padding:12px 20px;text-align:center}
        .sum-box b{font-size:24px;color:#1a2744}
        .sum-box span{font-size:13px;color:#6b7a99}
        table{width:100%;border-collapse:collapse;margin-top:20px}
        th,td{border:1px solid #ddd;padding:10px;text-align:center;font-size:14px}
        th{background:#1a2744;color:white}
        .footer{margin-top:20px;font-size:12px;color:#6b7a99;border-top:1px solid #e2e8f0;padding-top:10px}
        @media print{body{padding:10px}}
        </style></head><body>
        <h1>تقرير حضور يوم ${exportDate}</h1>
        <p style="color:#6b7a99;font-size:14px">اليوم: ${dayName}</p>
        <div class="summary">
          <div class="sum-box"><b>${activeGirls.length}</b><br><span>عدد المخدومات</span></div>
          <div class="sum-box"><b>${totalPresent}</b><br><span>حاضر</span></div>
          <div class="sum-box"><b>${totalAbsent}</b><br><span>غائب</span></div>
        </div>
        <table>
          <tr><th>#</th><th>الاسم</th><th>السنة</th><th>دراسي</th><th>قبطي</th><th>ألحان</th><th>محفوظات</th></tr>
          ${rows}
        </table>
        <div class="footer">تاريخ التصدير: ${new Date().toLocaleDateString('ar-EG')} | نظام متابعة المخدومات</div>
        </body></html>`;
    }

    const w = window.open('', '_blank');
    if (!w) { showToast('تم حجب النافذة من المتصفح', 'error'); return; }
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
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ============================================================
// PENDING SYNC
// ============================================================


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

$$('.modal-overlay').forEach(overlay => overlay.addEventListener('click', e => {
  if (e.target === overlay) closeModal(overlay.id);
}));

// ============================================================
// EVENT DELEGATION
// ============================================================
function setupDelegation() {
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
      // Handle inline star rating clicks
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
      const delBtn = e.target.closest('.att-delete-btn');
      if (delBtn) {
        e.stopPropagation();
        e.preventDefault();
        deleteAttendanceRecord(delBtn.dataset.attKey);
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
        const g = state.girls.find(x => x.id === item.dataset.girlId);
        if (g && DOM.attendanceDate) toggleAttendanceStatus(g.id, g.name, DOM.attendanceDate.value);
      }
    });

    DOM.attendanceList.addEventListener('mousedown', e => {
      const item = e.target.closest('.att-item');
      if (!item) return;
      state.isLongPress = false;
      state.longPressTimer = setTimeout(() => {
        state.isLongPress = true;
        const g = state.girls.find(x => x.id === item.dataset.girlId);
        if (g && DOM.attendanceDate) openAttendanceEntry(g.id, g.name, DOM.attendanceDate.value);
      }, 500);
    });
    DOM.attendanceList.addEventListener('mouseup', () => {
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
      setTimeout(() => { state.isLongPress = false; }, 100);
    });
    DOM.attendanceList.addEventListener('mouseleave', () => {
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
    });

    DOM.attendanceList.addEventListener('touchstart', e => {
      const item = e.target.closest('.att-item');
      if (!item) return;
      state.isLongPress = false;
      state.longPressTimer = setTimeout(() => {
        state.isLongPress = true;
        const g = state.girls.find(x => x.id === item.dataset.girlId);
        if (g && DOM.attendanceDate) openAttendanceEntry(g.id, g.name, DOM.attendanceDate.value);
      }, 500);
    }, { passive: true });
    DOM.attendanceList.addEventListener('touchend', () => {
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
      setTimeout(() => { state.isLongPress = false; }, 100);
    });
    DOM.attendanceList.addEventListener('touchcancel', () => {
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

// Subscribe all major render functions to TimeContext date changes
TimeContext.subscribe(() => {
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'girls') renderGirlsList();
  if (state.currentPage === 'stats') renderStats();
  if (state.currentPage === 'export') renderExport();
});

// ============================================================
// BOOTSTRAP — Fixed with proper error handling
// ============================================================
async function bootstrap() {
  initDarkMode();
  TimeContext.init(); // Initialize unified date source

  // Initialize IndexedDB first (always, even without Firebase)
  try {
    await IDB.init();
    state.idb = true;
  } catch (e) {
    console.warn('IndexedDB init failed:', e);
    state.idb = false;
  }

  // Initialize Firebase modules
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
