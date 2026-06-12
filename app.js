// ============================================================
// نظام متابعة المخدومات — REFACTORED v3.0
// Clean Architecture: config → utils → cache → export-engine → app
// Fixed: All 21 bugs from audit (Critical + Logic + Performance + Design)
// ============================================================

import {
  SERVICE_DAYS, SERVICE_DAY_NUMBERS, DAY_NAMES, ACTIVITIES, ACTIVITY_ICONS,
  PERIOD_LABELS, GRADE_ORDER, getGradeOrder, compareGrades, PAGE_TITLES,
  FIREBASE_CONFIG, isValidActivity, getActivityData
} from './config.js';

import {
  esc, xmlEsc, parseDateStr, compareDateStr, isDateInRange,
  DateUtil, makeAttKey, getMonthStr, isServiceDayDate,
  normalizeArabic, normalizeName, isValidDateStr,
  downloadFile, gradeFileSuffix, validateDateStr
} from './utils.js';

import {
  Cache, AttendanceStore, buildAbsenceCache, hasConsecutiveAbsences,
  setCacheStateRefs
} from './cache.js';

import {
  buildExportData, exportToExcel, exportToJSON, exportToPrint
} from './export-engine.js';

// ============================================================
// FB MODULE — Firebase proxy singleton
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
      throw new Error(`FB.${String(prop)} accessed before Firebase initialization.`);
    }
    return target[prop];
  }
});

function ensureFB() {
  if (!firebaseReady) throw new Error('Firebase not initialized');
}

// ============================================================
// GLOBAL ERROR HANDLER + SPLASH STATE
// ============================================================
const SplashState = {
  _done: false, _forceHidden: false, _locked: false,
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
setTimeout(hideSplashForced, 6000);

function hideSplashForced() {
  if (SplashState.done) return;
  SplashState.markForceHidden();
  if (DOM.splash) {
    DOM.splash.classList.add('fade-out');
    setTimeout(() => DOM.splash?.remove(), 500);
  }
  setTimeout(() => {
    if (DOM.loginScreen && DOM.mainApp && DOM.mainApp.classList.contains('hidden') && DOM.loginScreen.classList.contains('hidden')) {
      DOM.loginScreen.classList.remove('hidden');
      showLogin();
    }
  }, 600);
}

// ============================================================
// FIREBASE INIT
// ============================================================
let firebaseApp, auth, db, provider;
let firebaseReady = false;
let XLSX = null;

const _unsubscribers = [];
let _listenersInitialized = false;

function clearAllSnapshots() {
  _unsubscribers.forEach(unsub => { try { unsub(); } catch (e) { } });
  _unsubscribers.length = 0;
  _listenersInitialized = false;
}

function pushUnsubscriber(unsub) { _unsubscribers.push(unsub); }

async function initModules() {
  try {
    const { initializeApp } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js');
    const { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, onAuthStateChanged, signOut } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js');
    const { getFirestore, collection, doc, setDoc, getDocs, deleteDoc, query, orderBy, onSnapshot, writeBatch, where } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js');

    firebaseApp = initializeApp(FIREBASE_CONFIG);
    auth = getAuth(firebaseApp);
    db = getFirestore(firebaseApp);
    provider = new GoogleAuthProvider();
    firebaseReady = true;

    try {
      const { getAnalytics } = await import('https://www.gstatic.com/firebasejs/10.7.1/firebase-analytics.js');
      getAnalytics(firebaseApp);
    } catch (e) { /* analytics optional */ }

    FB.collection = collection; FB.doc = doc; FB.setDoc = setDoc;
    FB.getDocs = getDocs; FB.deleteDoc = deleteDoc; FB.query = query;
    FB.orderBy = orderBy; FB.onSnapshot = onSnapshot; FB.writeBatch = writeBatch;
    FB.where = where; FB.signInWithPopup = signInWithPopup;
    FB.signInWithRedirect = signInWithRedirect; FB.getRedirectResult = getRedirectResult;
    FB.onAuthStateChanged = onAuthStateChanged; FB.signOut = signOut;

    try {
      const xlsxMod = await import('https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs');
      XLSX = xlsxMod;
    } catch (e) { console.warn('XLSX library failed to load:', e); }

    return true;
  } catch (e) {
    console.error('Firebase failed to initialize:', e);
    firebaseReady = false;
    const splashContent = document.querySelector('.splash-content');
    if (splashContent) {
      splashContent.innerHTML = '<h1>⚠️ خطأ في الاتصال</h1><p>تعذر تحميل نظام التسجيل</p><p style="font-size:14px;opacity:0.7">تحقق من اتصال الإنترنت وأعد تحميل الصفحة</p>';
    }
    return false;
  }
}

// ============================================================
// DOM CACHE
// ============================================================
const $ = (id) => document.getElementById(id);

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

const DOM = new Proxy(_domCache, {
  get(target, prop) { return target[prop] ?? null; }
});
_buildDOMCache();

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
  pendingAttendanceOps: new Set(),
  pendingSaveGirl: false,
  exportGradeFilter: 'أولى إعدادي',
};

// Wire cache to state (was: Bug #10 — mismatch between sources)
setCacheStateRefs(state.girls, state.attendanceData);

function setStateGirls(newGirls) {
  state.girls = newGirls;
  setCacheStateRefs(state.girls, state.attendanceData);
  Cache.invalidate();
}

function setStateAttendanceData(newData) {
  if (typeof newData === 'function') {
    state.attendanceData = newData(state.attendanceData);
  } else {
    state.attendanceData = newData;
  }
  setCacheStateRefs(state.girls, state.attendanceData);
  Cache.invalidate();
}

const HISTORY_PAGE_SIZE = 30;

// ============================================================
// THEME
// ============================================================
const Theme = {
  KEY: 'theme',
  init() {
    const saved = localStorage.getItem(this.KEY);
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    this._apply(saved || (prefersDark ? 'dark' : 'light'), false);
  },
  toggle() {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    this._apply(isDark ? 'light' : 'dark', true);
  },
  _apply(theme, animate) {
    if (!animate) document.body.classList.add('theme-switching');
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem(this.KEY, theme);
    if (DOM.darkToggleSwitch) DOM.darkToggleSwitch.classList.toggle('on', theme === 'dark');
    if (!animate) requestAnimationFrame(() => document.body.classList.remove('theme-switching'));
  },
  isDark() { return document.documentElement.getAttribute('data-theme') === 'dark'; }
};

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
  toastTimeout = setTimeout(() => {
    if (DOM.toast) DOM.toast.className = 'toast toast-out';
  }, 3000);
}

// ============================================================
// SPLASH
// ============================================================
function hideSplash() {
  if (SplashState.done) return;
  SplashState.markDone();
  if (DOM.splash) {
    DOM.splash.classList.add('fade-out');
    setTimeout(() => DOM.splash?.remove(), 500);
  }
}

// ============================================================
// INDEXEDDB
// ============================================================
const IDB = {
  db: null,
  DB_NAME: 'girlsTrackerDB',
  DB_VERSION: 2,

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
      const request = tx.objectStore(storeName).put(data);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  },

  async getAll(storeName) {
    if (!this.db) return [];
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readonly');
      const request = tx.objectStore(storeName).getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  },

  async get(storeName, id) {
    if (!this.db) return null;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readonly');
      const request = tx.objectStore(storeName).get(id);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  },

  async clear(storeName) {
    if (!this.db) return;
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, 'readwrite');
      const request = tx.objectStore(storeName).clear();
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  }
};

async function createBackup(operationId, data) {
  try { await IDB.add('backups', { id: operationId, data, timestamp: Date.now() }); } catch (e) { }
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
// TIMECONTEXT
// ============================================================
const TimeContext = {
  _selectedDate: null,
  _listeners: [],

  init() {
    const saved = localStorage.getItem('trackerSelectedDate');
    const today = DateUtil.toStr();
    if (saved && isValidDateStr(saved) && saved === today) {
      this._selectedDate = saved;
    } else {
      this._selectedDate = today;
      localStorage.setItem('trackerSelectedDate', today);
    }
  },

  getDate() { return this._selectedDate || DateUtil.toStr(); },

  setDate(dateStr) {
    if (!isValidDateStr(dateStr)) return;
    this._selectedDate = dateStr;
    localStorage.setItem('trackerSelectedDate', dateStr);
    this._notify(dateStr);
  },

  getMonth() { return getMonthStr(this._selectedDate || DateUtil.toStr()); },
  getYear() { return (this._selectedDate || DateUtil.toStr()).substring(0, 4); },

  resetToToday() {
    this._selectedDate = DateUtil.toStr();
    localStorage.removeItem('trackerSelectedDate');
    this._notify(this._selectedDate);
  },

  subscribe(fn) {
    this._listeners.push(fn);
    return () => { this._listeners = this._listeners.filter(l => l !== fn); };
  },

  _notify(dateStr) {
    this._listeners.forEach(fn => {
      try { fn(dateStr); } catch (e) { console.error('TimeContext listener error:', e); }
    });
  }
};

// ============================================================
// AUTH
// ============================================================
async function initAuth() {
  if (!firebaseReady) { hideSplash(); showLogin(); return; }

  try {
    try { await FB.getRedirectResult(auth); } catch (e) { }

    FB.onAuthStateChanged(auth, async (user) => {
      hideSplash();
      if (!user) {
        state.currentUser = null;
        state.appInitialized = false;
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

if (DOM.googleSignIn) {
  DOM.googleSignIn.addEventListener('click', async () => {
    if (!firebaseReady) { showToast('الإنترنت غير متاح', 'warning'); return; }
    DOM.googleSignIn.classList.add('is-loading');
    try {
      await FB.signInWithPopup(auth, provider);
    } catch (e) {
      DOM.googleSignIn.classList.remove('is-loading');
      if (['auth/popup-blocked', 'auth/popup-closed-by-user', 'auth/cancelled-popup-request'].includes(e.code)) {
        try { await FB.signInWithRedirect(auth, provider); } catch (e2) { showToast('فشل تسجيل الدخول', 'error'); }
      } else {
        showToast('فشل تسجيل الدخول: ' + e.message, 'error');
      }
    }
  });
}

if (DOM.signOutBtn) {
  DOM.signOutBtn.addEventListener('click', async () => {
    clearAllSnapshots();
    if (!firebaseReady) { state.currentUser = null; state.appInitialized = false; showLogin(); return; }
    await FB.signOut(auth);
  });
}

function showApp(user) {
  if (DOM.loginScreen) DOM.loginScreen.classList.add('hidden');
  if (DOM.mainApp) DOM.mainApp.classList.remove('hidden');
  if (DOM.googleSignIn) DOM.googleSignIn.classList.remove('is-loading');
  const initial = user?.displayName?.[0] || 'خ';
  if (DOM.userAvatar) DOM.userAvatar.textContent = initial;
  if (DOM.drawerAvatar) DOM.drawerAvatar.textContent = initial;
  if (DOM.drawerUserName) DOM.drawerUserName.textContent = user?.displayName || 'الخادم';
  if (DOM.drawerUserEmail) DOM.drawerUserEmail.textContent = user?.email || '';
}

function showLogin() {
  if (DOM.loginScreen) DOM.loginScreen.classList.remove('hidden');
  if (DOM.mainApp) DOM.mainApp.classList.add('hidden');
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      const card = document.getElementById('loginCard');
      if (card) {
        card.classList.add('animate-in');
        card.querySelectorAll('.login-cross-icon, .login-church-name, .login-system-title, .login-divider, .login-welcome, .btn-google').forEach(el => el.classList.add('animate-in'));
      }
    });
  });
}

// ============================================================
// DATA LISTENERS
// ============================================================
async function loadData() {
  if (!firebaseReady) return;
  if (_listenersInitialized) return;
  clearAllSnapshots();
  _listenersInitialized = true;

  try {
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
          if (change.type === 'removed') { delete newData[a.id]; changed = true; }
          else { newData[a.id] = a; changed = true; }
        }
        if (changed) { setStateAttendanceData(newData); scheduleRender(); }
      },
      (err) => console.error('Attendance snapshot error:', err)
    );
    pushUnsubscriber(unsub2);

    const unsub3 = FB.onSnapshot(
      FB.query(FB.collection(db, 'history'), FB.orderBy('timestamp', 'desc')),
      (snap) => {
        let changed = false;
        const idbOps = [];
        for (const change of snap.docChanges()) {
          const log = { id: change.doc.id, ...change.doc.data() };
          if (change.type === 'removed') { idbOps.push(IDB.delete('history', log.id).catch(() => {})); changed = true; }
          else { idbOps.push(IDB.add('history', log).catch(() => {})); changed = true; }
        }
        Promise.all(idbOps).catch(() => {});
        if (changed && state.currentPage === 'history') renderHistory(false);
      },
      (err) => console.error('History snapshot error:', err)
    );
    pushUnsubscriber(unsub3);

  } catch (e) {
    console.error('Load error:', e);
    _listenersInitialized = false;
  }
}

// ============================================================
// RENDER ENGINE
// ============================================================
function scheduleRender() {
  if (state.renderPending) return;
  state.renderPending = true;
  requestAnimationFrame(() => {
    queueMicrotask(() => {
      state.renderPending = false;
      renderPage();
    });
  });
}

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
function navigateTo(page) {
  const pageEl = $(`page-${page}`);
  if (!pageEl) return;
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  pageEl.classList.add('active');
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.page === page));
  document.querySelectorAll('.menu-item[data-page]').forEach(b => b.classList.toggle('active', b.dataset.page === page));
  const [title, sub] = PAGE_TITLES[page] || [page, ''];
  if (DOM.pageTitle) DOM.pageTitle.textContent = title;
  if (DOM.pageSubtitle) DOM.pageSubtitle.textContent = sub;
  state.currentPage = page;
  if (page === 'attendance') state.attendancePageInitialized = false;
  if (page !== 'calendar') hideDayDetail();
  renderPage();
  closeDrawer();
}

document.querySelectorAll('.nav-btn').forEach(btn => btn.addEventListener('click', () => navigateTo(btn.dataset.page)));
document.querySelectorAll('.menu-item[data-page]').forEach(item => item.addEventListener('click', e => { e.preventDefault(); navigateTo(item.dataset.page); }));

if (DOM.menuBtn) DOM.menuBtn.addEventListener('click', openDrawer);
if (DOM.drawerOverlay) DOM.drawerOverlay.addEventListener('click', closeDrawer);

function openDrawer() { DOM.drawer?.classList.add('open'); DOM.drawerOverlay?.classList.add('show'); }
function closeDrawer() { DOM.drawer?.classList.remove('open'); DOM.drawerOverlay?.classList.remove('show'); }

// ============================================================
// HOME PAGE
// ============================================================
function getBestGradeFiltered(monthStr, gradeFilter) {
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const [year, month] = monthStr.split('-').map(Number);
  const serviceDays = getServiceDaysInMonth(year, month - 1) || [];
  const totalServiceDays = serviceDays.length || 1;

  const gradeStats = {};
  for (const g of activeGirls) {
    if (gradeFilter && g.grade !== gradeFilter) continue;
    if (!gradeStats[g.grade]) gradeStats[g.grade] = { totalGirls: 0, presentDates: new Set() };
    gradeStats[g.grade].totalGirls++;
  }

  for (const a of Cache.getAllAttendance()) {
    if (!a.date?.startsWith(monthStr)) continue;
    if (a.status !== 'حاضر') continue;
    const girl = Cache.getGirl(a.girlId);
    if (!girl) continue;
    if (gradeFilter && girl.grade !== gradeFilter) continue;
    if (!gradeStats[girl.grade]) continue;
    gradeStats[girl.grade].presentDates.add(a.date + '_' + a.girlId);
  }

  let best = null;
  for (const [grade, data] of Object.entries(gradeStats)) {
    const maxPossible = data.totalGirls * totalServiceDays;
    const percent = maxPossible > 0 ? (data.presentDates.size / maxPossible) * 100 : 0;
    if (!best || percent > best.percent) best = { grade, percent };
  }
  return best;
}

function getTopActivityFiltered(monthStr, gradeFilter) {
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const activeGirlIds = gradeFilter
    ? new Set(activeGirls.filter(g => g.grade === gradeFilter).map(g => g.id))
    : Cache.getActiveGirlIds();

  const counts = {};
  for (const a of ACTIVITIES) counts[a] = 0;

  for (const a of Cache.getAllAttendance()) {
    if (!a.date?.startsWith(monthStr)) continue;
    if (!activeGirlIds.has(a.girlId)) continue;
    if (a.status === 'حاضر' && a.activity in counts) counts[a.activity]++;
  }

  let topName = ACTIVITIES[0], topValue = 0;
  for (const [name, count] of Object.entries(counts)) {
    if (count > topValue) { topName = name; topValue = count; }
  }
  return topValue > 0 ? { name: topName, count: topValue } : null;
}

function getMostRegularGirlFiltered(monthStr, gradeFilter) {
  let activeGirls = state.girls.filter(g => !g.isDeleted);
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  if (!activeGirls.length) return null;

  const [year, month] = monthStr.split('-').map(Number);
  const serviceDays = getServiceDaysInMonth(year, month - 1) || [];
  const totalServiceDays = serviceDays.length || 1;

  const presentDatesByGirl = {};
  for (const g of activeGirls) presentDatesByGirl[g.id] = new Set();

  for (const a of Cache.getAllAttendance()) {
    if (!a.date?.startsWith(monthStr)) continue;
    if (a.status === 'حاضر' && presentDatesByGirl[a.girlId] !== undefined) {
      presentDatesByGirl[a.girlId].add(a.date);
    }
  }

  let best = null;
  for (const [girlId, dateSet] of Object.entries(presentDatesByGirl)) {
    const count = dateSet.size;
    if (count === 0) continue;
    const percent = (count / totalServiceDays) * 100;
    const girl = Cache.getGirl(girlId);
    if (!girl) continue;
    if (!best || percent > best.percent || (percent === best.percent && count > best.count)) {
      best = { name: girl.name, count, percent };
    }
  }
  return best;
}

function renderHome() {
  const selectedDate = TimeContext.getDate();
  const now = parseDateStr(selectedDate);
  const dayName = DateUtil.dayName(now);
  const dateStr = selectedDate;
  const monthStr = TimeContext.getMonth();

  if (DOM.todayDay) DOM.todayDay.textContent = `${DateUtil.formatDateShort(now)} ${dayName}`;
  if (DOM.todayDate) DOM.todayDate.textContent = now.toLocaleDateString('ar-EG', { year: 'numeric', month: 'long', day: 'numeric' });

  const normalized = DateUtil.normalizeDay(dayName);
  const isService = SERVICE_DAYS[normalized];
  if (DOM.todayServiceBadge) {
    DOM.todayServiceBadge.textContent = isService ? 'يوم خدمة \u2713' : 'لا توجد خدمة اليوم';
    DOM.todayServiceBadge.classList.toggle('active', isService);
  }

  const gradeFilter = state.homeGradeFilter;
  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const filteredGirls = gradeFilter ? activeGirls.filter(g => g.grade === gradeFilter) : activeGirls;
  const activeGirlIds = Cache.getActiveGirlIds();
  const filteredGirlIds = gradeFilter ? new Set(filteredGirls.map(g => g.id)) : activeGirlIds;

  // Grade counts — single pass
  const gradeCounts = { 'أولى إعدادي': 0, 'تانية إعدادي': 0, 'تالتة إعدادي': 0 };
  for (const g of activeGirls) {
    if (gradeCounts[g.grade] !== undefined) gradeCounts[g.grade]++;
  }

  const hfcAll = $('homeFilterCountAll'), hfc1 = $('homeFilterCount1'), hfc2 = $('homeFilterCount2'), hfc3 = $('homeFilterCount3');
  if (hfcAll) hfcAll.textContent = activeGirls.length;
  if (hfc1) hfc1.textContent = gradeCounts['أولى إعدادي'];
  if (hfc2) hfc2.textContent = gradeCounts['تانية إعدادي'];
  if (hfc3) hfc3.textContent = gradeCounts['تالتة إعدادي'];

  document.querySelectorAll('#homeGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  if (DOM.statTotal) DOM.statTotal.textContent = filteredGirls.length;

  // Single-pass attendance scan
  const presentGirlIds = new Set();
  const absentGirlIds = new Set();
  const todayRecordsByGirl = {};
  const monthPresentsByGirl = {};
  let totalRating = 0, ratingCount = 0;

  for (const a of Cache.getAllAttendance()) {
    if (a.date === dateStr && filteredGirlIds.has(a.girlId)) {
      if (!todayRecordsByGirl[a.girlId]) todayRecordsByGirl[a.girlId] = [];
      todayRecordsByGirl[a.girlId].push(a);
    }
    if (a.date?.startsWith(monthStr) && filteredGirlIds.has(a.girlId)) {
      if (a.status === 'حاضر') {
        if (!monthPresentsByGirl[a.girlId]) monthPresentsByGirl[a.girlId] = new Set();
        monthPresentsByGirl[a.girlId].add(a.date);
      }
      if (a.rating > 0) { totalRating += a.rating; ratingCount++; }
    }
  }

  for (const g of filteredGirls) {
    const records = todayRecordsByGirl[g.id];
    if (records?.length > 0) {
      records.some(r => r.status === 'حاضر') ? presentGirlIds.add(g.id) : absentGirlIds.add(g.id);
    } else if (isService) {
      absentGirlIds.add(g.id);
    }
  }

  if (DOM.statPresentToday) DOM.statPresentToday.textContent = presentGirlIds.size;
  if (DOM.statAbsentToday) DOM.statAbsentToday.textContent = absentGirlIds.size;
  if (DOM.statAvgRating) DOM.statAvgRating.textContent = ratingCount ? (totalRating / ratingCount).toFixed(1) : '-';

  // Best grade
  const bestGrade = getBestGradeFiltered(monthStr, gradeFilter);
  if (DOM.bestGrade && DOM.bestGradePercent) {
    if (bestGrade?.percent > 0) {
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

  // Top attendees
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
      for (let i = 0; i < sorted.length; i++) {
        const [id, count] = sorted[i];
        const g = Cache.getGirl(id);
        if (!g) continue;
        const div = document.createElement('div');
        div.className = 'top-item';
        div.innerHTML = `<span class="top-rank">${i + 1}</span><span class="top-name">${esc(g.name)}</span><span class="top-count">${count} يوم</span>`;
        frag.appendChild(div);
      }
      DOM.topAttendees.innerHTML = '';
      DOM.topAttendees.appendChild(frag);
    }
  }

  // Needs followup
  const needs = [];
  for (const g of filteredGirls) {
    const result = hasConsecutiveAbsences(g.id, monthStr);
    if (result.hasConsecutive) needs.push({ girl: g, result });
  }

  if (DOM.needsFollowup) {
    if (!needs.length) {
      DOM.needsFollowup.innerHTML = '<div class="empty-state">لا توجد حالات تحتاج متابعة</div>';
    } else {
      const frag = document.createDocumentFragment();
      for (const { girl, result } of needs) {
        const div = document.createElement('div');
        div.className = 'followup-item';
        div.dataset.girlId = girl.id;
        div.innerHTML = `<span class="followup-name">${esc(girl.name)}</span><span class="followup-badge">${result.count} غياب</span>`;
        frag.appendChild(div);
      }
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
    const q = DOM.globalSearch?.value?.trim() || '';
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

  const gfcAll = $('girlsFilterCountAll'), gfc1 = $('girlsFilterCount1'), gfc2 = $('girlsFilterCount2'), gfc3 = $('girlsFilterCount3');
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

  // Precompute attendance counts per girl — single pass
  const monthStr = TimeContext.getMonth();
  const girlStats = {};
  for (const a of Cache.getAllAttendance()) {
    if (!a.date?.startsWith(monthStr)) continue;
    if (!girlStats[a.girlId]) girlStats[a.girlId] = { present: 0, absent: 0 };
    if (a.status === 'حاضر') girlStats[a.girlId].present++;
    else if (a.status === 'غائب') girlStats[a.girlId].absent++;
  }

  const frag = document.createDocumentFragment();
  for (const g of filtered) {
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
  }
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
// DELETE GIRL
// ============================================================
if (DOM.deleteGirlBtn) {
  DOM.deleteGirlBtn.addEventListener('click', async () => {
    if (!state.editingGirlId || state.deleteInProgress) return;
    const currentId = state.editingGirlId;
    const g = Cache.getGirl(currentId);
    if (!g) return;
    closeModal('girlModal');

    showConfirm({
      icon: '&#9888;', title: 'حذف مخدومة',
      msg: `هل أنت متأكد من حذف "${g.name}"؟`,
      okLabel: 'حذف', okClass: 'confirm-delete',
      onOk: async () => {
        if (state.deleteInProgress) return;
        if (state.editingGirlId !== currentId) { showToast('خطأ: تم تغيير المخدومة المحددة', 'error'); return; }
        state.deleteInProgress = true;
        await createBackup('delete_' + currentId + '_' + Date.now(), { girl: g, attendanceData: state.attendanceData });

        try {
          setStateGirls(state.girls.filter(x => x.id !== currentId));
          const newAttData = { ...state.attendanceData };
          for (const key of Object.keys(newAttData)) {
            if (newAttData[key].girlId === currentId) delete newAttData[key];
          }
          setStateAttendanceData(newAttData);

          if (firebaseReady) {
            try {
              await FB.setDoc(FB.doc(db, 'girls', currentId), { isDeleted: true, deletedAt: Date.now(), deletedBy: state.currentUser?.email || '', name: g.name, grade: g.grade }, { merge: true });
              const attSnap = await FB.getDocs(FB.query(FB.collection(db, 'attendance'), FB.where('girlId', '==', currentId)));
              if (!attSnap.empty) {
                const docs = attSnap.docs;
                for (let i = 0; i < docs.length; i += 500) {
                  const batch = FB.writeBatch(db);
                  docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                  await batch.commit();
                }
              }
            } catch (e) { console.error('Delete Firestore error:', e); }
          }
          await logHistory('حذف مخدومة', `${g.name} - ${g.grade}`);
          showToast(`تم حذف ${g.name}`, 'success');
          state.editingGirlId = null;
          scheduleRender();
        } catch (err) {
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
    if (state.savingGirl || state.pendingSaveGirl) return;
    state.savingGirl = true;
    state.pendingSaveGirl = true;
    await createBackup('saveGirl_' + Date.now(), { girls: state.girls, attendanceData: state.attendanceData });

    try {
      const name = DOM.girlName?.value?.trim() || '';
      const phone = DOM.girlPhone?.value?.trim() || '';
      const grade = DOM.girlGrade?.value || '';
      const notes = DOM.girlNotes?.value?.trim() || '';

      if (!name) { showToast('الرجاء إدخال اسم المخدومة', 'error'); return; }
      if (!grade) { showToast('الرجاء اختيار السنة الدراسية', 'error'); return; }

      const normalizedName = normalizeName(name);
      const existingGirl = state.girls.find(g => normalizeName(g.name) === normalizedName && g.id !== state.editingGirlId && !g.isDeleted);
      if (existingGirl) { showToast('هذه المخدومة موجودة بالفعل', 'error'); return; }

      const id = state.editingGirlId || 'girl_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
      const now = Date.now();
      const girlData = { id, name, phone, grade, notes, createdAt: state.editingGirlId ? (Cache.getGirl(id)?.createdAt || now) : now, updatedAt: now, updatedBy: state.currentUser?.displayName || 'خادم', updatedByEmail: state.currentUser?.email || '', isDeleted: false };
      const isNewGirl = !state.editingGirlId;
      const wasEditing = !!state.editingGirlId;

      if (firebaseReady) {
        try { await FB.setDoc(FB.doc(db, 'girls', id), girlData); }
        catch (e) { showToast('فشل الحفظ في السحابة', 'error'); return; }
      }

      if (state.editingGirlId) {
        setStateGirls(state.girls.map(g => g.id === id ? girlData : g));
      } else {
        setStateGirls([...state.girls, girlData]);
      }

      await logHistory(wasEditing ? 'تعديل مخدومة' : 'إضافة مخدومة', `${name} - ${grade}`);
      closeModal('girlModal');
      showToast(wasEditing ? 'تم تعديل البيانات' : 'تمت إضافة المخدومة', 'success');
      state.editingGirlId = null;
      renderPage();
    } catch (err) {
      showToast('حدث خطأ أثناء الحفظ: ' + (err.message || 'خطأ غير معروف'), 'error');
    } finally {
      state.savingGirl = false;
      state.pendingSaveGirl = false;
    }
  });
}

// ============================================================
// GIRL PROFILE
// ============================================================
function showGirlProfile(id) {
  const g = Cache.getGirl(id);
  if (!g) return;
  state.currentProfileGirlId = id;
  if (DOM.profileName) DOM.profileName.textContent = g.name;

  const girlAtt = Cache.getAllAttendance().filter(a => a.girlId === id);
  girlAtt.sort((a, b) => compareDateStr(b.date, a.date));

  const totalRecords = girlAtt.length;
  const presentCount = girlAtt.filter(a => a.status === 'حاضر').length;
  const absentCount = girlAtt.filter(a => a.status === 'غائب').length;
  const attendanceRate = totalRecords > 0 ? Math.round((presentCount / totalRecords) * 100) : 0;
  const ratings = girlAtt.filter(a => a.rating > 0).map(a => a.rating);
  const avgRating = ratings.length ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(1) : '0';

  const sortedAtt = [...girlAtt].sort((a, b) => compareDateStr(a.date, b.date));
  const lastAttendance = [...sortedAtt].reverse().find(a => a.status === 'حاضر');
  const lastDate = lastAttendance ? lastAttendance.date : '-';

  const months = {};
  for (const a of girlAtt) {
    const m = getMonthStr(a.date);
    if (!m) continue;
    if (!months[m]) months[m] = [];
    months[m].push(a);
  }

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
    for (const [month, records] of Object.entries(months).sort((a, b) => b[0].localeCompare(a[0]))) {
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
    }
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
// SHARE PROFILE
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

    const shareText = `${g.name}\n${g.grade}\n[H] حضور: ${presentCount}\n[G] غياب: ${absentCount}\n[%] نسبة: ${attendanceRate}%`.trim();

    if (navigator.share) {
      try { await navigator.share({ title: `ملف ${g.name}`, text: shareText }); } catch (e) { }
    } else {
      try { await navigator.clipboard.writeText(shareText); showToast('تم نسخ البيانات للمشاركة', 'success'); } catch (e) { showToast('المشاركة غير متوفرة', 'warning'); }
    }
  });
}

// ============================================================
// ATTENDANCE PAGE
// ============================================================
function getCurrentServiceDay() {
  const dayMap = { 6: 'السبت', 1: 'الاثنين', 3: 'الاربعاء' };
  return dayMap[new Date().getDay()] || null;
}

function renderAttendancePage() {
  if (!DOM.attendanceDate) return;
  DOM.attendanceDate.value = TimeContext.getDate();
  const currentServiceDay = getCurrentServiceDay();
  if (currentServiceDay && !state.attendancePageInitialized) state.selectedDay = currentServiceDay;
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

document.querySelectorAll('.day-btn').forEach(b => b.addEventListener('click', () => { setActiveDay(b.dataset.day); renderAttendanceList(); }));
document.querySelectorAll('.act-tab').forEach(b => b.addEventListener('click', () => { setActiveActivity(b.dataset.activity); state.attendancePageInitialized = false; renderAttendancePage(); }));
if (DOM.attendanceDate) {
  DOM.attendanceDate.addEventListener('change', () => { TimeContext.setDate(DOM.attendanceDate.value); state.attendancePageInitialized = false; renderAttendancePage(); });
}
if (DOM.selectAllPresent) DOM.selectAllPresent.addEventListener('click', () => selectAllStatus('حاضر'));
if (DOM.selectAllAbsent) DOM.selectAllAbsent.addEventListener('click', () => selectAllStatus('غائب'));

function debouncedAttSearch() {
  clearTimeout(state.attSearchDebounceTimer);
  state.attSearchDebounceTimer = setTimeout(() => renderAttendanceList(), 250);
}
if (DOM.attendanceSearch) DOM.attendanceSearch.addEventListener('input', debouncedAttSearch);

async function toggleAttendanceStatus(girlId, girlName, date) {
  const opKey = `toggle_${girlId}_${date}_${state.selectedActivity}`;
  if (state.pendingAttendanceOps.has(opKey)) return;
  state.pendingAttendanceOps.add(opKey);

  try {
    const key = makeAttKey(girlId, date, state.selectedActivity);
    const existing = state.attendanceData[key];
    const newStatus = existing?.status === 'حاضر' ? 'غائب' : 'حاضر';

    const rec = {
      id: key, girlId, date, day: state.selectedDay,
      activity: state.selectedActivity, status: newStatus,
      rating: newStatus === 'حاضر' ? (existing?.rating || 0) : 0,
      notes: existing?.notes || '', updatedAt: Date.now(),
      updatedBy: state.currentUser?.displayName || 'خادم',
      updatedByEmail: state.currentUser?.email || ''
    };

    setStateAttendanceData(prev => ({ ...prev, [key]: rec }));

    if (firebaseReady) {
      try { await FB.setDoc(FB.doc(db, 'attendance', key), rec); } catch (e) { console.error('Save attendance Firestore error:', e); }
    }
    debouncedRender(80);
  } finally {
    state.pendingAttendanceOps.delete(opKey);
  }
}

async function selectAllStatus(status) {
  if (!DOM.attendanceDate) return;
  const date = DOM.attendanceDate.value;
  if (!date) { showToast('الرجاء اختيار التاريخ أولاً', 'error'); return; }

  const activeGirls = state.girls.filter(g => !g.isDeleted);
  const filteredGirls = state.attendanceGradeFilter ? activeGirls.filter(g => g.grade === state.attendanceGradeFilter) : activeGirls;
  const newAttData = { ...state.attendanceData };
  const currentDateRecords = [];

  for (const g of filteredGirls) {
    const key = makeAttKey(g.id, date, state.selectedActivity);
    const rec = { id: key, girlId: g.id, date, day: DateUtil.dayName(parseDateStr(date)), activity: state.selectedActivity, status, rating: status === 'حاضر' ? (newAttData[key]?.rating || 0) : 0, notes: newAttData[key]?.notes || '', updatedAt: Date.now(), updatedBy: state.currentUser?.displayName || 'خادم', updatedByEmail: state.currentUser?.email || '' };
    newAttData[key] = rec;
    currentDateRecords.push(rec);
  }

  if (firebaseReady) {
    try {
      const batch = FB.writeBatch(db);
      currentDateRecords.forEach(rec => batch.set(FB.doc(db, 'attendance', rec.id), rec));
      await batch.commit();
    } catch (e) { console.error('Batch save error:', e); }
  }

  setStateAttendanceData(newAttData);
  await logHistory('تسجيل حضور', `${status === 'حاضر' ? 'تحديد الكل حاضر' : 'تحديد الكل غائب'} - ${state.selectedActivity} - ${date}`);
  showToast(status === 'حاضر' ? 'تم تحديد الكل حاضر' : 'تم تحديد الكل غائب', 'success');
  renderAttendanceList();
  if (state.currentPage === 'home') renderHome();
  if (state.currentPage === 'stats') renderStats();
  if (state.currentPage === 'calendar') renderCalendar();
}

async function saveInlineRating(attKey, rating) {
  const rec = state.attendanceData[attKey];
  if (!rec || rec.status !== 'حاضر') { showToast('التقييم متاح فقط للحاضرات', 'warning'); return; }
  const opKey = `rating_${attKey}`;
  if (state.pendingAttendanceOps.has(opKey)) return;
  state.pendingAttendanceOps.add(opKey);

  try {
    const updatedRec = { ...rec, rating, updatedAt: Date.now(), updatedBy: state.currentUser?.displayName || 'خادم', updatedByEmail: state.currentUser?.email || '' };
    setStateAttendanceData({ ...state.attendanceData, [attKey]: updatedRec });
    if (firebaseReady) {
      try { await FB.setDoc(FB.doc(db, 'attendance', attKey), updatedRec); } catch (e) { }
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

document.querySelectorAll('.attend-btn').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('.attend-btn').forEach(x => x.classList.remove('selected'));
  b.classList.add('selected');
  if (DOM.ratingSection) DOM.ratingSection.classList.toggle('hidden', b.dataset.status !== 'حاضر');
}));

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

    const key = makeAttKey(state.currentAttendanceGirlId, date, state.selectedActivity);
    const rec = { id: key, girlId: state.currentAttendanceGirlId, date, day: state.selectedDay, activity: state.selectedActivity, status: statusBtn.dataset.status, rating: statusBtn.dataset.status === 'حاضر' ? state.currentAttendanceRating : 0, notes: DOM.attendanceNotes?.value?.trim() || '', updatedAt: Date.now(), updatedBy: state.currentUser?.displayName || 'خادم', updatedByEmail: state.currentUser?.email || '' };

    setStateAttendanceData({ ...state.attendanceData, [key]: rec });
    if (firebaseReady) {
      try { await FB.setDoc(FB.doc(db, 'attendance', key), rec); } catch (e) { }
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

function renderAttendanceList() {
  if (!DOM.attendanceDate || !DOM.attendanceList) return;
  const date = DOM.attendanceDate.value;
  const el = DOM.attendanceList;
  if (!date) { el.innerHTML = '<div class="empty-state">الرجاء اختيار التاريخ</div>'; return; }

  let activeGirls = state.girls.filter(g => !g.isDeleted);
  const gradeFilter = state.attendanceGradeFilter;
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);

  const searchQuery = DOM.attendanceSearch?.value?.trim() || '';
  if (searchQuery) {
    const qNorm = normalizeArabic(searchQuery);
    activeGirls = activeGirls.filter(g => normalizeArabic(g.name).includes(qNorm));
  }

  // Update filter counts
  const allActive = state.girls.filter(g => !g.isDeleted);
  const fcAll = $('attFilterCountAll'), fc1 = $('attFilterCount1'), fc2 = $('attFilterCount2'), fc3 = $('attFilterCount3');
  if (fcAll) fcAll.textContent = allActive.length;
  if (fc1) fc1.textContent = allActive.filter(g => g.grade === 'أولى إعدادي').length;
  if (fc2) fc2.textContent = allActive.filter(g => g.grade === 'تانية إعدادي').length;
  if (fc3) fc3.textContent = allActive.filter(g => g.grade === 'تالتة إعدادي').length;

  document.querySelectorAll('#attendanceGradeFilters .grade-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.grade === gradeFilter);
  });

  if (!activeGirls.length) {
    el.innerHTML = searchQuery ? '<div class="empty-state">لا توجد نتائج للبحث</div>' : '<div class="empty-state">لا توجد مخدومات مسجلة<br><small>أضف مخدومات أولاً من صفحة المخدومات</small></div>';
    if (DOM.presentCount) DOM.presentCount.textContent = 0;
    if (DOM.absentCount) DOM.absentCount.textContent = 0;
    if (DOM.totalCount) DOM.totalCount.textContent = 0;
    return;
  }

  let present = 0, absent = 0;
  const frag = document.createDocumentFragment();

  // Pre-filter attendance by date + activity
  const dateAttendance = {};
  const currentActivity = state.selectedActivity;
  for (const a of Cache.getAllAttendance()) {
    if (a.date === date && a.activity === currentActivity) {
      dateAttendance[makeAttKey(a.girlId, a.date, a.activity)] = a;
    }
  }

  for (const g of activeGirls) {
    const key = makeAttKey(g.id, date, state.selectedActivity);
    const rec = dateAttendance[key];
    let statusClass = 'absent', statusIcon = '&#10007;', statusText = 'غائب';
    if (rec?.status === 'حاضر') { statusClass = 'present'; statusIcon = '&#10003;'; statusText = 'حاضر'; present++; }
    else { absent++; }

    const stars = rec?.rating ? '&#9733;'.repeat(rec.rating) + '&#9734;'.repeat(5 - rec.rating) : '';
    const currentRating = rec?.rating || 0;

    let inlineRatingHtml = '';
    if (statusClass === 'present') {
      let starsHtml = '';
      for (let i = 1; i <= 5; i++) {
        starsHtml += `<span class="att-inline-star ${i <= currentRating ? 'active' : ''}" data-val="${i}" role="button" aria-label="${i} نجمة">&#9733;</span>`;
      }
      inlineRatingHtml = `<div class="att-inline-rating" data-att-key="${esc(key)}"><span class="att-inline-rating-label">التقييم:</span><span class="att-inline-stars">${starsHtml}</span>${currentRating > 0 ? `<span class="att-inline-rating-val">${currentRating}/5</span>` : '<span class="att-inline-rating-hint">اضغط نجمة للتقييم</span>'}</div>`;
    }

    const div = document.createElement('div');
    div.className = `att-item ${statusClass}`;
    div.dataset.girlId = g.id;
    div.dataset.attKey = key;
    div.dataset.girlName = g.name;
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
  }

  el.innerHTML = '';
  el.appendChild(frag);
  if (DOM.presentCount) DOM.presentCount.textContent = present;
  if (DOM.absentCount) DOM.absentCount.textContent = absent;
  if (DOM.totalCount) DOM.totalCount.textContent = activeGirls.length;
}

// ============================================================
// CALENDAR PAGE
// ============================================================
function renderCalendar() {
  const year = state.calendarDate.getFullYear();
  const month = state.calendarDate.getMonth();
  if (DOM.calMonthYear) DOM.calMonthYear.textContent = state.calendarDate.toLocaleDateString('ar-EG', { year: 'numeric', month: 'long' });

  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const todayStr = TimeContext.getDate();
  const currentCalendarActivity = state.selectedActivity || '';

  // Build activity-aware date index
  const dateIndex = new Set();
  const activeGirlIds = Cache.getActiveGirlIds();
  for (const a of Cache.getAllAttendance()) {
    if (activeGirlIds.has(a.girlId)) {
      dateIndex.add(`${a.date}_${a.activity}`);
    }
  }

  let html = '<div class="cal-weekdays">';
  ['أح', 'إث', 'ثل', 'أر', 'خم', 'جم', 'سب'].forEach(d => html += `<div class="cal-wday">${d}</div>`);
  html += '</div><div class="cal-days">';
  for (let i = 0; i < firstDay; i++) html += '<div class="cal-day empty"></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${DateUtil.pad(month + 1)}-${DateUtil.pad(d)}`;
    const dayOfWeek = new Date(year, month, d).getDay();
    const isService = SERVICE_DAY_NUMBERS.includes(dayOfWeek);
    const hasData = dateIndex.has(`${dateStr}_${currentCalendarActivity}`);
    const isToday = dateStr === todayStr;
    html += `<div class="cal-day ${isService ? 'service-day' : ''} ${hasData ? 'has-data' : ''} ${isToday ? 'today' : ''}" data-date="${dateStr}"><span>${d}</span>${isService ? '<div class="service-dot"></div>' : ''}</div>`;
  }
  html += '</div>';
  if (DOM.calendarGrid) DOM.calendarGrid.innerHTML = html;

  const now = new Date();
  if (year === now.getFullYear() && month === now.getMonth() && !currentDayDetailDate) {
    currentDayDetailDate = todayStr;
    refreshDayDetail();
  } else if (currentDayDetailDate) {
    refreshDayDetail();
  }
}

let currentDayDetailDate = null;

function showDayDetail(dateStr) { currentDayDetailDate = dateStr; refreshDayDetail(); }
function hideDayDetail() { currentDayDetailDate = null; if (DOM.dayDetail) DOM.dayDetail.classList.remove('show'); }

function refreshDayDetail() {
  if (!currentDayDetailDate || !DOM.dayDetail) return;
  const activeGirlIds = Cache.getActiveGirlIds();
  const dayRecords = Cache.getAttendanceByDate(currentDayDetailDate).filter(r => activeGirlIds.has(r.girlId));
  const el = DOM.dayDetail;

  if (!dayRecords.length) {
    el.innerHTML = `<div class="day-detail-header">${currentDayDetailDate}</div><div class="empty-state">لا توجد سجلات لهذا اليوم</div>`;
  } else {
    const grouped = {};
    for (const r of dayRecords) {
      const act = r.activity || 'عام';
      if (!grouped[act]) grouped[act] = [];
      grouped[act].push(r);
    }
    let html = `<div class="day-detail-header">${currentDayDetailDate}</div>`;
    for (const [act, recs] of Object.entries(grouped)) {
      const presentCount = recs.filter(r => r.status === 'حاضر').length;
      const absentCount = recs.filter(r => r.status === 'غائب').length;
      html += `<div class="day-activity"><b>${esc(act)}</b>: <span class="green-text">${presentCount} حاضر</span> \u00B7 <span class="red-text">${absentCount} غائب</span> من ${recs.length}</div>`;
    }
    el.innerHTML = html;
  }
  el.classList.add('show');
}

if (DOM.calPrev) {
  DOM.calPrev.addEventListener('click', () => {
    hideDayDetail();
    state.calendarDate.setMonth(state.calendarDate.getMonth() - 1);
    const y = state.calendarDate.getFullYear(), m = state.calendarDate.getMonth() + 1;
    const d = parseInt(TimeContext.getDate().split('-')[2]) || 1;
    const daysInNewMonth = new Date(y, m, 0).getDate();
    TimeContext.setDate(`${y}-${String(m).padStart(2, '0')}-${String(Math.min(d, daysInNewMonth)).padStart(2, '0')}`);
    renderCalendar();
  });
}
if (DOM.calNext) {
  DOM.calNext.addEventListener('click', () => {
    hideDayDetail();
    state.calendarDate.setMonth(state.calendarDate.getMonth() + 1);
    const y = state.calendarDate.getFullYear(), m = state.calendarDate.getMonth() + 1;
    const d = parseInt(TimeContext.getDate().split('-')[2]) || 1;
    const daysInNewMonth = new Date(y, m, 0).getDate();
    TimeContext.setDate(`${y}-${String(m).padStart(2, '0')}-${String(Math.min(d, daysInNewMonth)).padStart(2, '0')}`);
    renderCalendar();
  });
}

// ============================================================
// STATS PAGE
// ============================================================
function getStatsBounds() {
  const selectedDate = validateDateStr(TimeContext.getDate(), DateUtil.toStr());
  const selYear = parseInt(selectedDate.substring(0, 4));
  const selMonth = parseInt(selectedDate.substring(5, 7));

  switch (state.statsTimeFilter) {
    case 'today': return { start: selectedDate, end: selectedDate };
    case 'month': {
      const lastDay = new Date(selYear, selMonth, 0).getDate();
      return { start: selectedDate.substring(0, 7) + '-01', end: selectedDate.substring(0, 7) + '-' + String(lastDay).padStart(2, '0') };
    }
    case 'year': return { start: selectedDate.substring(0, 4) + '-01-01', end: selectedDate.substring(0, 4) + '-12-31' };
    default: return { start: '2000-01-01', end: selectedDate };
  }
}

function renderStats() {
  const selectedDate = TimeContext.getDate();
  if (DOM.statsMonth) DOM.statsMonth.value = selectedDate;

  const { start, end } = getStatsBounds();
  document.querySelectorAll('#timeFilterTabs .time-filter-tab').forEach(btn => btn.classList.toggle('active', btn.dataset.period === state.statsTimeFilter));

  const gradeFilter = state.statsGradeFilter;
  document.querySelectorAll('#statsGradeFilter .stats-grade-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.grade === gradeFilter));

  let activeGirls = state.girls.filter(g => !g.isDeleted);
  if (gradeFilter) activeGirls = activeGirls.filter(g => g.grade === gradeFilter);
  const activeGirlIds = new Set(activeGirls.map(g => g.id));

  // Single-pass scan
  const monthAtt = [];
  const recordsByGirlDate = {};
  let ratingSum = 0, ratingCount = 0;
  const uniqueDates = new Set();

  for (const g of activeGirls) {
    recordsByGirlDate[g.id] = { girlId: g.id, hasPresent: false, hasAbsent: false };
  }

  for (const a of Cache.getAllAttendance()) {
    if (!isDateInRange(a.date, start, end)) continue;
    if (!activeGirlIds.has(a.girlId)) continue;
    monthAtt.push(a);
    uniqueDates.add(a.date);
    if (a.status === 'حاضر') { recordsByGirlDate[a.girlId].hasPresent = true; }
    if (a.status === 'غائب') { recordsByGirlDate[a.girlId].hasAbsent = true; }
    if (a.rating > 0) { ratingSum += a.rating; ratingCount++; }
  }

  let presents = 0, absents = 0;
  for (const day of Object.values(recordsByGirlDate)) {
    if (day.hasPresent) presents++;
    else if (day.hasAbsent) absents++;
  }

  const avgRating = ratingCount > 0 ? (ratingSum / ratingCount).toFixed(1) : '-';

  let followupCount = 0;
  for (const g of activeGirls) {
    if (hasConsecutiveAbsences(g.id, TimeContext.getMonth()).hasConsecutive) followupCount++;
  }

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
  if (DOM.activityStatsGrade) DOM.activityStatsGrade.textContent = gradeFilter ? `· ${gradeFilter}` : '';

  // Absence chart
  const absenceCounts = {};
  for (const a of monthAtt) {
    if (a.status === 'غائب') {
      absenceCounts[a.girlId] = (absenceCounts[a.girlId] || 0) + 1;
    }
  }
  let maxAbs = 1;
  for (const v of Object.values(absenceCounts)) { if (v > maxAbs) maxAbs = v; }
  const sortedAbs = Object.entries(absenceCounts).filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]).slice(0, 10);

  if (DOM.absenceChart) {
    DOM.absenceChart.innerHTML = sortedAbs.length
      ? sortedAbs.map(([id, count]) => {
        const g = Cache.getGirl(id);
        if (!g) return '';
        return `<div class="chart-row"><span class="chart-name">${esc(g.name)}</span><div class="chart-bar-wrap"><div class="chart-bar" style="width:${Math.round((count / maxAbs) * 100)}%"></div></div><span class="chart-val">${count}</span></div>`;
      }).join('')
      : `<div class="empty-state">لا توجد غيابات حتى ${dateLabel} \u127881;</div>`;
  }

  // Attendance ranking
  const presentCounts = {};
  for (const a of monthAtt) {
    if (a.status === 'حاضر') presentCounts[a.girlId] = (presentCounts[a.girlId] || 0) + 1;
  }
  const sortedPresents = Object.entries(presentCounts).filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]);

  if (DOM.attendanceRanking) {
    DOM.attendanceRanking.innerHTML = sortedPresents.length
      ? sortedPresents.map(([id, count], i) => {
        const g = Cache.getGirl(id);
        if (!g) return '';
        return `<div class="rank-item"><span class="rank-num">${i + 1}</span><span class="rank-name">${esc(g.name)}</span><span class="rank-grade">${esc(g.grade)}</span><span class="rank-count">${count} يوم</span></div>`;
      }).join('')
      : `<div class="empty-state">لا توجد بيانات حضور حتى ${dateLabel}</div>`;
  }
}

if (DOM.statsMonth) DOM.statsMonth.addEventListener('change', () => { TimeContext.setDate(DOM.statsMonth.value); renderStats(); });
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
// ACTIVITY STATS
// ============================================================
function getActivityStats(period, gradeFilter = '') {
  const activeGirlIds = gradeFilter
    ? new Set(state.girls.filter(g => !g.isDeleted && g.grade === gradeFilter).map(g => g.id))
    : Cache.getActiveGirlIds();

  const { start, end } = getPeriodBounds(period);
  const stats = { 'دراسي': { present: 0, absent: 0 }, 'ألحان': { present: 0, absent: 0 }, 'قبطي': { present: 0, absent: 0 }, 'محفوظات': { present: 0, absent: 0 } };

  for (const a of Cache.getAllAttendance()) {
    if (!activeGirlIds.has(a.girlId)) continue;
    if (!isDateInRange(a.date, start, end)) continue;
    if (a.activity in stats) {
      if (a.status === 'حاضر') stats[a.activity].present++;
      else if (a.status === 'غائب') stats[a.activity].absent++;
    }
  }

  return Object.entries(stats).filter(([, data]) => data.present > 0 || data.absent > 0).sort((a, b) => (b[1].present + b[1].absent) - (a[1].present + a[1].absent));
}

function getPeriodBounds(period, customDate) {
  const selectedDate = validateDateStr(customDate || TimeContext.getDate(), DateUtil.toStr());
  const selYear = parseInt(selectedDate.substring(0, 4));
  const selMonth = parseInt(selectedDate.substring(5, 7));
  switch (period) {
    case 'today': return { start: selectedDate, end: selectedDate };
    case 'month': {
      const lastDay = new Date(selYear, selMonth, 0).getDate();
      return { start: selectedDate.substring(0, 7) + '-01', end: selectedDate.substring(0, 7) + '-' + String(lastDay).padStart(2, '0') };
    }
    case 'year': return { start: selectedDate.substring(0, 4) + '-01-01', end: selectedDate.substring(0, 4) + '-12-31' };
    default: return { start: '2000-01-01', end: selectedDate };
  }
}

function renderActivityStats(period, gradeFilter = '') {
  const stats = getActivityStats(period, gradeFilter);
  const el = DOM.activityStatsGrid;
  if (!el) return;
  if (!stats.length) { el.innerHTML = '<div class="empty-state" style="grid-column:1/-1;">لا توجد بيانات حضور للفترة المحددة</div>'; return; }

  const medals = ['&#129351;', '&#129352;', '&#129353;', '&#127941;'];
  el.innerHTML = stats.map(([activity, data], i) => `
    <div class="activity-stat-card" data-activity="${esc(activity)}" role="button" tabindex="0">
      <div class="activity-stat-rank">${medals[i] || (i + 1)}</div>
      <div class="activity-stat-icon">${ACTIVITY_ICONS[activity] || '&#128202;'}</div>
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
    if (!card?.dataset.activity) return;
    const selectedDate = DOM.statsMonth?.value || DateUtil.toStr();
    openActivityDetailModal(card.dataset.activity, state.statsTimeFilter, state.statsGradeFilter, selectedDate);
  });
}

// ============================================================
// ACTIVITY DETAIL MODAL
// ============================================================
function openActivityDetailModal(activity, period, gradeFilter = '', customDate) {
  const { start, end } = getPeriodBounds(period, customDate);
  const activeGirlIds = gradeFilter
    ? new Set(state.girls.filter(g => !g.isDeleted && g.grade === gradeFilter).map(g => g.id))
    : Cache.getActiveGirlIds();

  const records = Cache.getAllAttendance().filter(a =>
    a.activity === activity && activeGirlIds.has(a.girlId) && isDateInRange(a.date, start, end)
  );

  const byGirl = {};
  for (const a of records) { if (!byGirl[a.girlId]) byGirl[a.girlId] = []; byGirl[a.girlId].push(a); }

  const presentGirls = [], absentGirls = [];
  for (const [girlId, girlRecords] of Object.entries(byGirl)) {
    girlRecords.sort((a, b) => compareDateStr(b.date, a.date));
    const girl = Cache.getGirl(girlId);
    if (!girl) continue;
    const pCount = girlRecords.filter(r => r.status === 'حاضر').length;
    const aCount = girlRecords.filter(r => r.status === 'غائب').length;
    const total = girlRecords.length;
    const rate = total > 0 ? Math.round((pCount / total) * 100) : 0;
    const entry = { girl, presentCount: pCount, absentCount: aCount, totalRecords: total, attendanceRate: rate, latestRecord: girlRecords[0] };
    if (total > 0 && rate >= 50) presentGirls.push(entry);
    else absentGirls.push(entry);
  }

  presentGirls.sort((a, b) => b.attendanceRate - a.attendanceRate || a.girl.name.localeCompare(b.girl.name, 'ar'));
  absentGirls.sort((a, b) => b.attendanceRate - a.attendanceRate || a.girl.name.localeCompare(b.girl.name, 'ar'));

  state.currentActivityDetail = { activity, period, presentGirls, absentGirls };
  state.activityDetailTab = 'present';

  if (DOM.activityDetailTitle) DOM.activityDetailTitle.textContent = `تفاصيل ${activity}`;
  if (DOM.activityDetailIcon) DOM.activityDetailIcon.innerHTML = ACTIVITY_ICONS[activity] || '&#128202;';
  if (DOM.activityDetailName) DOM.activityDetailName.textContent = activity;
  if (DOM.activityDetailPeriod) DOM.activityDetailPeriod.textContent = PERIOD_LABELS[period] || '';
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
    el.innerHTML = `<div class="empty-state">${isPresentTab ? 'لا يوجد حاضرون للفترة المحددة' : 'لا يوجد غائبون للفترة المحددة'}</div>`;
    return;
  }

  const frag = document.createDocumentFragment();
  for (const { girl, presentCount, absentCount, totalRecords, attendanceRate, latestRecord } of list) {
    const div = document.createElement('div');
    div.className = 'detail-girl-item';
    div.dataset.girlId = girl.id;
    div.innerHTML = `
      <div class="detail-girl-avatar">${esc(girl.name[0])}</div>
      <div class="detail-girl-info">
        <div class="detail-girl-name">${esc(girl.name)}</div>
        <div class="detail-girl-grade">${esc(girl.grade)} \u00B7 ${presentCount} حضور \u00B7 ${absentCount} غياب \u00B7 ${attendanceRate}% نسبة \u00B7 آخر: ${esc(latestRecord.date)}</div>
      </div>
      <div class="detail-status-icon ${isPresentTab ? 'present' : 'absent'}">${isPresentTab ? '&#10003;' : '&#10007;'}</div>`;
    frag.appendChild(div);
  }
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
    if (item?.dataset.girlId) { closeModal('activityDetailModal'); showGirlProfile(item.dataset.girlId); }
  });
}
if (DOM.closeActivityDetailModal) DOM.closeActivityDetailModal.addEventListener('click', () => closeModal('activityDetailModal'));

// ============================================================
// HISTORY PAGE
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

    if (firebaseReady) {
      try {
        const snap = await FB.getDocs(FB.query(FB.collection(db, 'history'), FB.orderBy('timestamp', 'desc')));
        snap.docs.forEach(d => {
          const log = { id: d.id, ...d.data() };
          if (!seenIds.has(log.id)) { seenIds.add(log.id); allLogs.push(log); }
        });
      } catch (e) { }
    }

    try {
      const idbLogs = await IDB.getAll('history');
      idbLogs.forEach(log => { if (!seenIds.has(log.id)) { seenIds.add(log.id); allLogs.push(log); } });
    } catch (e) { }

    allLogs.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
    state.historyAllLogs = filter ? allLogs.filter(l => l.action?.includes(filter)) : allLogs;
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
        <span class="history-meta">${esc(log.by || 'خادم')} \u00B7 ${new Date(log.timestamp).toLocaleString('ar-EG')}</span>
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
                const batch = FB.writeBatch(db);
                snap.docs.slice(i, i + 500).forEach(d => batch.delete(d.ref));
                await batch.commit();
              }
            }
          } catch (e) { }
        }
        showToast('تم مسح السجل التاريخي', 'success');
        renderHistory(false);
      }
    });
  });
}

function getHistoryIcon(action) {
  if (action?.includes('إضافة')) return '&#10133;';
  if (action?.includes('تعديل')) return '&#9999;';
  if (action?.includes('حذف')) return '&#10060;';
  if (action?.includes('حضور')) return '&#128203;';
  return '&#128221;';
}

async function logHistory(action, detail) {
  const log = { id: 'log_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7), action, detail, by: state.currentUser?.displayName || 'خادم', byEmail: state.currentUser?.email || '', timestamp: Date.now() };
  try { await IDB.add('history', log); } catch (e) { }
  if (firebaseReady) { try { await FB.setDoc(FB.doc(db, 'history', log.id), log); } catch (e) { } }
}

// ============================================================
// EXPORT PAGE — Uses unified export-engine (was: 3x duplicated code)
// ============================================================
function renderExport() {
  if (DOM.exportMonth) DOM.exportMonth.value = TimeContext.getDate();
}

if (DOM.exportMonth) {
  DOM.exportMonth.addEventListener('change', () => { if (DOM.exportMonth.value) TimeContext.setDate(DOM.exportMonth.value); });
}

// Export grade filter — unified with state (was: Bug #5 inconsistency)
if (DOM.exportGradeFilter) {
  DOM.exportGradeFilter.addEventListener('click', e => {
    const btn = e.target.closest('.export-grade-btn');
    if (!btn) return;
    state.exportGradeFilter = btn.dataset.grade;
    document.querySelectorAll('#exportGradeFilter .export-grade-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.grade === state.exportGradeFilter);
    });
  });
}

/** Get filtered + sorted girls for export — single source (was: duplicated ×3) */
function getExportGirls() {
  let girls = state.girls.filter(g => !g.isDeleted);
  if (state.exportGradeFilter) girls = girls.filter(g => g.grade === state.exportGradeFilter);
  return [...girls].sort((a, b) => compareGrades(a.grade, b.grade) || a.name.localeCompare(b.name, 'ar'));
}

/** Get export date range */
function getExportRange(exportMode, exportDate) {
  if (exportMode === 'month') {
    const [year, month] = getMonthStr(exportDate).split('-').map(Number);
    const daysInMonth = new Date(year, month, 0).getDate();
    return {
      start: getMonthStr(exportDate) + '-01',
      end: getMonthStr(exportDate) + '-' + String(daysInMonth).padStart(2, '0')
    };
  }
  return { start: exportDate, end: exportDate };
}

// ---- Excel Export ----
if (DOM.exportCSV) {
  DOM.exportCSV.addEventListener('click', () => {
    try {
      const exportMode = document.querySelector('input[name="exportMode"]:checked')?.value || 'day';
      const exportDate = DOM.exportMonth.value || TimeContext.getDate();
      const exportGirls = getExportGirls();
      const { start, end } = getExportRange(exportMode, exportDate);
      const activeGirlIds = new Set(exportGirls.map(g => g.id));
      const exportAtt = Cache.getAllAttendance().filter(a => isDateInRange(a.date, start, end) && activeGirlIds.has(a.girlId));

      // ONE data build for ALL formats (was: 3 separate builds)
      const data = buildExportData({ girls: exportGirls, attendance: exportAtt, exportDate, exportMode, gradeFilter: state.exportGradeFilter });

      exportToExcel(data, { exportDate, exportMode, gradeFilter: state.exportGradeFilter });
      showToast(exportMode === 'month' ? 'تم تصدير ملف Excel للشهر' : 'تم تصدير ملف Excel لليوم', 'success');
    } catch (err) {
      showToast(err.message || 'فشل التصدير', 'error');
    }
  });
}

// ---- JSON Export ----
if (DOM.exportJSON) {
  DOM.exportJSON.addEventListener('click', () => {
    try {
      const exportDate = DOM.exportMonth.value || TimeContext.getDate();
      const { start } = getExportRange('month', exportDate);
      const exportGirls = getExportGirls();
      const activeGirlIds = new Set(exportGirls.map(g => g.id));
      const exportAtt = Cache.getAllAttendance().filter(a => isDateInRange(a.date, start, exportDate) && activeGirlIds.has(a.girlId));

      const data = buildExportData({ girls: exportGirls, attendance: exportAtt, exportDate, exportMode: 'month', gradeFilter: state.exportGradeFilter });
      exportToJSON(data, { exportDate, gradeFilter: state.exportGradeFilter });
      showToast('تم تصدير JSON', 'success');
    } catch (err) {
      showToast(err.message || 'فشل التصدير', 'error');
    }
  });
}

// ---- Print Export ----
if (DOM.exportPrint) {
  DOM.exportPrint.addEventListener('click', () => {
    try {
      const exportMode = document.querySelector('input[name="exportMode"]:checked')?.value || 'day';
      const exportDate = DOM.exportMonth.value || TimeContext.getDate();
      const exportGirls = getExportGirls();
      const { start, end } = getExportRange(exportMode, exportDate);
      const activeGirlIds = new Set(exportGirls.map(g => g.id));
      const exportAtt = Cache.getAllAttendance().filter(a => isDateInRange(a.date, start, end) && activeGirlIds.has(a.girlId));

      const data = buildExportData({ girls: exportGirls, attendance: exportAtt, exportDate, exportMode, gradeFilter: state.exportGradeFilter });
      const result = exportToPrint(data, { exportDate, exportMode, gradeFilter: state.exportGradeFilter });

      if (result === 'popup_blocked') {
        showToast('تم حفظ التقرير كملف HTML (تم حجب النافذة المنبثقة)', 'success');
      } else {
        showToast('تم فتح نافذة الطباعة', 'success');
      }
    } catch (err) {
      showToast(err.message || 'فشل التصدير', 'error');
    }
  });
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
  DOM.confirmCancel.addEventListener('click', () => { DOM.confirmOverlay?.classList.remove('show'); confirmResolve = null; });
}
if (DOM.confirmOverlay) {
  DOM.confirmOverlay.addEventListener('click', e => { if (e.target === DOM.confirmOverlay) { DOM.confirmOverlay.classList.remove('show'); confirmResolve = null; } });
}

if (DOM.closeGirlModal) DOM.closeGirlModal.addEventListener('click', () => closeModal('girlModal'));
if (DOM.cancelGirlModal) DOM.cancelGirlModal.addEventListener('click', () => closeModal('girlModal'));
if (DOM.closeAttendanceModal) DOM.closeAttendanceModal.addEventListener('click', () => closeModal('attendanceModal'));
if (DOM.cancelAttendanceModal) DOM.cancelAttendanceModal.addEventListener('click', () => closeModal('attendanceModal'));

document.querySelectorAll('.modal-overlay').forEach(overlay => overlay.addEventListener('click', e => {
  if (e.target === overlay) closeModal(overlay.id);
}));

// ============================================================
// EVENT DELEGATION
// ============================================================
function setupDelegation() {
  if (window.__delegationInit) return;
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
      if (item?.dataset.girlId) showGirlProfile(item.dataset.girlId);
    });
  }

  if (DOM.attendanceList) {
    // Click handling — with long press support (was: Bug #19)
    DOM.attendanceList.addEventListener('click', e => {
      const star = e.target.closest('.att-inline-star');
      if (star) {
        e.stopPropagation(); e.preventDefault();
        const ratingWrap = star.closest('.att-inline-rating');
        if (ratingWrap) saveInlineRating(ratingWrap.dataset.attKey, parseInt(star.dataset.val));
        return;
      }
      if (state.isLongPress) { state.isLongPress = false; e.preventDefault(); e.stopPropagation(); return; }
      const item = e.target.closest('.att-item');
      if (item) {
        const g = Cache.getGirl(item.dataset.girlId);
        if (g && DOM.attendanceDate) toggleAttendanceStatus(g.id, g.name, DOM.attendanceDate.value);
      }
    });

    // Long press — safer handling (was: Bug #19 race condition)
    let longPressActive = false;

    const startLongPress = (e) => {
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
    };

    const cancelLongPress = () => {
      longPressActive = false;
      if (state.longPressTimer) { clearTimeout(state.longPressTimer); state.longPressTimer = null; }
      setTimeout(() => { state.isLongPress = false; }, 150);
    };

    DOM.attendanceList.addEventListener('mousedown', startLongPress);
    DOM.attendanceList.addEventListener('mouseup', cancelLongPress);
    DOM.attendanceList.addEventListener('mouseleave', cancelLongPress);
    DOM.attendanceList.addEventListener('touchstart', startLongPress, { passive: true });
    DOM.attendanceList.addEventListener('touchend', cancelLongPress);
    DOM.attendanceList.addEventListener('touchcancel', cancelLongPress);
  }

  if (DOM.calendarGrid) {
    DOM.calendarGrid.addEventListener('click', e => {
      const day = e.target.closest('.cal-day');
      if (day && !day.classList.contains('empty')) showDayDetail(day.dataset.date);
    });
  }
}

// Grade filter handlers
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
const girlsSearchInput = $('girlsSearch');
if (girlsSearchInput) {
  let girlsSearchTimer = null;
  girlsSearchInput.addEventListener('input', () => {
    clearTimeout(girlsSearchTimer);
    girlsSearchTimer = setTimeout(() => { state.girlsSearchQuery = girlsSearchInput.value; renderGirlsList(); }, 250);
  });
}

setupDelegation();

// ============================================================
// PAGE RENDER SCHEDULER
// ============================================================
const PageRenderScheduler = {
  _pending: false,
  _lastRenderedDate: null,
  schedule() {
    const currentDate = TimeContext.getDate();
    if (this._lastRenderedDate === currentDate) return;
    if (this._pending) return;
    this._pending = true;
    requestAnimationFrame(() => {
      this._pending = false;
      const date = TimeContext.getDate();
      if (this._lastRenderedDate === date) return;
      this._lastRenderedDate = date;
      renderPage();
    });
  }
};
TimeContext.subscribe(() => PageRenderScheduler.schedule());

// ============================================================
// BOOTSTRAP
// ============================================================
async function bootstrap() {
  Theme.init();
  TimeContext.init();

  try { await IDB.init(); state.idb = true; } catch (e) { state.idb = false; }

  const modulesReady = await initModules();
  if (modulesReady) await initAuth();
  else { hideSplash(); showLogin(); }
}

bootstrap();
