// ============================================================
// CACHE — Indexed data layer + AttendanceStore + Absence cache
// All O(1) lookups. Safe null handling throughout (was: Bug #2)
// ============================================================

import { isValidActivity } from './config.js';
import { getMonthStr, isValidDateStr } from './utils.js';

// ---- STATE REFERENCES (injected from app) ----
let _stateGirlsRef = [];
let _stateAttendanceDataRef = {};

export function setCacheStateRefs(girlsRef, attendanceDataRef) {
  _stateGirlsRef = girlsRef;
  _stateAttendanceDataRef = attendanceDataRef;
}

// ---- CORE CACHE ----

export const Cache = {
  girlsById: null,
  allAttendance: null,
  attendanceByGirl: null,
  attendanceByDate: null,
  attendanceByMonth: null,
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
    AbsenceCache.clear();
  },

  build() {
    if (!this._dirty) return;

    // FULL rebuild from source of truth
    this.girlsById = {};
    const activeIds = new Set();
    for (const g of _stateGirlsRef) {
      if (!g.isDeleted && g.id) {
        this.girlsById[g.id] = g;
        activeIds.add(g.id);
      }
    }
    this.activeGirlIds = activeIds;

    // Deduplicate attendance by ID — keep most recent version
    const attMap = new Map();
    for (const a of Object.values(_stateAttendanceDataRef)) {
      if (!a || !a.id) continue;
      const existing = attMap.get(a.id);
      if (!existing || (a.updatedAt || 0) >= (existing.updatedAt || 0)) {
        attMap.set(a.id, a);
      }
    }
    const allAtt = Array.from(attMap.values());
    this.allAttendance = allAtt;

    // Build indexed structures for O(1) lookups
    this.attendanceByGirl = {};
    this.attendanceByDate = {};
    this.attendanceByMonth = {};

    for (const a of allAtt) {
      // By girl
      if (!this.attendanceByGirl[a.girlId]) this.attendanceByGirl[a.girlId] = [];
      this.attendanceByGirl[a.girlId].push(a);
      // By date
      if (!this.attendanceByDate[a.date]) this.attendanceByDate[a.date] = [];
      this.attendanceByDate[a.date].push(a);
      // By month — with validation
      const month = getMonthStr(a.date);
      if (month) {
        if (!this.attendanceByMonth[month]) this.attendanceByMonth[month] = [];
        this.attendanceByMonth[month].push(a);
      }
    }

    this._dirty = false;
  },

  /** Get girl by ID — always returns null if not found (safe) */
  getGirl(id) {
    this.build();
    return this.girlsById?.[id] ?? null;
  },

  /** Require a girl — throws descriptive error for debugging (was: Bug #2) */
  requireGirl(id, context = '') {
    const g = this.getGirl(id);
    if (!g) {
      const ctx = context ? ` (${context})` : '';
      console.warn(`Cache.getGirl(${id}) returned null${ctx}`);
    }
    return g;
  },

  getAllAttendance() {
    this.build();
    return this.allAttendance || [];
  },

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

  getActiveGirlIds() {
    this.build();
    return this.activeGirlIds || new Set();
  },

  /** Build a lookup map: girlId → girl object for a set of IDs */
  getGirlMap(girlIds) {
    this.build();
    const map = {};
    for (const id of girlIds) {
      map[id] = this.girlsById?.[id] ?? null;
    }
    return map;
  }
};

// ---- ATTENDANCE STORE (memoized snapshot) ----

export const AttendanceStore = {
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
const _origInvalidate = Cache.invalidate.bind(Cache);
Cache.invalidate = function () {
  _origInvalidate();
  AttendanceStore.invalidate();
};

// ---- ABSENCE CACHE ----

import { getServiceDaysInMonth } from './utils.js';

const AbsenceCache = {
  _data: {},
  _lastMonth: null,

  clear() {
    this._data = {};
    this._lastMonth = null;
  },

  get(monthStr, girlId) {
    return this._data[monthStr]?.[girlId] ?? { hasConsecutive: false, count: 0, dates: [] };
  },

  has(monthStr) {
    return !!this._data[monthStr];
  },

  set(monthStr, data) {
    this._data[monthStr] = data;
    this._lastMonth = monthStr;
  }
};

/** Build absence cache for a month in O(n) single pass */
export function buildAbsenceCache(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  const serviceDays = getServiceDaysInMonth(year, month - 1);
  const monthAtt = Cache.getAttendanceByMonth(monthStr);

  // Group absence records by girl
  const absByGirl = {};
  for (const a of monthAtt) {
    if (a.status === 'غائب') {
      if (!absByGirl[a.girlId]) absByGirl[a.girlId] = new Set();
      absByGirl[a.girlId].add(a.date);
    }
  }

  const cache = {};
  for (const [girlId, absDateSet] of Object.entries(absByGirl)) {
    const absDates = [...absDateSet].sort();
    if (absDates.length < 2) {
      cache[girlId] = { hasConsecutive: false, count: absDates.length, dates: absDates };
      continue;
    }

    // Build absent service indices
    const absentIndices = [];
    for (let i = 0; i < serviceDays.length; i++) {
      if (absDateSet.has(serviceDays[i])) absentIndices.push(i);
    }

    if (absentIndices.length < 2) {
      cache[girlId] = { hasConsecutive: false, count: absDates.length, dates: absDates };
      continue;
    }

    // Check for consecutive service day absences
    let consecutiveCount = 1;
    let maxConsecutive = 1;
    for (let i = 0; i < absentIndices.length - 1; i++) {
      if (absentIndices[i + 1] - absentIndices[i] === 1) {
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
  }

  AbsenceCache.set(monthStr, cache);
}

/** O(1) lookup — builds cache on first access if needed */
export function hasConsecutiveAbsences(girlId, monthStr) {
  if (!AbsenceCache.has(monthStr)) {
    buildAbsenceCache(monthStr);
  }
  return AbsenceCache.get(monthStr, girlId);
}
