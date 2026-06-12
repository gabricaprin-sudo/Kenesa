// ============================================================
// UTILS — Pure functions, no side effects, no DOM
// ============================================================
import { DAY_NAMES, SERVICE_DAY_NUMBERS } from './config.js';

// ---- XSS PROTECTION ----

const _escDiv = document.createElement('div');
const _escText = document.createTextNode('');
_escDiv.appendChild(_escText);

/** Escape HTML entities — prevents injection (was: Bug #18) */
export function esc(str) {
  _escText.nodeValue = String(str ?? '');
  return _escDiv.innerHTML;
}

/** Escape XML attributes */
export function xmlEsc(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** CSV-safe escaping */
export function csvEscape(v) {
  return '"' + String(v ?? '').replace(/"/g, '""') + '"';
}

// ---- DATE UTILITIES ----

/**
 * Safely parse YYYY-MM-DD without timezone bugs.
 * Returns Invalid Date (isNaN) for malformed input.
 */
export function parseDateStr(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return new Date(NaN);
  const parts = dateStr.split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return new Date(NaN);
  const [year, month, day] = parts;
  if (month < 1 || month > 12 || day < 1 || day > 31) return new Date(NaN);
  const d = new Date(year, month - 1, day);
  // Verify no silent JS correction (e.g. 2024-02-31 → 2024-03-02)
  if (d.getMonth() !== month - 1 || d.getDate() !== day) return new Date(NaN);
  return d;
}

/** Compare two date strings: -1 if a < b, 0 if equal, 1 if a > b */
export function compareDateStr(a, b) {
  if (a === b) return 0;
  const ta = parseDateStr(a).getTime();
  const tb = parseDateStr(b).getTime();
  if (isNaN(ta) || isNaN(tb)) return String(a).localeCompare(String(b));
  return ta < tb ? -1 : ta > tb ? 1 : 0;
}

/** Check if dateStr is within range [start, end] inclusive */
export function isDateInRange(dateStr, start, end) {
  return compareDateStr(dateStr, start) >= 0 && compareDateStr(dateStr, end) <= 0;
}

/** Generate attendance record key — centralized format */
export function makeAttKey(girlId, date, activity) {
  return `${girlId}_${date}_${activity}`;
}

/** Validate date string format */
export function isValidDateStr(dateStr) {
  return /^\d{4}-\d{2}-\d{2}$/.test(dateStr) && !isNaN(parseDateStr(dateStr).getTime());
}

/** Safe month extraction — validated (was: Bug #8 format uncertainty) */
export function getMonthStr(dateStr) {
  if (!isValidDateStr(dateStr)) return '';
  return dateStr.substring(0, 7);
}

/** Pad number to 2 digits */
function pad2(n) {
  return String(n).padStart(2, '0');
}

export const DateUtil = {
  pad: pad2,

  toStr(d = new Date()) {
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  },

  getMonthStr(d = new Date()) {
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
  },

  formatMonth(str) {
    if (!str) return '';
    const [y, m] = str.split('-').map(Number);
    return new Date(y, m - 1, 1).toLocaleDateString('ar-EG', { year: 'numeric', month: 'long' });
  },

  formatDateShort(d = new Date()) {
    return `${d.getDate()}/${d.getMonth() + 1}`;
  },

  dayName(d = new Date()) {
    return DAY_NAMES[d.getDay()] || '';
  },

  /** Normalize day name variants (hamza forms) */
  normalizeDay(d) {
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

// ---- ARABIC TEXT NORMALIZATION ----

export function normalizeArabic(str) {
  if (!str) return '';
  return str
    .replace(/[\u064B-\u065F\u0670\u0640]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ؤ/g, 'و')
    .replace(/ئ/g, 'ي')
    .toLowerCase();
}

export function normalizeName(name) {
  return name
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[أإآ]/g, 'ا')
    .replace(/ى/g, 'ي')
    .toLowerCase();
}

// ---- SERVICE DAY HELPERS ----

/** Get all service day dates in a given month */
export function getServiceDaysInMonth(year, month) {
  const days = [];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    const dayOfWeek = new Date(year, month, d).getDay();
    if (SERVICE_DAY_NUMBERS.includes(dayOfWeek)) {
      days.push(`${year}-${pad2(month + 1)}-${pad2(d)}`);
    }
  }
  return days;
}

/** Count service days up to a specific date */
export function getServiceDaysUpToDate(fromYear, fromMonth, toDateStr) {
  const to = parseDateStr(toDateStr);
  if (isNaN(to.getTime())) return 0;
  const toYear = to.getFullYear();
  const toMonth = to.getMonth();
  const toDay = to.getDate();
  const lastDay = (fromYear === toYear && fromMonth === toMonth)
    ? toDay
    : new Date(fromYear, fromMonth + 1, 0).getDate();
  let count = 0;
  for (let d = 1; d <= lastDay; d++) {
    if (SERVICE_DAY_NUMBERS.includes(new Date(fromYear, fromMonth, d).getDay())) {
      count++;
    }
  }
  return count;
}

/** Check if a date is a service day */
export function isServiceDayDate(dateStr) {
  const d = parseDateStr(dateStr);
  if (isNaN(d.getTime())) return false;
  return SERVICE_DAY_NUMBERS.includes(d.getDay());
}

// ---- STATS BOUNDS ----

/** Validate date string with fallback */
export function validateDateStr(dateStr, fallback) {
  if (!dateStr || typeof dateStr !== 'string' || dateStr.length < 10) return fallback;
  const parts = dateStr.split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return fallback;
  const [year, month, day] = parts;
  if (year < 2000 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) return fallback;
  const d = new Date(year, month - 1, day);
  if (d.getMonth() !== month - 1 || d.getDate() !== day) return fallback;
  return dateStr;
}

// ---- URL / POPUP HELPERS ----

/** Safer popup open — detects blockers (was: Bug #4) */
export function openPopup(url, name = '_blank') {
  let w;
  try {
    w = window.open(url, name);
  } catch (e) {
    return { window: null, blocked: true };
  }
  if (!w || w.closed || typeof w.closed === 'undefined') {
    return { window: null, blocked: true };
  }
  return { window: w, blocked: false };
}

/** Download a file via blob */
export function downloadFile(filename, content, mimeType) {
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

/** Generate a safe filename suffix from grade */
export function gradeFileSuffix(grade) {
  return grade ? '_' + grade.replace(/\s/g, '_') : '_الكل';
}
