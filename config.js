// ============================================================
// CONFIG — All constants in one place (was: magic strings + scattered constants)
// ============================================================

/** Service days mapping */
export const SERVICE_DAYS = { 'السبت': true, 'الاثنين': true, 'الاربعاء': true };
export const SERVICE_DAY_NUMBERS = [1, 3, 6]; // Mon, Wed, Sat

/** Day names (Sunday = index 0 in JS Date, but we display in Arabic order) */
export const DAY_NAMES = ['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];

/** Activities — single source of truth */
export const ACTIVITIES = Object.freeze(['دراسي', 'محفوظات', 'قبطي', 'ألحان']);

/** Activity icons */
export const ACTIVITY_ICONS = Object.freeze({
  'دراسي': '&#128216;',
  'ألحان': '&#127925;',
  'قبطي': '&#9961;',
  'محفوظات': '&#128221;'
});

/** Activity-safe accessor — prevents undefined crash (was: Bug #1)
 * Usage: getActivityData(obj, activityKey) instead of obj[activityKey]
 */
export function getActivityData(container, activity, fallback = null) {
  if (!container || typeof container !== 'object') return fallback;
  if (!isValidActivity(activity)) return fallback;
  return container[activity] ?? fallback;
}

/** Validate activity is known */
export function isValidActivity(activity) {
  return ACTIVITIES.includes(activity);
}

/** Period labels */
export const PERIOD_LABELS = Object.freeze({
  today: 'اليوم',
  month: 'هذا الشهر',
  year: 'هذه السنة',
  all: 'كل الفترات'
});

/** Grade ordering for export — Map is safer than plain object (was: Bug #3) */
export const GRADE_ORDER = Object.freeze({
  'تالتة إعدادي': 1,
  'تانية إعدادي': 2,
  'أولى إعدادي': 3
});

/** Valid grades */
export const VALID_GRADES = Object.freeze(Object.keys(GRADE_ORDER));

/** Grade-safe accessor with fallback (was: Bug #3 silent failure) */
export function getGradeOrder(grade, fallback = 99) {
  return GRADE_ORDER[grade] ?? fallback;
}

/** Sort comparator for grades (تالته → تانية → أولى) */
export function compareGrades(a, b) {
  return getGradeOrder(a) - getGradeOrder(b);
}

/** History page size */
export const HISTORY_PAGE_SIZE = 30;

/** Page titles */
export const PAGE_TITLES = Object.freeze({
  home: ['الرئيسية', ''],
  attendance: ['الحضور اليومي', 'تسجيل وإدارة الحضور'],
  girls: ['المخدومات', 'قائمة المخدومات'],
  calendar: ['التقويم الشهري', 'أيام الخدمة'],
  stats: ['الإحصائيات', 'تحليلات وتقارير'],
  history: ['السجل التاريخي', 'سجل التعديلات'],
  export: ['التصدير', 'تصدير البيانات']
});

/** Attendance statuses */
export const STATUS = Object.freeze({
  PRESENT: 'حاضر',
  ABSENT: 'غائب'
});

/** Firebase config (kept for reference — loaded from server in production) */
export const FIREBASE_CONFIG = Object.freeze({
  apiKey: "AIzaSyB2cycBTKMjVg8S_fBYN8C-hwUk5FUF81Q",
  authDomain: "kenesa-e5efd.firebaseapp.com",
  projectId: "kenesa-e5efd",
  storageBucket: "kenesa-e5efd.firebasestorage.app",
  messagingSenderId: "227273753184",
  appId: "1:227273753184:web:ecdf258142ad55ed5cf905",
  measurementId: "G-6HS8KNW1GZ"
});
