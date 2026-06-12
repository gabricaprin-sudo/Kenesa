// ============================================================
// EXPORT ENGINE — Zero duplication, safe, fast
// All 3 formats (Excel/JSON/Print) share ONE data pipeline
// Fixes: Bugs #1, #2, #5, #6, #11, #12, #13, #15, #18
// ============================================================

import {
  ACTIVITIES, isValidActivity, ACTIVITY_ICONS,
  GRADE_ORDER, compareGrades, getGradeOrder, VALID_GRADES
} from './config.js';
import { Cache } from './cache.js';
import {
  esc, xmlEsc, csvEscape, parseDateStr, compareDateStr, isDateInRange,
  DateUtil, getMonthStr, isValidDateStr, makeAttKey,
  downloadFile, gradeFileSuffix, openPopup
} from './utils.js';

// ---- ACTIVITY DATA BUILDER (was: Bug #1 crash site) ----

/** Create a safe activity stats container with all activities preset */
function createActivityContainer(initialValue = 0) {
  const c = {};
  for (const act of ACTIVITIES) c[act] = initialValue;
  return c;
}

/** Create a safe activity pair container { present, absent } */
function createActivityPairContainer() {
  const c = {};
  for (const act of ACTIVITIES) c[act] = { present: 0, absent: 0 };
  return c;
}

/** Safe activity increment — validated (was: Bug #1) */
function incActivity(container, activity, field) {
  if (!container || !isValidActivity(activity)) return;
  if (!container[activity]) container[activity] = { present: 0, absent: 0 };
  if (field === 'present' || field === 'absent') {
    container[activity][field]++;
  }
}

// ---- UNIFIED DATA PIPELINE (fixes: Bugs #6, #12, #15) ----

/**
 * Build export data ONCE, used by all formats.
 * Pre-computes girl map (no Cache.getGirl in loops).
 * Validates activities (no undefined crash).
 */
export function buildExportData({
  girls,          // filtered + sorted girls array
  attendance,     // filtered attendance records
  exportDate,     // string YYYY-MM-DD
  exportMode,     // 'day' | 'month'
  gradeFilter     // string or ''
}) {
  if (!girls || !girls.length) {
    return { girls: [], girlMap: {}, byGrade: {}, sortedGrades: [], attData: [], attByGirl: {}, hasData: false };
  }

  // Pre-compute girl lookup map ONCE (was: Bug #11 — Cache.getGirl in loops)
  const girlMap = {};
  const girlIds = new Set();
  for (const g of girls) {
    girlMap[g.id] = g;
    girlIds.add(g.id);
  }

  // Group girls by grade (was: duplicated in Excel/Print/JSON)
  const byGrade = {};
  for (const g of girls) {
    if (!byGrade[g.grade]) byGrade[g.grade] = [];
    byGrade[g.grade].push(g);
  }

  // Sort grades: تالثة → تانية → أولى (was: Bug #3 — GRADE_ORDER silently fails)
  const sortedGrades = Object.keys(byGrade).sort(compareGrades);

  // Build attendance data per girl — validated activities (was: Bug #1)
  const attByGirl = {};
  for (const a of attendance) {
    if (!girlIds.has(a.girlId)) continue;

    if (!attByGirl[a.girlId]) {
      const g = girlMap[a.girlId];
      attByGirl[a.girlId] = {
        name: g?.name || '',
        grade: g?.grade || '',
        activities: createActivityPairContainer(),
        totalPresent: 0,
        totalAbsent: 0
      };
    }

    // Safe activity access — validated (was: Bug #1)
    if (isValidActivity(a.activity)) {
      if (a.status === 'حاضر') {
        attByGirl[a.girlId].activities[a.activity].present++;
        attByGirl[a.girlId].totalPresent++;
      } else if (a.status === 'غائب') {
        attByGirl[a.girlId].activities[a.activity].absent++;
        attByGirl[a.girlId].totalAbsent++;
      }
    }
  }

  // For day mode: include ALL girls (even those with no attendance)
  if (exportMode === 'day') {
    for (const g of girls) {
      if (!attByGirl[g.id]) {
        attByGirl[g.id] = {
          name: g.name,
          grade: g.grade,
          activities: createActivityPairContainer(),
          totalPresent: 0,
          totalAbsent: 0
        };
      }
    }
  }

  return {
    girls,
    girlMap,         // id → girl (precomputed, no Cache lookups later)
    byGrade,         // grade → girls[]
    sortedGrades,    // ['تالتة إعدادي', 'تانية إعدادي', ...]
    attByGirl,       // girlId → { activities, totalPresent, totalAbsent }
    attendance,      // filtered attendance records
    hasData: girls.length > 0
  };
}

// ---- EXCEL EXPORT ----

export function exportToExcel(data, { exportDate, exportMode, gradeFilter }) {
  // Dynamic import XLSX (available as global in app)
  const XLSX = window.XLSX;
  if (!XLSX) {
    throw new Error('مكتبة Excel غير محملة، حاول تحديث الصفحة');
  }

  const { byGrade, sortedGrades, attByGirl, attendance, girls } = data;
  const gradeSuffix = gradeFilter ? ` — ${gradeFilter}` : '';
  const wb = XLSX.utils.book_new();

  if (exportMode === 'month') {
    const monthName = DateUtil.formatMonth(getMonthStr(exportDate));

    // ===== Sheet 1: Summary by grade =====
    const wsData = [];
    wsData.push(['تقرير حضور شهر ' + monthName + gradeSuffix]);
    wsData.push([]);
    wsData.push(['عدد المخدومات', girls.length]);
    if (gradeFilter) wsData.push(['السنة المحددة', gradeFilter]);
    wsData.push([]);

    for (const grade of sortedGrades) {
      wsData.push([`═══ ${grade} ═══`]);
      wsData.push(['الاسم', 'دراسي', 'قبطي', 'ألحان', 'محفوظات', 'إجمالي الحضور', 'إجمالي الغياب']);

      const gradeGirls = byGrade[grade];
      for (const g of gradeGirls) {
        const r = attByGirl[g.id];
        wsData.push([
          g.name,
          r?.activities?.['دراسي']?.present ?? 0,
          r?.activities?.['قبطي']?.present ?? 0,
          r?.activities?.['ألحان']?.present ?? 0,
          r?.activities?.['محفوظات']?.present ?? 0,
          r?.totalPresent ?? 0,
          r?.totalAbsent ?? 0
        ]);
      }
      wsData.push([]);
    }

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{ wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 14 }];
    ws['!dir'] = 'rtl';
    XLSX.utils.book_append_sheet(wb, ws, 'ملخص الشهر');

    // ===== Sheet 2: Detailed daily records =====
    // Sort once (was: Bug #9 — sort inside loop)
    const sortedAtt = [...attendance].sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      const gA = data.girlMap[a.girlId];
      const gB = data.girlMap[b.girlId];
      const oA = getGradeOrder(gA?.grade);
      const oB = getGradeOrder(gB?.grade);
      if (oA !== oB) return oA - oB;
      return (gA?.name || '').localeCompare(gB?.name || '', 'ar');
    });

    const detailData = [];
    detailData.push(['تقرير تفصيلي — ' + monthName + gradeSuffix]);
    detailData.push([]);
    detailData.push(['التاريخ', 'اليوم', 'المخدومة', 'السنة', 'النشاط', 'الحالة', 'التقييم', 'ملاحظات']);

    for (const a of sortedAtt) {
      const g = data.girlMap[a.girlId];
      const dayName = DateUtil.dayName(parseDateStr(a.date));
      const stars = a.rating ? '\u2605'.repeat(a.rating) + '\u2606'.repeat(5 - a.rating) : '';
      detailData.push([
        a.date, dayName, g?.name || '', g?.grade || '',
        a.activity || '', a.status === 'حاضر' ? '\u2713' : '\u2717',
        stars, a.notes || ''
      ]);
    }

    const wsDetail = XLSX.utils.aoa_to_sheet(detailData);
    wsDetail['!cols'] = [{ wch: 14 }, { wch: 10 }, { wch: 24 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 24 }];
    wsDetail['!dir'] = 'rtl';
    XLSX.utils.book_append_sheet(wb, wsDetail, 'تفاصيل يومية');

  } else {
    // ===== Day export =====
    const dayName = DateUtil.dayName(parseDateStr(exportDate));
    const wsData = [];
    wsData.push([`تقرير حضور يوم ${exportDate} (${dayName})${gradeSuffix}`]);
    if (gradeFilter) wsData.push(['السنة المحددة:', gradeFilter]);
    wsData.push([]);
    wsData.push(['الاسم', 'السنة', 'دراسي', 'قبطي', 'ألحان', 'محفوظات']);

    for (const g of girls) {
      const row = [g.name, g.grade];
      for (const act of ACTIVITIES) {
        const r = attByGirl[g.id]?.activities?.[act];
        if (r && (r.present > 0 || r.absent > 0)) {
          row.push(r.present > 0 ? '\u2713' : '\u2717');
        } else {
          row.push('\u2014');
        }
      }
      wsData.push(row);
    }

    const totalPresent = attendance.filter(a => a.status === 'حاضر').length;
    const totalAbsent = attendance.filter(a => a.status === 'غائب').length;
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
  a.download = `حضور_${exportDate}${exportMode === 'month' ? '_شهر' : '_يوم'}${gradeFileSuffix(gradeFilter)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---- JSON EXPORT ----

export function exportToJSON(data, { exportDate, gradeFilter }) {
  const monthStr = getMonthStr(exportDate);
  const exportStart = monthStr + '-01';
  const exportEnd = exportDate;

  const payload = {
    dateRange: { start: exportStart, end: exportEnd },
    girls: data.girls,
    attendance: data.attendance,
    exportedAt: new Date().toISOString(),
    gradeFilter: gradeFilter || 'all'
  };

  downloadFile(
    `بيانات_${exportDate}${gradeFileSuffix(gradeFilter)}.json`,
    JSON.stringify(payload, null, 2),
    'application/json'
  );
}

// ---- PRINT/PDF EXPORT ----

export function exportToPrint(data, { exportDate, exportMode, gradeFilter }) {
  const { byGrade, sortedGrades, attByGirl, attendance, girls } = data;
  const gradeLabel = gradeFilter ? ` — ${gradeFilter}` : '';
  const totalPresent = attendance.filter(a => a.status === 'حاضر').length;
  const totalAbsent = attendance.filter(a => a.status === 'غائب').length;

  let html;

  if (exportMode === 'month') {
    const monthName = DateUtil.formatMonth(getMonthStr(exportDate));
    const exportStart = getMonthStr(exportDate) + '-01';
    const exportEnd = exportDate;

    // Build grade sections (was: Bug #13 — string concat in loop)
    const gradeSections = [];
    for (const grade of sortedGrades) {
      const gradeGirls = byGrade[grade];
      const rows = [];
      for (let i = 0; i < gradeGirls.length; i++) {
        const g = gradeGirls[i];
        const r = attByGirl[g.id];
        rows.push(`<tr>
          <td>${i + 1}</td>
          <td>${esc(g.name)}</td>
          <td>${r?.activities?.['دراسي']?.present ?? 0}</td>
          <td>${r?.activities?.['قبطي']?.present ?? 0}</td>
          <td>${r?.activities?.['ألحان']?.present ?? 0}</td>
          <td>${r?.activities?.['محفوظات']?.present ?? 0}</td>
          <td style="color:green;font-weight:700">${r?.totalPresent ?? 0}</td>
          <td style="color:red;font-weight:700">${r?.totalAbsent ?? 0}</td>
        </tr>`);
      }

      gradeSections.push(`
        <h2 style="color:#1a2744;margin-top:24px;margin-bottom:12px;padding:8px 12px;background:#f0f2f8;border-radius:8px;font-size:18px;">
          ${esc(grade)} — ${gradeGirls.length} مخدومة
        </h2>
        <table>
          <tr><th>#</th><th>الاسم</th><th>دراسي</th><th>قبطي</th><th>ألحان</th><th>محفوظات</th><th>إجمالي الحضور</th><th>إجمالي الغياب</th></tr>
          ${rows.join('')}
        </table>
      `);
    }

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
      <h1>تقرير حضور شهر ${monthName}${esc(gradeLabel)}</h1>
      <p style="color:#6b7a99;font-size:14px">الفترة: من ${exportStart} إلى ${exportEnd}</p>
      <div class="summary">
        <div class="sum-box"><b>${girls.length}</b><br><span>عدد المخدومات</span></div>
        <div class="sum-box"><b>${totalPresent}</b><br><span>إجمالي الحضور</span></div>
        <div class="sum-box"><b>${totalAbsent}</b><br><span>إجمالي الغياب</span></div>
      </div>
      ${gradeSections.join('')}
      <div class="footer">تاريخ التصدير: ${new Date().toLocaleDateString('ar-EG')} | نظام متابعة المخدومات</div>
      </body></html>`;

  } else {
    // Day export
    const dayName = DateUtil.dayName(parseDateStr(exportDate));
    const gradeSections = [];

    for (const grade of sortedGrades) {
      const gradeGirls = byGrade[grade];
      const rows = [];
      for (let i = 0; i < gradeGirls.length; i++) {
        const g = gradeGirls[i];
        const r = attByGirl[g.id];
        const cells = [];
        for (const act of ACTIVITIES) {
          const actData = r?.activities?.[act];
          if (actData && (actData.present > 0 || actData.absent > 0)) {
            cells.push(actData.present > 0
              ? '<td style="color:green;font-weight:700;font-size:16px">\u2713</td>'
              : '<td style="color:red;font-weight:700;font-size:16px">\u2717</td>');
          } else {
            cells.push('<td style="color:#ccc">\u2014</td>');
          }
        }
        rows.push(`<tr>
          <td>${i + 1}</td>
          <td>${esc(g.name)}</td>
          <td>${esc(g.grade)}</td>
          ${cells.join('')}
        </tr>`);
      }

      gradeSections.push(`
        <h2 style="color:#1a2744;margin-top:20px;margin-bottom:10px;padding:6px 10px;background:#f0f2f8;border-radius:8px;font-size:16px;">
          ${esc(grade)} — ${gradeGirls.length} مخدومة
        </h2>
        <table>
          <tr><th>#</th><th>الاسم</th><th>السنة</th><th>دراسي</th><th>قبطي</th><th>ألحان</th><th>محفوظات</th></tr>
          ${rows.join('')}
        </table>
      `);
    }

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
      <h1>تقرير حضور يوم ${exportDate}${esc(gradeLabel)}</h1>
      <p style="color:#6b7a99;font-size:14px">اليوم: ${dayName}</p>
      <div class="summary">
        <div class="sum-box"><b>${girls.length}</b><br><span>عدد المخدومات</span></div>
        <div class="sum-box"><b>${totalPresent}</b><br><span>حاضر</span></div>
        <div class="sum-box"><b>${totalAbsent}</b><br><span>غائب</span></div>
      </div>
      ${gradeSections.join('')}
      <div class="footer">تاريخ التصدير: ${new Date().toLocaleDateString('ar-EG')} | نظام متابعة المخدومات</div>
      </body></html>`;
  }

  // Safer popup handling (was: Bug #4)
  const { window: w, blocked } = openPopup('', '_blank');
  if (blocked || !w) {
    downloadFile(
      `تقرير_${exportDate}${exportMode === 'month' ? '_شهر' : '_يوم'}${gradeFileSuffix(gradeFilter)}.html`,
      html,
      'text/html;charset=utf-8'
    );
    return 'popup_blocked';
  }

  w.document.write(html);
  w.document.close();
  w.print();
  return 'success';
}
