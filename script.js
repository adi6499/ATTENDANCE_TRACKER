/* ═══════════════════════════════════════════════════════════════
   SMART ATTENDANCE SYSTEM v2.0 — Production-Grade Architecture
   ═══════════════════════════════════════════════════════════════ */

// ── STRUCTURED APPLICATION STATE ──
const AppState = {
  raw: { excelFiles: [], datFiles: [], holidayFiles: [], shiftMap: {} },
  processed: {
    records: [], filtered: [],
    lateRecs: [], absentRecs: [], earlyRecs: [],
    dataReliability: 100
  },
  config: {
    holidays: [],
    failureDates: [],
    version: 'v5'
  },
  ui: {
    sortCol: 'date', sortDir: 1,
    page: 1, perPage: 100,
    quickFilter: '',
    mobileFiltersOpen: false,
    activeProfile: { uid: '', filter: 'all' }
  },
  validation: { warnings: [], duplicateUids: [] }
};

// Legacy compatibility — all old S.xxx references map here
window.S = {
  get excelFiles() { return AppState.raw.excelFiles; },
  set excelFiles(v) { AppState.raw.excelFiles = v; },
  get datFiles() { return AppState.raw.datFiles; },
  set datFiles(v) { AppState.raw.datFiles = v; },
  get shiftMap() { return AppState.raw.shiftMap; },
  set shiftMap(v) { AppState.raw.shiftMap = v; },
  get records() { return AppState.processed.records; },
  set records(v) { AppState.processed.records = v; },
  get filtered() { return AppState.processed.filtered; },
  set filtered(v) { AppState.processed.filtered = v; },
  get lateRecs() { return AppState.processed.lateRecs; },
  set lateRecs(v) { AppState.processed.lateRecs = v; },
  get absentRecs() { return AppState.processed.absentRecs; },
  set absentRecs(v) { AppState.processed.absentRecs = v; },
  get earlyRecs() { return AppState.processed.earlyRecs; },
  set earlyRecs(v) { AppState.processed.earlyRecs = v; },
  get sortCol() { return AppState.ui.sortCol; },
  set sortCol(v) { AppState.ui.sortCol = v; },
  get sortDir() { return AppState.ui.sortDir; },
  set sortDir(v) { AppState.ui.sortDir = v; },
  get page() { return AppState.ui.page; },
  set page(v) { AppState.ui.page = v; },
  get perPage() { return AppState.ui.perPage; },
  set perPage(v) { AppState.ui.perPage = v; },
  get quickFilter() { return AppState.ui.quickFilter; },
  set quickFilter(v) { AppState.ui.quickFilter = v; },
  get mobileFiltersOpen() { return AppState.ui.mobileFiltersOpen; },
  set mobileFiltersOpen(v) { AppState.ui.mobileFiltersOpen = v; },
  get activeProfile() { return AppState.ui.activeProfile; },
  set activeProfile(v) { AppState.ui.activeProfile = v; },
  get holidays() { return AppState.config.holidays; },
  set holidays(v) { AppState.config.holidays = v; },
  get failureDates() { return AppState.config.failureDates; },
  set failureDates(v) { AppState.config.failureDates = v; }
};

const FILTER_CONFIG = {
  'f-branch': { label: 'Branches' },
  'f-dept': { label: 'Departments' },
  'f-status': { label: 'Statuses' },
  'f-emp': { label: 'Employees' }
};

const STATUS_OPTIONS = ['Present', 'Late', 'Late (Comp)', 'Half Day', 'Missed Punch', 'Absent', 'Holiday', 'Week Off', 'System Error'];
const THEMES = [
  { id: 'light', label: '☀️ Light' },
  { id: 'obsidian', label: '🌑 Obsidian' },
  { id: 'midnight', label: '🌙 Midnight' },
  { id: 'sapphire', label: '💎 Sapphire' },
  { id: 'paper', label: '📄 Paper' }
];
const STORAGE_KEYS = {
  records: 'hr_att_v5',
  ui: 'hr_att_ui_v3',
  holidays: 'hr_att_holidays_v1',
  config: 'hr_att_config_v1'
};
const LEGACY_STORAGE_KEYS = {
  records: ['hr_att_v3', 'hr_att_v4'],
  ui: ['hr_att_ui_v1', 'hr_att_ui_v2']
};

// ── EMPLOYEE KEY: uid + branch to avoid cross-branch collisions ──
function makeEmployeeKey(uid, branch) {
  return uid + '_' + (branch || '').replace(/\s+/g, '_').toLowerCase();
}

// ── DATA NORMALIZATION HUB ──
const DataNormalizer = {
  date: (v) => {
    if (!v || v === 'NULL') return null;
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
    }
    const s = String(v).trim();
    if (!s || s === 'NULL' || s === '--' || s === '-') return null;

    const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;

    // Handle DD-Mon-YYYY or DD-Mon-YY
    const dmy = s.match(/^(\d{1,2})[\/\-]([A-Za-z]+)[\/\-](\d{2,4})$/);
    if (dmy) {
      const months = {
        jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
        jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
      };
      const mo = months[dmy[2].toLowerCase().slice(0, 3)];
      if (mo) {
        let yr = parseInt(dmy[3]);
        if (yr < 100) yr += 2000;
        return `${yr}-${String(mo).padStart(2, '0')}-${String(parseInt(dmy[1])).padStart(2, '0')}`;
      }
    }

    // Handle DD/MM/YYYY or DD/MM/YY
    const dmy2 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmy2) {
      let yr = parseInt(dmy2[3]);
      if (yr < 100) yr += 2000;
      return `${yr}-${String(parseInt(dmy2[2])).padStart(2, '0')}-${String(parseInt(dmy2[1])).padStart(2, '0')}`;
    }

    return null;
  },

  time: (v) => {
    if (v == null || v === '' || v === '--') return null;
    if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
    if (typeof v === 'number') return Math.round(v * 24 * 60);
    const m = String(v).match(/(\d+):(\d+)/);
    return m ? parseInt(m[1]) * 60 + parseInt(m[2]) : null;
  },

  branch: (raw) => {
    if (!raw) return '';
    return String(raw)
      .trim()
      .toLowerCase()
      .replace(/^[a-z]+\s*-\s*/i, '')  // Remove prefixes like 'JISPL - '
      .replace(/\s+/g, ' ')
      .trim();
  },

  holidayValue: (raw) => {
    const val = String(raw || '').trim().toLowerCase();
    if (!val || val === 'null' || val === '--' || val === '') return null;
    if (val.includes('holiday') || ['y', 'yes', '1', 'x', '✓', 'ph'].includes(val)) return 'holiday';
    return null;
  }
};

// Legacy alias for compatibility during refactor
function normalizeBranch(raw) {
  return DataNormalizer.branch(raw);
}

// ── HOLIDAY SYSTEM ──
function isHoliday(date, branch) {
  const holidays = AppState.config.holidays;
  if (!holidays.length) return null;

  // STRICT: missing or blank branch = no holiday (prevents global spread via empty string match)
  const empBranch = normalizeBranch(branch);
  if (!empBranch || empBranch === '--') return null;

  const match = holidays.find(h => {
    // Date range check first (applies to both global and branch-specific)
    const dateMatch = h.d2 ? (date >= h.d && date <= h.d2) : (date === h.d);
    if (!dateMatch) return false;

    // __global__ sentinel: holiday has no specific branch column → applies to ALL branches
    if (h.b.includes('__global__')) return true;

    // STRICT branch match: normalize both sides, both must be non-empty
    return h.b.some(b => {
      const hBranch = normalizeBranch(b);
      if (!hBranch) return false;
      return empBranch.includes(hBranch) || hBranch.includes(empBranch);
    });
  }) || null;

  if (match) {
    console.log(`[Holiday Match] ${date} "${branch}"("${empBranch}") → "${match.name}"`);
  }
  return match;
}

function loadHolidays() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.holidays);
    if (raw) {
      AppState.config.holidays = JSON.parse(raw);
      updateHolidayStatus();
    }
  } catch (e) { console.warn('Failed to load holidays', e); }
}

function persistHolidays() {
  if (AppState.config.holidays.length) {
    localStorage.setItem(STORAGE_KEYS.holidays, JSON.stringify(AppState.config.holidays));
  } else {
    localStorage.removeItem(STORAGE_KEYS.holidays);
  }
  updateHolidayStatus();
}

function clearHolidayData() {
  if (!confirm('Clear all holiday data?')) return;
  AppState.config.holidays = [];
  AppState.raw.holidayFiles = [];
  persistHolidays();
  renderPills('holiday');
  showToast('Holiday data cleared.');
}

function updateHolidayStatus() {
  const el = document.getElementById('holiday-status');
  if (!el) return;
  const count = AppState.config.holidays.length;
  el.innerHTML = count
    ? `<span class="holiday-loaded">✓ ${count} holidays loaded</span>`
    : '<span class="holiday-empty">No holidays loaded</span>';
  const clearBtn = document.getElementById('btn-clear-holidays');
  if (clearBtn) clearBtn.style.display = count ? 'inline-flex' : 'none';
}

async function parseHolidaySheet(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        // Read with cellDates:true to get proper Date objects for Excel date serials
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, dateNF: 'yyyy-mm-dd' });
        if (rows.length < 2) { res([]); return; }

        const hdr = rows[0].map(c => String(c || '').trim().toLowerCase());

        // Column detection — name col must NOT be the holiday-value columns
        // "Name of Holiday" → col 0, "Day" → col 1, "Date" → col 2, branches → col 3+
        const nameCol = hdr.findIndex(h => h.includes('name'));
        const dateCol = hdr.findIndex(h => h.includes('date') && !h.includes('end'));
        const endDateCol = hdr.findIndex(h => h.includes('end') && h.includes('date'));

        // Branch columns: any column NOT name/day/date whose header is a known location
        // or just any remaining column after the date column
        const knownBranches = ['mumbai', 'borivali', 'nagpur', 'gujarat', 'goa',
          'delhi', 'pune', 'chennai', 'hyderabad', 'bangalore'];
        const skipCols = new Set([nameCol, dateCol, endDateCol].filter(c => c >= 0));
        // Also skip "day" column
        hdr.forEach((h, i) => { if (h === 'day') skipCols.add(i); });

        const branchCols = [];
        hdr.forEach((h, i) => {
          if (skipCols.has(i)) return;
          if (!h) return;
          branchCols.push({ idx: i, name: h }); // keep all remaining columns as branch columns
        });

        console.log('[parseHolidaySheet] Headers:', hdr);
        console.log('[parseHolidaySheet] nameCol:', nameCol, 'dateCol:', dateCol, 'branchCols:', branchCols.map(b => b.name));

        // Date parser — handles all formats found in this file:
        // - "19-Mar-2026"   (string with month abbreviation)
        // - "14-Jan-2026"
        // - "2026-03-03 00:00:00"  (datetime string from Excel)
        // - Excel serial number (handled by cellDates:true above, becomes formatted string)
        // Use unified DataNormalizer
        const parseDate = (v) => DataNormalizer.date(v);

        let invalidDates = 0;
        let skippedRows = 0;

        const holidays = [];
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          const rawDate = row[dateCol >= 0 ? dateCol : 2];
          const dateStr = parseDate(rawDate);

          if (!dateStr && rawDate) invalidDates++;
          if (!dateStr) continue;

          const endStr = endDateCol >= 0 ? parseDate(row[endDateCol]) : null;
          const name = nameCol >= 0 ? String(row[nameCol] || '').trim() : 'Holiday';
          if (!name) continue;

          // Branch detection
          const branches = [];
          branchCols.forEach(bc => {
            if (DataNormalizer.holidayValue(row[bc.idx]) === 'holiday') {
              branches.push(DataNormalizer.branch(bc.name));
            }
          });

          // Skip row if no branches are marked
          if (!branches.length) {
            skippedRows++;
            console.log(`[parseHolidaySheet] Skipping "${name}" (${dateStr}) — no branches marked.`);
            continue;
          }

          const holiday = { name, d: dateStr, b: branches };
          if (endStr && endStr !== dateStr) holiday.d2 = endStr;
          holidays.push(holiday);
        }

        if (invalidDates > 0) AppState.validation.warnings.push(`ℹ ${invalidDates} unrecognized dates in holiday sheet (skipped).`);
        if (skippedRows > 0) AppState.validation.warnings.push(`ℹ ${skippedRows} holidays skipped due to missing branch markings (does not impact summary).`);


        console.log(`[parseHolidaySheet] Total holidays parsed: ${holidays.length}`);
        res(holidays);
      } catch (err) {
        console.error('[parseHolidaySheet] Error:', err);
        rej(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── VALIDATION LAYER ──
function runValidation(shiftMap, punches) {
  const warnings = AppState.validation.warnings || [];
  const uidBranches = {};

  // Detect duplicate UIDs across branches
  Object.entries(shiftMap).forEach(([uid, info]) => {
    if (!uidBranches[uid]) uidBranches[uid] = new Set();
    if (info.branch) uidBranches[uid].add(info.branch);
  });
  const duplicateUids = [];
  Object.entries(uidBranches).forEach(([uid, branches]) => {
    if (branches.size > 1) {
      duplicateUids.push({ uid, branches: [...branches] });
      warnings.push(`ℹ User ${uid} exists in multiple branches: ${[...branches].join(', ')} (first match used).`);
    }
  });

  // Detect unknown IDs in DAT logs
  const shiftUids = new Set(Object.keys(shiftMap));
  const datUids = new Set(punches.map(p => p.uid));
  const unknownIds = [...datUids].filter(uid => !shiftUids.has(uid));
  if (unknownIds.length > 0) {
    warnings.push(`ℹ ${unknownIds.length} unrecognized ID(s) in logs ignored (does not impact core results). ${unknownIds.slice(0, 5).join(', ')}${unknownIds.length > 5 ? '...' : ''}`);
  }

  // Detect employees in shift master with no punches
  const missingPunches = [...shiftUids].filter(uid => !datUids.has(uid));
  if (missingPunches.length > 0) {
    warnings.push(`ℹ ${missingPunches.length} employee(s) in Shift Master with no attendance logs.`);
  }

  // Already checking missing branches in parseShift, no need to duplicate here

  AppState.validation = { warnings, duplicateUids };
  renderValidationCenter(warnings);
  return warnings;
}

function renderValidationCenter(warnings) {
  const el = document.getElementById('validation-center');
  const countEl = document.getElementById('vc-count');
  const listEl = document.getElementById('vc-list');
  if (!el || !countEl || !listEl) return;

  if (!warnings || warnings.length === 0) {
    el.style.display = 'none';
    return;
  }

  countEl.textContent = warnings.length;
  listEl.innerHTML = warnings.map(w => `<li>${w}</li>`).join('');
  el.style.display = 'flex';
}

function toggleValidationCenter() {
  const body = document.getElementById('vc-body');
  const chev = document.getElementById('vc-chevron');
  if (!body || !chev) return;

  if (body.classList.contains('collapsed')) {
    body.classList.remove('collapsed');
    body.style.display = 'block';
    chev.classList.add('vc-chevron-up');
  } else {
    body.classList.add('collapsed');
    body.style.display = 'none';
    chev.classList.remove('vc-chevron-up');
  }
}

function dismissValidation(e) {
  if (e) e.stopPropagation();
  const el = document.getElementById('validation-center');
  if (el) el.style.display = 'none';
}

const MOBILE_FILTER_BREAKPOINT = 720;
const layoutObservers = [];

function isMobileFilterViewport() {
  return window.matchMedia(`(max-width: ${MOBILE_FILTER_BREAKPOINT}px)`).matches;
}

const CLR = ['#1B4FD8', '#0B7B60', '#8A5A00', '#5B3FA6', '#C0280C', '#3E4B66'];
function avatarCol(n) { return CLR[n.charCodeAt(0) % CLR.length] }
function initials(n) { return n.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase() }

function getRecordDateBounds() {
  const dates = S.records.map(r => r.date).filter(Boolean).sort();
  return {
    min: dates[0] || '',
    max: dates[dates.length - 1] || ''
  };
}

function countActiveFilters() {
  let active = 0;
  const searchValue = document.getElementById('search')?.value.trim();
  const { min, max } = getRecordDateBounds();
  const from = document.getElementById('date-from')?.value || '';
  const to = document.getElementById('date-to')?.value || '';

  if (searchValue) active++;
  if (getSelectedValues('f-branch').length) active++;
  if (getSelectedValues('f-dept').length) active++;
  if (getSelectedValues('f-status').length || S.quickFilter === 'Late') active++;
  if (getSelectedValues('f-emp').length) active++;
  if ((from && from !== min) || (to && to !== max)) active++;

  return active;
}

function refreshMobileFilterUi() {
  const shell = document.getElementById('daily-filter-shell');
  const toggle = document.getElementById('mobile-filter-toggle');
  const summary = document.getElementById('mobile-filter-summary');
  const state = document.getElementById('mobile-filter-state');

  if (!shell || !toggle || !summary || !state) return;

  const mobile = isMobileFilterViewport();
  const open = !mobile || S.mobileFiltersOpen;
  const activeCount = countActiveFilters();

  shell.classList.toggle('collapsed', mobile && !open);
  toggle.setAttribute('aria-expanded', String(open));
  summary.textContent = activeCount ? `${activeCount} active` : 'No filters';
  state.textContent = open ? 'Hide' : 'Show';

  if (!open) {
    document.querySelectorAll('.mfilter.open').forEach(el => el.classList.remove('open'));
  }
}

function toggleMobileFilters(forceOpen) {
  S.mobileFiltersOpen = typeof forceOpen === 'boolean' ? forceOpen : !S.mobileFiltersOpen;
  refreshMobileFilterUi();
  queueStickyLayoutSync();
}

function syncResponsiveUi() {
  refreshMobileFilterUi();
  queueStickyLayoutSync();
}

function updateDailyTitle() {
  const title = document.getElementById('daily-title');
  if (!title) return;

  const count = S.filtered.length;
  const suffix = count === S.records.length ? '' : ' filtered';

  // Intelligent Context: Find the Month and Year from the active date range
  const fv = document.getElementById('date-from')?.value;
  let contextStr = '';
  if (fv) {
    const d = new Date(fv + 'T12:00:00');
    const monthName = d.toLocaleDateString('en-US', { month: 'short' });
    const year = d.getFullYear();
    contextStr = ` - <span style="color:var(--teal)">${monthName} ${year}</span>`;
  }

  title.innerHTML = `Daily Attendance Records${contextStr} (<span id="tb-daily-title-count">${count}</span>${suffix})`;
}

function syncStickyLayout() {
  const root = document.documentElement;
  const topbar = document.querySelector('.topbar');
  const stickyHeader = document.querySelector('.sticky-header-daily');

  const topbarBottom = (topbar && topbar.offsetParent !== null)
    ? Math.max(0, Math.ceil(topbar.getBoundingClientRect().bottom))
    : 0;
  const headerRect = (stickyHeader && stickyHeader.offsetParent !== null)
    ? stickyHeader.getBoundingClientRect()
    : null;
  const headerHeight = headerRect ? Math.ceil(headerRect.height) : 0;
  const dailyTableTop = Math.max(0, topbarBottom + headerHeight - 1);

  root.style.setProperty('--sticky-panel-top', topbarBottom + 'px');
  root.style.setProperty('--sticky-filter-height', headerHeight + 'px');
  root.style.setProperty('--sticky-daily-table-top', dailyTableTop + 'px');
}

function queueStickyLayoutSync() {
  window.requestAnimationFrame(syncStickyLayout);
}

function getShiftDurationMinutes(shiftStart, shiftEnd) {
  if (shiftStart === null || shiftEnd === null || (shiftStart === 0 && shiftEnd === 0)) return 0;
  return shiftEnd >= shiftStart ? shiftEnd - shiftStart : (24 * 60 - shiftStart) + shiftEnd;
}

function getLateMinutes(shiftStart, inM) {
  if (shiftStart == null || inM == null) return 0;
  return Math.max(0, inM - shiftStart);
}

function getEarlyMinutes(shiftStart, shiftEnd, outM) {
  if (shiftEnd == null || outM == null) return 0;
  if (shiftStart != null && shiftEnd < shiftStart) {
    const normalizedOut = outM < shiftStart ? outM + 1440 : outM;
    return Math.max(0, (shiftEnd + 1440) - normalizedOut);
  }
  return Math.max(0, shiftEnd - outM);
}

function getEarlyArrivalMinutes(shiftStart, inM) {
  if (shiftStart === null || inM === null) return 0;
  const sStart = Number(shiftStart);
  const inTime = Number(inM);
  if (isNaN(sStart) || isNaN(inTime)) return 0;
  // If inTime is before shiftStart, calculate the difference
  return Math.max(0, sStart - inTime);
}

function isAttendedStatus(status) {
  return ['Present', 'Late', 'Late (Comp)'].includes(status);
}

function hasLateArrival(record) {
  return Number(record?.lateMins || 0) > 0;
}

function getWorkedMinutesFromPunches(punches) {
  if (!Array.isArray(punches) || punches.length < 2) return { mins: 0, isOdd: punches?.length === 1 };

  const isOdd = punches.length % 2 !== 0;
  let total = 0;

  if (isOdd) {
    // Fallback logic for odd punches: Last OUT - First IN
    const first = punches[0];
    const last = punches[punches.length - 1];
    let grossMins = Math.max(0, Math.round((last - first) / 60000));

    // Treat intermediate punches as breaks: 
    // E.g. [p0, p1, p2] -> break = p2 - p1? No, p1 and p2 could be anything.
    // If [p0, p1, p2, p3, p4] -> break1 = p2 - p1, break2 = p4 - p3.
    let breakMins = 0;
    for (let i = 1; i + 1 < punches.length - 1; i += 2) {
      breakMins += Math.max(0, Math.round((punches[i + 1] - punches[i]) / 60000));
    }

    total = Math.max(0, grossMins - breakMins);
  } else {
    // Standard strict pairing
    for (let i = 0; i + 1 < punches.length; i += 2) {
      total += Math.max(0, Math.round((punches[i + 1] - punches[i]) / 60000));
    }
  }

  return { mins: total, isOdd };
}

function getOvertimeMinutes(shiftStart, shiftEnd, workedMins, lateMins, earlyExitMins) {
  if (shiftStart === null || shiftEnd === null || (shiftStart === 0 && shiftEnd === 0)) return 0;
  const scheduledMins = getShiftDurationMinutes(shiftStart, shiftEnd);
  // Corrected OT: effective work = worked - late penalty - early exit penalty
  const effectiveWork = Math.max(0, Number(workedMins || 0) - Number(lateMins || 0) - Number(earlyExitMins || 0));
  const rawOT = Math.max(0, effectiveWork - scheduledMins);
  // Floor to nearest 30-minute block, minimum 30 min threshold
  return rawOT >= 30 ? Math.floor(rawOT / 30) * 30 : 0;
}

function initApp() {
  window.closeSidebar();
  bindDatePickers();
  bindFilterMenus();

  // Load holidays from localStorage
  loadHolidays();

  const savedTheme = localStorage.getItem('theme') || 'obsidian';
  applyTheme(savedTheme);

  const topbar = document.querySelector('.topbar');
  const stickyHeader = document.querySelector('.sticky-header-daily');
  if ('ResizeObserver' in window) {
    [topbar, stickyHeader].filter(Boolean).forEach(target => {
      const observer = new ResizeObserver(() => queueStickyLayoutSync());
      observer.observe(target);
      layoutObservers.push(observer);
    });
  }

  document.querySelector('.logo-img')?.addEventListener('load', queueStickyLayoutSync, { once: true });
  window.addEventListener('load', syncResponsiveUi, { once: true });
  window.addEventListener('orientationchange', syncResponsiveUi);
  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', syncResponsiveUi);
  }
  if (document.fonts?.ready) {
    document.fonts.ready.then(syncResponsiveUi).catch(() => { });
  }

  migrateStoredSession();
  restoreSavedSession();
  syncResponsiveUi();
}

window.addEventListener('resize', syncResponsiveUi);

function toggleTheme() {
  const current = document.documentElement.getAttribute('data-theme') || 'obsidian';
  const idx = THEMES.findIndex(t => t.id === current);
  const next = THEMES[(idx + 1 + THEMES.length) % THEMES.length];
  applyTheme(next.id);
}

function timeToMins(t) { if (!t) return null; const [h, m] = t.split(':').map(Number); return h * 60 + m }
function minsToTime(m) { if (m === null) return '--'; const h = Math.floor(m / 60), mm = Math.floor(m % 60); return `${h.toString().padStart(2, '0')}:${mm.toString().padStart(2, '0')}` }

function applyTheme(themeId) {
  const theme = THEMES.find(t => t.id === themeId) || THEMES[0];
  document.documentElement.setAttribute('data-theme', theme.id);
  // Support both dropdown and button label
  const selector = document.getElementById('theme-selector');
  const label = document.getElementById('theme-btn-label');
  if (selector) selector.value = theme.id;
  if (label) label.textContent = theme.label;
  localStorage.setItem('theme', theme.id);
  syncAppMode();
  syncResponsiveUi();
}

function getActiveTabName() {
  const activeTab = document.querySelector('.tab.active');
  const action = activeTab?.getAttribute('onclick') || '';
  const match = action.match(/switchTab\('([^']+)'/);
  return match ? match[1] : 'daily';
}

function getTabButton(name) {
  return document.querySelector(`.tab[onclick*="'${name}'"]`) || document.querySelector(`.tab[onclick*="${name}"]`);
}

function persistRecords() {
  if (!S.records.length) {
    localStorage.removeItem(STORAGE_KEYS.records);
    return;
  }
  localStorage.setItem(STORAGE_KEYS.records, JSON.stringify(S.records));
}

function migrateStoredSession() {
  const hasCurrentSession = !!localStorage.getItem(STORAGE_KEYS.records);
  const hadLegacySession = !hasCurrentSession && LEGACY_STORAGE_KEYS.records.some(key => localStorage.getItem(key));

  LEGACY_STORAGE_KEYS.records.forEach(key => {
    if (key !== STORAGE_KEYS.records) localStorage.removeItem(key);
  });
  LEGACY_STORAGE_KEYS.ui.forEach(key => {
    if (key !== STORAGE_KEYS.ui) localStorage.removeItem(key);
  });

  if (hadLegacySession) {
    showToast('Data format updated (v5). Please regenerate reports from source files for best accuracy.');
  }
}

function persistUiState() {
  if (!S.records.length) {
    localStorage.removeItem(STORAGE_KEYS.ui);
    return;
  }

  const state = {
    search: document.getElementById('search')?.value || '',
    branches: getSelectedValues('f-branch'),
    departments: getSelectedValues('f-dept'),
    statuses: getSelectedValues('f-status'),
    employees: getSelectedValues('f-emp'),
    dateFrom: document.getElementById('date-from')?.value || '',
    dateTo: document.getElementById('date-to')?.value || '',
    activeTab: getActiveTabName(),
    quickFilter: S.quickFilter || '',
    mobileFiltersOpen: !!S.mobileFiltersOpen
  };

  localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state));
}

function restoreUiState() {
  const raw = localStorage.getItem(STORAGE_KEYS.ui);
  if (!raw) return;

  try {
    const state = JSON.parse(raw);
    const search = document.getElementById('search');
    const dateFrom = document.getElementById('date-from');
    const dateTo = document.getElementById('date-to');
    S.quickFilter = state.quickFilter || '';
    S.mobileFiltersOpen = !!state.mobileFiltersOpen;

    if (search) search.value = state.search || '';
    if (dateFrom && state.dateFrom) dateFrom.value = state.dateFrom;
    if (dateTo && state.dateTo) dateTo.value = state.dateTo;

    setSelectedValues('f-branch', Array.isArray(state.branches) ? state.branches : []);
    setSelectedValues('f-dept', Array.isArray(state.departments) ? state.departments : []);
    setSelectedValues('f-status', Array.isArray(state.statuses) ? state.statuses : []);
    setSelectedValues('f-emp', Array.isArray(state.employees) ? state.employees : []);

    applyFilters();

    const targetTab = ['daily', 'summary', 'late', 'absent', 'early'].includes(state.activeTab)
      ? state.activeTab
      : 'daily';
    const tabBtn = getTabButton(targetTab);
    if (tabBtn) switchTab(targetTab, tabBtn);
    refreshMobileFilterUi();
  } catch (err) {
    console.error('Failed to restore UI state.', err);
    localStorage.removeItem(STORAGE_KEYS.ui);
  }
}

function restoreSavedSession() {
  const raw = localStorage.getItem(STORAGE_KEYS.records);
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    const records = Array.isArray(parsed)
      ? parsed
      : (parsed && Array.isArray(parsed.records) ? parsed.records : []);

    if (!records.length) return;

    records.forEach(r => {
      if (r.branch) {
        r.branch = r.branch.replace(/\s*-\s*/g, ' - ').replace(/\s+/g, ' ').trim();
      }
    });

    S.records = records;
    buildReport();
    renderSparklines();
    restoreUiState();
    showToast('Saved session restored.');
  } catch (err) {
    console.error('Failed to restore saved session.', err);
    localStorage.removeItem(STORAGE_KEYS.records);
    localStorage.removeItem(STORAGE_KEYS.ui);
  }
}

function bindDatePickers() {
  ['date-from', 'date-to'].forEach(id => {
    const input = document.getElementById(id);
    if (!input || input.dataset.pickerBound === '1') return;
    const openPicker = () => {
      if (typeof input.showPicker === 'function') {
        try { input.showPicker(); } catch (e) { }
      }
    };
    input.addEventListener('click', openPicker);
    input.addEventListener('focus', openPicker);
    input.dataset.pickerBound = '1';
  });
}

function getSelectedValues(id) {
  const el = document.getElementById(id);
  return (el.dataset.selected || '').split('||').filter(Boolean);
}

function setSelectedValues(id, values) {
  const el = document.getElementById(id);
  el.dataset.selected = values.join('||');
  refreshFilterButton(id);
  syncFilterChecks(id);
}

function bindFilterMenus() {
  document.addEventListener('click', e => {
    if (e.target.closest('.mfilter')) return;
    document.querySelectorAll('.mfilter.open').forEach(el => el.classList.remove('open'));
  });
}

function renderFilterMenu(id, items) {
  const el = document.getElementById(id);
  const label = FILTER_CONFIG[id].label;
  const selected = getSelectedValues(id);
  const summary = selected.length ? `${label} (${selected.length})` : label;
  el.className = 'mfilter';
  el.dataset.label = label;
  el.dataset.items = JSON.stringify(items);
  el.innerHTML = `
    <button class="mfilter-btn" type="button" onclick="toggleFilterMenu('${id}')">
      <span>${summary}</span>
      <span class="mfilter-arrow">▼</span>
    </button>
    <div class="mfilter-panel">
      <div class="mfilter-actions">
        <button type="button" class="mfilter-action" onclick="setFilterAll('${id}', true)">All</button>
        <button type="button" class="mfilter-action" onclick="setFilterAll('${id}', false)">Clear</button>
      </div>
      <div class="mfilter-list">
        ${items.map(item => `
          <label class="mfilter-option">
            <input type="checkbox" value="${item.value}" onchange="toggleFilterValue('${id}', this.value, this.checked)">
            <span>${item.label}</span>
          </label>
        `).join('')}
      </div>
    </div>
  `;
  syncFilterChecks(id);
}

function refreshFilterButton(id) {
  const el = document.getElementById(id);
  const btn = el.querySelector('.mfilter-btn span');
  if (!btn) return;
  const label = FILTER_CONFIG[id].label;
  const selected = getSelectedValues(id);
  if (id === 'f-status' && S.quickFilter === 'Late') {
    btn.textContent = 'Late Arrivals';
    return;
  }
  btn.textContent = selected.length ? `${label} (${selected.length})` : label;
}

function syncFilterChecks(id) {
  const selected = new Set(getSelectedValues(id));
  document.querySelectorAll(`#${id} input[type="checkbox"]`).forEach(box => {
    box.checked = selected.has(box.value);
  });
}

function toggleFilterMenu(id) {
  document.querySelectorAll('.mfilter.open').forEach(el => {
    if (el.id !== id) el.classList.remove('open');
  });
  document.getElementById(id).classList.toggle('open');
}

function toggleFilterValue(id, value, checked) {
  const selected = new Set(getSelectedValues(id));
  if (id === 'f-status') S.quickFilter = '';
  if (checked) selected.add(value); else selected.delete(value);
  setSelectedValues(id, [...selected]);
  applyFilters();
}

function setFilterAll(id, selectAll) {
  const el = document.getElementById(id);
  const items = JSON.parse(el.dataset.items || '[]');
  if (id === 'f-status') S.quickFilter = '';
  setSelectedValues(id, selectAll ? items.map(item => item.value) : []);
  applyFilters();
}

function clearData() {
  if (!confirm('Clear all data and restart?')) return;
  localStorage.removeItem(STORAGE_KEYS.records);
  localStorage.removeItem(STORAGE_KEYS.ui);
  // Note: Holiday data is preserved across sessions by default
  location.reload();
}

async function handleFiles(files, type) {
  const fileArr = Array.from(files);
  if (type === 'holiday') {
    AppState.raw.holidayFiles.push(...fileArr);
    AppState.validation.warnings = []; // Reset validation state for new sheet

    // Parse and merge holidays immediately
    for (const f of fileArr) {
      try {
        const parsed = await parseHolidaySheet(f);
        AppState.config.holidays.push(...parsed);
      } catch (e) { console.error('Holiday parse error', e); }
    }
    persistHolidays();
    renderPills('holiday');
    renderValidationCenter(AppState.validation.warnings);
    showToast(`${AppState.config.holidays.length} holidays loaded.`);
    return;
  }
  (type === 'excel' ? S.excelFiles : S.datFiles).push(...fileArr);
  renderPills(type); checkReady();
}
function dzDrag(e, id) { e.preventDefault(); document.getElementById(id).classList.add('drag-over') }
function dzLeave(id) { document.getElementById(id).classList.remove('drag-over') }
function dzDrop(e, type) {
  e.preventDefault();
  const dzMap = { excel: 'dz-excel', dat: 'dz-dat', holiday: 'dz-holiday' };
  const el = document.getElementById(dzMap[type]);
  if (el) el.classList.remove('drag-over');
  handleFiles(e.dataTransfer.files, type);
}
function removeFile(type, idx) {
  if (type === 'holiday') {
    AppState.raw.holidayFiles.splice(idx, 1);
  } else {
    (type === 'excel' ? S.excelFiles : S.datFiles).splice(idx, 1);
  }
  renderPills(type); checkReady();
}
function renderPills(type) {
  const arr = type === 'excel' ? S.excelFiles : type === 'holiday' ? AppState.raw.holidayFiles : S.datFiles;
  const el = document.getElementById('fl-' + type);
  if (!el) return;
  el.innerHTML = arr.map((f, i) =>
    `<div class="pill"><span>${f.name}</span><button class="pill-rm" onclick="removeFile('${type}',${i})">x</button></div>`
  ).join('');
}
function checkReady() {
  document.getElementById('btn-gen').disabled = !(S.excelFiles.length && S.datFiles.length);
}

function setProg(pct, msg) {
  document.getElementById('prog-wrap').style.display = 'block';
  document.getElementById('prog-fill').style.width = pct + '%';
  document.getElementById('prog-pct').textContent = pct + '%';
  document.getElementById('prog-msg').textContent = msg;
}
function hideProg() { document.getElementById('prog-wrap').style.display = 'none' }
function showToast(msg) {
  const t = document.getElementById('toast');
  document.getElementById('toast-msg').textContent = msg;
  t.style.display = 'flex'; setTimeout(() => t.style.display = 'none', 5000);
}

async function parseShift(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const map = {}; let hi = 0;
        for (let i = 0; i < Math.min(6, rows.length); i++) {
          const r2 = rows[i].map(c => String(c || '').toLowerCase());
          if (r2.some(c => c.includes('userid') || c.includes('user id') || c.includes('emp'))) { hi = i; break; }
        }
        const hdr = rows[hi].map(c => String(c || '').toLowerCase().trim());
        const col = k => hdr.findIndex(h => h.includes(k));

        const idC = col('userid') !== -1 ? col('userid') : col('user') !== -1 ? col('user') : col('emp');
        const nmC = col('particular') !== -1 ? col('particular') : col('name');
        const dpC = col('department') !== -1 ? col('department') : col('dept');
        const stC = col('shift start') !== -1 ? col('shift start') : col('start');
        const enC = col('shift end') !== -1 ? col('shift end') : col('end');

        // Branch column: try multiple common header variants
        let brC = col('branch');
        if (brC === -1) brC = col('site');
        if (brC === -1) brC = col('location');
        if (brC === -1) brC = col('office');
        if (brC === -1) brC = col('division');
        if (brC === -1) console.warn('[parseShift] No branch column found in Shift Master. Employees will have blank branch.');

        let skippedRows = 0;
        let missingBranch = 0;

        for (let i = hi + 1; i < rows.length; i++) {
          const row = rows[i];
          const uid = row[idC];
          if (!uid && !row[nmC]) continue; // completely empty row

          const uidStr = String(uid || '').trim();
          if (!uidStr || uidStr === '--' || uidStr === '-') {
            skippedRows++;
            continue; // skip invalid rows
          }

          const branchVal = (brC >= 0 && row[brC]) ? DataNormalizer.branch(row[brC]) : '';
          if (!branchVal) missingBranch++;

          if (uidStr) {
            map[uidStr] = {
              name: row[nmC] ? String(row[nmC]).trim() : 'User ' + uidStr,
              branch: branchVal,
              department: row[dpC] ? String(row[dpC]).trim() : '',
              shiftStart: DataNormalizer.time(row[stC]),
              shiftEnd: DataNormalizer.time(row[enC])
            };
          }
        }

        if (skippedRows > 0) AppState.validation.warnings.push(`ℹ ${skippedRows} malformed rows skipped in Shift Master (does not impact summary).`);
        if (missingBranch > 0) AppState.validation.warnings.push(`ℹ ${missingBranch} employees missing valid branch mapping.`);

        res(map);
      } catch (err) { rej(err); }
    };
    r.readAsArrayBuffer(file);
  });
}


async function parseDat(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => {
      const lines = e.target.result.split(/\r?\n/);
      const ps = [];
      let invalidDates = 0;

      for (const line of lines) {
        if (!line.trim()) continue;
        const pts = line.trim().split(/\s+/); if (pts.length < 2) continue;
        const ds = (pts[1] && pts[2] && !pts[1].includes(' ')) ? pts[1] + ' ' + pts[2] : pts[1];
        const d = new Date(ds);
        if (!isNaN(d.getTime())) {
          ps.push({ uid: String(pts[0]).trim(), dt: d });
        } else {
          invalidDates++;
        }
      }

      if (invalidDates > 0) AppState.validation.warnings.push(`ℹ ${invalidDates} unparseable timestamps skipped in ${file.name} (does not impact core summary).`);
      res(ps);
    };
    r.onerror = rej;
    r.readAsText(file);
  });
}

async function generateReport() {
  document.getElementById('toast').style.display = 'none';
  document.getElementById('btn-gen').disabled = true;
  AppState.validation.warnings = []; // Reset validation warnings on new run

  try {
    setProg(10, 'Reading shift master...');
    S.shiftMap = {};
    for (const f of S.excelFiles) Object.assign(S.shiftMap, await parseShift(f));
    setProg(30, 'Parsing attendance logs...');
    let punches = [];
    for (const f of S.datFiles) punches.push(...await parseDat(f));
    if (!punches.length) throw new Error('No valid punch records found in the uploaded files.');

    // ── VALIDATION LAYER ──
    setProg(40, 'Validating data...');
    const warnings = runValidation(S.shiftMap, punches);
    if (warnings.length) {
      console.warn('Validation warnings:', warnings);
      showToast(warnings[0]); // Show first warning
    }

    setProg(50, 'Deduplicating punches...');
    // ── PUNCH DEDUPLICATION: Remove duplicates within 2-minute window ──
    punches.sort((a, b) => (a.uid + a.dt.getTime()).localeCompare(b.uid + b.dt.getTime()) || a.dt - b.dt);
    const dedupPunches = [];
    let lastKey = '', lastTime = 0;
    for (const p of punches) {
      const key = p.uid;
      const time = p.dt.getTime();
      if (key === lastKey && Math.abs(time - lastTime) < 120000) continue; // 2-min window
      dedupPunches.push(p);
      lastKey = key;
      lastTime = time;
    }
    punches = dedupPunches;

    setProg(60, 'Calculating attendance...');
    const grouped = {};
    for (const p of punches) {
      const dk = p.dt.toISOString().split('T')[0];
      const key = p.uid + '|' + dk;
      if (!grouped[key]) grouped[key] = { uid: p.uid, date: dk, punches: [] };
      grouped[key].punches.push(p.dt);
    }
    const allPunchedDates = punches.map(p => p.dt.getTime());
    let minD = new Date(Math.min(...allPunchedDates)), maxD = new Date(Math.max(...allPunchedDates));

    // SAFETY GUARD: If logs span too many years, limit to current month only.
    if ((maxD.getTime() - minD.getTime()) > (62 * 24 * 60 * 60 * 1000)) {
      minD = new Date(maxD.getFullYear(), maxD.getMonth(), 1);
    }

    const dts = [];
    for (let d = new Date(minD); d <= maxD; d.setDate(d.getDate() + 1)) {
      dts.push(d.toISOString().split('T')[0]);
    }
    const uids = [...new Set([...Object.keys(S.shiftMap), ...punches.map(p => p.uid)])];
    const pad2 = n => String(n).padStart(2, '0');
    const m2t = m => m == null || isNaN(m) ? '--' : pad2(Math.floor(m / 60)) + ':' + pad2(m % 60);
    const fmtD = m => m <= 0 ? '--' : (Math.floor(m / 60) ? Math.floor(m / 60) + 'h ' : '') + ((m % 60) ? m % 60 + 'm' : '');

    let oddPunchCount = 0;
    S.records = [];
    for (const dt of dts) {
      for (const uid of uids) {
        const k = uid + '|' + dt;
        const g = grouped[k];
        const info = S.shiftMap[uid] || { name: 'User ' + uid, branch: '', department: '', shiftStart: null, shiftEnd: null };
        let ps = g ? g.punches.sort((a, b) => a - b) : [];

        // ── EDGE CASE: Filter invalid timestamps ──
        ps = ps.filter(p => !isNaN(p.getTime()));

        const dy = new Date(dt + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'short' });
        const hasShift = info.shiftStart !== null && info.shiftEnd !== null && (info.shiftStart !== 0 || info.shiftEnd !== 0);
        const sd = hasShift ? m2t(info.shiftStart) + ' - ' + m2t(info.shiftEnd) : '--';

        let first = null, last = null, hrs = 0, inM = null, outM = null, status = 'Absent', lateMins = 0, earlyMins = 0, earlyArrivalMins = 0, otMins = 0;
        const sDur = hasShift ? getShiftDurationMinutes(info.shiftStart, info.shiftEnd) : 0;

        // ── HOLIDAY CHECK: Use new isHoliday() function ──
        const isHol = isHoliday(dt, info.branch);

        if (ps.length === 0) {
          // Holiday overrides Absent
          if (isHol) status = 'Holiday';
          else if (dy === 'Sun') status = 'Week Off';
        }

        if (ps.length > 0) {
          first = ps[0]; last = ps[ps.length - 1];
          const workData = getWorkedMinutesFromPunches(ps);
          const workedMins = workData.mins;
          if (workData.isOdd && ps.length > 1) oddPunchCount++;

          hrs = workedMins / 60;
          inM = first.getHours() * 60 + first.getMinutes();
          outM = last.getHours() * 60 + last.getMinutes();
          status = 'Present';

          const sSt = info.shiftStart;
          const sEn = info.shiftEnd;

          lateMins = hasShift ? getLateMinutes(sSt, inM) : 0;
          if (hasShift && lateMins > 15) status = 'Late';

          earlyArrivalMins = hasShift ? getEarlyArrivalMinutes(sSt, inM) : 0;
          earlyMins = hasShift ? getEarlyMinutes(sSt, sEn, outM) : 0;

          // ── EDGE CASE: Odd punch count → Missed Punch ──
          if (ps.length === 1 || workData.isOdd) status = 'Missed Punch';
          else if (hrs < 0.25) status = 'Missed Punch';
          else if (hrs < 4.5) status = 'Half Day';

          // ── CORRECTED OT: late and early exit reduce effective work ──
          if (status !== 'Missed Punch') {
            otMins = hasShift ? getOvertimeMinutes(sSt, sEn, workedMins, lateMins, earlyMins) : 0;
          }

          // ── HOLIDAY OVERRIDE: Holiday overrides Late status ──
          if (isHol && ['Late', 'Absent'].includes(status)) {
            status = 'Holiday';
          }
        }

        let gapMins = 0;
        if (hasShift) {
          if (status === 'Absent') gapMins = sDur;
          else if (['Present', 'Late', 'Late (Comp)', 'Half Day'].includes(status)) gapMins = sDur - Math.round(hrs * 60);
        }

        if (status === 'Late' && gapMins <= 0) status = 'Late (Comp)';

        const gapFmt = gapMins === 0 ? '0m' : fmtD(Math.abs(gapMins));
        const gapClass = gapMins <= 0 ? 'g-ok' : 'g-err';

        const lateStr = lateMins > 0 ? fmtD(lateMins) : '--';
        const earlyStr = earlyMins > 0 ? fmtD(earlyMins) : '--';
        const earlyArriveStr = earlyArrivalMins > 0 ? fmtD(earlyArrivalMins) : '--';
        const empKey = makeEmployeeKey(uid, info.branch);

        S.records.push({
          uid, empKey, name: info.name, branch: info.branch, department: info.department,
          date: dt, day: dy, shiftDisplay: sd,
          firstIn: ps.length ? m2t(inM) : '--', lastOut: ps.length ? m2t(outM) : '--',
          hoursWorked: Math.round(hrs * 100) / 100, status,
          lateMins, earlyMins, earlyArrivalMins,
          lateBy: lateStr, earlyBy: earlyStr, earlyArrivalBy: earlyArriveStr,
          otMins, overtime: fmtD(otMins), punchCount: ps.length,
          gapMins, gapFmt, gapClass
        });
      }
    }

    // Aggregated Validation Alert
    if (oddPunchCount > 0) {
      AppState.validation.warnings.push(`ℹ ${oddPunchCount} records have a Missed Punch detected — hours calculated using absolute first and last punch.`);
      renderValidationCenter(AppState.validation.warnings); // Re-render to catch immediate additions
    }

    setProg(90, 'Saving session...');
    buildReport();
    renderSparklines();
    setProg(100, `Done — ${S.records.length} records processed.`);
    setTimeout(hideProg, 1200);
    document.getElementById('upload-section').style.display = 'none';
    document.getElementById('workspace-head').style.display = 'none';
    document.getElementById('btn-gen').style.display = 'none';
    document.getElementById('btn-clear').style.display = 'flex';
    syncAppMode();
    queueStickyLayoutSync();
  } catch (err) {
    showToast(err.message || 'Processing failed.');
    console.error(err);
    document.getElementById('btn-gen').disabled = false;
    hideProg();
  }
}

function buildReport() {
  detectMachineFailures();
  const uniq = k => [...new Set(S.records.map(r => r[k]).filter(Boolean))].sort();
  renderFilterMenu('f-branch', uniq('branch').map(v => ({ value: v, label: v })));
  renderFilterMenu('f-dept', uniq('department').map(v => ({ value: v, label: v })));
  renderFilterMenu('f-status', STATUS_OPTIONS.map(v => ({ value: v, label: v })));
  const emps = [...new Set(S.records.map(r => `${r.uid}|${r.name}`))].sort().map(e => { const [uid, nm] = e.split('|'); return { uid, name: nm } });
  renderFilterMenu('f-emp', emps.map(e => ({ value: e.uid, label: `${e.name} (${e.uid})` })));

  S.lateRecs = S.records.filter(r => r.lateMins > 0);
  S.absentRecs = S.records.filter(r => r.status === 'Absent');
  S.earlyRecs = S.records.filter(r => r.earlyMins > 0);

  // Auto-set dates to data range
  const dates = S.records.map(r => r.date).filter(Boolean).sort();
  if (dates.length) {
    document.getElementById('date-from').value = dates[0];
    document.getElementById('date-to').value = dates[dates.length - 1];
  }

  document.getElementById('dl-wrap').style.display = 'flex';
  document.getElementById('stats-row').style.display = 'grid';
  document.getElementById('tabs-row').style.display = 'flex';
  document.getElementById('tab-daily').style.display = 'block';
  const dBadge = document.getElementById('tb-daily-badge');
  if (dBadge) dBadge.textContent = S.records.length;
  document.getElementById('tb-summary').textContent = [...new Set(S.records.map(r => r.uid))].size;
  document.getElementById('tb-late').textContent = S.lateRecs.length;
  document.getElementById('tb-absent').textContent = S.absentRecs.length;
  document.getElementById('tb-early').textContent = S.earlyRecs.length;
  S.filtered = [...S.records]; S.page = 1;
  renderAuditSelects();
  updateDailyTitle();
  renderStats(S.records); renderTable(); renderSubTables(); renderInsights();
  persistRecords();
  persistUiState();
  syncAppMode();
  syncResponsiveUi();
}

function detectMachineFailures() {
  const r = S.records; if (!r.length) return;

  // Group by branch+date
  const grouped = {};
  r.forEach(x => {
    const bk = x.date + '|' + (x.branch || 'Default');
    if (!grouped[bk]) grouped[bk] = {
      total: 0, absent: 0, punched: 0, totalPunches: 0,
      date: x.date, branch: x.branch
    };
    grouped[bk].total++;
    if (x.status === 'Absent') grouped[bk].absent++;
    if (!['Absent', 'Week Off', 'Holiday', 'System Error'].includes(x.status)) grouped[bk].punched++;
    grouped[bk].totalPunches += (x.punchCount || 0);
  });

  // Calculate branch averages for comparison
  const branchAvgPunches = {};
  const branchDayCounts = {};
  Object.values(grouped).forEach(g => {
    const br = g.branch || 'Default';
    if (!branchAvgPunches[br]) { branchAvgPunches[br] = 0; branchDayCounts[br] = 0; }
    branchAvgPunches[br] += g.totalPunches;
    branchDayCounts[br]++;
  });
  Object.keys(branchAvgPunches).forEach(br => {
    branchAvgPunches[br] = branchDayCounts[br] ? branchAvgPunches[br] / branchDayCounts[br] : 0;
  });

  S.failureDates = [];
  let affectedRecords = 0;

  Object.keys(grouped).forEach(bk => {
    const g = grouped[bk];
    const dy = new Date(g.date + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'short' });
    if (dy === 'Sun') return;

    // Holiday check with debug logging
    const holCheck = isHoliday(g.date, g.branch);
    if (holCheck) {
      console.log(`[MachineFailure] Skipping ${g.date} ${g.branch} — holiday: ${holCheck.name}`);
      return;
    }

    if (g.total < 3) return;

    const absentRate = g.absent / g.total;
    const br = g.branch || 'Default';
    const avgPunches = branchAvgPunches[br] || 0;
    const punchRatio = avgPunches > 0 ? g.totalPunches / avgPunches : 1;

    // Enhanced detection conditions
    const isZeroPunchDay = g.punched === 0 && g.absent >= 3;
    const isNearZeroDay = g.punched <= 2 && g.absent >= 3 && absentRate >= 0.80;
    const isHighSpike = absentRate >= 0.40 && g.absent >= 3;
    const isModerateSpike = g.total >= 10 && absentRate >= 0.30 && g.absent >= 4;
    // NEW: Absence >15% AND punch volume <50% of branch average
    const isLowPunchAnomaly = absentRate > 0.15 && punchRatio < 0.5 && g.absent >= 3;

    const isAnomaly = isZeroPunchDay || isNearZeroDay || isHighSpike || isModerateSpike || isLowPunchAnomaly;

    if (isAnomaly) {
      let reason = 'high absence spike';
      if (isZeroPunchDay) reason = 'zero punch day (total machine failure)';
      else if (isNearZeroDay) reason = 'near-zero punches (likely machine failure)';
      else if (isLowPunchAnomaly) reason = `low punch volume (${Math.round(punchRatio * 100)}% of avg) + ${Math.round(absentRate * 100)}% absent`;
      else if (isHighSpike) reason = `${Math.round(absentRate * 100)}% absence spike`;
      else if (isModerateSpike) reason = `${Math.round(absentRate * 100)}% absence in large branch`;

      S.failureDates.push({
        date: g.date,
        branch: g.branch || 'Unknown Branch',
        reason,
        affected: g.absent,
        total: g.total,
        pct: Math.round(absentRate * 100)
      });

      r.forEach(x => {
        if (x.date === g.date && x.branch === g.branch && x.status === 'Absent') {
          x.status = 'System Error';
          x.gapMins = 0; x.gapFmt = '0m'; x.gapClass = 'g-ok';
          x.lateMins = 0; x.earlyMins = 0; x.lateBy = '--'; x.earlyBy = '--';
          affectedRecords++;
        }
      });
    }
  });

  // ── SAFETY POST-PASS: Holiday ALWAYS overrides System Error ──
  // This catches individual records whose branch normalization differs from the group
  r.forEach(x => {
    if (x.status === 'System Error') {
      const hol = isHoliday(x.date, x.branch);
      if (hol) {
        console.log(`[Holiday Override] ${x.date} ${x.uid} ${x.branch} → Holiday (was System Error): ${hol.name}`);
        x.status = 'Holiday';
      }
    }
  });

  // Calculate Data Reliability
  const totalWorkingRecords = r.filter(x => !['Week Off', 'Holiday'].includes(x.status)).length;
  const sysErrorRecords = r.filter(x => x.status === 'System Error').length;
  AppState.processed.dataReliability = totalWorkingRecords > 0
    ? Math.round(((totalWorkingRecords - sysErrorRecords) / totalWorkingRecords) * 100)
    : 100;
}

function quickFilter(s) {
  S.quickFilter = s === 'Late' ? 'Late' : '';
  setSelectedValues('f-status', s === 'Late' ? [] : [s]);
  document.querySelectorAll('.mfilter.open').forEach(el => el.classList.remove('open'));
  switchTab('daily', document.querySelector('[onclick*="switchTab(\'daily\'"]'));
  applyFilters();
}

function renderInsights() {
  const r = S.records; if (!r.length) return;
  const ins = document.getElementById('insights-panel'); ins.style.display = 'grid';
  const dash = v => v || '<span class="dash">--</span>';
  const brData = {};
  r.forEach(x => {
    if (!x.branch) return;
    if (!brData[x.branch]) brData[x.branch] = { days: 0, present: 0 };
    brData[x.branch].days++;
    if (isAttendedStatus(x.status)) brData[x.branch].present++;
  });
  let topBr = '', topPct = -1;
  Object.keys(brData).forEach(b => {
    const pct = Math.round((brData[b].present / brData[b].days) * 100);
    if (pct > topPct) { topPct = pct; topBr = b; }
  });

  const latePct = Math.round((S.lateRecs.length / r.length) * 100);
  const totalAtt = Math.round((r.filter(x => isAttendedStatus(x.status)).length / r.length) * 100);
  const reliability = AppState.processed.dataReliability;
  const holidayCount = AppState.config.holidays.length;
  const warnings = AppState.validation.warnings;

  let html = `
    <div class="insight-item">
      <div class="insight-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 14.8 8.7 21 9.5l-4.5 4.3 1.1 6.2L12 17.1 6.4 20l1.1-6.2L3 9.5l6.2-.8z"></path></svg></div>
      <div class="insight-copy"><span class="insight-kicker">Top Branch</span><div class="insight-txt"><strong>${dash(topBr)}</strong> is leading with <strong>${topPct}%</strong> attendance stability.</div></div>
    </div>
    <div class="insight-item">
      <div class="insight-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 18h16"></path><path d="M7 14 10 11l3 2 4-5"></path></svg></div>
      <div class="insight-copy"><span class="insight-kicker">Stability</span><div class="insight-txt"><strong>${totalAtt}%</strong> overall attendance. ${totalAtt > 85 ? 'Very healthy.' : 'Needs review.'}</div></div>
    </div>`;

  if (S.failureDates.length) {
    html += `
    <div class="insight-item" onclick="quickFilter('System Error')" style="cursor:pointer">
      <div class="insight-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 9v4"></path><path d="M12 17h.01"></path><path d="M10.3 4.9 2.6 18.2a1 1 0 0 0 .9 1.5h17a1 1 0 0 0 .9-1.5L13.7 4.9a1 1 0 0 0-1.7 0z"></path></svg></div>
      <div class="insight-copy"><span class="insight-kicker">System Alert</span><div class="insight-txt"><strong>${S.failureDates.length} incident${S.failureDates.length > 1 ? 's' : ''}</strong> across <strong>${[...new Set(S.failureDates.map(f => f.branch))].length} branch${[...new Set(S.failureDates.map(f => f.branch))].length > 1 ? 'es' : ''}</strong>. Data reliability: <strong>${reliability}%</strong>.</div></div>
    </div>`;
  } else {
    html += `
    <div class="insight-item">
      <div class="insight-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"></circle><path d="M12 8v4l3 2"></path></svg></div>
      <div class="insight-copy"><span class="insight-kicker">Punctuality</span><div class="insight-txt"><strong>${latePct}%</strong> of records are late. ${latePct < 10 ? 'Excellent discipline.' : 'Review shift overlaps.'}</div></div>
    </div>`;
  }

  // Validation warnings insight
  if (warnings.length) {
    html += `
    <div class="insight-item">
      <div class="insight-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 9v4"></path><path d="M12 17h.01"></path><circle cx="12" cy="12" r="9"></circle></svg></div>
      <div class="insight-copy"><span class="insight-kicker">Data Quality</span><div class="insight-txt">${warnings[0]}${warnings.length > 1 ? ` (+${warnings.length - 1} more)` : ''}</div></div>
    </div>`;
  }

  // Holiday info
  if (holidayCount) {
    html += `
    <div class="insight-item">
      <div class="insight-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="18" rx="3"></rect><path d="M8 2v4"></path><path d="M16 2v4"></path><path d="M3 10h18"></path></svg></div>
      <div class="insight-copy"><span class="insight-kicker">Holiday Calendar</span><div class="insight-txt"><strong>${holidayCount}</strong> holidays loaded for branch-wise detection.</div></div>
    </div>`;
  }

  ins.innerHTML = html;
}

function switchTab(name, btn) {
  ['daily', 'summary', 'late', 'absent', 'early'].forEach(t => document.getElementById('tab-' + t).style.display = t === name ? 'block' : 'none');
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  if (name === 'summary') renderSummary();
  persistUiState();
  syncResponsiveUi();
}

function syncAppMode() {
  const hasData = S.records.length > 0;

  // Hide instructions when data is active
  const head = document.getElementById('workspace-head');
  if (head) head.style.display = hasData ? 'none' : 'flex';

  // Hide upload section when data is active
  const up = document.getElementById('upload-section');
  if (up) up.style.display = hasData ? 'none' : 'grid';

  const btnGen = document.getElementById('btn-gen');
  if (btnGen) btnGen.style.display = hasData ? 'none' : '';

  const btnClear = document.getElementById('btn-clear');
  if (btnClear) btnClear.style.display = hasData ? 'flex' : 'none';

  if (!hasData) return;

  // Highlight active filters
  ['f-branch', 'f-dept', 'f-status', 'f-emp'].forEach(id => {
    const wrap = document.getElementById(id);
    if (!wrap) return;
    const btn = wrap.querySelector('.mfilter-btn');
    if (btn) {
      const selected = getSelectedValues(id);
      const isQuickLate = id === 'f-status' && S.quickFilter === 'Late';
      btn.classList.toggle('active', selected.length > 0 || isQuickLate);
    }
  });

  const dlWrap = document.getElementById('dl-wrap');
  if (dlWrap) dlWrap.style.display = hasData ? 'flex' : 'none';
}

function applyFilters() {
  const q = document.getElementById('search').value.toLowerCase();
  const branches = getSelectedValues('f-branch');
  const depts = getSelectedValues('f-dept');
  const statuses = getSelectedValues('f-status');
  const empIds = getSelectedValues('f-emp');
  const fv = document.getElementById('date-from').value;
  const tv = document.getElementById('date-to').value;
  const hasData = S.records.length > 0;
  const dlWrap = document.getElementById('dl-wrap');
  if (dlWrap) dlWrap.style.display = hasData ? 'flex' : 'none';
  let from = -Infinity, to = Infinity;
  if (fv) from = new Date(fv).setHours(0, 0, 0, 0);
  if (tv) to = new Date(tv).setHours(23, 59, 59, 999);

  S.filtered = S.records.filter(r => {
    if (q && !(r.name.toLowerCase().includes(q) || r.uid.includes(q))) return false;
    if (branches.length && !branches.includes(r.branch)) return false;
    if (depts.length && !depts.includes(r.department)) return false;
    if (statuses.length && !statuses.includes(r.status)) return false;
    if (S.quickFilter === 'Late' && !(Number(r.lateMins || 0) > 0)) return false;
    if (empIds.length && !empIds.includes(r.uid)) return false;
    const rt = new Date(r.date).getTime();
    if (fv && rt < from) return false;
    if (tv && rt > to) return false;
    return true;
  });

  S.lateRecs = S.filtered.filter(r => r.lateMins > 0);
  S.absentRecs = S.filtered.filter(r => r.status === 'Absent');
  S.earlyRecs = S.filtered.filter(r => r.earlyMins > 0);

  const dBadge = document.getElementById('tb-daily-badge');
  if (dBadge) dBadge.textContent = S.filtered.length;
  document.getElementById('tb-summary').textContent = [...new Set(S.filtered.map(r => r.uid))].size;
  document.getElementById('tb-late').textContent = S.lateRecs.length;
  document.getElementById('tb-absent').textContent = S.absentRecs.length;
  document.getElementById('tb-early').textContent = S.earlyRecs.length;

  S.page = 1; renderStats(S.filtered); renderTable();
  renderSubTables(); renderSummary();
  updateDailyTitle();
  persistUiState();
  syncAppMode();
  refreshMobileFilterUi();
}

function sortBy(col) {
  S.sortDir = S.sortCol === col ? S.sortDir * -1 : 1; S.sortCol = col;
  S.filtered.sort((a, b) => {
    const av = a[col] ?? '', bv = b[col] ?? '';
    if (['hoursWorked', 'lateMins', 'earlyMins', 'gapMins', 'earlyArrivalMins'].includes(col)) return (Number(av) - Number(bv)) * S.sortDir;
    return String(av).localeCompare(String(bv)) * S.sortDir;
  });
  renderTable();
}
function sortSub(type, col) {
  const arr = type === 'late' ? S.lateRecs : type === 'absent' ? S.absentRecs : S.earlyRecs;
  arr.sort((a, b) => ['lateMins', 'earlyMins', 'earlyArrivalMins'].includes(col) ? (Number(a[col] || 0) - Number(b[col] || 0)) : String(a[col] || '').localeCompare(String(b[col] || '')));
  renderSubTables();
}

function renderTable() {
  const total = S.filtered.length, start = (S.page - 1) * S.perPage;
  const slice = S.filtered.slice(start, start + S.perPage);
  const bdg = s => { const m = { Present: 'present', Late: 'late', Absent: 'absent', 'Half Day': 'half', 'Missed Punch': 'missed', 'Week Off': 'weekoff', 'Holiday': 'holiday', 'Late (Comp)': 'late-comp', 'System Error': 'syserr' }; return `<span class="badge b-${m[s] || 'present'}" data-status="${s}">${s}</span>` };
  const hc = h => h >= 8 ? 'h-ok' : h >= 4 ? 'h-low' : 'h-zero';
  document.getElementById('table-body').innerHTML = slice.map((r, i) => `
<tr onclick="toggleExp(${start + i})" id="row-${start + i}" style="animation-delay: ${i * 0.04}s">
  <td class="td-emp" data-label="Employee" onclick="event.stopPropagation();showEmpProfile('${r.uid}')" style="cursor:pointer">
    <strong>${r.name}</strong><small>ID ${r.uid}</small>
  </td>
  <td title="${r.branch} / ${r.department}" data-label="Branch/Dept"><div class="org-cell"><strong>${r.branch || '--'}</strong><small>${r.department || '--'}</small></div></td>
  <td class="date-cell" data-label="Date"><span class="date-main mono">${r.date}</span><span class="day-span">${r.day}</span></td>
  <td class="mono" data-label="In">${r.firstIn}</td>
  <td class="mono" data-label="Out">${r.lastOut}</td>
  <td class="mono ${hc(r.hoursWorked)}" data-label="Hrs">${r.hoursWorked}h</td>
  <td class="overtime-v ${r.otMins > 120 ? 'high-ot' : ''}" data-label="OT">${r.otMins > 0 ? r.overtime : '--'}</td>
  <td data-label="Status">${bdg(r.status)}</td>
  <td class="${r.gapClass}" data-label="Gap">${r.gapMins <= 0 ? '-' : ''}${r.gapFmt}</td>
  <td class="mono early-in-v" data-label="Early In">${r.earlyArrivalBy}</td>
  <td class="late-v" data-label="Late">${r.lateMins > 0 ? r.lateBy : '--'}</td>
  <td class="early-v" data-label="Early Out">${r.earlyMins > 0 ? r.earlyBy : '--'}</td>
</tr>
<tr class="exp-row" id="exp-${start + i}">
  <td class="exp-cell" colspan="14">
    <div class="exp-inner">
      <div class="exp-stat"><label>Shift</label><span>${r.shiftDisplay}</span></div>
      <div class="exp-stat"><label>Arrived Early</label><span class="early-in-v">${r.earlyArrivalMins > 0 ? r.earlyArrivalBy : '0m'}</span></div>
      <div class="exp-stat"><label>Overtime</label><span class="overtime-v">${r.otMins > 0 ? r.overtime : '0m'}</span></div>
      <div class="exp-stat"><label>Hours Worked</label><span class="${hc(r.hoursWorked)}">${r.hoursWorked}h</span></div>
      <div class="exp-stat"><label>Gap Time</label><span class="${r.gapClass}">${r.gapFmt}</span></div>
    </div>
  </td>
</tr>`).join('');
  const end = Math.min(start + S.perPage, total);
  document.getElementById('page-info').textContent = total === 0 ? 'No records' : `Showing ${start + 1}-${end} of ${total} records`;
  const pages = Math.ceil(total / S.perPage); let html = '';
  if (pages > 1) {
    html += `<button class="pBtn" onclick="goPage(${S.page - 1})" ${S.page === 1 ? 'disabled' : ''}>&lt;</button>`;
    for (let i = 1; i <= pages; i++) {
      if (i === 1 || i === pages || Math.abs(i - S.page) <= 1) html += `<button class="pBtn ${i === S.page ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
      else if (i === S.page - 2 || i === S.page + 2) html += `<button class="pBtn" style="pointer-events:none">...</button>`;
    }
    html += `<button class="pBtn" onclick="goPage(${S.page + 1})" ${S.page === pages ? 'disabled' : ''}>&gt;</button>`;
  }
  document.getElementById('page-ctrls').innerHTML = html;
}

function toggleExp(idx) {
  if (window.innerWidth <= 768) return;
  const row = document.getElementById('row-' + idx);
  const exp = document.getElementById('exp-' + idx);
  const open = exp.classList.contains('open');
  document.querySelectorAll('.exp-row.open').forEach(r => r.classList.remove('open'));
  document.querySelectorAll('tbody tr.expanded').forEach(r => r.classList.remove('expanded'));
  if (!open) { exp.classList.add('open'); row.classList.add('expanded') }
}
function goPage(p) { const pages = Math.ceil(S.filtered.length / S.perPage); if (p >= 1 && p <= pages) { S.page = p; renderTable() } }

function renderStats(recs) {
  const empSet = new Set(recs.map(r => r.uid));
  const dateSet = new Set(recs.map(r => r.date));
  const present = recs.filter(r => ['Present', 'Late', 'Late (Comp)'].includes(r.status)).length;
  const late = recs.filter(r => r.lateMins > 0).length;
  const absent = recs.filter(r => r.status === 'Absent').length;
  const validRecs = recs.filter(r => r.status !== 'System Error').length;
  const pct = validRecs ? Math.round(present / validRecs * 100) : 0;
  document.getElementById('st-emp').textContent = empSet.size;
  document.getElementById('st-emp-sub').textContent = recs.length + ' total records';
  document.getElementById('st-days').textContent = dateSet.size;
  document.getElementById('st-days-sub').textContent = 'unique dates';
  document.getElementById('st-att').textContent = pct + '%';
  document.getElementById('st-att-sub').textContent = present + ' present days';
  document.getElementById('st-late').textContent = late;
  document.getElementById('st-late-sub').textContent = recs.length ? Math.round(late / recs.length * 100) + '% of records' : '';
  document.getElementById('st-abs').textContent = absent;
  document.getElementById('st-abs-sub').textContent = recs.length ? Math.round(absent / recs.length * 100) + '% of records' : '';

  // Data Reliability Indicator
  const reliability = AppState.processed.dataReliability;
  const reliClass = reliability > 90 ? 'high' : reliability > 70 ? 'medium' : 'low';
  const reliLabel = reliability > 90 ? '✓ Reliable' : reliability > 70 ? '⚠ Moderate' : '✗ Low';
  const reliEl = document.getElementById('data-reliability');
  if (reliEl) {
    reliEl.className = 'data-confidence ' + reliClass;
    reliEl.textContent = `${reliLabel} (${reliability}%)`;
    reliEl.style.display = recs.length ? 'inline-flex' : 'none';
  }
}



function renderSummary() {
  const r = S.filtered;
  const container = document.getElementById('summary-content');
  if (!container) return;
  const summarySearchValue = document.getElementById('search-summary')?.value || '';
  const searchVal = summarySearchValue.toLowerCase();

  if (!r.length) {
    container.innerHTML = '<div class="empty-state">No data for summary</div>';
    return;
  }

  // Generate Branch Leaderboard
  const brs = {};
  r.forEach(x => {
    if (!x.branch) return;
    if (!brs[x.branch]) brs[x.branch] = { name: x.branch, total: 0, att: 0, ot: 0 };
    brs[x.branch].total++;
    if (isAttendedStatus(x.status)) brs[x.branch].att++;
    brs[x.branch].ot += (x.otMins || 0);
  });

  const sortedBr = Object.values(brs).map(b => ({
    ...b,
    score: Math.round((b.att / b.total) * 100),
    avgOt: Math.round(b.ot / b.total)
  })).sort((a, b) => b.score - a.score);

  // Group by Employee for the Summary Table
  const empSummary = {};
  r.forEach(x => {
    if (!empSummary[x.uid]) {
      empSummary[x.uid] = {
        uid: x.uid,
        name: x.name,
        branch: x.branch,
        dept: x.department,
        total: 0,
        present: 0,
        late: 0,
        ot: 0,
        _h: 0
      };
    }
    const s = empSummary[x.uid];
    s.total++;
    s._h += (x.hoursWorked || 0);
    if (['Present', 'Late', 'Late (Comp)'].includes(x.status)) s.present++;
    if (x.lateMins > 0) s.late++;
    s.ot += (x.otMins || 0);
  });

  let sortedEmps = Object.values(empSummary).map(e => ({
    ...e,
    attPct: Math.round((e.present / e.total) * 100),
    avgHrs: (e._h / e.total).toFixed(1)
  })).sort((a, b) => b.attPct - a.attPct);

  if (searchVal) {
    sortedEmps = sortedEmps.filter(e =>
      e.name.toLowerCase().includes(searchVal) ||
      e.uid.toLowerCase().includes(searchVal) ||
      (e.branch || '').toLowerCase().includes(searchVal) ||
      (e.dept || '').toLowerCase().includes(searchVal)
    );
  }

  // Graphical Helper: Circular Progress Ring
  const ring = (pct, color, size = 64) => {
    const r = size * 0.4;
    const circ = 2 * Math.PI * r;
    const off = circ - (pct / 100) * circ;
    return `
      <div class="stat-ring-wrap" style="width:${size}px; height:${size}px">
        <svg viewBox="0 0 ${size} ${size}">
          <circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="var(--border)" stroke-width="4" />
          <circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="${color}" stroke-width="4" stroke-linecap="round" 
            style="stroke-dasharray:${circ}; stroke-dashoffset:${off}; transition: stroke-dashoffset 1s ease" />
        </svg>
        <span class="ring-label">${pct}%</span>
      </div>
    `;
  };

  const avgStability = Math.round(sortedBr.reduce((s, b) => s + b.score, 0) / Math.max(1, sortedBr.length));
  const totalOtHrs = Math.round(sortedEmps.reduce((s, e) => s + e.ot, 0) / 60);

  container.innerHTML = `
    <div class="summary-shell anim-fade">
      <div class="briefing-kpis">
        <div class="b-kpi glass-tile">
          ${ring(avgStability, 'var(--teal)', 72)}
          <div class="b-kpi-data"><span class="b-kpi-label">Network Stability</span><small>Avg Attendance %</small></div>
        </div>
        <div class="b-kpi glass-tile">
          <div class="b-stat-box" style="color:var(--violet)"><span class="b-stat-val">${totalOtHrs}h</span><small>Total OT Captured</small></div>
          <div class="b-kpi-data"><span class="b-kpi-label">Labor Utilization</span></div>
        </div>
        <div class="b-kpi glass-tile">
          <div class="b-stat-box" style="color:var(--red)"><span class="b-stat-val">${S.failureDates?.length || 0}</span><small>System Alerts</small></div>
          <div class="b-kpi-data"><span class="b-kpi-label">Terminal Health</span></div>
        </div>
      </div>

      <div class="summary-section">
        <div class="section-header-row briefing-header">
          <h3>Branch Performance Leaderboard</h3>
          <span class="badge pills-info">${sortedBr.length} Branches</span>
        </div>
        <div class="summary-table-wrap">
          <table class="summary-table">
            <thead><tr><th>Branch</th><th>Stability</th><th>Avg OT</th><th>Grade</th></tr></thead>
            <tbody>
              ${sortedBr.map(b => `
                <tr class="clickable-row" onclick="filterByBranch('${b.name.replace(/'/g, "\\'")}')">
                  <td class="summary-main"><strong>${b.name}</strong></td>
                  <td class="summary-progress">
                    <div class="progress-mini"><div class="pm-bar"><div style="width:${b.score}%"></div></div><span>${b.score}%</span></div>
                  </td>
                  <td class="mono">${b.avgOt}m</td>
                  <td><span class="badge ${b.score > 85 ? 'pills-success' : b.score > 70 ? 'pills-warning' : 'pills-danger'}">${b.score > 85 ? 'A' : b.score > 70 ? 'B' : 'C'}</span></td>
                </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>

      <div class="section-header-row briefing-header" style="margin-top:32px">
        <h3>Efficiency Intelligence Grid</h3>
        <span class="badge pills-glass">${sortedEmps.length} ANALYTICAL PROFILES</span>
      </div>

      <div class="intelligence-grid-dashboard">
        ${sortedEmps.slice(0, 72).map(e => {
    const tags = [];
    if (e.attPct > 95) tags.push('<span class="tag-stable">STABLE</span>');
    if (e.ot > 300) tags.push('<span class="tag-ot">OT WARRIOR</span>');
    if (e.late > 2) tags.push('<span class="tag-late">LATE PATTERN</span>');
    if (Number(e.avgHrs) > 9) tags.push('<span class="tag-early">MORNING LARK</span>');

    return `
            <div class="dashboard-tile anim-fade" onclick="window.showEmpProfile('${e.uid}')">
              <div class="tile-top">
                <div class="tile-avatar" style="background:${avatarCol(e.name)}">${initials(e.name)}</div>
                <div class="tile-title"><strong>${e.name}</strong><small>ID: ${e.uid} | ${e.branch || 'General'}</small></div>
                <div class="tile-ring-mini">${ring(e.attPct, e.attPct < 85 ? 'var(--orange)' : 'var(--teal)', 44)}</div>
              </div>
              <div class="tile-mid">
                <div class="tm-stat"><small>OT (HRS)</small><strong>${(e.ot / 60).toFixed(1)}h</strong></div>
                <div class="tm-stat"><small>LATE</small><strong style="${e.late > 3 ? 'color:var(--red)' : ''}">${e.late} Days</strong></div>
                <div class="tm-stat"><small>AVG DAY</small><strong>${e.avgHrs}h</strong></div>
              </div>
              <div class="tile-tags">${tags.length ? tags.join('') : '<span class="tag-consistent">CONSISTENT</span>'}</div>
            </div>
          `;
  }).join('')}
      </div>
      ${sortedEmps.length > 72 ? `<div class="dashboard-more">Searching top performers. Use search to view ${sortedEmps.length - 72} more.</div>` : ''}
    </div>
  `;
  const summarySearch = document.getElementById('search-summary');
  if (summarySearch) summarySearch.value = searchVal;
}

function renderSubTables() {
  const row = (x, i, type) => {
    const isComp = x.earlyArrivalMins >= x.earlyMins && x.earlyMins > 0;

    // Safety check for missing properties
    const earlyArrival = x.earlyArrivalBy || '--';
    const lastOut = x.lastOut || '--';
    const earlyBy = x.earlyBy || '--';

    if (type === 'early') {
      return `
      <tr class="anim-fade ${isComp ? 'is-comp' : ''}" style="animation-delay: ${i * 0.04}s">
        <td class="td-emp" data-label="Employee"><strong>${x.name}</strong><small>ID ${x.uid}</small></td>
        <td title="${x.branch} / ${x.department}" data-label="Branch/Dept">
          <div class="org-cell"><strong>${x.branch || '--'}</strong><small>${x.department || '--'}</small></div>
        </td>
        <td class="mono" data-label="Date">${x.date} <span class="day-span">${x.day}</span></td>
        <td data-label="Shift"><span class="shift-chip">${x.shiftDisplay}</span></td>
        <td class="mono" data-label="IN">${x.firstIn}</td>
        <td class="mono early-in-v" data-label="Arrived Early">${earlyArrival}</td>
        <td class="mono early-v" data-label="Left At">${lastOut}</td>
        <td class="early-v" data-label="Left Early By">${earlyBy} ${isComp ? '<span class="comp-tag">Compensated</span>' : ''}</td>
      </tr>`;
    }

    return `
    <tr class="anim-fade" style="animation-delay: ${i * 0.04}s">
      <td class="td-emp" data-label="Employee"><strong>${x.name}</strong><small>ID ${x.uid}</small></td>
      <td title="${x.branch}" data-label="Branch">${x.branch || '--'}</td>
      <td title="${x.department}" data-label="Department">${x.department || '--'}</td>
      <td class="mono" data-label="Date">${x.date}</td>
      <td style="color:var(--ink3);font-size:12px" data-label="Day">${x.day}</td>
      <td data-label="Shift"><span class="shift-chip">${x.shiftDisplay}</span></td>
      ${type === 'late' ? `
        <td class="mono late-v" data-label="Arrived">${x.firstIn}</td>
        <td class="late-v" data-label="Late By">${x.lateBy}</td>
      `: `
        <td data-label="Status"><span class="badge b-absent">Absent</span></td>
      `}
    </tr>`;
  };

  document.getElementById('late-body').innerHTML = S.lateRecs.length ? S.lateRecs.slice(0, 50).map((x, i) => row(x, i, 'late')).join('') : '<tr><td colspan="8" style="padding:24px;text-align:center;color:var(--ink3)">No late arrivals</td></tr>';
  document.getElementById('absent-body').innerHTML = S.absentRecs.length ? S.absentRecs.slice(0, 50).map((x, i) => row(x, i, 'absent')).join('') : '<tr><td colspan="7" style="padding:24px;text-align:center;color:var(--ink3)">No absences</td></tr>';
  document.getElementById('early-body').innerHTML = S.earlyRecs.length ? S.earlyRecs.slice(0, 50).map((x, i) => row(x, i, 'early')).join('') : '<tr><td colspan="9" style="padding:24px;text-align:center;color:var(--ink3)">No early departures</td></tr>';
}

function downloadExcel() {
  const data = S.filtered;
  if (!data.length) return showToast('No records to export.');
  const wb = XLSX.utils.book_new();

  // 1. DASHBOARD OVERVIEW SHEET
  const attPct = Math.round((data.filter(x => isAttendedStatus(x.status)).length / data.length) * 100);
  const overview = [
    { 'Metric': 'Report Period', 'Value': `${document.getElementById('date-from').value} to ${document.getElementById('date-to').value}` },
    { 'Metric': 'Total Selected Records', 'Value': data.length },
    { 'Metric': 'Overall Attendance %', 'Value': attPct + '%' },
    { 'Metric': 'Late Arrivals Count', 'Value': S.lateRecs.length },
    { 'Metric': 'Absenteeism Count', 'Value': S.absentRecs.length },
    { 'Metric': 'Early Departures Count', 'Value': S.earlyRecs.length },
    { 'Metric': 'Detected Machine Failures', 'Value': S.failureDates.length }
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(overview), 'Dashboard Overview');

  // 2. SUMMARY BY EMPLOYEE
  const sm = {};
  data.forEach(r => {
    if (!sm[r.uid]) sm[r.uid] = {
      'Emp ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Department': r.department,
      'Total Days': 0, 'Present': 0, 'Late': 0, 'Half Day': 0, 'Absent': 0, 'Avg Hrs/Day': 0, 'Att %': '', 'Total OT (Mins)': 0, '_h': 0
    };
    const s = sm[r.uid]; s['Total Days']++; s['_h'] += r.hoursWorked;
    s['Total OT (Mins)'] += (r.otMins || 0);
    if (isAttendedStatus(r.status)) s['Present']++;
    if (hasLateArrival(r)) s['Late']++;
    if (r.status === 'Half Day') s['Half Day']++;
    if (['Absent', 'System Error'].includes(r.status)) s['Absent']++;
  });
  const sumRows = Object.values(sm).map(s => {
    s['Avg Hrs/Day'] = Math.round(s['_h'] / s['Total Days'] * 100) / 100;
    s['Att %'] = Math.round(s['Present'] / s['Total Days'] * 100) + '%';
    delete s['_h']; return s;
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sumRows), 'Summary');

  // 3. DAILY DETAIL (Numeric)
  const daily = data.map(r => ({
    'Emp ID': r.uid,
    'Name': r.name,
    'Branch': r.branch,
    'Department': r.department,
    'Date': r.date,
    'Day': r.day,
    'Shift': r.shiftDisplay,
    'Status': r.status,
    'First In': r.firstIn,
    'Last Out': r.lastOut,
    'Hours Worked': r.hoursWorked,
    'OT Mins': r.otMins || 0,
    'Late Mins': r.lateMins || 0,
    'Early Out Mins': r.earlyMins || 0,
    'Early In Mins': r.earlyArrivalMins || 0,
    'Punches': r.punchCount
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(daily), 'Daily Detail');

  // 4. LATE REPORT
  const lr = data.filter(r => hasLateArrival(r));
  if (lr.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lr.map(r => ({ 'Emp ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Department': r.department, 'Date': r.date, 'Day': r.day, 'Shift': r.shiftDisplay, 'Arrived': r.firstIn, 'Late By': r.lateBy, 'Late Mins': r.lateMins }))), 'Late Report');

  // 5. ABSENT REPORT
  const ar = data.filter(r => r.status === 'Absent');
  if (ar.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ar.map(r => ({ 'Emp ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Department': r.department, 'Date': r.date, 'Day': r.day, 'Shift': r.shiftDisplay }))), 'Absent Report');

  // 6. EARLY DEPARTURE
  const er = data.filter(r => Number(r.earlyMins || 0) > 0);
  if (er.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(er.map(r => ({ 'Emp ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Department': r.department, 'Date': r.date, 'Day': r.day, 'Shift': r.shiftDisplay, 'Early In Mins': r.earlyArrivalMins, 'Left At': r.lastOut, 'Left Early By': r.earlyBy, 'Early Out Mins': r.earlyMins }))), 'Early Departure');

  // 7. SYSTEM AUDIT (Machine Failures)
  if (S.failureDates && S.failureDates.length > 0) {
    const auditRows = S.failureDates.map(f => ({
      'Date': f.date, 'Branch': f.branch, 'Failure Intensity': f.pct + '% Absent', 'Reason': f.reason, 'Affected Count': f.affected, 'Action Taken': 'Auto-Tagged'
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(auditRows), 'System Audit');
  }

  const from = document.getElementById('date-from').value || 'start', to = document.getElementById('date-to').value || 'end';
  XLSX.writeFile(wb, `Attendance_Analytics_Report_${from}_to_${to}.xlsx`);
}

function renderAuditSelects() {
  const years = [...new Set(S.records.map(r => new Date(r.date + 'T12:00:00').getFullYear()))].sort((a, b) => b - a);
  const yp = document.getElementById('year-picker');
  const mp = document.getElementById('month-picker');
  if (!yp || !mp) return;

  yp.innerHTML = '<option value="" disabled>Year</option>' +
    years.map(y => `<option value="${y}">${y}</option>`).join('');

  if (years.length > 0) {
    yp.value = years[0];
  }
}

function jumpToDateContext() {
  const yp = document.getElementById('year-picker');
  const mp = document.getElementById('month-picker');
  if (!yp || !mp) return;

  const year = yp.value;
  const month = mp.value;

  if (!year) return; // Need at least a year

  const targetYear = parseInt(year);
  const pad = n => String(n).padStart(2, '0');

  let dFrom, dTo;

  if (month !== '') {
    // Both Year and Month
    const targetMonth = parseInt(month);
    dFrom = `${targetYear}-${pad(targetMonth + 1)}-01`;
    const lastDay = new Date(targetYear, targetMonth + 1, 0).getDate();
    dTo = `${targetYear}-${pad(targetMonth + 1)}-${pad(lastDay)}`;
  } else {
    // Year only
    dFrom = `${targetYear}-01-01`;
    dTo = `${targetYear}-12-31`;
  }

  document.getElementById('date-from').value = dFrom;
  document.getElementById('date-to').value = dTo;
  applyFilters();
}

function setDateRange(type) {
  const to = new Date();
  let from = new Date();
  if (type === '7d') from.setDate(to.getDate() - 7);
  else if (type === '30d') from.setDate(to.getDate() - 30);
  else if (type === 'month') {
    from.setDate(1);
    to.setMonth(to.getMonth() + 1);
    to.setDate(0);
  } else if (type === 'clear') {
    document.getElementById('date-from').value = '';
    document.getElementById('date-to').value = '';
    const mp = document.getElementById('month-picker');
    if (mp) mp.value = '';
    const yp = document.getElementById('year-picker');
    if (yp) yp.value = '';
    applyFilters();
    return;
  }

  // Clear month picker if a preset button is clicked instead
  const mp = document.getElementById('month-picker');
  if (mp) mp.value = '';

  document.getElementById('date-from').value = from.toISOString().split('T')[0];
  document.getElementById('date-to').value = to.toISOString().split('T')[0];
  applyFilters();
}

function renderSparklines() {
  const r = S.records;
  if (!r.length) return;
  const last7 = Array.from({ length: 7 }, (_, i) => {
    const d = new Date();
    d.setDate(d.getDate() - i);
    return d.toISOString().split('T')[0];
  }).reverse();

  const getPoints = (type, id) => {
    return last7.map((d, i) => {
      const dayRecs = r.filter(x => x.date === d);
      if (!dayRecs.length) return { x: i, y: 50 };
      let val = 0;
      if (type === 'att') val = (dayRecs.filter(x => isAttendedStatus(x.status)).length / dayRecs.length) * 100;
      else if (type === 'late') val = (dayRecs.filter(x => hasLateArrival(x)).length / dayRecs.length) * 100;
      else if (type === 'abs') val = (dayRecs.filter(x => x.status === 'Absent').length / dayRecs.length) * 100;
      else if (id === 'days') val = 100;
      return { x: i, y: 100 - (val || 0) };
    });
  };

  ['att', 'late', 'abs', 'days'].forEach(id => {
    const svg = document.getElementById('sp-' + id);
    if (!svg) return;
    const pts = getPoints(id, id);
    const path = `M ${pts.map(p => `${p.x * 25},${p.y * 0.3}`).join(' L ')}`;
    svg.innerHTML = `<path d="${path}" fill="none" stroke-width="2" />`;
  });
}

function renderSparklines() {
  const r = S.records;
  if (!r.length) return;
  const last7 = Array.from({ length: 7 }, (_, i) => {
    const d = new Date();
    d.setDate(d.getDate() - i);
    return d.toISOString().split('T')[0];
  }).reverse();

  const getPoints = (type, id) => {
    return last7.map((d, i) => {
      const dayRecs = r.filter(x => x.date === d);
      if (!dayRecs.length) return { x: i, y: 50 };
      let val = 0;
      if (type === 'att') val = (dayRecs.filter(x => isAttendedStatus(x.status)).length / dayRecs.length) * 100;
      else if (type === 'late') val = (dayRecs.filter(x => hasLateArrival(x)).length / dayRecs.length) * 100;
      else if (type === 'abs') val = (dayRecs.filter(x => x.status === 'Absent').length / dayRecs.length) * 100;
      else if (id === 'days') val = 100;
      return { x: i, y: 100 - (val || 0) };
    });
  };

  ['att', 'late', 'abs', 'days'].forEach(id => {
    const svg = document.getElementById('sp-' + id);
    if (!svg) return;
    const pts = getPoints(id, id);
    const path = `M ${pts.map(p => `${p.x * 25},${p.y * 0.3}`).join(' L ')}`;
    svg.innerHTML = `<path d="${path}" fill="none" stroke-width="2" stroke-linecap="round" />`;
  });
}

function renderHeatmap(uid) {
  const r = S.records.filter(x => x.uid === uid);
  const grid = document.getElementById('heatmap-grid');
  const today = new Date();
  const days = Array.from({ length: 28 }, (_, i) => {
    const d = new Date(today);
    d.setDate(today.getDate() - (27 - i));
    return d.toISOString().split('T')[0];
  });

  grid.innerHTML = days.map(d => {
    const rec = r.find(x => x.date === d);
    const status = rec ? rec.status : 'None';
    return `<div class="hm-box" data-status="${status}" data-date="${d}"></div>`;
  }).join('');
}

function getProfileFilterMeta(filter) {
  if (!filter || filter === 'all') {
    return { key: 'all', label: 'All activity', match: () => true };
  }
  if (filter === 'attended') {
    return {
      key: filter,
      label: 'Present Days',
      match: rec => ['Present', 'Late', 'Late (Comp)'].includes(rec.status)
    };
  }
  if (filter === 'exceptions') {
    return {
      key: filter,
      label: 'Exceptions',
      match: rec => ['Missed Punch', 'System Error'].includes(rec.status)
    };
  }
  if (filter.startsWith('status:')) {
    const status = filter.slice(7);
    return {
      key: filter,
      label: status,
      match: rec => rec.status === status
    };
  }
  return { key: 'all', label: 'All activity', match: () => true };
}

function setProfileFilter(filter) {
  const uid = S.activeProfile?.uid;
  if (!uid) return;
  showEmpProfile(uid, filter, true);
}

function formatHeatmapRange(start, end) {
  const startDate = new Date(start + 'T12:00:00');
  const endDate = new Date(end + 'T12:00:00');
  const startLabel = startDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
  const endLabel = endDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
  return start === end ? startLabel : `${startLabel} - ${endLabel}`;
}

function animateProfileRings(scope, delay = 0) {
  scope.querySelectorAll('.ring-fg').forEach(r => {
    const circ = parseFloat(r.style.strokeDasharray);
    const val = parseFloat(r.parentElement?.nextElementSibling?.textContent);
    const offset = circ - ((isNaN(val) ? 0 : val) / 100) * circ;
    setTimeout(() => {
      r.style.strokeDashoffset = isNaN(offset) ? circ : offset;
    }, delay);
  });
}

function selectHeatmapDay(el) {
  if (!el) return;
  const wrap = el.closest('.heatmap-wrap');
  const detail = document.getElementById('heatmap-detail');
  if (!wrap || !detail) return;
  wrap.querySelectorAll('.hm-box.active').forEach(box => box.classList.remove('active'));
  el.classList.add('active');
  const date = el.dataset.date || '--';
  const day = el.dataset.day || '--';
  const status = el.dataset.status || 'None';
  detail.innerHTML = `<div class="heatmap-detail-date">${date} <span>${day}</span></div><div class="heatmap-detail-status">${status}</div>`;
}

function genSpark(f) {
  const last14 = f.slice(-14);
  if (!last14.length) return '';
  const pts = last14.map((x, i) => ({ x: i * 4.5, y: isAttendedStatus(x.status) ? 5 : 20 }));
  const d = `M ${pts.map(p => `${p.x},${p.y}`).join(' L ')}`;
  return `<svg class="sparkline-trend" viewBox="0 0 60 25"><path d="${d}" fill="none" stroke="var(--blue)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" /></svg>`;
}

function showEmpProfile(uid, filter = 'all', preserveScroll = false) {
  const r = S.records.filter(x => x.uid === uid);
  const f = S.filtered.filter(x => x.uid === uid);
  if (!r.length) return;
  const profileRecs = f.length ? f : r;
  const name = r[0].name;
  const dept = r[0].department;
  const branch = r[0].branch;
  const overlay = document.getElementById('sidebar-overlay');
  const sidebar = document.getElementById('emp-sidebar');
  const content = document.getElementById('sidebar-content');
  const sidebarBody = sidebar?.querySelector('.sidebar-body');
  const wasOpen = sidebar?.classList.contains('open');
  const savedScroll = preserveScroll && sidebarBody ? sidebarBody.scrollTop : 0;
  const filterMeta = getProfileFilterMeta(filter);
  const filteredProfileRecs = profileRecs.filter(filterMeta.match);
  S.activeProfile = { uid, filter: filterMeta.key };

  const totalRecords = profileRecs.length;
  const total = totalRecords || 1;
  const presentRecs = profileRecs.filter(x => isAttendedStatus(x.status));
  const attPct = Math.round((presentRecs.length / total) * 100);
  const lateRecs = profileRecs.filter(x => hasLateArrival(x));
  const latePct = Math.round((lateRecs.length / total) * 100);
  const otHours = Math.round(profileRecs.reduce((sum, x) => sum + (x.otMins || 0), 0) / 60);
  const avgHours = totalRecords ? Math.round((profileRecs.reduce((sum, x) => sum + Number(x.hoursWorked || 0), 0) / totalRecords) * 100) / 100 : 0;
  const exceptionCount = profileRecs.filter(x => ['Missed Punch', 'System Error'].includes(x.status)).length;

  const statusOrder = ['Present', 'Late', 'Late (Comp)', 'Half Day', 'Missed Punch', 'Absent', 'Holiday', 'Week Off', 'System Error'];
  const statusCounts = Object.fromEntries(statusOrder.map(status => [
    status,
    profileRecs.filter(x => x.status === status).length
  ]));

  const validIn = (presentRecs.length ? presentRecs : []).map(x => timeToMins(x.firstIn)).filter(m => m !== null);
  const avgIn = validIn.length ? Math.round(validIn.reduce((a, b) => a + b) / validIn.length) : null;
  const avgInStr = minsToTime(avgIn);

  const fv = document.getElementById('date-from').value;
  const tv = document.getElementById('date-to').value;
  let dates = [];
  if (fv && tv) {
    let curr = new Date(fv);
    const end = new Date(tv);
    while (curr <= end) {
      dates.push(new Date(curr).toISOString().split('T')[0]);
      curr.setDate(curr.getDate() + 1);
    }
  }
  const weekRows = [];
  for (let i = 0; i < dates.length; i += 7) weekRows.push(dates.slice(i, i + 7));
  const defaultHeatmapDate = filteredProfileRecs.length ? [...dates].reverse().find(d => filteredProfileRecs.some(x => x.date === d)) || '' : '';
  const defaultHeatmapRec = defaultHeatmapDate ? profileRecs.find(x => x.date === defaultHeatmapDate) : null;
  const defaultHeatmapDay = defaultHeatmapDate ? (defaultHeatmapRec?.day || new Date(defaultHeatmapDate + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'short' })) : '';
  const defaultHeatmapStatus = defaultHeatmapRec ? defaultHeatmapRec.status : (defaultHeatmapDate ? 'None' : '');
  const recentActivity = filteredProfileRecs.slice().reverse().slice(0, 10);

  const ring = (pct, lbl, color, extraClass = '') => {
    const circ = 2 * Math.PI * 28;
    return `
      <div class="ring-card ${extraClass}">
        <div class="ring-svg-wrap">
          <svg class="ring-svg" viewBox="0 0 64 64"><circle class="ring-bg" cx="32" cy="32" r="28" /><circle class="ring-fg" cx="32" cy="32" r="28" style="stroke-dasharray:${circ}; stroke-dashoffset:${circ}; stroke:${color}" /></svg>
          <span class="ring-val">${pct}%</span>
        </div>
        <span class="ring-lbl">${lbl}</span>
      </div>
    `;
  };

  content.innerHTML = `
    <div class="profile-hero">
      <div class="hero-avatar" style="background:${avatarCol(name)}">${initials(name)}</div>
      <div class="hero-info"><h2><span>${name}</span>${genSpark(profileRecs)}</h2><p>ID ${uid} • ${dept} • ${branch}</p></div>
    </div>
    <div class="ring-grid">${ring(attPct, 'Stability', 'var(--teal)')}${ring(100 - latePct, 'Punctuality', 'var(--blue)')}<div class="ring-card ring-card-wide"><div class="ring-svg-wrap"><span class="ring-stat-value">${otHours}h</span></div><span class="ring-lbl">Overtime</span></div></div>
    <div class="dna-capsule"><div class="dna-item"><span class="dna-lbl">Shift DNA: Usual In</span><span class="dna-val">${avgInStr}</span></div><div class="dna-item"><span class="dna-lbl">Engagement</span><span class="dna-val">${attPct > 80 ? 'High' : 'Moderate'}</span></div></div>
    <div class="summary-mini-grid">
      <button class="summary-mini-card profile-action-card${filterMeta.key === 'all' ? ' active' : ''}" type="button" onclick="setProfileFilter('all')"><span class="summary-mini-lbl">Records</span><strong class="summary-mini-val">${totalRecords}</strong></button>
      <button class="summary-mini-card card-teal profile-action-card${filterMeta.key === 'attended' ? ' active' : ''}" type="button" onclick="setProfileFilter('attended')"><span class="summary-mini-lbl">Present Days</span><strong class="summary-mini-val">${presentRecs.length}</strong></button>
      <button class="summary-mini-card card-violet profile-action-card${filterMeta.key === 'status:Holiday' ? ' active' : ''}" type="button" onclick="setProfileFilter('status:Holiday')"><span class="summary-mini-lbl">Holidays</span><strong class="summary-mini-val">${statusCounts['Holiday']}</strong></button>
      <button class="summary-mini-card card-slate profile-action-card${filterMeta.key === 'status:Week Off' ? ' active' : ''}" type="button" onclick="setProfileFilter('status:Week Off')"><span class="summary-mini-lbl">Week Off</span><strong class="summary-mini-val">${statusCounts['Week Off']}</strong></button>
      <button class="summary-mini-card card-amber profile-action-card${filterMeta.key === 'status:Late' ? ' active' : ''}" type="button" onclick="setProfileFilter('status:Late')"><span class="summary-mini-lbl">Late Days</span><strong class="summary-mini-val">${statusCounts['Late']}</strong></button>
      <div class="summary-mini-card"><span class="summary-mini-lbl">Avg Hours</span><strong class="summary-mini-val">${avgHours}h</strong></div>
      <button class="summary-mini-card card-red profile-action-card${filterMeta.key === 'exceptions' ? ' active' : ''}" type="button" onclick="setProfileFilter('exceptions')"><span class="summary-mini-lbl">Exceptions</span><strong class="summary-mini-val">${exceptionCount}</strong></button>
    </div>
    <div class="section-head">Attendance History</div>
    <div class="heatmap-legend">
      <div class="legend-item"><div class="legend-dot p"></div>Present</div>
      <div class="legend-item"><div class="legend-dot l"></div>Late</div>
      <div class="legend-item"><div class="legend-dot a"></div>Absent</div>
      <div class="legend-item"><div class="legend-dot h"></div>Holiday</div>
      <div class="legend-item"><div class="legend-dot w"></div>Week Off</div>
      <div class="legend-item"><div class="legend-dot s"></div>System Error</div>
      <div class="legend-item"><div class="legend-dot n"></div>No Data</div>
    </div>
    <div class="heatmap-wrap"><div class="heatmap-week-list">${weekRows.map(week => `<div class="heatmap-week-row"><div class="heatmap-week-label">${formatHeatmapRange(week[0], week[week.length - 1])}</div><div class="heatmap-grid heatmap-row-grid" style="grid-template-columns:repeat(${week.length},1fr)">${week.map(d => {
    const dayRec = profileRecs.find(x => x.date === d);
    const status = dayRec ? dayRec.status : 'None';
    const day = dayRec?.day || new Date(d + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'short' });
    return `<div class="hm-box${d === defaultHeatmapDate ? ' active' : ''}${filterMeta.key !== 'all' && (!dayRec || !filterMeta.match(dayRec)) ? ' is-muted' : ''}" data-status="${status}" data-date="${d}" data-day="${day}" onclick="selectHeatmapDay(this)"></div>`;
  }).join('')}</div></div>`).join('')}</div><div id="heatmap-detail" class="heatmap-detail">${defaultHeatmapDate ? `<div class="heatmap-detail-date">${defaultHeatmapDate} <span>${defaultHeatmapDay}</span></div><div class="heatmap-detail-status">${defaultHeatmapStatus}</div>` : ''}</div></div>
    <div class="section-head">Recent activity</div>
    <div class="timeline-wrap">${recentActivity.length ? recentActivity.map(x => `<div class="tl-item"><div class="tl-date"><span>${x.date.split('-').slice(1).reverse().join('/')}</span><small>${x.day}</small></div><div class="tl-shift">${x.firstIn || '--'} → ${x.lastOut || '--'}<small>${x.shiftDisplay}</small></div><div class="tl-res"><span class="badge pills-glass">${x.status}</span></div></div>`).join('') : ''}</div>
    <button class="btn btn-primary" onclick="window.filterByEmployeeGlobal('${uid}')" style="margin-top:20px; width:100%">View All History</button>
  `;
  if (sidebarBody) sidebarBody.scrollTop = savedScroll;
  overlay.style.display = 'block';
  setTimeout(() => { overlay.classList.add('open'); sidebar.classList.add('open'); animateProfileRings(content, 600); }, 10);
}

function filterByBranch(branch) {
  setSelectedValues('f-branch', [branch]);
  switchTab('daily', document.querySelector('.tab[onclick*="daily"]'));
  applyFilters();
}

window.showEmpProfile = showEmpProfile;
window.filterByEmployeeGlobal = function (uid) {
  const overlay = document.getElementById('sidebar-overlay');
  const sidebar = document.getElementById('emp-sidebar');
  if (overlay) overlay.classList.remove('open');
  if (sidebar) sidebar.classList.remove('open');
  setTimeout(() => overlay.style.display = 'none', 300);
  switchTab('daily', document.querySelector('.tab[onclick*="daily"]'));
  setSelectedValues('f-emp', [uid]);
  applyFilters();
};
window.closeSidebar = function () {
  const sidebar = document.getElementById('emp-sidebar');
  const overlay = document.getElementById('sidebar-overlay');
  if (sidebar) sidebar.classList.remove('open');
  if (overlay) { overlay.classList.remove('open'); setTimeout(() => overlay.style.display = 'none', 300); }
};

// Expose new v2.0 functions for HTML event handlers
window.clearHolidayData = clearHolidayData;
window.handleFiles = handleFiles;
window.applyTheme = applyTheme;
window.toggleTheme = toggleTheme;
window.AppState = AppState;
// Debug exports (for console verification)
window.normalizeBranch = normalizeBranch;
window.isHoliday = isHoliday;

initApp();
