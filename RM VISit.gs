/*********************************************************************
 * RM Visit Tracker — Google Apps Script Backend (v3: Swipe-based login)
 * ------------------------------------------------------------------
 * Sheets used:
 *   1. Store Mapping     : Store Name | RM Name | CM Name
 *   2. RM Login Logout   : Month | Emp No | Emp Name | Swipe Date | Time | Store (Alias) | Store Name | Designation
 *   3. Distance Matrix   : From Store | To Store | Distance (KM)
 *   4. Visit Plan        : RM Name | Month | Store Name | Planned Date | Visit Type | Status | Submitted On
 *********************************************************************/

// ======= CONFIGURATION =======
const SHEET_ID = '1FQ1slvNeW-8N09miSWC5d4e8ymFcM-jGbMPOmWnio9g';
const RATE_PER_KM = 12;   // ₹ per KM for travel reimbursement

/**
 * Minimum swipes at a (store, date) combination for it to count as a visit.
 *   1 = lenient: any swipe = visit (recommended if people forget to swipe out)
 *   2 = strict : must have both an in-swipe and an out-swipe
 */
const MIN_SWIPES_FOR_VISIT = 1;

const SHEETS = {
  STORE_MAPPING  : 'Store Mapping',
  LOGIN_LOGOUT   : 'RM Login Logout',
  DISTANCE_MATRIX: 'Distance Matrix',
  VISIT_PLAN     : 'Visit Plan'
};

// Column indices for RM Login Logout (0-based)
const LC = {
  MONTH: 0, EMP_NO: 1, EMP_NAME: 2, SWIPE_DATE: 3,
  TIME: 4, STORE_ALIAS: 5, STORE_NAME: 6, DESIGNATION: 7
};

const VISIT_TYPES = {
  R  : { label: 'Routine',       color: '#34a853', isVisit: true  },
  A  : { label: 'Audit',         color: '#1a73e8', isVisit: true  },
  CM : { label: 'CM Accompany',  color: '#d01884', isVisit: true, isCM: true },
  E  : { label: 'Escalation',    color: '#ea4335', isVisit: true  },
  O  : { label: 'Outstation',    color: '#8430ce', isVisit: true  },
  S  : { label: 'Strategic',     color: '#00897b', isVisit: true  },
  H  : { label: 'Holiday',       color: '#9aa0a6', isVisit: false },
  OFC: { label: 'Office',        color: '#f9ab00', isVisit: false },
  L  : { label: 'Leave',         color: '#fb8c00', isVisit: false },
  WO : { label: 'Week Off',      color: '#616161', isVisit: false }
};

// ======= ENTRY POINT =======
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('RM Visit Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ======= UTILITIES =======
function _ss()         { return SpreadsheetApp.openById(SHEET_ID); }
function _sheet(name)  { return _ss().getSheetByName(name); }
function _norm(v)      { return String(v || '').trim().toLowerCase(); }
function _fmtDate(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function _daysInMonth(ym) {
  const [y, m] = ym.split('-').map(Number);
  return new Date(y, m, 0).getDate();
}
function _dayFromDate(d) {
  return (d instanceof Date ? d : new Date(d)).getDate();
}
/**
 * Convert a swipe-time cell to seconds-since-midnight for ordering.
 * Handles Date objects, "HH:MM" strings, and numbers.
 */
function _timeValue(t) {
  if (t instanceof Date) {
    return t.getHours() * 3600 + t.getMinutes() * 60 + t.getSeconds();
  }
  if (typeof t === 'string' && t.indexOf(':') >= 0) {
    const p = t.split(':').map(Number);
    return (p[0] || 0) * 3600 + (p[1] || 0) * 60 + (p[2] || 0);
  }
  if (typeof t === 'number') return t * 86400; // fractional day → seconds
  return 0;
}

// ======= INITIALIZATION =======
function initializeSheets() {
  const ss = _ss();
  const required = {
    'Store Mapping'   : ['Store Name', 'RM Name', 'CM Name'],
    'RM Login Logout' : ['Month', 'Emp No', 'Emp Name', 'Swipe Date', 'Time',
                         'Store (Alias)', 'Store Name', 'Designation'],
    'Distance Matrix' : ['From Store', 'To Store', 'Distance (KM)'],
    'Visit Plan'      : ['RM Name', 'Month', 'Store Name', 'Planned Date',
                         'Visit Type', 'Status', 'Submitted On']
  };
  const log = [];
  Object.keys(required).forEach(name => {
    let sh = ss.getSheetByName(name);
    const headers = required[name];
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
      sh.setFrozenRows(1);
      log.push('Created ' + name);
    } else {
      const lastCol = Math.max(sh.getLastColumn(), 1);
      const curr = sh.getRange(1, 1, 1, Math.max(lastCol, headers.length)).getValues()[0];
      for (let i = 0; i < headers.length; i++) {
        if (String(curr[i] || '').trim() !== headers[i]) {
          sh.getRange(1, i + 1).setValue(headers[i])
            .setFontWeight('bold').setBackground('#e8f0fe');
          log.push('Updated ' + name + ' col ' + (i + 1) + ' → ' + headers[i]);
        }
      }
    }
  });
  return log.length ? log.join('\n') : 'All sheets OK.';
}

// ======= AUTH =======
function authenticateRM(rmName) {
  if (!rmName || !String(rmName).trim()) {
    return { success: false, message: 'Please enter your name.' };
  }
  const sh = _sheet(SHEETS.STORE_MAPPING);
  if (!sh) return { success: false, message: 'Store Mapping sheet missing. Run initializeSheets().' };
  const data = sh.getDataRange().getValues();
  const target = _norm(rmName);
  for (let i = 1; i < data.length; i++) {
    if (_norm(data[i][1]) === target) {
      return { success: true, rmName: String(data[i][1]).trim() };
    }
  }
  return { success: false, message: 'RM name not found. Please contact admin.' };
}

// ======= STORE MAPPING =======
function getStoresForRM(rmName) {
  const sh = _sheet(SHEETS.STORE_MAPPING);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const target = _norm(rmName);
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (_norm(data[i][1]) === target) {
      list.push({
        storeName: String(data[i][0]).trim(),
        cmName   : String(data[i][2] || '').trim()
      });
    }
  }
  return list;
}

// ======= VISIT PLAN =======
function isPlanLocked(rmName, month) {
  const sh = _sheet(SHEETS.VISIT_PLAN);
  if (!sh || sh.getLastRow() < 2) return false;
  const data = sh.getDataRange().getValues();
  const target = _norm(rmName);
  for (let i = 1; i < data.length; i++) {
    if (_norm(data[i][0]) === target
        && String(data[i][1]) === month
        && String(data[i][5]) === 'Locked') {
      return true;
    }
  }
  return false;
}

function getVisitPlan(rmName, month) {
  const sh = _sheet(SHEETS.VISIT_PLAN);
  const out = { locked: false, grid: {}, submittedOn: null };
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getDataRange().getValues();
  const target = _norm(rmName);
  for (let i = 1; i < data.length; i++) {
    if (_norm(data[i][0]) !== target || String(data[i][1]) !== month) continue;
    const store = String(data[i][2]).trim();
    const day = _dayFromDate(data[i][3]);
    const vt = String(data[i][4] || 'R').trim().toUpperCase();
    if (!out.grid[store]) out.grid[store] = {};
    out.grid[store][day] = vt;
    if (String(data[i][5]) === 'Locked') out.locked = true;
    if (data[i][6]) out.submittedOn = _fmtDate(data[i][6]);
  }
  return out;
}

function submitVisitPlan(rmName, month, cells) {
  if (!rmName) return { success: false, message: 'Missing RM name.' };
  if (!/^\d{4}-\d{2}$/.test(month)) return { success: false, message: 'Invalid month.' };
  if (!Array.isArray(cells)) return { success: false, message: 'Invalid cells.' };

  const maxDay = _daysInMonth(month);
  const [yr, mo] = month.split('-').map(Number);
  const validTypes = Object.keys(VISIT_TYPES);

  const rows = [];
  const now = new Date();
  for (const c of cells) {
    if (!c.storeName || !c.visitType) continue;
    const day = Number(c.day);
    if (!day || day < 1 || day > maxDay) {
      return { success: false, message: 'Invalid day ' + c.day + ' for ' + c.storeName + '.' };
    }
    const vt = String(c.visitType).trim().toUpperCase();
    if (validTypes.indexOf(vt) < 0) {
      return { success: false, message: 'Unknown visit type "' + vt + '".' };
    }
    const dateStr = yr + '-' + String(mo).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    rows.push([rmName, month, c.storeName, dateStr, vt, 'Locked', now]);
  }
  if (rows.length === 0) {
    return { success: false, message: 'Please fill in at least one cell before submitting.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (isPlanLocked(rmName, month)) {
      return { success: false, message: 'Plan for ' + month + ' is already locked.' };
    }
    const sh = _sheet(SHEETS.VISIT_PLAN);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    return { success: true, message: 'Plan submitted and locked. ' + rows.length + ' entries saved.' };
  } finally {
    lock.releaseLock();
  }
}

// ======= PLAN KPIs =======
function getPlanKPIs(rmName, month) {
  const stores = getStoresForRM(rmName);
  const plan   = getVisitPlan(rmName, month);

  let totalVisits = 0, cmCount = 0;
  const coveredStores = new Set();

  Object.keys(plan.grid).forEach(store => {
    const days = plan.grid[store];
    Object.keys(days).forEach(d => {
      const vt = days[d];
      const cfg = VISIT_TYPES[vt];
      if (!cfg) return;
      if (cfg.isVisit) { totalVisits++; coveredStores.add(store); }
      if (cfg.isCM)    { cmCount++; }
    });
  });

  const plannedKM = _computePlannedKM(plan.grid);

  return {
    totalVisits   : totalVisits,
    storesCovered : coveredStores.size,
    totalStores   : stores.length,
    cmAccompany   : cmCount,
    notVisited    : Math.max(0, stores.length - coveredStores.size),
    plannedKM     : Math.round(plannedKM * 100) / 100,
    travelReimb   : Math.round(plannedKM * RATE_PER_KM * 100) / 100
  };
}

function _computePlannedKM(grid) {
  const visits = [];
  Object.keys(grid).forEach(store => {
    Object.keys(grid[store]).forEach(d => {
      const vt = grid[store][d];
      if (VISIT_TYPES[vt] && VISIT_TYPES[vt].isVisit) {
        visits.push({ store: store, day: Number(d) });
      }
    });
  });
  visits.sort((a, b) => a.day - b.day);

  const dist = _buildDistanceMap();
  let total = 0;
  for (let i = 1; i < visits.length; i++) {
    const from = visits[i-1].store, to = visits[i].store;
    if (from === to) continue;
    total += dist[from + '→' + to] || 0;
  }
  return total;
}

function _buildDistanceMap() {
  const distSh = _sheet(SHEETS.DISTANCE_MATRIX);
  const dist = {};
  if (distSh && distSh.getLastRow() >= 2) {
    const dd = distSh.getDataRange().getValues();
    for (let i = 1; i < dd.length; i++) {
      const from = String(dd[i][0]).trim();
      const to   = String(dd[i][1]).trim();
      const km   = Number(dd[i][2]) || 0;
      if (from && to) {
        dist[from + '→' + to] = km;
        if (dist[to + '→' + from] == null) dist[to + '→' + from] = km;
      }
    }
  }
  return dist;
}

// ======= SWIPE AGGREGATION (new schema) =======
/**
 * Reads the RM Login Logout sheet and groups swipes by (store, date)
 * for the given RM and month.
 * Returns: [{ store, date:Date, firstTime:seconds, swipeCount }], sorted
 * by date + firstTime. Only groups with swipeCount >= MIN_SWIPES_FOR_VISIT are kept.
 */
function _getVisitEvents(rmName, month) {
  const target = _norm(rmName);
  const [yr, mo] = month.split('-').map(Number);
  const loginSh = _sheet(SHEETS.LOGIN_LOGOUT);
  if (!loginSh || loginSh.getLastRow() < 2) return [];

  const data = loginSh.getDataRange().getValues();
  const groups = {}; // key: store|yyyy-mm-dd -> { store, date, firstTime, swipeCount }

  for (let i = 1; i < data.length; i++) {
    if (_norm(data[i][LC.EMP_NAME]) !== target) continue;
    const dt = data[i][LC.SWIPE_DATE] instanceof Date
      ? data[i][LC.SWIPE_DATE]
      : new Date(data[i][LC.SWIPE_DATE]);
    if (isNaN(dt)) continue;
    if (dt.getFullYear() !== yr || (dt.getMonth() + 1) !== mo) continue;

    const store = String(data[i][LC.STORE_NAME] || '').trim();
    if (!store) continue;

    const dateKey = _fmtDate(dt);
    const key = store + '|' + dateKey;
    const t = _timeValue(data[i][LC.TIME]);

    if (!groups[key]) {
      groups[key] = { store: store, date: dt, firstTime: t, swipeCount: 1 };
    } else {
      groups[key].swipeCount++;
      if (t < groups[key].firstTime) groups[key].firstTime = t;
    }
  }

  const list = Object.values(groups)
    .filter(g => g.swipeCount >= MIN_SWIPES_FOR_VISIT);

  list.sort((a, b) => {
    const da = a.date.getTime(), db = b.date.getTime();
    if (da !== db) return da - db;
    return a.firstTime - b.firstTime;
  });
  return list;
}

// ======= PLANNED vs ACTUAL =======
function getPlannedVsActual(rmName, month) {
  const target = _norm(rmName);

  // Planned (from Visit Plan sheet)
  const planned = {};
  const planSh = _sheet(SHEETS.VISIT_PLAN);
  if (planSh && planSh.getLastRow() >= 2) {
    const pd = planSh.getDataRange().getValues();
    for (let i = 1; i < pd.length; i++) {
      if (_norm(pd[i][0]) !== target || String(pd[i][1]) !== month) continue;
      const vt = String(pd[i][4] || 'R').trim().toUpperCase();
      if (!VISIT_TYPES[vt] || !VISIT_TYPES[vt].isVisit) continue;
      const store = String(pd[i][2]).trim();
      planned[store] = (planned[store] || 0) + 1;
    }
  }

  // Actual (from swipe events)
  const events = _getVisitEvents(rmName, month);
  const actual = {};
  events.forEach(ev => {
    actual[ev.store] = (actual[ev.store] || 0) + 1;
  });

  const mapped = getStoresForRM(rmName).map(s => s.storeName);
  const all = new Set([...Object.keys(planned), ...Object.keys(actual), ...mapped]);

  const rows = [];
  let totalPlanned = 0, totalActual = 0;
  all.forEach(store => {
    const p = planned[store] || 0;
    const a = actual[store]  || 0;
    const completion = p > 0 ? Math.round((a / p) * 100) : (a > 0 ? 100 : 0);
    totalPlanned += p;
    totalActual  += a;
    rows.push({ storeName: store, planned: p, actual: a, completion, missed: p > 0 && a < p });
  });

  rows.sort((a, b) => {
    if (a.missed !== b.missed) return a.missed ? -1 : 1;
    return a.completion - b.completion;
  });

  return {
    rows: rows,
    kpi: {
      totalPlanned: totalPlanned,
      totalActual : totalActual,
      completion  : totalPlanned > 0 ? Math.round((totalActual / totalPlanned) * 100) : 0
    }
  };
}

// ======= TRAVEL LOG (actuals) =======
function getTravelLog(rmName, month) {
  const events = _getVisitEvents(rmName, month);
  const dist = _buildDistanceMap();

  const log = [];
  let totalKM = 0;
  for (let i = 1; i < events.length; i++) {
    const from = events[i-1].store, to = events[i].store;
    if (!from || !to || from === to) continue;
    const km = dist[from + '→' + to];
    const distance = (km != null) ? km : 0;
    totalKM += distance;
    log.push({
      date     : _fmtDate(events[i].date),
      fromStore: from,
      toStore  : to,
      distance : distance,
      missing  : km == null
    });
  }
  return { log: log, totalKM: Math.round(totalKM * 100) / 100 };
}

// ======= TOP PERFORMER =======
function getTopPerformer(month) {
  const planSh = _sheet(SHEETS.VISIT_PLAN);
  if (!planSh || planSh.getLastRow() < 2) return null;
  const data = planSh.getDataRange().getValues();
  const rmSet = new Set();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === month) rmSet.add(String(data[i][0]).trim());
  }
  if (!rmSet.size) return null;

  let top = null, topPct = -1, topActual = -1;
  rmSet.forEach(rm => {
    const r = getPlannedVsActual(rm, month);
    if (r.kpi.totalPlanned <= 0) return;
    if (r.kpi.completion > topPct
        || (r.kpi.completion === topPct && r.kpi.totalActual > topActual)) {
      topPct = r.kpi.completion;
      topActual = r.kpi.totalActual;
      top = {
        rmName      : rm,
        completion  : r.kpi.completion,
        totalPlanned: r.kpi.totalPlanned,
        totalActual : r.kpi.totalActual
      };
    }
  });
  return top;
}

// ======= DASHBOARD BOOTSTRAP =======
function getDashboardData(rmName, month) {
  const stores     = getStoresForRM(rmName);
  const plan       = getVisitPlan(rmName, month);
  const planKPI    = getPlanKPIs(rmName, month);
  const comparison = getPlannedVsActual(rmName, month);
  const travel     = getTravelLog(rmName, month);
  const top        = getTopPerformer(month);
  const uniqueCMs  = Array.from(new Set(stores.map(s => s.cmName).filter(Boolean))).sort();

  const days = _daysInMonth(month);
  const [yr, mo] = month.split('-').map(Number);
  const shortWk = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const dayMeta = [];
  for (let d = 1; d <= days; d++) {
    const dt = new Date(yr, mo - 1, d);
    const w = dt.getDay();
    dayMeta.push({ day: d, weekday: shortWk[w], weekend: (w === 0 || w === 6) });
  }

  return {
    rmName       : rmName,
    month        : month,
    stores       : stores,
    uniqueCMs    : uniqueCMs,
    visitTypes   : VISIT_TYPES,
    dayMeta      : dayMeta,
    ratePerKM    : RATE_PER_KM,
    plan         : plan,
    planKPI      : planKPI,
    comparison   : comparison,
    travel       : travel,
    topPerformer : top,
    serverTime   : new Date().toISOString()
  };
}
