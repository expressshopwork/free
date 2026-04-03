// ============================================================
// Smart 5G Dashboard — app.js
// ============================================================

// ------------------------------------------------------------
// State & Constants
// ------------------------------------------------------------
let currentRole = 'admin'; // 'admin', 'supervisor', 'user'
let currentUser = null;
let currentPage = 'dashboard';
let currentSaleTab = 'unit';
let currentSaleModalType = 'unit'; // 'unit' or 'point'
let currentCustomerTab = 'new-customer';
let currentSettingsTab = 'permission';
let currentCoverageTab = 'smart-home';
let currentPromoView = 'new'; // 'new' or 'expired'
let currentReportView = 'table'; // 'table' or 'summary'
let filteredSales = [];
let saleUnitShowAll = false;
let salePointShowAll = false;
let itemGroupSelected = 'unit'; // 'unit' or 'dollar'
let kpiValueMode = 'unit'; // 'unit' or 'currency'
let kpiTypeSelected = 'Sales';
let kpiForSelected = 'shop'; // 'shop' or 'agent'
let kpiModeSelected = 'volume'; // 'volume' or 'point'
let kpiSelectedMonth = ''; // '' means no filter (show all)
let stSelectedMonth = ''; // Sale Tracking selected month
let currentPerfTab = 'tracking'; // 'tracking' or 'performance'

// Chart instances
let _cTrend = null, _cMix = null, _cAgent = null, _cGrowth = null;
let _cSaleMix = null, _cSaleAgent = null;
let _cDepositPerf = null;
let _cKpiGauges = [];

// Coverage map instances
var _covMapSmartHome = null, _covMapSmartHome5G = null, _covMapSmartFiber = null;
var _covPickerMap = null, _covPickerMarker = null, _covPickerHighlight = null;
var _tuPickerMap = null, _tuPickerMarker = null;

// Commune autocomplete state
var _covCommuneDebounce = null;

// Approval form state
var _approvalFormData = null;
var _sigCanvas = null, _sigCtx = null, _sigDrawing = false;

// Constants
const TAB_PERM = { admin: ['permission', 'google-sheets'], cluster: ['permission'], supervisor: [], agent: [], user: [] };
const TAB_LBL = { permission: 'Permission', 'google-sheets': 'Google Sheets' };
const AV_COLORS = ['#E53935','#8E24AA','#1565C0','#00838F','#2E7D32','#F57F17','#4E342E','#37474F'];
const CHART_PAL = ['#1B7D3D','#2196F3','#FF9800','#9C27B0','#F44336','#00BCD4','#FFEB3B','#795548'];
const KNOWN_CURS = ['USD','KHR','THB','VND'];
const KNOWN_UNITS = ['Unit','SIM','GB','MB','Minutes','SMS','Voucher'];
const GAS_URL_PREFIX = 'https://script.google.com/macros/s/';
const SYNCED_SHEETS = ['DailySale','Customers','TopUp','Terminations','OutCoverage','Promotions','Deposits','KPI','Items','Coverage','Staff'];
// Target spreadsheet ID — included in every GAS request so all roles (admin,
// supervisor, agent) always write to the same spreadsheet, regardless of which
// GAS deployment URL a given browser has cached in localStorage.
// Note: this identifier is safe to include in client-side code for a team-internal
// tool — it merely names the destination; the GAS script enforces all write access.
const SPREADSHEET_ID = '1K1cp9nhuHIu9iRGHCjk8hhjkL-CANfaC61qoEI-HT90';

// Item ID constants for key KPI calculations
const ITEM_ID_REVENUE = 'i8';
const ITEM_ID_RECHARGE = 'i7';
const ITEM_ID_GROSS_ADS = 'i1';
const ITEM_ID_SMART_HOME = 'i2';
const ITEM_ID_SMART_FIBER = 'i3';
const ITEM_ID_BUY_NUMBER = 'i9';

// Fixed sale form items (Dollar Group and Unit Group)
const DOLLAR_SALE_ITEMS = [
  { id: 'i10', name: 'Buy Number' },
  { id: 'i6',  name: 'ChangeSIM' },
  { id: 'i7',  name: 'Recharge' },
  { id: 'i11', name: 'Device + Accessories' },
];
const UNIT_SALE_ITEMS = [
  { id: 'i1', name: 'Gross Ads' },
  { id: 'i2', name: 'Smart@Home' },
  { id: 'i3', name: 'Smart Fiber+' },
  { id: 'i4', name: 'SmartNas' },
  { id: 'i5', name: 'Monthly Upsell' },
];

const BRANCHES = ['Phnom Penh', 'Siem Reap', 'Battambang', 'Sihanoukville', 'Kampong Cham', 'Express_Tramkak'];

// Smart Shop branches (Cambodia)
const SMART_SHOPS = [
  'Smart Shop Phnom Penh Central','Smart Shop Toul Kork','Smart Shop Boeung Keng Kang',
  'Smart Shop Tuol Sangke','Smart Shop Chbar Ampov','Smart Shop Meanchey',
  'Smart Shop Dangkor','Smart Shop Russey Keo','Smart Shop Prek Pnov',
  'Smart Shop Pur Senchey','Smart Shop Sensok','Smart Shop Chraoy Chongvar',
  'Smart Shop Siem Reap Central','Smart Shop Siem Reap Airport','Smart Shop Svay Dongkum',
  'Smart Shop Battambang Central','Smart Shop Battambang Market','Smart Shop Pailin',
  'Smart Shop Sihanoukville Central','Smart Shop Sihanoukville Beach','Smart Shop Kampot',
  'Smart Shop Kep','Smart Shop Kampong Cham Central','Smart Shop Kampong Cham Market',
  'Smart Shop Kampong Thom','Smart Shop Kampong Chhnang','Smart Shop Kampong Speu',
  'Smart Shop Kandal','Smart Shop Ta Khmau','Smart Shop Takeo',
  'Smart Shop Svay Rieng','Smart Shop Prey Veng','Smart Shop Kratié',
  'Smart Shop Stung Treng','Smart Shop Ratanakiri','Smart Shop Mondulkiri',
  'Smart Shop Pursat','Smart Shop Koh Kong','Smart Shop Preah Vihear',
  'Smart Shop Oddor Meanchey','Smart Shop Banteay Meanchey','Smart Shop Poipet',
  'Smart Shop Sisophon','Smart Shop Preah Sihanouk','Smart Shop Tboung Khmum',
  'Smart Shop Kep City','Smart Shop Kampong Leaeng','Smart Shop Kampong Tralach',
  'Smart Shop Chamkar Mon','Smart Shop Daun Penh','Smart Shop 7 Makara',
  'Smart Shop Tuol Kork 2','Smart Shop Prampi Makara','Smart Shop Ek Phnom',
  'Smart Shop Banan','Smart Shop Sangke','Smart Shop Moung Ruessei',
  'Smart Shop Thma Koul','Smart Shop Phnom Sruoch','Smart Shop Aoral',
  'Smart Shop Oral','Smart Shop Basedth','Express_Tramkak'
];

// ── KPI Point-Based Services ──────────────────────────────────────────────
const KPI_SERVICES = [
  { id: 'ks1',  name: 'Gross Add',                  category: 'MBB Pre-Paid', rate: 1.0, addOn: 0 },
  { id: 'ks2',  name: 'Change SIM',                 category: 'MBB Pre-Paid', rate: 1.0, addOn: 0 },
  { id: 'ks3',  name: 'Pre-Paid sub Recharge',      category: 'MBB Pre-Paid', rate: 1.0, addOn: 0 },
  { id: 'ks4',  name: 'SC Selling',                 category: 'MBB Pre-Paid', rate: 1.0, addOn: 0 },
  { id: 'ks5',  name: 'Home Internet Gross Add',    category: 'FBB Home',     rate: 2.0, addOn: 0 },
  { id: 'ks6',  name: 'FWBB Deposit / FTTx Signup', category: 'FBB Home',     rate: 2.0, addOn: 0 },
  { id: 'ks7',  name: 'Home Internet Migration',    category: 'FBB Home',     rate: 2.0, addOn: 0 },
  { id: 'ks8',  name: 'Home Internet Recharge',     category: 'FBB Home',     rate: 2.0, addOn: 0 },
  { id: 'ks9',  name: 'SME New Sub Gross Add',      category: 'FBB SME',      rate: 2.0, addOn: 0 },
  { id: 'ks10', name: 'SME Existing Recharge',      category: 'FBB SME',      rate: 2.0, addOn: 0 },
  { id: 'ks11', name: 'ICT Solution',               category: 'MBB ICT',      rate: 2.0, addOn: 0 },
  { id: 'ks12', name: 'Device Handset/Accessory',   category: 'Other',        rate: 0.5, addOn: 0 },
  { id: 'ks13', name: 'eSIM',                       category: 'Other',        rate: 0.0, addOn: 2.0 },
  { id: 'ks14', name: 'Smart NAS Download',         category: 'Other',        rate: 0.0, addOn: 2.0 },
];

// ── KPI Tier Thresholds ───────────────────────────────────────────────────
const KPI_TIERS = {
  agent:      { min: 800,  gate: 1000, otb: 1100, oab: 1300 },
  supervisor: { min: 5600, gate: 7000, otb: 7500, oab: 8000 }
};

function calcKpiPoints(serviceId, amount) {
  var svc = KPI_SERVICES.find(function(s) { return s.id === serviceId; });
  if (!svc) return 0;
  return (parseFloat(amount) || 0) * svc.rate + svc.addOn;
}

function getKpiTierLabel(points, role) {
  var tiers = (role === 'supervisor') ? KPI_TIERS.supervisor : KPI_TIERS.agent;
  if (points >= tiers.oab)  return { label: 'OAB',  color: '#4CAF50', rank: 4 };
  if (points >= tiers.otb)  return { label: 'OTB',  color: '#2196F3', rank: 3 };
  if (points >= tiers.gate) return { label: 'Gate', color: '#FF9800', rank: 2 };
  if (points >= tiers.min)  return { label: 'Min',  color: '#9C27B0', rank: 1 };
  return { label: 'Below Min', color: '#F44336', rank: 0 };
}

const SUPPORT_CONTACT = { email: 'support@smart5g.com', phone: '+855 23 123 456' };

// ── Google Sheets Sync ──────────────────────────────────────
const GS_URL_DEFAULT = 'https://script.google.com/macros/s/AKfycbxokMtXAEvhyUrtfCSFCDmmdv6Cr6rOFVxkBxtH_eUbQc4okwCcVNVVvOv02nmanfPdTA/exec';
// Runtime URL — loaded from localStorage on startup; falls back to the default above.
var gsUrl = (function() {
  try { return localStorage.getItem('smart5g_gas_url') || GS_URL_DEFAULT; } catch(e) { return GS_URL_DEFAULT; }
})();

function _gsPost(payload, retries) {
  if (!gsUrl) return Promise.resolve({});
  retries = retries === undefined ? 2 : retries;
  // Always include the spreadsheet ID so every role writes to the same target sheet,
  // regardless of which GAS deployment URL is stored in this browser's localStorage.
  var body = Object.assign({}, payload);
  body.spreadsheetId = SPREADSHEET_ID;
  return fetch(gsUrl, {
    method: 'POST',
    body: JSON.stringify(body)
  })
  .then(function(resp) { return resp.json(); })
  .catch(function(err) {
    if (retries > 0) {
      return new Promise(function(resolve) { setTimeout(resolve, 1500); })
        .then(function() { return _gsPost(payload, retries - 1); });
    }
    console.warn('GS post failed after retries:', err);
    throw err;
  });
}

function normalizeStaffRecord(u) {
  var out = {};
  Object.keys(u).forEach(function(k) {
    var val = u[k];
    // Normalize key to lowercase so that column headers from Google Sheets
    // (e.g. 'Username', 'Password') are accessible via the expected lowercase keys.
    var key = k.toLowerCase();
    // Trim all string values and handle null/undefined
    if (typeof val === 'string') {
      out[key] = val.trim();
    } else {
      out[key] = val;
    }
  });
  // Lowercase status so login comparison works regardless of how the sheet stores it
  // (role VALUE is kept as-is because roleMap lookups depend on the original casing, e.g. 'Admin')
  if (out.status) out.status = out.status.toLowerCase();
  // Ensure username exists and is not empty
  if (!out.username || out.username === '') {
    console.warn('Staff record missing username:', out);
  }
  return out;
}

function isAdminUser(u) {
  return (u.username || '').toLowerCase() === 'admin' && (u.role || '').toLowerCase() === 'admin';
}

function fetchStaffFromSheet() {
  if (!gsUrl) {
    console.warn('[SYNC] No Google Sheets URL configured');
    return Promise.resolve();
  }

  console.log('[SYNC] Fetching staff from Google Sheets...');

  return readSheet('Staff')
    .then(function(data) {
      console.log('[SYNC] Raw staff data received:', data.length, 'records');

      if (!Array.isArray(data) || data.length === 0) {
        console.warn('[SYNC] No staff data returned from sheet');
        return;
      }

      console.log('[SYNC] Processing', data.length, 'staff records');

      // Normalize staff records from Google Sheets (trim whitespace, lowercase status)
      var normalized = data.map(normalizeStaffRecord);
      console.log('[SYNC] Normalized staff records:', normalized.map(function(u) { return { username: u.username, status: u.status, role: u.role, hasPassword: !!u.password }; }));

      // Keep local admin as fallback if sheet doesn't contain one
      if (!normalized.some(isAdminUser)) {
        var localAdmin = staffList.find(isAdminUser);
        if (localAdmin) {
          console.log('[SYNC] Adding local admin as fallback');
          normalized.unshift(localAdmin);
        }
      }
      // Keep default agent users as fallback if sheet doesn't contain them
      DEFAULT_AGENT_USERS.forEach(function(du) {
        if (!normalized.some(function(u) { return u.username === du.username; })) {
          console.log('[SYNC] Adding default user:', du.username);
          normalized.push(du);
        }
      });
      staffList = normalized;
      lsSave(LS_KEYS.staff, staffList);
      console.log('[SYNC] ✓ Staff list updated:', staffList.length, 'total users');
    })
    .catch(function(e) {
      console.error('[SYNC] ✗ Staff fetch error:', e);
      var errEl = g('login-error');
      if (errEl) {
        var hasCachedStaff = localStorage.getItem(LS_KEYS.staff) !== null;
        errEl.textContent = hasCachedStaff
          ? 'Could not reach the server. Signing in with cached data.'
          : 'Could not reach the server. Please check your internet connection and try again.';
        errEl.style.display = '';
        setTimeout(function() { if (errEl) errEl.style.display = 'none'; }, 5000);
      }
    });
}

// Serialize all object/array field values to JSON strings so rows can be stored
// as flat strings in Google Sheets. Used by both syncSheet and syncUpAll.
function normalizeRowForSheet(row) {
  var out = {};
  Object.keys(row).forEach(function(k) {
    var v = row[k];
    out[k] = (v !== null && v !== undefined && typeof v === 'object') ? JSON.stringify(v) : (v !== undefined ? v : '');
  });
  return out;
}

function normalizeArrayForSheet(dataArray) {
  return Array.isArray(dataArray) ? dataArray.map(normalizeRowForSheet) : [];
}

function syncSheet(sheetName, dataArray) {
  if (!gsUrl) return;
  var ind = document.getElementById('gs-sync-indicator');
  var lbl = document.getElementById('gs-sync-status');
  if (ind) ind.className = 'syncing';
  if (lbl) lbl.textContent = 'Syncing\u2026';

  _gsPost({ sheet: sheetName, action: 'sync', data: normalizeArrayForSheet(dataArray) })
    .then(function() {
      if (ind) ind.className = '';
      if (lbl) lbl.textContent = 'Synced \u2713';
      setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
    })
    .catch(function(err) {
      console.warn('GS sync error:', err);
      if (ind) ind.className = 'error';
      if (lbl) lbl.textContent = 'Sync failed';
    });
}

function deleteFromSheet(sheetName, id) {
  if (!gsUrl) return;
  _gsPost({ sheet: sheetName, action: 'delete', data: { id: id } })
    .catch(function(err) { console.warn('GS delete error:', err); });
}

// Read a single sheet from Google Apps Script (returns a Promise<array>).
// Sends { sheet, action:'read' } via POST and accepts both a plain array
// and a { status:'ok', data:[...] } envelope in the response.
function readSheet(sheetName) {
  if (!gsUrl) return Promise.resolve([]);

  console.log('[SYNC] Reading sheet:', sheetName);

  var body = { sheet: sheetName, action: 'read' };
  body.spreadsheetId = SPREADSHEET_ID;

  return fetch(gsUrl, {
    method: 'POST',
    body: JSON.stringify(body)
  })
    .then(function(r) {
      if (!r.ok) {
        console.error('[SYNC] HTTP error:', r.status);
        throw new Error('HTTP ' + r.status);
      }
      return r.json();
    })
    .then(function(resp) {
      console.log('[SYNC] Sheet response for', sheetName + ':', resp);

      if (Array.isArray(resp)) {
        console.log('[SYNC] Received array with', resp.length, 'items');
        return resp;
      }
      if (resp && Array.isArray(resp.data)) {
        console.log('[SYNC] Received data envelope with', resp.data.length, 'items');
        return resp.data;
      }
      console.warn('[SYNC] Unexpected response format');
      return [];
    })
    .catch(function(e) {
      console.error('[SYNC] readSheet(' + sheetName + ') failed:', e);
      return [];
    });
}

// Push all local data to Google Sheets. Exposed globally so admin can call it
// from the browser console or a "Sync Up" button.
// NOTE: Running Sync Up once will also cause ensureHeaders (in gas/Code.gs) to
// create any missing columns (e.g. status, approvedBy, approvedAt) in the
// Deposits sheet if they were absent from an older sheet setup.
function syncUpAll() {
  if (!gsUrl) {
    showToast('No sync URL configured.', 'error');
    return;
  }
  var ind = document.getElementById('gs-sync-indicator');
  var lbl = document.getElementById('gs-sync-status');
  var upBtn = document.getElementById('sync-up-btn');
  if (ind) ind.className = 'syncing';
  if (lbl) lbl.textContent = 'Uploading\u2026';
  if (upBtn) upBtn.disabled = true;

  var sheets = [
    { name: 'DailySale',    data: function() { return saleRecords; } },
    { name: 'Customers',    data: function() { return newCustomers; } },
    { name: 'TopUp',        data: function() { return topUpList; } },
    { name: 'Terminations', data: function() { return terminationList; } },
    { name: 'OutCoverage',  data: function() { return outCoverageList; } },
    { name: 'Promotions',   data: function() { return promotionList; } },
    { name: 'Deposits',     data: function() { return depositList; } },
    { name: 'KPI',          data: function() { return kpiList; } },
    { name: 'Items',        data: function() { return itemCatalogue; } },
    { name: 'Coverage',     data: function() { return coverageLocations; } },
    { name: 'Staff',        data: function() { return staffList; } },
  ];

  var promises = sheets.map(function(s) {
    return _gsPost({ sheet: s.name, action: 'sync', data: normalizeArrayForSheet(s.data()) });
  });

  Promise.all(promises).then(function() {
    if (ind) ind.className = '';
    if (lbl) lbl.textContent = 'Uploaded \u2713';
    if (upBtn) upBtn.disabled = false;
    setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
    showToast('All data uploaded to Google Sheets.', 'success');
  }).catch(function(err) {
    console.warn('Sync up error:', err);
    if (ind) ind.className = 'error';
    if (lbl) lbl.textContent = 'Upload failed';
    if (upBtn) upBtn.disabled = false;
    showToast('Upload failed. Please try again.', 'error');
  });
}

// Sync all sheets down from Google Sheets into local memory and localStorage,
// then re-render the current page. Exposed globally so admin can call it from
// the browser console or a "Sync Down" button.
function syncDownAll() {
  if (!gsUrl) {
    showToast('No sync URL configured.', 'error');
    return Promise.resolve();
  }
  var ind = document.getElementById('gs-sync-indicator');
  var lbl = document.getElementById('gs-sync-status');
  var btn = document.getElementById('sync-down-btn');
  if (ind) ind.className = 'syncing';
  if (lbl) lbl.textContent = 'Syncing\u2026';
  if (btn) btn.disabled = true;

  var sheets = [
    { name: 'DailySale',    lsKey: LS_KEYS.sales,        assign: function(d) { saleRecords = d.map(normalizeSaleRecord); } },
    { name: 'Customers',    lsKey: LS_KEYS.customers,    assign: function(d) { newCustomers = d; } },
    { name: 'TopUp',        lsKey: LS_KEYS.topup,        assign: function(d) { topUpList = d; } },
    { name: 'Terminations', lsKey: LS_KEYS.terminations, assign: function(d) { terminationList = d; } },
    { name: 'OutCoverage',  lsKey: LS_KEYS.outCoverage,  assign: function(d) { outCoverageList = d; } },
    { name: 'Promotions',   lsKey: LS_KEYS.promotions,   assign: function(d) { promotionList = d; } },
    { name: 'Deposits',     lsKey: LS_KEYS.deposits,     assign: function(d) { depositList = d.map(normalizeDeposit); } },
    { name: 'KPI',          lsKey: LS_KEYS.kpis,         assign: function(d) { kpiList = d; } },
    { name: 'Items',        lsKey: LS_KEYS.items,        assign: function(d) { if (d.length) itemCatalogue = d; } },
    { name: 'Coverage',     lsKey: LS_KEYS.coverage,     assign: function(d) { coverageLocations = d; } },
    { name: 'Staff',        lsKey: LS_KEYS.staff,        assign: function(d) {
      var normalizedStaff = d.map(normalizeStaffRecord);
      if (!normalizedStaff.some(isAdminUser)) {
        var localAdmin = staffList.find(isAdminUser);
        if (localAdmin) normalizedStaff.unshift(localAdmin);
      }
      staffList = normalizedStaff;
    }}
  ];

  var promises = sheets.map(function(s) {
    return readSheet(s.name)
      .then(function(data) {
        if (!Array.isArray(data) || data.length === 0) return;
        // Parse JSON-serialized object fields back to objects (e.g. items, dollarItems, cashDetail)
        var parsed = data.map(function(row) {
          var out = {};
          Object.keys(row).forEach(function(k) {
            var v = row[k];
            if (typeof v === 'string' && v.length > 0 && (v[0] === '{' || v[0] === '[')) {
              try { out[k] = JSON.parse(v); } catch(parseErr) { out[k] = v; }
            } else {
              out[k] = v;
            }
          });
          return out;
        });
        s.assign(parsed);
        lsSave(s.lsKey, parsed);
      });
  });

  return Promise.all(promises).then(function() {
    // Re-sync currentUser from the refreshed staffList
    if (currentUser) {
      var freshUser = staffList.find(function(u) { return u.id === currentUser.id; });
      if (freshUser) {
        currentUser = freshUser;
        var nameEl = g('topbar-name'); if (nameEl) nameEl.textContent = freshUser.name;
        var avatarEl = g('topbar-avatar'); if (avatarEl) avatarEl.textContent = ini(freshUser.name);
      } else if (!isAdminUser(currentUser)) {
        currentUser = null;
        var as2 = g('app-shell'); if (as2) as2.style.display = 'none';
        var ls2 = g('login-screen'); if (ls2) ls2.style.display = 'flex';
        var lf2 = g('login-form'); if (lf2) lf2.reset();
        showToast('Your account was not found. You have been signed out.', 'error');
        return;
      }
    }
    if (ind) ind.className = '';
    if (lbl) lbl.textContent = 'Synced \u2713';
    if (btn) btn.disabled = false;
    setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
    // Re-render current page with fresh data
    if (currentPage === 'dashboard') renderDashboard();
    else if (currentPage === 'promotionPage') renderPromotionCards();
    else if (currentPage === 'kpi') renderKpiTable();
    else if (currentPage === 'sale') applyReportFilters();
    else if (currentPage === 'customer') { renderNewCustomerTable(); renderTopUpTable(); renderTerminationTable(); renderOutCoverageTable(); }
    else if (currentPage === 'deposit') { renderDepositTable(); updateDepositKpis(); }
    else if (currentPage === 'settings') { renderStaffTable(); renderAccessContent(currentSettingsTab); }
    showToast('Data synced from Google Sheets.', 'success');
  }).catch(function(err) {
    console.warn('Sync down error:', err);
    if (ind) ind.className = 'error';
    if (lbl) lbl.textContent = 'Sync failed';
    if (btn) btn.disabled = false;
    showToast('Sync failed. App will use cached data.', 'error');
  });
}

// Expose sync functions globally so they can be called from buttons or console
window.syncDownAll = syncDownAll;
window.syncUpAll = syncUpAll;

function refreshAllData() {
  return syncDownAll();
}

// ------------------------------------------------------------
// Sample Data
// ------------------------------------------------------------
let itemCatalogue = [
  { id: 'i1',  name: 'Gross Ads',            shortcut: 'GA', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Gross Ads' },
  { id: 'i2',  name: 'Smart@Home',            shortcut: 'SH', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Smart@Home package' },
  { id: 'i3',  name: 'Smart Fiber+',          shortcut: 'SF', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Smart Fiber+' },
  { id: 'i4',  name: 'SmartNas',              shortcut: 'SN', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'SmartNas' },
  { id: 'i5',  name: 'Monthly Upsell',        shortcut: 'MU', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Monthly Upsell' },
  { id: 'i6',  name: 'ChangeSIM',             shortcut: 'CS', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Change SIM ($)' },
  { id: 'i7',  name: 'Recharge',              shortcut: 'RC', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Recharge ($)' },
  { id: 'i9',  name: 'Buy Number (Legacy)',   shortcut: 'BN', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'inactive', desc: 'Buy Number ($) - replaced by i10' },
  { id: 'i10', name: 'Buy Number',             shortcut: 'BN', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Buy Number ($)' },
  { id: 'i11', name: 'Device + Accessories',  shortcut: 'DA', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Device + Accessories ($)' },
  { id: 'i8',  name: 'Total Revenue',         shortcut: 'RV', group: 'dollar', currency: '$', price: 1, noAutoSum: true, category: 'Sales', status: 'active', desc: 'Total Revenue ($)' },
];

let saleRecords = [];

let newCustomers = [];

let topUpList = [];

let terminationList = [];

let outCoverageList = [];

const DEFAULT_AGENT_USERS = [
  { id: 'u2', name: 'RIM SARAY', username: 'rim.saray', password: 'Tramkak@2026', role: 'Supervisor', branch: 'Express_Tramkak', status: 'active', email: 'rim.saray1@smart.com.kh' },
  { id: 'u3', name: 'KUN CHAMNAN', username: 'kun.chamnan', password: 'Tramkak@2026', role: 'Agent', branch: '', status: 'active', email: 'kun.chamnan@smart.com.kh' },
];

let staffList = [
  { id: 'u1', name: 'Admin', username: 'admin', password: 'Tramkak@2026', role: 'Admin', branch: '', status: 'active' },
].concat(DEFAULT_AGENT_USERS);

let kpiList = [];

let promotionList = [];

let depositList = [];

let coverageLocations = [];

let notifications = [];

// ------------------------------------------------------------
// localStorage Persistence
// ------------------------------------------------------------
const LS_KEYS = {
  items: 'smart5g_items',
  sales: 'smart5g_sales',
  customers: 'smart5g_customers',
  topup: 'smart5g_topup',
  terminations: 'smart5g_terminations',
  outCoverage: 'smart5g_out_coverage',
  staff: 'smart5g_staff',
  kpis: 'smart5g_kpis',
  promotions: 'smart5g_promotions',
  deposits: 'smart5g_deposits',
  coverage: 'smart5g_coverage',
  session: 'smart5g_session',
  gasUrl: 'smart5g_gas_url'
};

const SESSION_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes
var sessionTimer = null;

function startSessionTimer(remainingMs) {
  clearSessionTimer();
  var delay = (remainingMs !== undefined) ? remainingMs : SESSION_TIMEOUT_MS;
  sessionTimer = setTimeout(function() {
    performAutoLogout();
  }, delay);
}

function clearSessionTimer() {
  if (sessionTimer) { clearTimeout(sessionTimer); sessionTimer = null; }
}

function resetSessionTimer() {
  if (!currentUser) return;
  // Update loginTime in localStorage so page reloads also get the refreshed timestamp
  var saved = lsLoad(LS_KEYS.session, null);
  if (saved) { saved.loginTime = Date.now(); lsSave(LS_KEYS.session, saved); }
  startSessionTimer();
}

function performAutoLogout() {
  clearSessionTimer();
  currentUser = null;
  localStorage.removeItem(LS_KEYS.session);
  var as = g('app-shell'); if (as) as.style.display = 'none';
  var ls = g('login-screen'); if (ls) ls.style.display = 'flex';
  var lf = g('login-form'); if (lf) lf.reset();
  var errEl = g('login-error'); if (errEl) errEl.style.display = 'none';
  var btn = g('login-submit-btn'); if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-right-to-bracket"></i> Sign In'; }
  showToast('Your session has expired. Please sign in again.', 'warning');
}

function lsSave(key, data) {
  try { localStorage.setItem(key, JSON.stringify(data)); } catch (e) { console.warn('localStorage save failed:', e); }
}

function lsLoad(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) { return fallback; }
}

function saveAllData() {
  lsSave(LS_KEYS.items, itemCatalogue);
  lsSave(LS_KEYS.sales, saleRecords);
  lsSave(LS_KEYS.customers, newCustomers);
  lsSave(LS_KEYS.topup, topUpList);
  lsSave(LS_KEYS.terminations, terminationList);
  lsSave(LS_KEYS.outCoverage, outCoverageList);
  lsSave(LS_KEYS.staff, staffList);
  lsSave(LS_KEYS.kpis, kpiList);
  lsSave(LS_KEYS.promotions, promotionList);
  lsSave(LS_KEYS.deposits, depositList);
  lsSave(LS_KEYS.coverage, coverageLocations);
}

function loadAllData() {
  itemCatalogue = lsLoad(LS_KEYS.items, itemCatalogue);
  saleRecords = lsLoad(LS_KEYS.sales, saleRecords).map(normalizeSaleRecord);
  newCustomers = lsLoad(LS_KEYS.customers, newCustomers);
  topUpList = lsLoad(LS_KEYS.topup, topUpList);
  terminationList = lsLoad(LS_KEYS.terminations, terminationList);
  outCoverageList = lsLoad(LS_KEYS.outCoverage, outCoverageList);
  staffList = lsLoad(LS_KEYS.staff, staffList);
  kpiList = lsLoad(LS_KEYS.kpis, kpiList);
  promotionList = lsLoad(LS_KEYS.promotions, promotionList);
  depositList = lsLoad(LS_KEYS.deposits, depositList).map(normalizeDeposit);
  coverageLocations = lsLoad(LS_KEYS.coverage, coverageLocations);
  // Ensure admin user always exists
  var hasAdmin = staffList.some(function(u) { return u.username === 'admin' && u.role === 'Admin'; });
  if (!hasAdmin) {
    staffList.unshift({ id: 'u1', name: 'Admin', username: 'admin', password: 'Tramkak@2026', role: 'Admin', branch: '', status: 'active' });
  }
  // Ensure default agent users always exist
  DEFAULT_AGENT_USERS.forEach(function(du) {
    if (!staffList.some(function(u) { return u.username === du.username; })) {
      staffList.push(du);
    }
  });
}

// ------------------------------------------------------------
// Helper Functions
// ------------------------------------------------------------
function g(id) { return document.getElementById(id); }
function rv(id) { const el = g(id); return el ? el.value.trim() : ''; }
function rt(id) { const el = g(id); return el ? el.value : ''; }
function $$(sel) { return document.querySelectorAll(sel); }
function uid() { return '_' + Math.random().toString(36).substr(2, 9); }
function ini(name) { return name.split(' ').map(w => w[0]).join('').toUpperCase().substring(0, 2); }
function fmtMoney(n, cur) { cur = cur !== undefined ? cur : '$'; return cur + Number(n).toFixed(2); }
function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function showToast(message, type) {
  var toast = document.createElement('div');
  var bg = (type === 'success') ? '#1B7D3D' : (type === 'error') ? '#C62828' : (type === 'warning') ? '#E65100' : '#333';
  toast.style.cssText = 'position:fixed;bottom:24px;right:24px;background:' + bg + ';color:#fff;padding:12px 20px;border-radius:10px;font-size:.875rem;font-weight:500;box-shadow:0 4px 16px rgba(0,0,0,.2);z-index:10000;opacity:0;transition:opacity .25s;max-width:320px;';
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(function() { toast.style.opacity = '1'; }, 10);
  setTimeout(function() { toast.style.opacity = '0'; setTimeout(function() { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 300); }, 3000);
}

// Professional alert dialog (replaces native alert)
function showAlert(message, type, title) {
  type = type || 'error';
  var iconMap = { error: 'fa-circle-xmark', success: 'fa-circle-check', warning: 'fa-triangle-exclamation', info: 'fa-circle-info' };
  var titleMap = { error: 'Error', success: 'Success', warning: 'Warning', info: 'Information' };
  var classMap = { error: 'dialog-icon-error', success: 'dialog-icon-success', warning: 'dialog-icon-warning', info: 'dialog-icon-info' };
  var iconEl = g('alert-icon');
  var iconWrap = g('alert-icon-wrap');
  var titleEl = g('alert-title');
  var msgEl = g('alert-message');
  if (iconEl) { iconEl.className = 'fas ' + (iconMap[type] || iconMap.error); }
  if (iconWrap) { iconWrap.className = 'dialog-icon-wrap ' + (classMap[type] || classMap.error); }
  if (titleEl) { titleEl.textContent = title || titleMap[type] || 'Alert'; }
  if (msgEl) { msgEl.textContent = message; }
  openModal('modal-alert');
}

// Professional confirm dialog (replaces native confirm)
var _confirmCallback = null;
function showConfirm(message, onConfirm, title, confirmText, isDanger) {
  _confirmCallback = onConfirm || null;
  var titleEl = g('confirm-title');
  var msgEl = g('confirm-message');
  var btn = g('confirm-ok-btn');
  var iconEl = g('confirm-icon');
  var iconWrap = g('confirm-icon-wrap');
  if (titleEl) titleEl.textContent = title || 'Confirm Action';
  if (msgEl) msgEl.textContent = message;
  if (btn) {
    btn.textContent = confirmText || 'Confirm';
    btn.className = isDanger === false ? 'btn btn-primary' : 'btn btn-danger';
  }
  if (iconEl) { iconEl.className = isDanger === false ? 'fas fa-circle-check' : 'fas fa-trash-can'; }
  if (iconWrap) { iconWrap.className = isDanger === false ? 'dialog-icon-wrap dialog-icon-info' : 'dialog-icon-wrap dialog-icon-delete'; }
  openModal('modal-confirm');
}

function _onConfirmAction() {
  closeModal('modal-confirm');
  if (typeof _confirmCallback === 'function') {
    var cb = _confirmCallback;
    _confirmCallback = null;
    cb();
  }
}

// Populate branch dropdowns
function getBranches() {
  var fromStaff = [...new Set(staffList.map(function(u) { return u.branch; }).filter(Boolean))].sort();
  // Merge staff branches with SMART_SHOPS to ensure all known shops are available
  return [...new Set([...fromStaff, ...SMART_SHOPS])].sort();
}

function populateBranchSelects() {
  const branches = getBranches();
  const branchSelectIds = ['sale-branch', 'nc-branch', 'tu-branch', 'term-branch', 'dep-branch'];
  branchSelectIds.forEach(function(id) {
    const sel = g(id);
    if (!sel) return;
    const current = sel.value;
    sel.innerHTML = '<option value="">Select branch</option>' +
      branches.map(function(b) { return '<option value="' + esc(b) + '"' + (current === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  });
}

// ------------------------------------------------------------
// Date Helpers
// ------------------------------------------------------------
function ymOf(dateStr) { return dateStr ? dateStr.substring(0, 7) : ''; }
function dateOf(ts) { return ts && typeof ts === 'string' ? ts.substring(0, 10) : ''; }
function ymNow() { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); }
function ymPrev() { const d = new Date(); d.setMonth(d.getMonth() - 1); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); }
function ymLabel(ym) {
  const parts = ym.split('-');
  return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+parts[1] - 1] + ' ' + parts[0];
}
function last7Months() {
  const r = [];
  const d = new Date();
  for (let i = 6; i >= 0; i--) {
    const t = new Date(d.getFullYear(), d.getMonth() - i, 1);
    r.push(t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0'));
  }
  return r;
}
// Parse a date string safely as a local-timezone Date to avoid UTC off-by-one.
// Handles "YYYY-MM-DD", ISO timestamps ("2026-03-19T17:00:00.000Z"), and "DD-MM-YYYY".
function parseLocalDate(str) {
  if (!str) return null;
  // YYYY-MM-DD — parse directly as local date (avoids UTC midnight shift)
  var iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str);
  if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  // ISO timestamp — convert UTC instant to local date components
  if (str.length > 10) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  // DD-MM-YYYY (Google Sheets locale variant)
  var dmy = /^(\d{1,2})-(\d{1,2})-(\d{4})$/.exec(str);
  if (dmy) return new Date(+dmy[3], +dmy[2] - 1, +dmy[1]);
  return null;
}

// Normalise any date value (string or timestamp) to a "YYYY-MM-DD" local date string.
function toLocalDateStr(str) {
  if (!str) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str; // already normalised
  var d = parseLocalDate(str);
  if (!d) return str;
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

// Returns today's date as a "YYYY-MM-DD" string in the local timezone.
function todayStr() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

// Returns the first day of the current month as a "YYYY-MM-DD" string.
function currentMonthStart() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-01';
}

// Auto-calculate expiry date for Top Up: start date + period * 30 days.
function calcTuExpiry() {
  var dateEl = g('tu-date');
  var periodEl = g('tu-period');
  var endDateEl = g('tu-end-date');
  if (!dateEl || !periodEl || !endDateEl) return;
  var dateVal = dateEl.value;
  var periodVal = parseInt(periodEl.value, 10);
  if (!dateVal || !periodVal || periodVal < 1) return;
  var parts = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateVal);
  if (!parts) return;
  var start = new Date(+parts[1], +parts[2] - 1, +parts[3]);
  start.setDate(start.getDate() + periodVal * 30);
  endDateEl.value = start.getFullYear() + '-' +
    String(start.getMonth() + 1).padStart(2, '0') + '-' +
    String(start.getDate()).padStart(2, '0');
}

function pctChange(curr, prev) {
  if (!prev) return null;
  return Math.round((curr - prev) / prev * 100);
}
function setTrend(elId, curr, prev) {
  const el = g(elId);
  if (!el) return;
  const pct = pctChange(curr, prev);
  if (pct === null) {
    el.innerHTML = '<i class="fas fa-minus"></i> N/A';
    el.className = 'trend-badge trend-flat';
  } else if (pct > 0) {
    el.innerHTML = '<i class="fas fa-arrow-up"></i> ' + pct + '%';
    el.className = 'trend-badge trend-up';
  } else if (pct < 0) {
    el.innerHTML = '<i class="fas fa-arrow-down"></i> ' + Math.abs(pct) + '%';
    el.className = 'trend-badge trend-down';
  } else {
    el.innerHTML = '<i class="fas fa-minus"></i> 0%';
    el.className = 'trend-badge trend-flat';
  }
}
function destroyChart(ref) {
  if (ref) { try { ref.destroy(); } catch (e) {} }
  return null;
}
function clearCanvas(id) {
  const c = g(id);
  if (c) { const ctx = c.getContext('2d'); ctx.clearRect(0, 0, c.width, c.height); }
}

// ------------------------------------------------------------
// Navigation
// ------------------------------------------------------------
function navigateTo(page, btn) {
  $$('.page').forEach(function(p) { p.classList.remove('active'); });
  const pg = g('page-' + page);
  if (pg) pg.classList.add('active');
  currentPage = page;

  $$('.nav-item').forEach(function(li) { li.classList.remove('active'); });
  if (btn) {
    const li = btn.closest ? btn.closest('.nav-item') : null;
    if (li) li.classList.add('active');
  }

  const titles = {
    dashboard: 'Dashboard',
    promotionPage: currentPromoView === 'expired' ? 'Expired Promotion' : 'New Promotion',
    deposit: 'Deposit',
    sale: 'Sale',
    performance: 'Performance Tracking',
    kpi: 'KPI Setting',
    customer: 'Customer',
    settings: 'Settings',
    'inv-sale-stock': 'Inventory – Sale Stock',
    'inv-shop-stock': 'Inventory – Shop Stock',
    coverage: 'Coverage'
  };
  const titleEl = g('page-title');
  if (titleEl) titleEl.textContent = titles[page] || page;

  populateBranchSelects();

  if (page === 'dashboard') renderDashboard();
  if (page === 'promotionPage') renderPromotionCards();
  if (page === 'kpi') { initKpiMonthPicker(); renderKpiTable(); }
  if (page === 'performance') {
    // Reset to Sale Tracking tab when (re-)entering the page
    currentPerfTab = 'tracking';
    $$('.perf-tab-content').forEach(function(c) { c.style.display = 'none'; });
    var tc = g('perf-tab-content-tracking'); if (tc) tc.style.display = '';
    $$('[id^="perf-tab-btn-"]').forEach(function(b) { b.classList.remove('active'); });
    var defBtn = g('perf-tab-btn-tracking'); if (defBtn) defBtn.classList.add('active');
    initSaleTracking();
  }
  if (page === 'deposit') { renderDepositTable(); updateDepositKpis(); }
  if (page === 'sale') { renderItemChips(); applyReportFilters(); applySaleTabTableVisibility(); }
  if (page === 'customer') {
    renderNewCustomerTable();
    renderTopUpTable();
    renderTerminationTable();
    renderOutCoverageTable();
  }
  if (page === 'settings') {
    renderStaffTable();
    renderAccessContent(currentSettingsTab);
  }
  if (page === 'inv-sale-stock') {
    // Reset to stock tab
    $$('.inv-tab-content').forEach(function(c) { c.classList.remove('active'); });
    var tc = g('invtab-content-stock'); if (tc) tc.classList.add('active');
    $$('.tab-btn[id^="invtab-"]').forEach(function(b) { b.classList.remove('active'); });
    var stockTabBtn = g('invtab-stock'); if (stockTabBtn) stockTabBtn.classList.add('active');
    renderInvSaleStock();
  }
  if (page === 'inv-shop-stock') {
    renderShopStockTable();
  }
  if (page === 'coverage') {
    initCoveragePage();
  }
}

function toggleSubmenu(id, elOrId) {
  const sub = g(id);
  if (!sub) return;
  const isOpen = sub.classList.contains('open');
  $$('.submenu').forEach(function(s) { s.classList.remove('open'); });
  $$('.has-submenu').forEach(function(li) { li.classList.remove('submenu-open'); });
  if (!isOpen) {
    sub.classList.add('open');
    const li = typeof elOrId === 'string' ? g(elOrId) : (elOrId && elOrId.closest ? elOrId.closest('.has-submenu') : null);
    if (li) li.classList.add('submenu-open');
  }
}

function openSaleTab(tab, btn) {
  currentSaleTab = tab;
  $$('.sale-tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  // Show/hide KPI card groups
  var unitKpi = g('sale-kpi-unit-cards');
  var pointKpi = g('sale-kpi-point-cards');
  if (unitKpi) unitKpi.style.display = (tab === 'unit') ? '' : 'none';
  if (pointKpi) pointKpi.style.display = (tab === 'point') ? '' : 'none';
  // Show/hide table cards (only matters in table view)
  applySaleTabTableVisibility();
}

function applySaleTabTableVisibility() {
  if (currentReportView !== 'table') return;
  var unitCard = g('sale-unit-table-card');
  var pointCard = g('sale-point-table-card');
  if (unitCard) unitCard.style.display = (currentSaleTab === 'point') ? 'none' : '';
  if (pointCard) pointCard.style.display = (currentSaleTab === 'unit') ? 'none' : '';
}

function switchSaleTab(tab) {
  openSaleTab(tab, g('tab-sale-btn-' + tab));
}

function openCustomerTab(tab, btn) {
  switchCustomerTab(tab);
  $$('.customer-tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
}

function openPerfTab(tab, btn) {
  currentPerfTab = tab;
  $$('.perf-tab-content').forEach(function(c) { c.style.display = 'none'; });
  var tc = g('perf-tab-content-' + tab);
  if (tc) tc.style.display = '';
  $$('[id^="perf-tab-btn-"]').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  if (tab === 'performance') renderPerformancePage();
}

function switchCustomerTab(tab) {
  currentCustomerTab = tab;
  $$('.customer-tab-content').forEach(function(c) { c.classList.remove('active'); });
  $$('.tab-content').forEach(function(c) { c.classList.remove('active'); });
  // Update tab button states
  $$('.tab-btn').forEach(function(b) {
    if (b.getAttribute('data-tab') === tab) b.classList.add('active');
    else if (['new-customer','topup','termination','out-coverage'].includes(b.getAttribute('data-tab'))) b.classList.remove('active');
  });
  const tc = g('tab-content-' + tab);
  if (tc) tc.classList.add('active');
}

function openSettingsTab(tab, btn) {
  switchSettingsTab(tab);
  $$('.settings-tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  renderAccessContent(tab);
}

function switchSettingsTab(tab) {
  currentSettingsTab = tab;
  $$('.settings-tab-content').forEach(function(c) { c.classList.remove('active'); });
  // Update tab button states
  $$('.tab-btn').forEach(function(b) {
    if (b.getAttribute('data-tab') === tab) b.classList.add('active');
    else if (['permission', 'google-sheets'].includes(b.getAttribute('data-tab'))) b.classList.remove('active');
  });
  const tc = g('stab-content-' + tab);
  if (tc) tc.classList.add('active');
}

function renderAccessContent(tab) {
  const allowed = TAB_PERM[currentRole] || [];
  var banner = g('settings-contact-banner');
  if (banner) banner.style.display = (currentRole !== 'admin' && currentRole !== 'cluster') ? '' : 'none';
  if (!allowed.includes(tab)) {
    const tc = g('stab-content-' + tab);
    if (tc) {
      tc.innerHTML = '<div class="access-denied"><i class="fas fa-lock fa-3x" style="color:#BDBDBD;margin-bottom:12px;"></i><h3 style="color:#555;">Access Denied</h3><p style="color:#999;">You do not have permission to access this section.</p></div>';
    }
  } else {
    if (tab === 'permission') renderStaffTable();
    if (tab === 'google-sheets') renderGoogleSheetsSettings();
  }
}

// ── Google Sheets Settings Panel ─────────────────────────────
function renderGoogleSheetsSettings() {
  var tc = g('stab-content-google-sheets');
  if (!tc) return;
  var current = gsUrl || GS_URL_DEFAULT;
  var isDefault = current === GS_URL_DEFAULT;
  tc.innerHTML =
    '<div class="page-header">' +
      '<h2 class="page-header-title">Google Sheets Integration</h2>' +
    '</div>' +
    '<div class="table-card" style="max-width:680px;">' +
      '<div style="padding:20px 24px;">' +
        '<p style="font-size:.875rem;color:#555;margin-bottom:18px;">' +
          'Enter the Google Apps Script Web App URL that connects this dashboard to your Google Spreadsheet. ' +
          'Deploy the <code>gas/Code.gs</code> script as a Web App (<em>Execute as: Me, Who has access: Anyone</em>), then paste the generated URL below.' +
        '</p>' +
        '<label style="display:block;font-weight:600;font-size:.875rem;color:#1A1A2E;margin-bottom:6px;" for="gs-url-input">' +
          'GAS Web App URL' +
        '</label>' +
        '<div style="display:flex;gap:10px;align-items:center;">' +
          '<input id="gs-url-input" type="url" class="form-input" style="flex:1;font-size:.8125rem;" ' +
            'placeholder="https://script.google.com/macros/s/…/exec" ' +
            'value="' + escHtml(current) + '" />' +
          '<button class="btn btn-primary" onclick="saveGasUrl()">' +
            '<i class="fas fa-floppy-disk"></i> Save' +
          '</button>' +
        '</div>' +
        (isDefault
          ? '<p style="font-size:.75rem;color:#888;margin-top:8px;"><i class="fas fa-circle-info" style="color:#1B7D3D;margin-right:4px;"></i>Currently using the default URL.</p>'
          : '<p style="font-size:.75rem;color:#1B7D3D;margin-top:8px;"><i class="fas fa-circle-check" style="margin-right:4px;"></i>Custom URL saved in this browser.</p>') +
        '<hr style="margin:20px 0;border:none;border-top:1px solid #e8ecf0;" />' +
        '<p style="font-weight:600;font-size:.875rem;color:#1A1A2E;margin-bottom:10px;">Sheets synced by this dashboard</p>' +
        '<div style="display:flex;flex-wrap:wrap;gap:8px;">' +
          SYNCED_SHEETS.map(function(s) {
            return '<span style="background:#e8f5e9;color:#1B7D3D;border-radius:4px;padding:3px 10px;font-size:.8125rem;font-weight:500;">' + s + '</span>';
          }).join('') +
        '</div>' +
        '<hr style="margin:20px 0;border:none;border-top:1px solid #e8ecf0;" />' +
        '<div style="display:flex;gap:12px;">' +
          '<button class="btn btn-primary" onclick="syncDownAll()">' +
            '<i class="fas fa-cloud-arrow-down"></i> Sync Down (Pull from Sheets)' +
          '</button>' +
          '<button class="btn btn-outline" onclick="syncUpAll()">' +
            '<i class="fas fa-cloud-arrow-up"></i> Sync Up (Push to Sheets)' +
          '</button>' +
        '</div>' +
        '<div style="margin-top:12px;">' +
          '<button class="btn btn-outline" style="color:#888;border-color:#ddd;font-size:.8125rem;" onclick="resetGasUrl()">' +
            '<i class="fas fa-rotate-left"></i> Reset to Default URL' +
          '</button>' +
        '</div>' +
      '</div>' +
    '</div>';
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function saveGasUrl() {
  var input = g('gs-url-input');
  if (!input) return;
  var val = input.value.trim();
  if (!val) {
    showAlert('Please enter a valid URL.', 'error');
    return;
  }
  if (!val.startsWith(GAS_URL_PREFIX)) {
    showAlert('URL must start with ' + GAS_URL_PREFIX, 'error');
    return;
  }
  gsUrl = val;
  try { localStorage.setItem(LS_KEYS.gasUrl, val); } catch(e) { console.warn('Could not save GAS URL:', e); }
  showToast('Google Sheets URL saved.', 'success');
  renderGoogleSheetsSettings();
}

function resetGasUrl() {
  gsUrl = GS_URL_DEFAULT;
  try { localStorage.removeItem(LS_KEYS.gasUrl); } catch(e) {}
  showToast('Reset to default Google Sheets URL.', 'success');
  renderGoogleSheetsSettings();
}

function setPromoView(view) {
  currentPromoView = view;
  var sectionAvailable = g('promo-section-available');
  var sectionExpired = g('promo-section-expired');
  var addBtn = g('promo-new-btn');
  var titleEl = g('promo-page-title');
  if (view === 'new') {
    if (sectionAvailable) sectionAvailable.style.display = '';
    if (sectionExpired) sectionExpired.style.display = 'none';
    if (addBtn) addBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';
    if (titleEl) titleEl.innerHTML = '<i class="fas fa-tags" style="color:#1B7D3D;margin-right:8px"></i>New Promotion';
    var pageTitleEl = g('page-title');
    if (pageTitleEl) pageTitleEl.textContent = 'New Promotion';
  } else {
    if (sectionAvailable) sectionAvailable.style.display = 'none';
    if (sectionExpired) sectionExpired.style.display = '';
    if (addBtn) addBtn.style.display = 'none';
    if (titleEl) titleEl.innerHTML = '<i class="fas fa-clock-rotate-left" style="color:#1B7D3D;margin-right:8px"></i>Expired Promotion';
    var pageTitleEl = g('page-title');
    if (pageTitleEl) pageTitleEl.textContent = 'Expired Promotion';
  }
}

function openPromoSubMenu(view, el) {
  navigateTo('promotionPage', null);
  setPromoView(view);
  setActiveSubItem(el);
}

function openSettingsMenu(el) {
  navigateTo('settings', null);
  // Derive tab name from element id: 'nav-permission' → 'permission', 'nav-google-sheets' → 'google-sheets'
  var tab = el && el.id ? el.id.replace(/^nav-/, '') : 'permission';
  switchSettingsTab(tab);
  renderAccessContent(tab);
  setActiveSubItem(el);
}

function setActiveSubItem(el) {
  $$('.submenu-item').forEach(function(li) { li.classList.remove('active'); });
  if (el) el.classList.add('active');
}

// ------------------------------------------------------------
// Role Switcher
// ------------------------------------------------------------
function switchRole(role) {
  currentRole = role;
  const roleNames = { admin: 'Admin User', cluster: 'Cluster User', supervisor: 'Supervisor User', agent: 'Agent User', user: 'Agent User' };
  const roleBadges = { admin: 'Admin', cluster: 'Cluster', supervisor: 'Supervisor', agent: 'Agent', user: 'Agent' };
  const roleColors = { admin: '#1B7D3D', cluster: '#6A1B9A', supervisor: '#1565C0', agent: '#E65100', user: '#E65100' };

  // For demo role switching: pick an appropriate representative user from staffList
  if (role === 'supervisor') {
    var sup = staffList.find(function(u) { return u.role === 'Supervisor' && u.status === 'active'; });
    if (sup) currentUser = sup;
  } else if (role === 'agent') {
    var agent = staffList.find(function(u) { return u.role === 'Agent' && u.status === 'active'; });
    if (agent) currentUser = agent;
  } else if (role === 'admin' || role === 'cluster') {
    var adminUser = staffList.find(function(u) { return (u.role === 'Admin' || u.role === 'Cluster') && u.status === 'active'; });
    if (adminUser) currentUser = adminUser;
  }

  const nameEl = g('topbar-name');
  const roleEl = g('topbar-role');
  const avatarEl = g('topbar-avatar');

  if (nameEl) nameEl.textContent = currentUser ? currentUser.name : roleNames[role];
  if (roleEl) { roleEl.textContent = roleBadges[role]; roleEl.style.background = roleColors[role]; }
  if (avatarEl) { avatarEl.textContent = ini(currentUser ? currentUser.name : roleNames[role]); avatarEl.style.background = roleColors[role]; }

  const rb = g('role-widget-btn');
  if (rb) { const lbl = rb.querySelector('#role-widget-label'); if (lbl) lbl.textContent = roleBadges[role]; }

  const wd = g('role-widget-dropdown');
  if (wd) wd.style.display = 'none';

  var banner = g('settings-contact-banner');
  if (banner) banner.style.display = (currentRole !== 'admin' && currentRole !== 'cluster') ? '' : 'none';

  var newBtn = g('promo-new-btn');
  if (newBtn) newBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';

  var saleNewBtn = g('sale-new-btn');
  if (saleNewBtn) saleNewBtn.style.display = (currentRole === 'cluster') ? 'none' : '';

  var covAddBtn = g('cov-add-btn');
  if (covAddBtn) covAddBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';

  if (currentPage === 'dashboard') renderDashboard();
  if (currentPage === 'settings') renderAccessContent(currentSettingsTab);
  if (currentPage === 'kpi') renderKpiTable();
  if (currentPage === 'sale') applyReportFilters();
  if (currentPage === 'coverage') initCoveragePage();
  if (currentPage === 'customer') { renderNewCustomerTable(); renderTopUpTable(); renderTerminationTable(); renderOutCoverageTable(); }
  if (currentPage === 'deposit') { renderDepositTable(); updateDepositKpis(); }

  // Always refresh item chips so add-button and chip interactivity reflect new role
  renderItemChips();
}

function toggleRoleWidget() {
  const wd = g('role-widget-dropdown');
  if (!wd) return;
  wd.style.display = wd.style.display === 'block' ? 'none' : 'block';
}

// ------------------------------------------------------------
// Modal Helpers
// ------------------------------------------------------------
function openModal(id) {
  populateBranchSelects();
  const el = g(id);
  if (el) { el.style.display = 'flex'; setTimeout(function() { el.classList.add('active'); }, 10); }
}

function closeModal(id) {
  const el = g(id);
  if (el) { el.classList.remove('active'); setTimeout(function() { el.style.display = 'none'; }, 300); }
  if (id === 'modal-topUp') {
    if (_tuPickerMap) { _tuPickerMap.remove(); _tuPickerMap = null; _tuPickerMarker = null; }
  }
}

function handleOverlay(e, id) {
  if (e.target === e.currentTarget) closeModal(id);
}

function openAddModal(type) {
  if (type === 'item') openItemModal();
  else if (type === 'sale') openNewSaleModal();
  else if (type === 'new-customer') openCustomerModal('new-customer');
  else if (type === 'topup') openCustomerModal('topup');
  else if (type === 'termination') openCustomerModal('termination');
  else if (type === 'kpi') openKpiModal();
  else if (type === 'user') openUserModal();
  else if (type === 'promotion') openPromotionModal();
  else if (type === 'deposit') openDepositModal();
}

function togglePwd(inputId, eyeId) {
  const inp = g(inputId);
  const eye = g(eyeId);
  if (!inp) return;
  if (inp.type === 'password') {
    inp.type = 'text';
    if (eye) eye.className = 'fas fa-eye-slash';
  } else {
    inp.type = 'password';
    if (eye) eye.className = 'fas fa-eye';
  }
}

// ------------------------------------------------------------
// Item Catalogue
// ------------------------------------------------------------
function openItemModal(item) {
  if (currentRole !== 'admin') { showAlert('Only admin can manage items.', 'error'); return; }
  const form = g('form-addItem');
  if (form) form.reset();

  selectItemGroup('unit');
  g('item-edit-id').value = '';

  const title = g('modal-addItem-title');
  const btn = g('item-submit-btn');

  if (item) {
    if (title) title.textContent = 'Edit Item';
    if (btn) btn.textContent = 'Update Item';
    g('item-edit-id').value = item.id;
    g('item-name').value = item.name || '';
    g('item-shortcut').value = item.shortcut || '';
    g('item-category').value = item.category || '';
    g('item-status').value = item.status || 'active';
    g('item-desc').value = item.desc || '';

    selectItemGroup(item.group || 'unit');

    if (item.group === 'unit') {
      const unitSel = g('item-unit');
      if (unitSel) {
        if (KNOWN_UNITS.includes(item.unit)) {
          unitSel.value = item.unit;
        } else {
          unitSel.value = 'custom';
          const cu = g('item-custom-unit');
          if (cu) { cu.value = item.unit || ''; cu.style.display = ''; }
        }
      }
    } else {
      const curSel = g('item-currency');
      if (curSel) {
        if (KNOWN_CURS.includes(item.currency)) {
          curSel.value = item.currency;
        } else {
          curSel.value = 'custom';
          const cc = g('item-custom-currency');
          if (cc) { cc.value = item.currency || ''; cc.style.display = ''; }
        }
      }
      const priceEl = g('item-price');
      if (priceEl) priceEl.value = item.price || '';
    }
  } else {
    if (title) title.textContent = 'Add Item';
    if (btn) btn.textContent = 'Add Item';
  }

  openModal('modal-addItem');
}

function selectItemGroup(grp) {
  itemGroupSelected = grp;
  const unitBtn = g('grp-btn-unit');
  const dollarBtn = g('grp-btn-dollar');
  const unitFields = g('item-unit-fields');
  const dollarFields = g('item-dollar-fields');

  if (grp === 'unit') {
    if (unitBtn) unitBtn.classList.add('active');
    if (dollarBtn) dollarBtn.classList.remove('active');
    if (unitFields) unitFields.style.display = '';
    if (dollarFields) dollarFields.style.display = 'none';
  } else {
    if (dollarBtn) dollarBtn.classList.add('active');
    if (unitBtn) unitBtn.classList.remove('active');
    if (unitFields) unitFields.style.display = 'none';
    if (dollarFields) dollarFields.style.display = '';
  }
}

function handleCurrencySelectChange() {
  const sel = g('item-currency');
  const inp = g('item-custom-currency');
  if (!inp) return;
  inp.style.display = sel && sel.value === 'custom' ? '' : 'none';
}

function handleUnitSelectChange() {
  const sel = g('item-unit');
  const inp = g('item-custom-unit');
  if (!inp) return;
  inp.style.display = sel && sel.value === 'custom' ? '' : 'none';
}

function submitItem(e) {
  e.preventDefault();
  const editId = rv('item-edit-id');
  const name = rv('item-name');
  const shortcut = rv('item-shortcut');
  const category = rv('item-category');
  const status = rv('item-status');
  const desc = rv('item-desc');
  const grp = itemGroupSelected;

  if (!name) { showAlert('Please enter item name'); return; }

  let unit = '', currency = '', price = 0;
  if (grp === 'unit') {
    const unitSel = g('item-unit');
    unit = unitSel && unitSel.value === 'custom' ? rv('item-custom-unit') : rv('item-unit');
  } else {
    const curSel = g('item-currency');
    currency = curSel && curSel.value === 'custom' ? rv('item-custom-currency') : rv('item-currency');
    price = parseFloat(rv('item-price')) || 0;
  }

  const obj = { id: editId || uid(), name: name, shortcut: shortcut, group: grp, unit: unit, currency: currency, price: price, category: category, status: status, desc: desc };

  if (editId) {
    const idx = itemCatalogue.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) itemCatalogue[idx] = obj;
  } else {
    itemCatalogue.push(obj);
  }

  closeModal('modal-addItem');
  renderItemChips();
  if (currentPage === 'dashboard') renderDashboard();
  saveAllData();
  showToast(editId ? 'Item updated.' : 'Item added successfully.', 'success');
}

function editItem(id) {
  if (currentRole !== 'admin') { showAlert('Only admin can edit items.', 'error'); return; }
  const item = itemCatalogue.find(function(x) { return x.id === id; });
  if (item) openItemModal(item);
}

function deleteItem(id) {
  if (currentRole !== 'admin') { showAlert('Only admin can delete items.', 'error'); return; }
  showConfirm('Are you sure you want to delete this item? This action cannot be undone.', function() {
    itemCatalogue = itemCatalogue.filter(function(x) { return x.id !== id; });
    renderItemChips();
    if (currentPage === 'dashboard') renderDashboard();
    saveAllData();
    showToast('Item deleted.', 'success');
  }, 'Delete Item', 'Delete');
}

function renderItemChips() {
  const strip = g('items-strip');
  if (!strip) return;
  const isAdmin = currentRole === 'admin';
  const addBtn = g('item-add-btn');
  if (addBtn) addBtn.style.display = isAdmin ? '' : 'none';
  const active = itemCatalogue.filter(function(x) { return x.status === 'active'; });
  if (!active.length) {
    strip.innerHTML = isAdmin
      ? '<span style="color:#999;font-size:0.85rem;">No items in catalogue. <a href="#" onclick="openAddModal(\'item\');return false;">Add Item</a></span>'
      : '<span style="color:#999;font-size:0.85rem;">No items in catalogue.</span>';
    return;
  }
  strip.innerHTML = active.map(function(item) {
    const chipClass = item.group === 'unit' ? 'item-chip-unit' : 'item-chip-dollar';
    if (isAdmin) {
      return '<span class="item-chip ' + chipClass + '" onclick="editItem(\'' + esc(item.id) + '\')" title="Click to edit: ' + esc(item.name) + '">' +
        esc(item.shortcut || item.name) + '</span>';
    }
    return '<span class="item-chip ' + chipClass + '" style="cursor:default;" title="' + esc(item.name) + '">' +
      esc(item.shortcut || item.name) + '</span>';
  }).join('');
}

// ------------------------------------------------------------
// New Sale Modal
// ------------------------------------------------------------
function switchSaleModalType(type) {
  currentSaleModalType = type;
  var unitBtn = g('sale-type-btn-unit');
  var pointBtn = g('sale-type-btn-point');
  var unitSection = g('sale-unit-section');
  var pointSection = g('sale-point-section');
  if (unitBtn) unitBtn.classList.toggle('active', type === 'unit');
  if (pointBtn) pointBtn.classList.toggle('active', type === 'point');
  if (unitSection) unitSection.style.display = (type === 'unit') ? '' : 'none';
  if (pointSection) pointSection.style.display = (type === 'point') ? '' : 'none';
  // Adjust required attributes
  var kpiServiceSel = g('sale-kpi-service');
  var kpiAmountInp = g('sale-kpi-amount');
  if (kpiServiceSel) kpiServiceSel.required = (type === 'point');
  if (kpiAmountInp) kpiAmountInp.required = (type === 'point');
}

function updateKpiPointPreview() {
  var serviceId = rv('sale-kpi-service');
  var amtInp = g('sale-kpi-amount');
  // eSIM and Smart NAS Download count as 1 unit transaction, not a dollar amount
  var isUnitCounted = (serviceId === 'ks13' || serviceId === 'ks14');
  if (amtInp) amtInp.required = !isUnitCounted;
  if (isUnitCounted && amtInp && amtInp.value === '') amtInp.value = '1';
  var amount = rv('sale-kpi-amount');
  var pts = calcKpiPoints(serviceId, amount);
  var el = g('kpi-points-value');
  if (el) el.textContent = pts % 1 === 0 ? pts.toString() : pts.toFixed(2);
}

function openNewSaleModal(sale) {
  const form = g('form-newSale');
  if (form) form.reset();
  g('sale-edit-id').value = '';

  const title = g('modal-newSale-title');
  const btn = g('sale-submit-btn');

  // Populate KPI service dropdown
  var kpiServiceSel = g('sale-kpi-service');
  if (kpiServiceSel) {
    var grouped = {};
    KPI_SERVICES.forEach(function(s) {
      if (!grouped[s.category]) grouped[s.category] = [];
      grouped[s.category].push(s);
    });
    kpiServiceSel.innerHTML = '<option value="">Select service</option>' +
      Object.keys(grouped).map(function(cat) {
        return '<optgroup label="' + esc(cat) + '">' +
          grouped[cat].map(function(s) {
            return '<option value="' + esc(s.id) + '">' + esc(s.name) +
              ' (×' + s.rate + (s.addOn ? ' +' + s.addOn : '') + ' pts)</option>';
          }).join('') + '</optgroup>';
      }).join('');
  }

  const unitContainer = g('sale-unit-items');
  const dollarContainer = g('sale-dollar-items');

  if (unitContainer) {
    unitContainer.innerHTML = '<div class="sale-items-grid">' + UNIT_SALE_ITEMS.map(function(item) {
      return '<div class="sic-card sic-card-unit">' +
        '<div class="sic-label">' + esc(item.name) + '</div>' +
        '<input type="number" class="sic-input" id="sic-' + esc(item.id) + '" min="0" value="" placeholder="0">' +
        '</div>';
    }).join('') + '</div>';
  }

  if (dollarContainer) {
    dollarContainer.innerHTML = '<div class="sale-items-grid">' + DOLLAR_SALE_ITEMS.map(function(item) {
      return '<div class="sic-card sic-card-dollar">' +
        '<div class="sic-label">' + esc(item.name) + '</div>' +
        '<input type="number" class="sic-input sic-dollar-input" id="sic-' + esc(item.id) + '" min="0" step="0.01" value="" placeholder="0.00">' +
        '</div>';
    }).join('') + '</div>';
  }

  const revTotalEl = g('sale-revenue-total');
  if (revTotalEl) {
    revTotalEl.style.display = '';
    revTotalEl.innerHTML = '<div class="sale-revenue-total-bar"><i class="fas fa-calculator"></i> Total Revenue (Auto Sum): <span id="sale-revenue-total-value">$0.00</span></div>';
  }

  function updateSaleRevenueTotal() {
    var sum = 0;
    DOLLAR_SALE_ITEMS.forEach(function(item) {
      var inp = g('sic-' + item.id);
      if (inp) sum += parseFloat(inp.value) || 0;
    });
    var el = g('sale-revenue-total-value');
    if (el) el.textContent = '$' + sum.toFixed(2);
  }

  if (dollarContainer) {
    dollarContainer.querySelectorAll('.sic-dollar-input').forEach(function(inp) {
      inp.addEventListener('input', function() { updateSaleRevenueTotal(); });
    });
  }

  populateBranchSelects();

  // Determine transaction type: use the sale's type when editing, otherwise default to the active tab
  var editType;
  if (sale) {
    editType = (sale.transactionType === 'point') ? 'point' : 'unit';
  } else {
    editType = (currentSaleTab === 'point') ? 'point' : 'unit';
  }
  switchSaleModalType(editType);

  if (sale) {
    if (title) title.textContent = 'Edit Sale';
    if (btn) btn.textContent = 'Update Sale';
    g('sale-edit-id').value = sale.id;
    g('sale-agent-name').value = sale.agent || '';
    const brSel = g('sale-branch');
    if (brSel) brSel.value = sale.branch || '';
    g('sale-date').value = sale.date || '';
    g('sale-remark').value = sale.remark || sale.note || '';

    if (editType === 'point') {
      if (kpiServiceSel) kpiServiceSel.value = sale.kpiService || '';
      var kpiAmtInp = g('sale-kpi-amount');
      if (kpiAmtInp) kpiAmtInp.value = sale.kpiAmount || '';
      var kpiPhoneInp = g('sale-kpi-phone');
      if (kpiPhoneInp) kpiPhoneInp.value = sale.customerPhone || '';
      updateKpiPointPreview();
    } else {
      if (sale.items) {
        Object.keys(sale.items).forEach(function(iid) {
          const inp = g('sic-' + iid);
          if (inp) inp.value = sale.items[iid];
        });
      }
      if (sale.dollarItems) {
        Object.keys(sale.dollarItems).forEach(function(iid) {
          const inp = g('sic-' + iid);
          if (inp) inp.value = sale.dollarItems[iid];
        });
      }
      updateSaleRevenueTotal();
    }
  } else {
    if (title) title.textContent = 'New Sale';
    if (btn) btn.textContent = 'Save Sale';
    g('sale-date').value = todayStr();
    if (currentUser) {
      const agentEl = g('sale-agent-name');
      if (agentEl) { agentEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agentEl.readOnly = true; }
      const brEl = g('sale-branch');
      if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
    }
  }

  openModal('modal-newSale');
}

function submitSale(e) {
  e.preventDefault();
  const editId = rv('sale-edit-id');
  const agent = rv('sale-agent-name');
  const branch = rv('sale-branch');
  const date = rv('sale-date');
  const note = rv('sale-remark');

  if (!agent) { showAlert('Please enter agent name'); return; }
  if (!date) { showAlert('Please select date'); return; }

  const now = new Date().toISOString();
  const existingRecord = editId ? saleRecords.find(function(x) { return x.id === editId; }) : null;
  var obj;

  if (currentSaleModalType === 'point') {
    // Point-Based transaction
    var kpiService = rv('sale-kpi-service');
    var kpiAmount = parseFloat(rv('sale-kpi-amount')) || 0;
    var customerPhone = rv('sale-kpi-phone');
    if (!kpiService) { showAlert('Please select a service type'); return; }
    // eSIM and Smart NAS Download count as 1 unit transaction, not a dollar amount
    if ((kpiService === 'ks13' || kpiService === 'ks14') && kpiAmount <= 0) { kpiAmount = 1; }
    if (kpiAmount <= 0) { showAlert('Please enter a valid amount'); return; }
    var kpiPoints = calcKpiPoints(kpiService, kpiAmount);
    obj = {
      id: editId || uid(),
      transactionType: 'point',
      agent: agent,
      branch: branch,
      date: date,
      submittedAt: (existingRecord && existingRecord.submittedAt) || now,
      note: note,
      kpiService: kpiService,
      kpiAmount: kpiAmount,
      kpiPoints: kpiPoints,
      customerPhone: customerPhone,
      items: {},
      dollarItems: {}
    };
  } else {
    // Unit-Based transaction (existing logic)
    const items = {}, dollarItems = {};
    let autoRevenue = 0;
    UNIT_SALE_ITEMS.forEach(function(item) {
      const inp = g('sic-' + item.id);
      if (!inp) return;
      const val = parseFloat(inp.value) || 0;
      if (val > 0) items[item.id] = val;
    });
    DOLLAR_SALE_ITEMS.forEach(function(item) {
      const inp = g('sic-' + item.id);
      if (!inp) return;
      const val = parseFloat(inp.value) || 0;
      if (val > 0) dollarItems[item.id] = val;
      autoRevenue += val;
    });
    if (autoRevenue > 0) dollarItems[ITEM_ID_REVENUE] = autoRevenue;
    obj = {
      id: editId || uid(),
      transactionType: 'unit',
      agent: agent,
      branch: branch,
      date: date,
      submittedAt: (existingRecord && existingRecord.submittedAt) || now,
      note: note,
      items: items,
      dollarItems: dollarItems
    };
  }

  if (editId) {
    if (existingRecord && !canModifySaleRecord(existingRecord)) { showSalePermissionError('edit'); return; }
    const idx = saleRecords.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) saleRecords[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a sale record.');
  } else {
    saleRecords.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' added a new sale record.');
  }

  closeModal('modal-newSale');
  applyReportFilters();
  if (currentPage === 'dashboard') renderDashboard();
  if (currentPage === 'performance') renderPerformancePage();
  syncSheet('DailySale', saleRecords);
  saveAllData();
  showToast(editId ? 'Sale record updated.' : 'Sale record added successfully.', 'success');
}

function editSale(id) {
  const sale = saleRecords.find(function(x) { return x.id === id; });
  if (!sale) return;
  if (!canModifySaleRecord(sale)) { showSalePermissionError('edit'); return; }
  openNewSaleModal(sale);
}

function deleteSale(id) {
  const sale = saleRecords.find(function(x) { return x.id === id; });
  if (!sale) return;
  if (!canModifySaleRecord(sale)) { showSalePermissionError('delete'); return; }
  showConfirm('Are you sure you want to delete this sale record? This action cannot be undone.', function() {
    saleRecords = saleRecords.filter(function(x) { return x.id !== id; });
    applyReportFilters();
    if (currentPage === 'dashboard') renderDashboard();
    syncSheet('DailySale', saleRecords);
    saveAllData();
    showToast('Sale record deleted.', 'success');
  }, 'Delete Sale Record', 'Delete');
}

// ------------------------------------------------------------
// Sale Filters & Table
// ------------------------------------------------------------

// Returns the base sale records filtered by the current user's role.
// Agents only see their own records; supervisors see all records in their branch/shop; admin/cluster see all.
function getSaleBaseRecords() {
  if (currentRole === 'agent' && currentUser) {
    return saleRecords.filter(function(s) { return s.agent === currentUser.name; });
  }
  if (currentRole === 'supervisor' && currentUser) {
    return saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
  }
  return saleRecords.slice();
}

// Returns true if the current user is permitted to edit or delete the given sale record.
function canModifySaleRecord(sale) {
  if (!sale) return false;
  if (currentRole === 'admin') return true;
  if (currentRole === 'cluster') return false;
  if (!currentUser) return false;
  if (currentRole === 'supervisor') return sale.branch === currentUser.branch;
  if (currentRole === 'agent') return sale.agent === currentUser.name;
  return false;
}

// Returns records filtered by the current user's role.
// Agents only see their own records; supervisors see all records in their branch/shop.
function getBaseRecordsForRole(list) {
  if (currentRole === 'agent' && currentUser) {
    return list.filter(function(r) { return r.agent === currentUser.name; });
  }
  if (currentRole === 'supervisor' && currentUser) {
    return list.filter(function(r) { return r.branch === currentUser.branch; });
  }
  return list.slice();
}

// Returns true if the current user can edit or delete the given customer/deposit record.
function canModifyRecord(record) {
  if (!record) return false;
  if (currentRole === 'admin') return true;
  if (currentRole === 'cluster') return false;
  if (!currentUser) return false;
  if (currentRole === 'supervisor') return record.branch === currentUser.branch;
  if (currentRole === 'agent') return record.agent === currentUser.name;
  return false;
}

// Returns true if the current user can create, edit, or delete KPIs.
// Admin and cluster have full access; supervisors can manage KPIs for their shop/agents; agents cannot.
function canManageKpis() {
  return currentRole === 'admin' || currentRole === 'cluster' || currentRole === 'supervisor';
}

// Returns true if a supervisor can modify the given KPI (must be their shop KPI or an agent KPI in their branch).
function canSupervisorModifyKpi(kpi) {
  if (!kpi || !currentUser) return false;
  return (kpi.kpiFor === 'shop' && kpi.assigneeId === currentUser.id) ||
         (kpi.kpiFor === 'agent' && kpi.assigneeBranch === currentUser.branch);
}

// Returns the subset of kpiList visible to the current user.
// Admin/cluster see all KPIs; supervisor/agent see only their own branch's KPIs.
function getVisibleKpis() {
  if (currentRole === 'admin' || currentRole === 'cluster') return kpiList.slice();
  if (!currentUser || !currentUser.branch) return [];
  var userBranch = currentUser.branch;
  return kpiList.filter(function(k) {
    if (k.kpiFor === 'shop') {
      var sup = staffList.find(function(u) { return u.id === k.assigneeId; });
      return sup && sup.branch === userBranch;
    } else if (k.kpiFor === 'agent') {
      return k.assigneeBranch === userBranch;
    }
    return false;
  });
}

// Shows a role-appropriate permission error for sale report modification attempts.
function showSalePermissionError(action) {
  if (currentRole === 'cluster') {
    showAlert('You do not have permission to ' + action + ' sale reports.', 'error');
  } else if (currentRole === 'agent') {
    showAlert('You can only ' + action + ' your own sale reports.', 'error');
  } else {
    showAlert('You can only ' + action + ' sale reports within your branch.', 'error');
  }
}

function applyReportFilters() {
  const baseRecords = getSaleBaseRecords();

  const dateFrom = rv('sale-date-from');
  const dateTo = rv('sale-date-to');
  const agent = rv('sale-filter-agent');
  const branch = rv('sale-filter-branch');

  filteredSales = baseRecords.filter(function(s) {
    if (dateFrom && s.date < dateFrom) return false;
    if (dateTo && s.date > dateTo) return false;
    if (agent && s.agent !== agent) return false;
    if (branch && s.branch !== branch) return false;
    return true;
  });

  // Reset show-all when filters change so user sees the latest 5 again
  saleUnitShowAll = false;
  salePointShowAll = false;

  renderSaleTable();
  updateSaleKpis();
  // If the summary/chart view is currently active, refresh it too so charts stay up to date
  if (currentReportView === 'summary') {
    var unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
    var dollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active'; });
    renderSummaryView(filteredSales, unitItems, dollarItems);
    // Defer chart render by one tick so the DOM update from renderSummaryView completes first
    setTimeout(renderSaleCharts, 50);
  }
}

function clearReportFilters() {
  var mStart = currentMonthStart();
  var mEnd = todayStr();
  var fromEl = g('sale-date-from'); if (fromEl) fromEl.value = mStart;
  var toEl = g('sale-date-to'); if (toEl) toEl.value = mEnd;
  ['sale-date-search', 'sale-filter-agent', 'sale-filter-branch'].forEach(function(id) {
    const el = g(id); if (el) el.value = '';
  });
  saleUnitShowAll = false;
  salePointShowAll = false;
  filteredSales = getSaleBaseRecords().filter(function(s) { return s.date >= mStart && s.date <= mEnd; });
  renderSaleTable();
  updateSaleKpis();
}

function setSaleDateSearch(dateVal) {
  // Set both From and To to the same date for a quick single-date search
  const fromEl = g('sale-date-from');
  const toEl = g('sale-date-to');
  if (fromEl) fromEl.value = dateVal;
  if (toEl) toEl.value = dateVal;
  applyReportFilters();
}

function clearSaleDateSearch() {
  const el = g('sale-date-search');
  if (el) el.value = '';
}

function setReportView(view) {
  currentReportView = view;
  $$('.view-toggle-btn').forEach(function(b) { b.classList.remove('active'); });
  const btn = g('btn-view-' + view);
  if (btn) btn.classList.add('active');

  const tableCard = g('sale-table-card');
  const summaryView = g('sale-summary-view');

  if (view === 'table') {
    if (tableCard) tableCard.style.display = '';
    if (summaryView) summaryView.style.display = 'none';
    applySaleTabTableVisibility();
  } else {
    if (tableCard) tableCard.style.display = 'none';
    if (summaryView) summaryView.style.display = '';
    const unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
    const dollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active'; });
    renderSummaryView(filteredSales, unitItems, dollarItems);
    renderSaleCharts();
  }
}

function updateSaleKpis() {
  const data = filteredSales;
  let totalRevenue = 0, totalRecharge = 0, totalGrossAds = 0, totalHomeInternet = 0, totalPoints = 0;

  data.forEach(function(s) {
    // Total Revenue: sum of Revenue (RV) dollar item
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) totalRevenue += s.dollarItems[ITEM_ID_REVENUE];
    // Total Recharge: sum of Recharge (RC) dollar item
    if (s.dollarItems && s.dollarItems[ITEM_ID_RECHARGE]) totalRecharge += s.dollarItems[ITEM_ID_RECHARGE];
    // Total Gross Ads: sum of Gross Ads (GA) unit item
    if (s.items && s.items[ITEM_ID_GROSS_ADS]) totalGrossAds += s.items[ITEM_ID_GROSS_ADS];
    // Total Home Internet: sum of Smart@Home (SH) + Smart Fiber+ (SF)
    if (s.items && s.items[ITEM_ID_SMART_HOME]) totalHomeInternet += s.items[ITEM_ID_SMART_HOME];
    if (s.items && s.items[ITEM_ID_SMART_FIBER]) totalHomeInternet += s.items[ITEM_ID_SMART_FIBER];
    // Total Points (Point-Based transactions)
    if (s.transactionType === 'point' && s.kpiPoints) totalPoints += parseFloat(s.kpiPoints) || 0;
  });

  const el1 = g('sale-kpi-revenue'); if (el1) el1.textContent = fmtMoney(totalRevenue);
  const el2 = g('sale-kpi-recharge'); if (el2) el2.textContent = fmtMoney(totalRecharge);
  const el3 = g('sale-kpi-gross-ads'); if (el3) el3.textContent = totalGrossAds;
  const el4 = g('sale-kpi-home-internet'); if (el4) el4.textContent = totalHomeInternet;
  const el5 = g('sale-kpi-points'); if (el5) el5.textContent = totalPoints % 1 === 0 ? totalPoints.toString() : totalPoints.toFixed(2);
}

function renderSaleTable() {
  var unitTable  = g('sale-unit-table');
  var pointTable = g('sale-point-table');
  if (!unitTable && !pointTable) return;

  // Role-based base records for filter dropdowns
  const baseRecords = getSaleBaseRecords();

  // Show/hide branch filter: hidden for agent and supervisor (supervisor is locked to their branch)
  const branchFilterWrap = g('sale-branch-filter-wrap');
  if (branchFilterWrap) {
    branchFilterWrap.style.display = (currentRole === 'agent' || currentRole === 'supervisor') ? 'none' : '';
  }

  // Show/hide New Sale button: cluster cannot create/edit reports
  const saleNewBtn = g('sale-new-btn');
  if (saleNewBtn) {
    saleNewBtn.style.display = (currentRole === 'cluster') ? 'none' : '';
  }

  // Populate filter dropdowns
  const agentFilter = g('sale-filter-agent');
  const branchFilter = g('sale-filter-branch');
  if (agentFilter) {
    const agents = [...new Set(baseRecords.map(function(s) { return s.agent; }))];
    const curAgent = agentFilter.value;
    agentFilter.innerHTML = '<option value="">All Agents</option>' +
      agents.map(function(a) { return '<option value="' + esc(a) + '"' + (curAgent === a ? ' selected' : '') + '>' + esc(a) + '</option>'; }).join('');
  }
  if (branchFilter) {
    const branches = [...new Set(baseRecords.map(function(s) { return s.branch; }))];
    const curBranch = branchFilter.value;
    branchFilter.innerHTML = '<option value="">All Branches</option>' +
      branches.map(function(b) { return '<option value="' + esc(b) + '"' + (curBranch === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  }

  // Update period badges to show current filter range
  var filterFrom = rv('sale-date-from');
  var filterTo   = rv('sale-date-to');
  var mStart = currentMonthStart();
  var mToday = todayStr();
  var isDefaultMonth = filterFrom === mStart && filterTo === mToday;
  var periodLabel;
  if (isDefaultMonth) {
    periodLabel = 'This Month';
  } else if (filterFrom && filterTo) {
    periodLabel = filterFrom + ' → ' + filterTo;
  } else if (filterFrom) {
    periodLabel = 'From ' + filterFrom;
  } else if (filterTo) {
    periodLabel = 'Until ' + filterTo;
  } else {
    periodLabel = 'All Time';
  }
  var unitBadge = g('sale-unit-period-badge');   if (unitBadge)  unitBadge.textContent  = periodLabel;
  var pointBadge = g('sale-point-period-badge'); if (pointBadge) pointBadge.textContent = periodLabel;

  const data = filteredSales;
  const unitItems   = itemCatalogue.filter(function(x) { return x.group === 'unit'   && x.status === 'active'; });
  const dollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active' && x.id !== ITEM_ID_REVENUE; });

  // Sort all data newest first
  var sortNewest = function(a, b) {
    var da = (a.submittedAt || a.date || '');
    var db = (b.submittedAt || b.date || '');
    return db.localeCompare(da);
  };

  const unitDataAll  = data.filter(function(s) { return s.transactionType !== 'point'; }).sort(sortNewest);
  const pointDataAll = data.filter(function(s) { return s.transactionType === 'point'; }).sort(sortNewest);

  // Pre-calculate totals from ALL filtered data (not just the displayed 5)
  let totalUnits = 0, totalRevenue = 0, totalPoints = 0, totalDollar = 0;
  unitDataAll.forEach(function(s) {
    unitItems.forEach(function(item) {
      if (s.items && s.items[item.id]) totalUnits += s.items[item.id];
    });
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) totalRevenue += s.dollarItems[ITEM_ID_REVENUE];
    dollarItems.forEach(function(item) {
      if (s.dollarItems && s.dollarItems[item.id]) totalDollar += s.dollarItems[item.id];
    });
  });
  pointDataAll.forEach(function(s) { totalPoints += parseFloat(s.kpiPoints) || 0; });

  // Update unit table header summary
  var unitHeaderTotals = g('sale-unit-header-totals');
  if (unitHeaderTotals) {
    unitHeaderTotals.innerHTML = unitDataAll.length
      ? '<span style="color:#2d6a4f;font-weight:600;">Units: ' + totalUnits + '</span>'
        + '<span style="margin:0 6px;color:#ccc;">|</span>'
        + '<span style="color:#555;">Rev: ' + fmtMoney(totalRevenue) + '</span>'
        + (unitDataAll.length <= 5 ? '' : '<span style="margin:0 6px;color:#ccc;">|</span><span style="color:#888;">' + unitDataAll.length + ' records</span>')
      : '<span style="color:#aaa;font-style:italic;">No data</span>';
  }

  // Update point table header summary
  var pointHeaderTotals = g('sale-point-header-totals');
  if (pointHeaderTotals) {
    pointHeaderTotals.innerHTML = pointDataAll.length
      ? '<span style="color:#FF9800;font-weight:600;">Points: ' + (totalPoints % 1 === 0 ? totalPoints : totalPoints.toFixed(2)) + '</span>'
        + (pointDataAll.length <= 5 ? '' : '<span style="margin:0 6px;color:#ccc;">|</span><span style="color:#888;">' + pointDataAll.length + ' records</span>')
      : '<span style="color:#aaa;font-style:italic;">No data</span>';
  }

  // Update toggle buttons
  var unitToggle = g('sale-unit-view-toggle');
  if (unitToggle) {
    if (unitDataAll.length <= 5) {
      unitToggle.style.display = 'none';
    } else {
      unitToggle.style.display = '';
      unitToggle.innerHTML = saleUnitShowAll
        ? '<i class="fas fa-compress-alt"></i> Last 5'
        : '<i class="fas fa-list"></i> View All (' + unitDataAll.length + ')';
    }
  }
  var pointToggle = g('sale-point-view-toggle');
  if (pointToggle) {
    if (pointDataAll.length <= 5) {
      pointToggle.style.display = 'none';
    } else {
      pointToggle.style.display = '';
      pointToggle.innerHTML = salePointShowAll
        ? '<i class="fas fa-compress-alt"></i> Last 5'
        : '<i class="fas fa-list"></i> View All (' + pointDataAll.length + ')';
    }
  }

  // Limit rows for display
  const unitData  = saleUnitShowAll  ? unitDataAll  : unitDataAll.slice(0, 5);
  const pointData = salePointShowAll ? pointDataAll : pointDataAll.slice(0, 5);

  // ---- Unit-Based Performance Table ----
  if (unitTable) {
    if (!unitDataAll.length) {
      unitTable.innerHTML = '<thead></thead><tbody><tr><td colspan="20" style="text-align:center;padding:30px;color:#999;"><i class="fas fa-inbox" style="font-size:1.5rem;display:block;margin-bottom:8px;"></i>No unit-based records found</td></tr></tbody>';
    } else {
      let uHeader1 = '<tr><th rowspan="2">Agent</th><th rowspan="2">Branch</th><th rowspan="2">Submit Date</th>';
      if (unitItems.length)   uHeader1 += '<th colspan="' + unitItems.length   + '" class="th-group-unit">Unit Group</th>';
      if (dollarItems.length) uHeader1 += '<th colspan="' + dollarItems.length + '" class="th-group-dollar">Dollar Group</th>';
      uHeader1 += '<th rowspan="2" class="td-buy-number">Revenue</th><th rowspan="2">Remark</th><th rowspan="2">Actions</th></tr>';

      let uHeader2 = '<tr>';
      unitItems.forEach(function(item)   { uHeader2 += '<th class="th-unit">'   + esc(item.shortcut || item.name) + '</th>'; });
      dollarItems.forEach(function(item) { uHeader2 += '<th class="th-dollar">' + esc(item.shortcut || item.name) + '</th>'; });
      uHeader2 += '</tr>';

      const uRows = unitData.map(function(s) {
        const canEdit    = canModifySaleRecord(s);
        const avIdx      = Math.abs((s.agent.charCodeAt(0) || 0)) % 8;
        const submitDate = dateOf(s.submittedAt) || dateOf(s.date);

        const unitCells = unitItems.map(function(item) {
          const qty = s.items && s.items[item.id] ? s.items[item.id] : 0;
          return '<td class="td-unit">' + (qty || '') + '</td>';
        }).join('');

        const dCells = dollarItems.map(function(item) {
          const amt = s.dollarItems && s.dollarItems[item.id] ? s.dollarItems[item.id] : 0;
          return '<td class="td-dollar">' + (amt > 0 ? fmtMoney(amt, esc(item.currency) + ' ') : '') + '</td>';
        }).join('');

        const saleRev  = s.dollarItems && s.dollarItems[ITEM_ID_REVENUE] ? s.dollarItems[ITEM_ID_REVENUE] : 0;
        const revCell  = '<td class="td-buy-number">' + fmtMoney(saleRev) + '</td>';
        const remCell  = '<td style="color:#888;font-size:0.8rem;">' + esc(s.note || s.remark || '') + '</td>';

        return '<tr>' +
          '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(s.agent)) + '</span>' + esc(s.agent) + '</div></td>' +
          '<td>' + esc(s.branch) + '</td>' +
          '<td style="white-space:nowrap;font-size:0.82rem;color:#555;">' + esc(submitDate) + '</td>' +
          unitCells + dCells + revCell + remCell +
          '<td style="white-space:nowrap;">' +
            (canEdit ? '<button class="btn-edit" onclick="editSale(\'' + esc(s.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
            (canEdit ? '<button class="btn-delete" onclick="deleteSale(\'' + esc(s.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
          '</td></tr>';
      }).join('');

      const uThead = unitTable.querySelector('thead') || document.createElement('thead');
      const uTbody = unitTable.querySelector('tbody') || document.createElement('tbody');
      uThead.innerHTML = uHeader1 + uHeader2;
      uTbody.innerHTML = uRows;
      if (!unitTable.contains(uTbody)) unitTable.appendChild(uTbody);
      if (unitTable.firstChild !== uThead) unitTable.insertBefore(uThead, unitTable.firstChild);
    }
  }

  // ---- Point-Based Performance Table ----
  if (pointTable) {
    if (!pointDataAll.length) {
      pointTable.innerHTML = '<thead></thead><tbody><tr><td colspan="9" style="text-align:center;padding:30px;color:#999;"><i class="fas fa-inbox" style="font-size:1.5rem;display:block;margin-bottom:8px;"></i>No point-based records found</td></tr></tbody>';
    } else {
      const pHeader = '<tr>' +
        '<th>Agent</th><th>Branch</th><th>Submit Date</th><th>Service</th>' +
        '<th style="color:#FF9800;">Points</th><th>Amount</th><th>Phone</th>' +
        '<th>Remark</th><th>Actions</th></tr>';

      const pRows = pointData.map(function(s) {
        const canEdit    = canModifySaleRecord(s);
        const avIdx      = Math.abs((s.agent.charCodeAt(0) || 0)) % 8;
        const submitDate = dateOf(s.submittedAt) || dateOf(s.date);
        var svc     = KPI_SERVICES.find(function(x) { return x.id === s.kpiService; });
        var svcName = svc ? svc.name + ' (' + svc.category + ')' : (s.kpiService || '');
        var pts     = parseFloat(s.kpiPoints) || 0;

        return '<tr>' +
          '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(s.agent)) + '</span>' + esc(s.agent) + '</div></td>' +
          '<td>' + esc(s.branch) + '</td>' +
          '<td style="white-space:nowrap;font-size:0.82rem;color:#555;">' + esc(submitDate) + '</td>' +
          '<td style="font-size:0.85rem;color:#555;">' + esc(svcName) + '</td>' +
          '<td style="text-align:center;"><span style="color:#FF9800;font-weight:700;">' + (pts % 1 === 0 ? pts : pts.toFixed(2)) + '</span> <small style="color:#aaa;">pts</small></td>' +
          '<td style="text-align:right;">' + (s.kpiAmount ? '<span style="color:#388e3c;">' + fmtMoney(parseFloat(s.kpiAmount)) + '</span>' : '') + '</td>' +
          '<td style="font-size:0.82rem;color:#888;">' + (s.customerPhone ? '<i class="fas fa-phone" style="font-size:0.75rem;margin-right:4px;"></i>' + esc(s.customerPhone) : '') + '</td>' +
          '<td style="color:#888;font-size:0.8rem;">' + esc(s.note || s.remark || '') + '</td>' +
          '<td style="white-space:nowrap;">' +
            (canEdit ? '<button class="btn-edit" onclick="editSale(\'' + esc(s.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
            (canEdit ? '<button class="btn-delete" onclick="deleteSale(\'' + esc(s.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
          '</td></tr>';
      }).join('');

      const pThead = pointTable.querySelector('thead') || document.createElement('thead');
      const pTbody = pointTable.querySelector('tbody') || document.createElement('tbody');
      pThead.innerHTML = pHeader;
      pTbody.innerHTML = pRows;
      if (!pointTable.contains(pTbody)) pointTable.appendChild(pTbody);
      if (pointTable.firstChild !== pThead) pointTable.insertBefore(pThead, pointTable.firstChild);
    }
  }

  updateTotalBar(totalUnits, totalDollar, totalPoints);
}

function updateTotalBar(units, dollar, points) {
  const bar = g('sale-total-bar');
  if (!bar) return;
  bar.innerHTML =
    '<span class="total-label"><strong>Total Units:</strong> ' + units + '</span>' +
    '<span class="total-label"><strong>Total $:</strong> ' + fmtMoney(dollar) + '</span>' +
    (points ? '<span class="total-label"><strong>Total Points:</strong> <span style="color:#FF9800;">' + (Number.isInteger(points) ? points : points.toFixed(2)) + '</span></span>' : '');
}

// ------------------------------------------------------------
// Sale CSV Download
// ------------------------------------------------------------
function openDownloadModal() {
  const today = todayStr();
  const fromEl = g('dl-date-from');
  const toEl = g('dl-date-to');
  if (fromEl && !fromEl.value) fromEl.value = today.slice(0, 7) + '-01';
  if (toEl && !toEl.value) toEl.value = today;

  // Show/hide role-based filter sections
  const shopSection = g('dl-shop-section');
  const agentSection = g('dl-agent-section');
  if (shopSection) shopSection.style.display = (currentRole === 'cluster') ? '' : 'none';
  if (agentSection) agentSection.style.display = (currentRole === 'supervisor') ? '' : 'none';

  // Populate cluster shop (branch) filter
  if (currentRole === 'cluster') {
    const shopFilter = g('dl-shop-filter');
    if (shopFilter) {
      const branches = [...new Set(saleRecords.map(function(s) { return s.branch; }).filter(Boolean))].sort();
      shopFilter.innerHTML = '<option value="">All Shops</option>' +
        branches.map(function(b) { return '<option value="' + esc(b) + '">' + esc(b) + '</option>'; }).join('');
    }
  }

  // Populate supervisor agent filter (scoped to their branch)
  if (currentRole === 'supervisor' && currentUser) {
    const agentFilter = g('dl-agent-filter');
    if (agentFilter) {
      const branchRecords = saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
      const agents = [...new Set(branchRecords.map(function(s) { return s.agent; }).filter(Boolean))].sort();
      agentFilter.innerHTML = '<option value="">All Agents</option>' +
        agents.map(function(a) { return '<option value="' + esc(a) + '">' + esc(a) + '</option>'; }).join('');
    }
  }

  openModal('modal-saleDownload');
}

function downloadSaleCSV() {
  const dateFrom = rv('dl-date-from');
  const dateTo = rv('dl-date-to');
  const shopFilter = (currentRole === 'cluster') ? rv('dl-shop-filter') : '';
  const agentFilter = (currentRole === 'supervisor') ? rv('dl-agent-filter') : '';

  const base = getSaleBaseRecords();
  const rows = base.filter(function(s) {
    if (dateFrom && s.date < dateFrom) return false;
    if (dateTo && s.date > dateTo) return false;
    if (shopFilter && s.branch !== shopFilter) return false;
    if (agentFilter && s.agent !== agentFilter) return false;
    return true;
  });

  if (!rows.length) { showAlert('No records found for the selected filters.'); return; }

  // Only include item columns that have at least one non-zero value in the filtered rows
  const allUnitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
  const allDollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active' && x.id !== ITEM_ID_REVENUE; });
  const unitItems = allUnitItems.filter(function(item) { return rows.some(function(s) { return s.items && s.items[item.id] > 0; }); });
  const dollarItems = allDollarItems.filter(function(item) { return rows.some(function(s) { return s.dollarItems && s.dollarItems[item.id] > 0; }); });

  const escape = function(v) {
    const s = String(v === null || v === undefined ? '' : v);
    return '"' + s.replace(/"/g, '""') + '"';
  };

  const unitHeaders = unitItems.map(function(x) { return escape(x.shortcut || x.name); });
  const dollarHeaders = dollarItems.map(function(x) { return escape(x.shortcut || x.name); });
  const header = ['Agent', 'Branch', 'Submit Date'].concat(unitHeaders).concat(dollarHeaders).concat(['Total Revenue', 'Remark']).map(escape).join(',');

  const lines = rows.map(function(s) {
    const submitDate = dateOf(s.submittedAt) || dateOf(s.date);
    const unitCols = unitItems.map(function(item) { return escape(s.items && s.items[item.id] ? s.items[item.id] : ''); });
    const dollarCols = dollarItems.map(function(item) { return escape(s.dollarItems && s.dollarItems[item.id] ? s.dollarItems[item.id] : ''); });
    const rev = s.dollarItems && s.dollarItems[ITEM_ID_REVENUE] ? s.dollarItems[ITEM_ID_REVENUE] : '';
    return [escape(s.agent), escape(s.branch), escape(submitDate)]
      .concat(unitCols).concat(dollarCols)
      .concat([escape(rev), escape(s.note || s.remark || '')])
      .join(',');
  });

  const csv = header + '\n' + lines.join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const suffix = (dateFrom || 'earliest') + '_to_' + (dateTo || 'latest');
  a.href = url;
  a.download = 'daily_sale_' + suffix + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  closeModal('modal-saleDownload');
  showToast('Download complete.', 'success');
}

function renderSummaryView(data, unitItems, dollarItems) {
  const container = g('sale-summary-view');
  if (!container) return;

  if (!data.length) {
    container.innerHTML = '<div style="text-align:center;padding:40px;color:#999;"><i class="fas fa-inbox fa-3x" style="display:block;margin-bottom:12px;"></i>No records found</div>';
    return;
  }

  const agentMap = {};
  data.forEach(function(s) {
    if (!agentMap[s.agent]) agentMap[s.agent] = { units: {}, dollars: {}, totalUnits: 0, totalRev: 0 };
    const ag = agentMap[s.agent];
    Object.keys(s.items || {}).forEach(function(iid) {
      ag.units[iid] = (ag.units[iid] || 0) + s.items[iid];
      ag.totalUnits += s.items[iid];
    });
    Object.keys(s.dollarItems || {}).forEach(function(iid) {
      ag.dollars[iid] = (ag.dollars[iid] || 0) + s.dollarItems[iid];
      if (iid === ITEM_ID_REVENUE) ag.totalRev += s.dollarItems[iid];
    });
  });

  const cards = Object.keys(agentMap).map(function(agent, idx) {
    const ag = agentMap[agent];
    const unitRows = unitItems.map(function(item) {
      const qty = ag.units[item.id] || 0;
      return qty ? '<div style="display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px solid #f0f0f0;font-size:0.8125rem;"><span>' + esc(item.name) + '</span><span style="font-weight:600;color:#1B7D3D;">' + qty + '</span></div>' : '';
    }).join('');
    const dollarRows = dollarItems.map(function(item) {
      const amt = ag.dollars[item.id] || 0;
      return amt ? '<div style="display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px solid #f0f0f0;font-size:0.8125rem;"><span>' + esc(item.name) + '</span><span style="font-weight:600;color:#E65100;">' + fmtMoney(amt, esc(item.currency) + ' ') + '</span></div>' : '';
    }).join('');
    return '<div class="summary-card">' +
      '<div class="summary-card-header">' +
        '<span class="sc-avatar av-' + (idx % 8) + '">' + esc(ini(agent)) + '</span>' +
        '<div><div class="sc-name">' + esc(agent) + '</div><div style="font-size:0.72rem;opacity:0.8;">Rev: ' + fmtMoney(ag.totalRev) + '</div></div>' +
      '</div>' +
      '<div class="summary-card-body">' + (unitRows + dollarRows || '<div style="color:#999;font-size:0.8rem;">No sales</div>') + '</div>' +
      '</div>';
  }).join('');

  const chartRow = '<div class="chart-row" style="margin-top:20px;">' +
    '<div class="chart-card"><div class="chart-card-header"><span class="chart-card-title">Sales by Item</span></div><div class="chart-card-body"><canvas id="cSaleMix"></canvas></div></div>' +
    '<div class="chart-card"><div class="chart-card-header"><span class="chart-card-title">Sales by Agent</span></div><div class="chart-card-body"><canvas id="cSaleAgent"></canvas></div></div>' +
    '</div>';

  container.innerHTML = '<div class="summary-grid">' + cards + '</div>' + chartRow;
  setTimeout(renderSaleCharts, 50);
}

function renderSaleCharts() {
  _cSaleMix = destroyChart(_cSaleMix);
  _cSaleAgent = destroyChart(_cSaleAgent);

  const data = filteredSales.length ? filteredSales : saleRecords;
  const unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });

  const mixLabels = unitItems.map(function(x) { return x.name; });
  const mixData = unitItems.map(function(item) {
    let t = 0; data.forEach(function(s) { t += (s.items && s.items[item.id]) ? s.items[item.id] : 0; }); return t;
  });

  const mixCtx = g('cSaleMix');
  if (mixCtx && typeof Chart !== 'undefined' && mixData.some(function(v) { return v > 0; })) {
    _cSaleMix = new Chart(mixCtx, {
      type: 'doughnut',
      data: { labels: mixLabels, datasets: [{ data: mixData, backgroundColor: CHART_PAL, borderWidth: 2, borderColor: '#fff', hoverOffset: 6 }] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '65%',
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, font: { size: 11 } } },
          tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 } }
        }
      }
    });
  }

  const agentMap = {};
  data.forEach(function(s) {
    if (!agentMap[s.agent]) agentMap[s.agent] = { rev: 0, branch: s.branch || '' };
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) {
      agentMap[s.agent].rev += s.dollarItems[ITEM_ID_REVENUE];
    }
  });

  // Group agents by branch for color coding
  const branchList = [];
  Object.keys(agentMap).forEach(function(a) {
    const br = agentMap[a].branch;
    if (br && branchList.indexOf(br) === -1) branchList.push(br);
  });

  const agentLabels = Object.keys(agentMap);
  const agentVals = agentLabels.map(function(a) { return agentMap[a].rev; });
  const agentColors = agentLabels.map(function(a) {
    const bIdx = branchList.indexOf(agentMap[a].branch);
    return CHART_PAL[bIdx >= 0 ? bIdx % CHART_PAL.length : 0];
  });
  const agentBranchLabels = agentLabels.map(function(a) {
    return agentMap[a].branch ? a + ' (' + agentMap[a].branch + ')' : a;
  });

  const agCtx = g('cSaleAgent');
  if (agCtx && typeof Chart !== 'undefined' && agentLabels.length) {
    _cSaleAgent = new Chart(agCtx, {
      type: 'bar',
      data: {
        labels: agentBranchLabels,
        datasets: [{
          label: 'Revenue ($)',
          data: agentVals,
          backgroundColor: agentColors,
          borderColor: agentColors,
          borderWidth: 1,
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 },
            callbacks: {
              label: function(ctx) { return 'Revenue: $' + Number(ctx.parsed.y).toFixed(2); }
            }
          }
        },
        scales: {
          x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 } } },
          y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, callback: function(v) { return '$' + v; } } }
        }
      }
    });
  }
}

// ------------------------------------------------------------
// Dashboard
// ------------------------------------------------------------
function renderDashboard() {
  const ym = ymNow();
  const ymP = ymPrev();

  // Role-based data filtering
  let viewSales = saleRecords;
  if (currentRole === 'agent' && currentUser) {
    viewSales = saleRecords.filter(function(s) { return s.agent === currentUser.name; });
  } else if (currentRole === 'supervisor' && currentUser) {
    viewSales = saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
  }
  // Apply branch filter (available for admin/cluster)
  const branchFilterVal = g('dash-branch-filter') ? g('dash-branch-filter').value : '';
  if (branchFilterVal) {
    viewSales = viewSales.filter(function(s) { return s.branch === branchFilterVal; });
  }
  // Hide branch filter for agent and supervisor (they auto-display their own branch)
  var branchFilterWrap = g('dash-branch-filter-wrap');
  if (branchFilterWrap) branchFilterWrap.style.display = (currentRole === 'agent' || currentRole === 'supervisor') ? 'none' : '';

  // Show branch summary table for admin, cluster, and supervisor (supervisor sees only their own branch via filtered viewSales)
  var branchSection = g('dash-branch-section');
  if (branchSection) branchSection.style.display = (currentRole === 'admin' || currentRole === 'cluster' || currentRole === 'supervisor') ? '' : 'none';

  const currSales = viewSales.filter(function(s) { return ymOf(s.date) === ym; });
  const prevSales = viewSales.filter(function(s) { return ymOf(s.date) === ymP; });

  let currRevenue = 0, prevRevenue = 0;
  let currRecharge = 0, prevRecharge = 0;
  let currGrossAds = 0, prevGrossAds = 0;
  let currHomeInternet = 0, prevHomeInternet = 0;

  currSales.forEach(function(s) {
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) currRevenue += s.dollarItems[ITEM_ID_REVENUE];
    if (s.dollarItems && s.dollarItems[ITEM_ID_RECHARGE]) currRecharge += s.dollarItems[ITEM_ID_RECHARGE];
    if (s.items && s.items[ITEM_ID_GROSS_ADS]) currGrossAds += s.items[ITEM_ID_GROSS_ADS];
    if (s.items && s.items[ITEM_ID_SMART_HOME]) currHomeInternet += s.items[ITEM_ID_SMART_HOME];
    if (s.items && s.items[ITEM_ID_SMART_FIBER]) currHomeInternet += s.items[ITEM_ID_SMART_FIBER];
  });

  prevSales.forEach(function(s) {
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) prevRevenue += s.dollarItems[ITEM_ID_REVENUE];
    if (s.dollarItems && s.dollarItems[ITEM_ID_RECHARGE]) prevRecharge += s.dollarItems[ITEM_ID_RECHARGE];
    if (s.items && s.items[ITEM_ID_GROSS_ADS]) prevGrossAds += s.items[ITEM_ID_GROSS_ADS];
    if (s.items && s.items[ITEM_ID_SMART_HOME]) prevHomeInternet += s.items[ITEM_ID_SMART_HOME];
    if (s.items && s.items[ITEM_ID_SMART_FIBER]) prevHomeInternet += s.items[ITEM_ID_SMART_FIBER];
  });

  const kr = g('kv-revenue'); if (kr) kr.textContent = fmtMoney(currRevenue);
  const krc = g('kv-recharge'); if (krc) krc.textContent = fmtMoney(currRecharge);
  const kg = g('kv-gross-ads'); if (kg) kg.textContent = currGrossAds;
  const kh = g('kv-home-internet'); if (kh) kh.textContent = currHomeInternet;

  setTrend('tr-revenue', currRevenue, prevRevenue);
  setTrend('tr-recharge', currRecharge, prevRecharge);
  setTrend('tr-gross-ads', currGrossAds, prevGrossAds);
  setTrend('tr-home-internet', currHomeInternet, prevHomeInternet);

  // Chart 1: Monthly Trend / Shop KPI Achievement Gauge
  _cTrend = destroyChart(_cTrend);
  clearCanvas('cTrend');
  var trendGaugeOverlay = g('trend-gauge-overlay');

  // For supervisor/agent: show KPI achievement gauge for their own shop
  var shopKpiForGauge = null;
  if ((currentRole === 'supervisor' || currentRole === 'agent') && currentUser) {
    var shopSup = (currentRole === 'supervisor') ? currentUser :
      staffList.find(function(u) { return u.role === 'Supervisor' && u.branch === currentUser.branch; });
    if (shopSup) {
      shopKpiForGauge = kpiList.find(function(k) { return k.kpiFor === 'shop' && k.assigneeId === shopSup.id; });
    }
  }

  const tCtx = g('cTrend');
  if (tCtx && shopKpiForGauge && typeof Chart !== 'undefined') {
    var kpiActualVal = calcKpiActual(shopKpiForGauge);
    var kpiPct = shopKpiForGauge.target > 0 ? Math.round(kpiActualVal / shopKpiForGauge.target * 100) : 0;
    var fillPct = Math.min(kpiPct, 100);
    var gColor = kpiPct >= 100 ? '#1B7D3D' : kpiPct >= 70 ? '#FF9800' : '#E53935';
    _cTrend = new Chart(tCtx, {
      type: 'doughnut',
      data: {
        datasets: [{
          data: [fillPct, 100 - fillPct],
          backgroundColor: [gColor, '#EEEEEE'],
          borderWidth: 0,
          circumference: 180,
          rotation: -90
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '70%',
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        animation: { animateRotate: true, duration: 700 }
      }
    });
    if (trendGaugeOverlay) {
      trendGaugeOverlay.style.display = '';
      var pctEl = g('trend-gauge-pct-text');
      var detailEl = g('trend-gauge-detail');
      if (pctEl) { pctEl.style.color = gColor; pctEl.textContent = kpiPct + '%'; }
      if (detailEl) {
        var tgValueDisplay = shopKpiForGauge.valueMode === 'currency'
          ? fmtMoney(kpiActualVal, shopKpiForGauge.currency + ' ') + ' / ' + fmtMoney(shopKpiForGauge.target, shopKpiForGauge.currency + ' ')
          : Math.round(kpiActualVal * 100) / 100 + ' / ' + shopKpiForGauge.target + (shopKpiForGauge.unit ? ' ' + esc(shopKpiForGauge.unit) : '');
        detailEl.textContent = esc(shopKpiForGauge.name) + ': ' + tgValueDisplay;
      }
    }
  } else {
    if (trendGaugeOverlay) trendGaugeOverlay.style.display = 'none';
    if (tCtx && typeof Chart !== 'undefined') {
      const months = last7Months();
      const monthLabels = months.map(ymLabel);
      const unitsPerMonth = months.map(function(m) {
        let u = 0;
        viewSales.filter(function(s) { return ymOf(s.date) === m; }).forEach(function(s) {
          Object.values(s.items || {}).forEach(function(v) { u += v; });
        });
        return u;
      });
      const revPerMonth = months.map(function(m) {
        let r = 0;
        viewSales.filter(function(s) { return ymOf(s.date) === m; }).forEach(function(s) {
          Object.keys(s.dollarItems || {}).forEach(function(iid) {
            const item = itemCatalogue.find(function(x) { return x.id === iid; });
            if (item && !item.noAutoSum && !item.noAutoRevenue) r += s.dollarItems[iid] * (item.price || 1);
          });
        });
        return r;
      });
      _cTrend = new Chart(tCtx, {
        type: 'line',
        data: {
          labels: monthLabels,
          datasets: [
            { label: 'Units', data: unitsPerMonth, borderColor: '#1B7D3D', backgroundColor: 'rgba(27,125,61,0.08)', yAxisID: 'y', tension: 0.4, fill: true, pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: '#1B7D3D', borderWidth: 2 },
            { label: 'Revenue ($)', data: revPerMonth, borderColor: '#FF9800', backgroundColor: 'rgba(255,152,0,0.08)', yAxisID: 'y1', tension: 0.4, fill: true, pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: '#FF9800', borderWidth: 2 }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, font: { size: 11 } } },
            tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, titleFont: { size: 12 }, bodyFont: { size: 11 } }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 11 } } },
            y: { position: 'left', title: { display: true, text: 'Units', font: { size: 11 } }, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 } } },
            y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Revenue ($)', font: { size: 11 } }, ticks: { font: { size: 11 } } }
          }
        }
      });
    }
  }

  // Chart 2: Item Mix (doughnut)
  _cMix = destroyChart(_cMix);
  clearCanvas('cMix');
  const unitItemsDash = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
  const mixData = unitItemsDash.map(function(item) {
    let total = 0;
    viewSales.forEach(function(s) { total += (s.items && s.items[item.id]) ? s.items[item.id] : 0; });
    return total;
  });
  const mCtx = g('cMix');
  if (mCtx && unitItemsDash.length && typeof Chart !== 'undefined') {
    _cMix = new Chart(mCtx, {
      type: 'doughnut',
      data: {
        labels: unitItemsDash.map(function(x) { return x.name; }),
        datasets: [{ data: mixData, backgroundColor: CHART_PAL, borderWidth: 2, borderColor: '#fff', hoverOffset: 6 }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '65%',
        plugins: {
          legend: { position: 'right', labels: { usePointStyle: true, pointStyle: 'circle', padding: 12, font: { size: 11 } } },
          tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 } }
        }
      }
    });
  }

  // Chart 3: Agent Performance (Revenue)
  _cAgent = destroyChart(_cAgent);
  clearCanvas('cAgent');
  const agentRevenue = {};
  currSales.forEach(function(s) {
    if (!(s.agent in agentRevenue)) agentRevenue[s.agent] = 0;
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) agentRevenue[s.agent] += s.dollarItems[ITEM_ID_REVENUE];
  });
  const agentNames = Object.keys(agentRevenue);
  const agentVals = agentNames.map(function(a) { return agentRevenue[a]; });
  const aCtx = g('cAgent');
  if (aCtx && agentNames.length && typeof Chart !== 'undefined') {
    _cAgent = new Chart(aCtx, {
      type: 'bar',
      data: {
        labels: agentNames,
        datasets: [{ label: 'Revenue This Month ($)', data: agentVals, backgroundColor: CHART_PAL, borderRadius: 4, borderSkipped: false }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 },
            callbacks: { label: function(ctx) { return ' ' + fmtMoney(ctx.parsed.x || 0); } }
          }
        },
        scales: {
          x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, callback: function(v) { return fmtMoney(v); } } },
          y: { grid: { display: false }, ticks: { font: { size: 11 } } }
        }
      }
    });
  }

  // Chart 4: Growth vs Last Month
  _cGrowth = destroyChart(_cGrowth);
  clearCanvas('cGrowth');
  // Sync dropdown state
  var growthViewSel = g('growth-chart-view');
  if (growthViewSel && growthViewSel.value !== _growthChartView) growthViewSel.value = _growthChartView;
  const growthLabels = unitItemsDash.map(function(x) { return x.shortcut || x.name; });
  const currItemUnits = unitItemsDash.map(function(item) {
    let t = 0; currSales.forEach(function(s) { t += (s.items && s.items[item.id]) || 0; }); return t;
  });
  const prevItemUnits = unitItemsDash.map(function(item) {
    let t = 0; prevSales.forEach(function(s) { t += (s.items && s.items[item.id]) || 0; }); return t;
  });
  const gCtx = g('cGrowth');
  if (gCtx && unitItemsDash.length && typeof Chart !== 'undefined') {
    if (_growthChartView === 'pct') {
      // Show percentage growth per item
      const pctData = currItemUnits.map(function(curr, i) {
        var prev = prevItemUnits[i];
        if (!prev) return null;
        return Math.round((curr - prev) / prev * 100);
      });
      const barColors = pctData.map(function(v) {
        if (v === null) return '#BDBDBD';
        return v >= 0 ? '#1B7D3D' : '#E53935';
      });
      _cGrowth = new Chart(gCtx, {
        type: 'bar',
        data: {
          labels: growthLabels,
          datasets: [{
            label: '% Growth vs Last Month',
            data: pctData,
            backgroundColor: barColors,
            borderRadius: 5,
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, font: { size: 11 } } },
            tooltip: {
              backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 },
              callbacks: {
                label: function(ctx) {
                  var v = ctx.parsed.y;
                  if (v === null) return 'N/A';
                  return (v >= 0 ? '+' : '') + v + '%';
                }
              }
            }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 11 } } },
            y: {
              grid: { color: 'rgba(0,0,0,0.05)' },
              ticks: {
                font: { size: 11 },
                callback: function(v) { return v + '%'; }
              }
            }
          }
        }
      });
    } else {
      _cGrowth = new Chart(gCtx, {
        type: 'line',
        data: {
          labels: growthLabels,
          datasets: [
            { label: 'This Month', data: currItemUnits, borderColor: '#1B7D3D', backgroundColor: 'rgba(27,125,61,0.08)', tension: 0.3, fill: false, pointRadius: 5, pointHoverRadius: 7, pointBackgroundColor: '#1B7D3D', pointBorderColor: '#fff', pointBorderWidth: 2, borderWidth: 2 },
            { label: 'Last Month', data: prevItemUnits, borderColor: '#A5D6A7', backgroundColor: 'rgba(165,214,167,0.08)', tension: 0.3, fill: false, pointRadius: 5, pointHoverRadius: 7, pointBackgroundColor: '#A5D6A7', pointBorderColor: '#fff', pointBorderWidth: 2, borderWidth: 2 }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, font: { size: 11 } } },
            tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 } }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 11 } } },
            y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 } }, beginAtZero: true }
          }
        }
      });
    }
  }

  // Branch summary table (visible for admin/cluster/supervisor)
  const branchTableBody = g('branch-table');
  if (branchTableBody) {
    const branches = [];
    viewSales.forEach(function(s) { if (branches.indexOf(s.branch) < 0) branches.push(s.branch); });
    if (!branches.length) {
      branchTableBody.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:20px;color:#999;">No data</td></tr>';
    } else {
      branchTableBody.innerHTML = branches.map(function(branch) {
        let cU = 0, pU = 0;
        currSales.filter(function(s) { return s.branch === branch; }).forEach(function(s) {
          Object.values(s.items || {}).forEach(function(v) { cU += v; });
        });
        prevSales.filter(function(s) { return s.branch === branch; }).forEach(function(s) {
          Object.values(s.items || {}).forEach(function(v) { pU += v; });
        });
        const pct = pctChange(cU, pU);
        const trendHtml = pct === null ? '<span class="pill pill-gray">N/A</span>' :
          pct > 0 ? '<span class="pill pill-green">+' + pct + '%</span>' :
          pct < 0 ? '<span class="pill pill-red">' + pct + '%</span>' : '<span class="pill pill-gray">0%</span>';
        return '<tr><td>' + esc(branch) + '</td><td>' + cU + '</td><td>' + pU + '</td><td>' + trendHtml + '</td></tr>';
      }).join('');
    }
  }

  // Branch filter dropdown
  const branchFilter = g('dash-branch-filter');
  if (branchFilter) {
    const branches = getBranches();
    const cur = branchFilter.value;
    branchFilter.innerHTML = '<option value="">All Branches</option>' +
      branches.map(function(b) { return '<option value="' + esc(b) + '"' + (cur === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  }

  // KPI Achievement section
  renderDashboardKpiSection();
}

// ------------------------------------------------------------
// KPI vs Actual Achievement Dashboard
// ------------------------------------------------------------
function getKpiUnitLabel(kpi) {
  if (kpi.kpiMode === 'point') return 'pts';
  if (kpi.itemId) {
    var item = itemCatalogue.find(function(x) { return x.id === kpi.itemId; });
    if (item) return item.name;
  }
  return kpi.unit || '';
}

function calcKpiActual(kpi, ym) {
  if (!ym) ym = ymNow();
  const currSales = saleRecords.filter(function(s) { return ymOf(s.date) === ym; });
  let filtered = currSales;
  if (kpi.kpiFor === 'agent') {
    const agent = staffList.find(function(u) { return u.id === kpi.assigneeId; });
    if (agent) filtered = currSales.filter(function(s) { return s.agent === agent.name; });
  } else if (kpi.kpiFor === 'shop') {
    const sup = staffList.find(function(u) { return u.id === kpi.assigneeId; });
    if (sup) filtered = currSales.filter(function(s) { return s.branch === sup.branch; });
  }
  let actual = 0;
  if (kpi.kpiMode === 'point') {
    // New-style point KPIs (with pointTiers) sum kpiPoints from sale records directly
    if (kpi.pointTiers) {
      filtered.forEach(function(s) {
        if (s.transactionType === 'point' && s.kpiPoints) {
          actual += parseFloat(s.kpiPoints) || 0;
        }
      });
    } else {
      // Legacy point KPIs using pointItems mapping
      var pointItems = kpi.pointItems || [];
      filtered.forEach(function(s) {
        pointItems.forEach(function(pi) {
          actual += (Number(s.items && s.items[pi.itemId]) || 0) * (Number(pi.points) || 0);
        });
      });
    }
    return actual;
  }
  if (kpi.valueMode === 'unit') {
    if (kpi.itemId) {
      filtered.forEach(function(s) { actual += (s.items && s.items[kpi.itemId]) || 0; });
    } else {
      filtered.forEach(function(s) { Object.values(s.items || {}).forEach(function(v) { actual += v; }); });
    }
  } else {
    if (kpi.itemId) {
      const kpiItem = itemCatalogue.find(function(x) { return x.id === kpi.itemId; });
      filtered.forEach(function(s) { actual += ((s.dollarItems && s.dollarItems[kpi.itemId]) || 0) * (kpiItem ? kpiItem.price || 1 : 1); });
    } else {
      filtered.forEach(function(s) {
        Object.keys(s.dollarItems || {}).forEach(function(iid) {
          const item = itemCatalogue.find(function(x) { return x.id === iid; });
          if (item && !item.noAutoSum && !item.noAutoRevenue) actual += s.dollarItems[iid] * (item.price || 1);
        });
      });
    }
  }
  return actual;
}

var _dashKpiView = 'all'; // 'all', 'shop', 'agent', 'branch'
var _growthChartView = 'compare'; // 'compare' or 'pct'

function setGrowthChartView(v) {
  _growthChartView = v;
  renderDashboard();
}

function setDashKpiView(view) {
  _dashKpiView = view;
  $$('.dash-kpi-view-btn').forEach(function(b) {
    b.classList.toggle('active', b.getAttribute('data-view') === view);
  });
  renderDashboardKpiSection();
}

function renderDashboardKpiSection() {
  var section = g('dash-kpi-section');
  if (!section) return;

  // Show/hide the "By Branch" button (only for cluster/admin)
  var branchBtn = g('dash-kpi-branch-btn');
  if (branchBtn) branchBtn.style.display = (currentRole === 'cluster' || currentRole === 'admin') ? '' : 'none';

  // Populate and show/hide branch dropdown for "By Branch" view
  var branchFilterWrap = g('dash-kpi-branch-filter-wrap');
  var branchFilterSel = g('dash-kpi-branch-filter');
  if (branchFilterWrap) branchFilterWrap.style.display = (_dashKpiView === 'branch' && (currentRole === 'cluster' || currentRole === 'admin')) ? '' : 'none';
  if (branchFilterSel && _dashKpiView === 'branch') {
    var allBranches = getBranches();
    var curBranch = branchFilterSel.value;
    branchFilterSel.innerHTML = '<option value="">All Branches</option>' +
      allBranches.map(function(b) { return '<option value="' + esc(b) + '"' + (curBranch === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  }

  // Determine relevant KPIs based on role
  var relevantKpis = [];
  if (currentRole === 'agent' && currentUser) {
    // Agents can view their own agent KPIs AND the shop KPIs for their branch
    var branchSups = staffList.filter(function(u) {
      return u.role === 'Supervisor' && u.branch === currentUser.branch;
    });
    var supIds = branchSups.map(function(s) { return s.id; });
    relevantKpis = kpiList.filter(function(k) {
      return (k.kpiFor === 'agent' && k.assigneeId === currentUser.id) ||
             (k.kpiFor === 'shop' && supIds.indexOf(k.assigneeId) !== -1);
    });
  } else if (currentRole === 'supervisor' && currentUser) {
    relevantKpis = kpiList.filter(function(k) {
      return (k.kpiFor === 'shop' && k.assigneeId === currentUser.id) ||
             (k.kpiFor === 'agent' && k.assigneeBranch === currentUser.branch);
    });
  } else if (currentRole === 'cluster' || currentRole === 'admin') {
    relevantKpis = kpiList.slice();
  }

  // Apply view filter
  if (_dashKpiView === 'shop') {
    relevantKpis = relevantKpis.filter(function(k) { return k.kpiFor === 'shop'; });
  } else if (_dashKpiView === 'agent') {
    relevantKpis = relevantKpis.filter(function(k) { return k.kpiFor === 'agent'; });
  } else if (_dashKpiView === 'branch' && branchFilterSel && branchFilterSel.value) {
    // Filter by branch: for shop KPIs use assignee's branch, for agent KPIs use assigneeBranch
    var selectedBranch = branchFilterSel.value;
    relevantKpis = relevantKpis.filter(function(k) {
      if (k.kpiFor === 'agent') return k.assigneeBranch === selectedBranch;
      if (k.kpiFor === 'shop') {
        var sup = staffList.find(function(u) { return u.id === k.assigneeId; });
        return sup ? sup.branch === selectedBranch : false;
      }
      return false;
    });
  }

  // Show/hide the view toggle buttons based on role
  var toggleWrap = g('dash-kpi-view-toggle');
  if (toggleWrap) {
    toggleWrap.style.display = (currentRole === 'agent' || currentRole === 'supervisor' || currentRole === 'cluster' || currentRole === 'admin') ? '' : 'none';
  }

  if (!relevantKpis.length) {
    section.style.display = 'none';
    return;
  }
  section.style.display = '';

  // Build KPI data with actuals and percentages
  var kpiData = relevantKpis.map(function(k) {
    var assignee = staffList.find(function(u) { return u.id === k.assigneeId; });
    var actual = calcKpiActual(k);
    var isPoint = k.kpiMode === 'point';
    var gateTarget = (isPoint && k.pointTiers) ? (k.pointTiers.gate || 0) : k.target;
    var pct = gateTarget > 0 ? Math.round(actual / gateTarget * 100) : 0;
    return { k: k, assignee: assignee, actual: Math.round(actual * 100) / 100, pct: pct };
  });

  // Destroy existing gauge charts
  _cKpiGauges.forEach(function(c) { if (c) { try { c.destroy(); } catch (e) { console.warn('Chart destroy error:', e); } } });
  _cKpiGauges = [];

  // Color helper for gauge achievement level
  function gaugeColor(pct) { return pct >= 100 ? '#1B7D3D' : pct >= 70 ? '#FF9800' : '#E53935'; }
  function gaugePillClass(pct) { return pct >= 100 ? 'pill-green' : pct >= 70 ? 'pill-orange' : 'pill-red'; }

  // Render gauge cards
  var gaugeGrid = g('dash-kpi-gauge-grid');
  if (gaugeGrid) {
    gaugeGrid.innerHTML = kpiData.map(function(d, i) {
      var color = gaugeColor(d.pct);
      var pctClass = gaugePillClass(d.pct);
      var forPill = d.k.kpiFor === 'shop'
        ? '<span class="pill pill-blue" style="font-size:.7rem;padding:2px 7px;">Shop</span>'
        : '<span class="pill pill-orange" style="font-size:.7rem;padding:2px 7px;">Agent</span>';
      var assigneeName = d.assignee ? esc(d.assignee.name) : (d.k.assigneeBranch ? esc(d.k.assigneeBranch) : '—');
      var isPointGauge = d.k.kpiMode === 'point';
      var valueDisplay = (isPointGauge && d.k.pointTiers)
        ? d.actual + ' / ' + (d.k.pointTiers.gate || 0) + ' pts'
        : d.k.valueMode === 'currency'
          ? fmtMoney(d.k.target, esc(d.k.currency) + ' ') + ' / ' + fmtMoney(d.actual, esc(d.k.currency) + ' ')
          : (function() { var ul = getKpiUnitLabel(d.k); return d.k.target + ' / ' + d.actual + (ul ? ' ' + esc(ul) : ''); })();
      return '<div class="kpi-gauge-card">' +
        '<div class="kpi-gauge-canvas-wrap">' +
        '<canvas id="kpiGauge_' + i + '" height="130"></canvas>' +
        '<div class="kpi-gauge-pct-text" style="color:' + color + '">' + d.pct + '%</div>' +
        '</div>' +
        '<div class="kpi-gauge-name">' + esc(d.k.name) + '</div>' +
        '<div class="kpi-gauge-assignee">' + forPill + ' <span>' + assigneeName + '</span></div>' +
        '<div class="kpi-gauge-value"><span class="pill ' + pctClass + '" style="font-size:.72rem;">' + valueDisplay + '</span></div>' +
        '</div>';
    }).join('');

    // Create Chart.js gauge (semicircle doughnut) for each KPI
    if (typeof Chart !== 'undefined') {
      kpiData.forEach(function(d, i) {
        var canvas = document.getElementById('kpiGauge_' + i);
        if (!canvas) return;
        var color = gaugeColor(d.pct);
        var fillPct = Math.min(d.pct, 100);
        var chart = new Chart(canvas, {
          type: 'doughnut',
          data: {
            datasets: [{
              data: [fillPct, 100 - fillPct],
              backgroundColor: [color, '#EEEEEE'],
              borderWidth: 0,
              circumference: 180,
              rotation: -90
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '72%',
            plugins: {
              legend: { display: false },
              tooltip: { enabled: false }
            },
            animation: { animateRotate: true, duration: 700 }
          }
        });
        _cKpiGauges.push(chart);
      });
    }
  }

  // Render summary table
  var tableWrap = g('dash-kpi-table-wrap');
  if (tableWrap) {
    var rows = kpiData.map(function(d) {
      var pct = d.pct;
      var pctClass = gaugePillClass(pct);
      var isPoint = d.k.kpiMode === 'point';
      var modePill = isPoint
        ? '<span class="pill pill-purple" style="font-size:.7rem;"><i class="fas fa-star"></i> Point</span>'
        : '<span class="pill pill-teal" style="font-size:.7rem;"><i class="fas fa-chart-bar"></i> Volume</span>';
      var forLabel = d.k.kpiFor === 'shop' ? '<span class="pill pill-blue">Shop</span>' : '<span class="pill pill-orange">Agent</span>';
      var assigneeName = d.assignee ? esc(d.assignee.name) : (d.k.assigneeBranch ? esc(d.k.assigneeBranch) : '—');
      var valueDisplay = (isPoint && d.k.pointTiers)
        ? d.actual + ' / ' + (d.k.pointTiers.gate || 0) + ' pts'
        : (!isPoint && d.k.valueMode === 'currency')
          ? fmtMoney(d.k.target, esc(d.k.currency) + ' ') + ' / ' + fmtMoney(d.actual, esc(d.k.currency) + ' ')
          : (function() { var ul = getKpiUnitLabel(d.k); return d.k.target + ' / ' + d.actual + (ul ? ' ' + esc(ul) : ''); })();
      return '<tr>' +
        '<td>' + esc(d.k.name) + '</td>' +
        '<td>' + modePill + '</td>' +
        '<td>' + forLabel + ' <small style="color:#888;">' + assigneeName + '</small></td>' +
        '<td>' + valueDisplay + '</td>' +
        '<td><span class="pill ' + pctClass + '">' + pct + '%</span></td>' +
        '</tr>';
    }).join('');
    tableWrap.innerHTML = '<table class="data-table" style="margin-top:8px;"><thead><tr><th>KPI</th><th>Mode</th><th>Assignee</th><th>Target / Actual</th><th>Achievement</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }
}
// ------------------------------------------------------------
function openCustomerModal(type, item) {
  if (type === 'new-customer') {
    const form = g('form-newCustomer');
    if (form) form.reset();
    g('nc-edit-id').value = '';
    const title = g('modal-newCustomer-title');
    populateBranchSelects();
    if (item) {
      if (title) title.textContent = 'Edit New Customer';
      g('nc-edit-id').value = item.id;
      g('nc-name').value = item.name || '';
      g('nc-phone').value = item.phone || '';
      g('nc-id').value = item.idNum || '';
      const tariffSel = g('nc-tariff'); if (tariffSel) tariffSel.value = item.tariff || item.pkg || '';
      g('nc-agent').value = item.agent || '';
      const bSel = g('nc-branch'); if (bSel) bSel.value = item.branch || '';
      g('nc-date').value = item.date || '';
      const statusSel = g('nc-status'); if (statusSel) statusSel.value = item.status || 'follow';
      if (g('nc-lat')) g('nc-lat').value = item.lat || '';
      if (g('nc-lng')) g('nc-lng').value = item.lng || '';
    } else {
      if (title) title.textContent = 'Add New Customer';
      g('nc-date').value = todayStr();
      if (currentUser) {
        const agEl = g('nc-agent'); if (agEl) { agEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true; }
        const brEl = g('nc-branch'); if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
      }
    }
    openModal('modal-newCustomer');
    // Initialize Leaflet map after modal animation completes (150ms matches CSS transition)
    setTimeout(function() {
      var latVal = parseFloat(g('nc-lat') ? g('nc-lat').value : '') || 0;
      var lngVal = parseFloat(g('nc-lng') ? g('nc-lng').value : '') || 0;
      // Default center: Phnom Penh, Cambodia
      var defaultCenter = (latVal && lngVal) ? [latVal, lngVal] : [11.5564, 104.9282];
      if (window._ncMap) { window._ncMap.remove(); window._ncMap = null; }
      var map = L.map('nc-map', {
        fullscreenControl: true,
        fullscreenControlOptions: { position: 'topleft' }
      }).setView(defaultCenter, latVal && lngVal ? 15 : 12);
      var streetLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19,
        attribution: '© OpenStreetMap contributors'
      });
      var satelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
        maxZoom: 19,
        attribution: 'Tiles &copy; Esri'
      });
      satelliteLayer.addTo(map);
      L.control.layers({ 'Satellite': satelliteLayer, 'Street Map': streetLayer }, null, { position: 'topright' }).addTo(map);
      var marker = null;
      if (latVal && lngVal) {
        marker = L.marker([latVal, lngVal]).addTo(map);
      }
      map.on('click', function(e) {
        var lat = e.latlng.lat.toFixed(6); // 6 decimal places ≈ 0.11 m precision
        var lng = e.latlng.lng.toFixed(6);
        if (g('nc-lat')) g('nc-lat').value = lat;
        if (g('nc-lng')) g('nc-lng').value = lng;
        if (marker) { map.removeLayer(marker); }
        marker = L.marker([e.latlng.lat, e.latlng.lng]).addTo(map);
      });
      window._ncMap = map;
    }, 150);

  } else if (type === 'topup') {
    const form = g('form-topUp');
    if (form) form.reset();
    g('tu-edit-id').value = '';
    const title = g('modal-topUp-title');
    populateBranchSelects();
    if (item) {
      if (title) title.textContent = 'Edit Top Up';
      g('tu-edit-id').value = item.id;
      g('tu-name').value = item.name || '';
      g('tu-phone').value = item.phone || '';
      g('tu-amount').value = item.amount || '';
      g('tu-agent').value = item.agent || '';
      const bSel = g('tu-branch'); if (bSel) bSel.value = item.branch || '';
      g('tu-date').value = item.date || '';
      const endDateEl = g('tu-end-date'); if (endDateEl) endDateEl.value = item.endDate || '';
      const tuStatusSel = g('tu-status'); if (tuStatusSel) tuStatusSel.value = item.tuStatus || 'active';
      const tuTariffEl = g('tu-tariff'); if (tuTariffEl) tuTariffEl.value = item.tariff || '';
      const tuRemarkEl = g('tu-remark'); if (tuRemarkEl) tuRemarkEl.value = item.remark || '';
      const tuLatEl = g('tu-lat'); if (tuLatEl) tuLatEl.value = item.lat || '';
      const tuLngEl = g('tu-lng'); if (tuLngEl) tuLngEl.value = item.lng || '';
      toggleTuLatlongRow(item.tuStatus || 'active', item.lat, item.lng);
    } else {
      if (title) title.textContent = 'Add Top Up';
      g('tu-date').value = todayStr();
      toggleTuLatlongRow('active');
      if (currentUser) {
        const agEl = g('tu-agent'); if (agEl) { agEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true; }
        const brEl = g('tu-branch'); if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
      }
    }
    openModal('modal-topUp');

  } else if (type === 'termination') {
    const form = g('form-termination');
    if (form) form.reset();
    g('term-edit-id').value = '';
    const title = g('modal-termination-title');
    populateBranchSelects();
    if (item) {
      if (title) title.textContent = 'Edit Termination';
      g('term-edit-id').value = item.id;
      g('term-name').value = item.name || '';
      g('term-phone').value = item.phone || '';
      g('term-reason').value = item.reason || '';
      g('term-agent').value = item.agent || '';
      const bSel = g('term-branch'); if (bSel) bSel.value = item.branch || '';
      g('term-date').value = item.date || '';
      const termLatEl = g('term-lat'); if (termLatEl) termLatEl.value = item.lat || '';
      const termLngEl = g('term-lng'); if (termLngEl) termLngEl.value = item.lng || '';
    } else {
      if (title) title.textContent = 'Add Termination';
      g('term-date').value = todayStr();
      if (currentUser) {
        const agEl = g('term-agent'); if (agEl) { agEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true; }
        const brEl = g('term-branch'); if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
      }
    }
    openModal('modal-termination');
  }
}

function submitNewCustomer(e) {
  e.preventDefault();
  const editId = rv('nc-edit-id');
  const tariffEl = g('nc-tariff');
  const statusEl = g('nc-status');
  const obj = {
    id: editId || uid(),
    name: rv('nc-name'), phone: rv('nc-phone'), idNum: rv('nc-id'),
    tariff: tariffEl ? tariffEl.value : '', pkg: tariffEl ? tariffEl.value : '',
    agent: rv('nc-agent'), branch: rv('nc-branch'), date: rv('nc-date'),
    status: statusEl ? statusEl.value : 'follow',
    lat: rv('nc-lat') || '', lng: rv('nc-lng') || ''
  };
  if (!obj.name) { showAlert('Please enter customer name'); return; }
  if (!obj.phone) { showAlert('Please enter phone number'); return; }
  if (!/^\d{6,15}$/.test(obj.phone.replace(/[\s\-+()]/g, ''))) { showAlert('Please enter a valid phone number (6–15 digits, separators allowed)'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  const prevStatus = editId ? (newCustomers.find(function(x){return x.id===editId;})||{}).status : null;
  if (editId) {
    const idx = newCustomers.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) newCustomers[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a customer record.');
  } else {
    newCustomers.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' added a new customer.');
  }
  // Auto-add to TopUp when status is Close and remove from New Customer list
  if (obj.status === 'close' && prevStatus !== 'close') {
    var existingTopUpRecord = topUpList.find(function(t) { return t.customerId === obj.id; });
    if (!existingTopUpRecord) {
      topUpList.push({
        id: uid(), customerId: obj.id, name: obj.name, phone: obj.phone,
        tariff: obj.tariff, agent: obj.agent, branch: obj.branch, date: obj.date,
        tuStatus: 'active', amount: 0, note: 'Auto-added (status: Close)'
      });
      syncSheet('TopUp', topUpList);
    }
    // Remove from new customer list and navigate to Top Up tab
    newCustomers = newCustomers.filter(function(x) { return x.id !== obj.id; });
    closeModal('modal-newCustomer');
    renderNewCustomerTable();
    renderTopUpTable();
    syncSheet('Customers', newCustomers);
    saveAllData();
    navigateTo('customer', null);
    switchCustomerTab('topup');
    // Update tab button active state
    $$('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    var topupBtn = g('tab-topup'); if (topupBtn) topupBtn.classList.add('active');
    showToast('Customer marked as Closed and moved to Top Up.', 'info');
    return;
  }
  // Move to Out Coverage List when status is Out Coverage
  if (obj.status === 'out-coverage' && prevStatus !== 'out-coverage') {
    var existingOutCovRecord = outCoverageList.find(function(t) { return t.customerId === obj.id; });
    if (!existingOutCovRecord) {
      outCoverageList.push({
        id: uid(), customerId: obj.id, name: obj.name, phone: obj.phone,
        idNum: obj.idNum, tariff: obj.tariff, agent: obj.agent, branch: obj.branch,
        date: obj.date, lat: obj.lat, lng: obj.lng, note: 'Out of coverage area'
      });
      syncSheet('OutCoverage', outCoverageList);
    }
    // Remove from new customer list and navigate to Out Coverage tab
    newCustomers = newCustomers.filter(function(x) { return x.id !== obj.id; });
    closeModal('modal-newCustomer');
    renderNewCustomerTable();
    renderOutCoverageTable();
    syncSheet('Customers', newCustomers);
    saveAllData();
    navigateTo('customer', null);
    switchCustomerTab('out-coverage');
    $$('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    var ocBtn = g('tab-out-coverage'); if (ocBtn) ocBtn.classList.add('active');
    showToast('Customer marked as Out Coverage and moved to Out Coverage List.', 'info');
    return;
  }
  closeModal('modal-newCustomer');
  renderNewCustomerTable();
  renderTopUpTable();
  syncSheet('Customers', newCustomers);
  saveAllData();
  showToast(editId ? 'Customer record updated.' : 'Customer added successfully.', 'success');
}

function editNewCustomer(id) {
  const item = newCustomers.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  openCustomerModal('new-customer', item);
}

function deleteNewCustomer(id) {
  const item = newCustomers.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Are you sure you want to delete this customer record? This action cannot be undone.', function() {
    newCustomers = newCustomers.filter(function(x) { return x.id !== id; });
    renderNewCustomerTable();
    syncSheet('Customers', newCustomers);
    saveAllData();
    showToast('Customer record deleted.', 'success');
  }, 'Delete Customer', 'Delete');
}

function renderNewCustomerTable(explicit) {
  const tbody = g('new-customer-table');
  if (!tbody) return;
  const searchVal = (rv('nc-search') || '').toLowerCase().trim();
  const baseList = getBaseRecordsForRole(newCustomers);
  const list = searchVal
    ? baseList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(searchVal) ||
               (c.phone || '').toLowerCase().includes(searchVal);
      })
    : baseList;
  if (!list.length) {
    if (explicit && searchVal) { showAlert('No results found for "' + searchVal + '"', 'warning'); }
    tbody.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-users" style="font-size:2rem;display:block;margin-bottom:8px;"></i>' + (searchVal ? 'No results found' : 'No customers yet') + '</td></tr>';
    return;
  }
  const statusPillMap = { follow: 'pill-gray', lead: 'pill-blue', 'hot-prospect': 'pill-orange', close: 'pill-green', 'out-coverage': 'pill-red' };
  const statusLabelMap = { follow: 'Follow', lead: 'Lead', 'hot-prospect': 'Hot Prospect', close: 'Close', 'out-coverage': 'Out Coverage' };
  tbody.innerHTML = list.map(function(c, i) {
    const avIdx = i % 8;
    const st = c.status || 'follow';
    const stPill = statusPillMap[st] || 'pill-gray';
    const stLabel = statusLabelMap[st] || esc(st);
    const canEdit = canModifyRecord(c);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + esc(c.idNum || '') + '</td>' +
      '<td>' + esc(c.tariff || c.pkg || '') + '</td>' +
      '<td><span class="pill ' + stPill + '">' + stLabel + '</span></td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      '<td style="white-space:nowrap;">' + (c.lat && c.lng
        ? '<a href="https://www.openstreetmap.org/?mlat=' + esc(c.lat) + '&mlon=' + esc(c.lng) + '&zoom=15" target="_blank" title="' + esc(c.lat) + ', ' + esc(c.lng) + '" style="color:#1B7D3D;text-decoration:none;"><i class="fas fa-map-marker-alt"></i> ' + esc(c.lat) + ', ' + esc(c.lng) + '</a>'
        : '<span style="color:#ccc;font-size:0.8rem;">—</span>') + '</td>' +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-edit" onclick="editNewCustomer(\'' + esc(c.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteNewCustomer(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function toggleTuLatlongRow(status, existingLat, existingLng) {
  const row = g('tu-latlong-row');
  if (!row) return;
  if (status === 'terminate') {
    row.style.display = '';
    setTimeout(function() {
      if (typeof L === 'undefined') return;
      if (_tuPickerMap) { _tuPickerMap.remove(); _tuPickerMap = null; _tuPickerMarker = null; }
      var pickerEl = g('tu-picker-map');
      if (!pickerEl) return;
      var lat0 = parseFloat(existingLat) || null;
      var lng0 = parseFloat(existingLng) || null;
      var center = (lat0 && lng0) ? [lat0, lng0] : KH_CENTER;
      var zoom = (lat0 && lng0) ? 14 : 7;
      _tuPickerMap = L.map('tu-picker-map').setView(center, zoom);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 18,
        attribution: '© OpenStreetMap contributors'
      }).addTo(_tuPickerMap);
      if (lat0 && lng0) {
        _tuPickerMarker = L.marker([lat0, lng0]).addTo(_tuPickerMap);
      }
      _tuPickerMap.on('click', function(e) {
        var lat = parseFloat(e.latlng.lat.toFixed(6));
        var lng = parseFloat(e.latlng.lng.toFixed(6));
        var latEl = g('tu-lat'); var lngEl = g('tu-lng');
        if (latEl) latEl.value = lat;
        if (lngEl) lngEl.value = lng;
        if (_tuPickerMarker) _tuPickerMap.removeLayer(_tuPickerMarker);
        _tuPickerMarker = L.marker([lat, lng]).addTo(_tuPickerMap);
      });
    }, 200);
  } else {
    row.style.display = 'none';
    if (_tuPickerMap) { _tuPickerMap.remove(); _tuPickerMap = null; _tuPickerMarker = null; }
    var latEl = g('tu-lat'); var lngEl = g('tu-lng');
    if (latEl) latEl.value = '';
    if (lngEl) lngEl.value = '';
  }
}

function submitTopUp(e) {
  e.preventDefault();
  const editId = rv('tu-edit-id');
  const tuStatusSel = g('tu-status');
  const tuStatus = tuStatusSel ? tuStatusSel.value : 'active';
  const existingRecord = editId ? topUpList.find(function(x) { return x.id === editId; }) : null;
  const obj = {
    id: editId || uid(),
    customerId: existingRecord ? (existingRecord.customerId || '') : '',
    name: rv('tu-name'), phone: rv('tu-phone'), amount: parseFloat(rv('tu-amount')) || 0,
    agent: rv('tu-agent'), branch: rv('tu-branch'), date: rv('tu-date'),
    endDate: rv('tu-end-date') || '',
    tariff: rv('tu-tariff') || '',
    remark: rv('tu-remark') || '',
    tuStatus: tuStatus,
    lat: rv('tu-lat') || '', lng: rv('tu-lng') || ''
  };
  if (!obj.name) { showAlert('Please enter customer name'); return; }
  if (!obj.phone) { showAlert('Please enter phone number'); return; }
  if (!/^\d{6,15}$/.test(obj.phone.replace(/[\s\-+()]/g, ''))) { showAlert('Please enter a valid phone number (6–15 digits, separators allowed)'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  const prevStatus = existingRecord ? existingRecord.tuStatus : null;
  if (editId) {
    const idx = topUpList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) topUpList[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a top-up record.');
  } else {
    topUpList.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' submitted a top-up.');
  }
  // Auto-add to termination when status = Terminate and remove from Top Up list
  if (tuStatus === 'terminate' && prevStatus !== 'terminate') {
    const existingTerminationRecord = terminationList.find(function(t) { return (obj.customerId && t.customerId === obj.customerId) || (t.name === obj.name && t.phone === obj.phone); });
    if (!existingTerminationRecord) {
      terminationList.push({
        id: uid(), customerId: obj.customerId || '', name: obj.name, phone: obj.phone, reason: 'Service terminated',
        agent: obj.agent, branch: obj.branch, date: obj.date, lat: obj.lat || '', lng: obj.lng || ''
      });
      syncSheet('Terminations', terminationList);
      renderTerminationTable();
    }
    // Remove from Top Up list and navigate to Termination tab
    topUpList = topUpList.filter(function(x) { return x.id !== obj.id; });
    closeModal('modal-topUp');
    renderTopUpTable();
    syncSheet('TopUp', topUpList);
    saveAllData();
    navigateTo('customer', null);
    switchCustomerTab('termination');
    // Update tab button active state
    $$('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    var termBtn = g('tab-termination'); if (termBtn) termBtn.classList.add('active');
    showToast('Customer terminated and moved to Termination list.', 'info');
    return;
  }
  closeModal('modal-topUp');
  renderTopUpTable();
  syncSheet('TopUp', topUpList);
  saveAllData();
  showToast(editId ? 'Top-up record updated.' : 'Top-up submitted successfully.', 'success');
}

function onTuExistingCustomerChange() {
  const sel = g('tu-existing-customer');
  if (!sel || !sel.value) return;
  const customer = newCustomers.find(function(c) { return c.id === sel.value; });
  if (!customer) return;
  const nameEl = g('tu-name'); if (nameEl) nameEl.value = customer.name || '';
  const phoneEl = g('tu-phone'); if (phoneEl) phoneEl.value = customer.phone || '';
  const agEl = g('tu-agent'); if (agEl && !agEl.value) agEl.value = customer.agent || '';
  const brEl = g('tu-branch'); if (brEl && !brEl.value) brEl.value = customer.branch || '';
}

function editTopUp(id) {
  const item = topUpList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  openCustomerModal('topup', item);
}

function deleteTopUp(id) {
  const item = topUpList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Are you sure you want to delete this top-up record? This action cannot be undone.', function() {
    topUpList = topUpList.filter(function(x) { return x.id !== id; });
    renderTopUpTable();
    syncSheet('TopUp', topUpList);
    saveAllData();
    showToast('Top-up record deleted.', 'success');
  }, 'Delete Top-Up Record', 'Delete');
}

function renderTopUpTable(explicit) {
  const tbody = g('topup-table');
  if (!tbody) return;

  // Build today / +7-day boundaries using local dates
  const MS_PER_DAY = 86400000;
  var _now = new Date();
  const today = new Date(_now.getFullYear(), _now.getMonth(), _now.getDate());
  const in7Days = new Date(today); in7Days.setDate(in7Days.getDate() + 7);

  const baseTopUpList = getBaseRecordsForRole(topUpList);

  // Collect expired (past due) and nearly-expired (≤ 7 days) customers
  const expired = baseTopUpList.filter(function(c) {
    if (!c.endDate || c.tuStatus === 'terminate') return false;
    const exp = parseLocalDate(c.endDate);
    return exp && exp < today;
  });
  const nearlyExpired = baseTopUpList.filter(function(c) {
    if (!c.endDate || c.tuStatus === 'terminate') return false;
    const exp = parseLocalDate(c.endDate);
    return exp && exp >= today && exp <= in7Days;
  });

  const noticeEl = g('topup-expiry-notice');
  if (noticeEl) {
    const totalAlert = expired.length + nearlyExpired.length;
    if (totalAlert) {
      var cards = '';
      if (expired.length) {
        cards += '<div class="expiry-notice-card expiry-notice-card-expired">' +
          '<div class="expiry-notice-card-icon"><i class="fas fa-circle-xmark"></i></div>' +
          '<div class="expiry-notice-card-body">' +
            '<div class="expiry-notice-card-count">' + expired.length + '</div>' +
            '<div class="expiry-notice-card-label">Expired</div>' +
          '</div>' +
          '<button class="btn expiry-view-btn expiry-view-btn-red" onclick="openExpiryNoticeModal(\'expired\')"><i class="fas fa-eye"></i> View</button>' +
        '</div>';
      }
      if (nearlyExpired.length) {
        cards += '<div class="expiry-notice-card expiry-notice-card-expiring">' +
          '<div class="expiry-notice-card-icon expiry-notice-card-icon-warn"><i class="fas fa-clock"></i></div>' +
          '<div class="expiry-notice-card-body">' +
            '<div class="expiry-notice-card-count expiry-notice-card-count-warn">' + nearlyExpired.length + '</div>' +
            '<div class="expiry-notice-card-label expiry-notice-card-label-warn">Expiring Soon</div>' +
          '</div>' +
          '<button class="btn expiry-view-btn" onclick="openExpiryNoticeModal(\'expiring\')"><i class="fas fa-eye"></i> View</button>' +
        '</div>';
      }
      noticeEl.innerHTML = '<div class="expiry-notice-cards">' + cards + '</div>';
    } else {
      noticeEl.innerHTML = '';
    }
  }

  if (!baseTopUpList.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-coins" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No top up records yet</td></tr>';
    return;
  }
  const tuSearchVal = (rv('tu-search') || '').toLowerCase().trim();
  const tuFiltered = tuSearchVal
    ? baseTopUpList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(tuSearchVal) ||
               (c.phone || '').toLowerCase().includes(tuSearchVal);
      })
    : baseTopUpList;
  // Sort by expiry date ascending: soonest-to-expire (including already-expired) at top,
  // records without an expiry date below, terminated records at the very bottom.
  const tuList = tuFiltered.slice().sort(function(a, b) {
    const isTermA = (a.tuStatus === 'terminate');
    const isTermB = (b.tuStatus === 'terminate');
    if (isTermA !== isTermB) return isTermA ? 1 : -1;
    const expA = parseLocalDate(a.endDate);
    const expB = parseLocalDate(b.endDate);
    if (!expA && !expB) return 0;
    if (!expA) return 1;
    if (!expB) return -1;
    return expA - expB;
  });
  if (!tuList.length) {
    if (explicit && tuSearchVal) { showAlert('No results found for "' + tuSearchVal + '"', 'warning'); }
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-coins" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No results found</td></tr>';
    return;
  }
  tbody.innerHTML = tuList.map(function(c, i) {
    const avIdx = i % 8;
    const tuSt = c.tuStatus || 'active';
    const stPill = tuSt === 'active' ? 'pill-green' : 'pill-red';
    const stLabel = tuSt === 'active' ? 'Active' : 'Terminate';
    const canEdit = canModifyRecord(c);
    var expiryCell = '<td>—</td>';
    var rowClass = '';
    var displayDate = toLocalDateStr(c.endDate);
    if (c.endDate) {
      const exp = parseLocalDate(c.endDate);
      const daysLeft = exp ? Math.round((exp - today) / MS_PER_DAY) : null;
      if (exp && tuSt !== 'terminate' && daysLeft >= 0 && daysLeft <= 7) {
        rowClass = ' class="tr-nearly-expired"';
        expiryCell = '<td><span class="expiry-badge expiry-badge-warn">' + esc(displayDate) + ' <span class="expiry-days-left">(' + (daysLeft === 0 ? 'Today' : daysLeft + 'd') + ')</span></span></td>';
      } else if (exp && tuSt !== 'terminate' && daysLeft < 0) {
        rowClass = ' class="tr-expired"';
        expiryCell = '<td><span class="expiry-badge expiry-badge-expired">' + esc(displayDate) + ' <span class="expiry-days-left">(Expired)</span></span></td>';
      } else {
        expiryCell = '<td>' + esc(displayDate) + '</td>';
      }
    }
    return '<tr' + rowClass + '>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + fmtMoney(c.amount) + '</td>' +
      '<td><span class="pill ' + stPill + '">' + stLabel + '</span></td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      expiryCell +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-edit" onclick="editTopUp(\'' + esc(c.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteTopUp(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function openExpiryNoticeModal(type) {
  const MS_PER_DAY = 86400000;
  var _now = new Date();
  const today = new Date(_now.getFullYear(), _now.getMonth(), _now.getDate());
  const in7Days = new Date(today); in7Days.setDate(in7Days.getDate() + 7);
  const baseTopUpList = getBaseRecordsForRole(topUpList);

  const expiredList = baseTopUpList.filter(function(c) {
    if (!c.endDate || c.tuStatus === 'terminate') return false;
    const exp = parseLocalDate(c.endDate);
    return exp && exp < today;
  });
  const nearlyExpiredList = baseTopUpList.filter(function(c) {
    if (!c.endDate || c.tuStatus === 'terminate') return false;
    const exp = parseLocalDate(c.endDate);
    return exp && exp >= today && exp <= in7Days;
  });

  // Sort: most overdue first for expired, soonest first for expiring
  expiredList.sort(function(a, b) { return parseLocalDate(a.endDate) - parseLocalDate(b.endDate); });
  nearlyExpiredList.sort(function(a, b) { return parseLocalDate(a.endDate) - parseLocalDate(b.endDate); });

  var rows = [];
  var listsToShow = [];
  if (type === 'expired') {
    listsToShow.push({ list: expiredList, rowType: 'expired' });
  } else if (type === 'expiring') {
    listsToShow.push({ list: nearlyExpiredList, rowType: 'expiring' });
  } else {
    listsToShow.push({ list: expiredList, rowType: 'expired' });
    listsToShow.push({ list: nearlyExpiredList, rowType: 'expiring' });
  }
  listsToShow.forEach(function(entry) {
    entry.list.forEach(function(c) {
      var exp = parseLocalDate(c.endDate);
      rows.push({ c: c, exp: exp, type: entry.rowType, daysLeft: Math.round((exp - today) / MS_PER_DAY) });
    });
  });

  // Update modal title based on type
  var modalTitle = g('expiry-notice-modal-title');
  if (modalTitle) {
    if (type === 'expired') {
      modalTitle.innerHTML = '<i class="fas fa-circle-xmark" style="color:#ef4444;margin-right:8px;"></i>Expired Customers';
    } else if (type === 'expiring') {
      modalTitle.innerHTML = '<i class="fas fa-clock" style="color:#f59e0b;margin-right:8px;"></i>Expiring Customers';
    } else {
      modalTitle.innerHTML = '<i class="fas fa-bell" style="color:#f59e0b;margin-right:8px;"></i>Expired &amp; Expiring Customers';
    }
  }

  var tbodyHtml = rows.map(function(row, i) {
    var c = row.c;
    var displayDate = toLocalDateStr(c.endDate);
    var statusBadge = row.type === 'expired'
      ? '<span class="expiry-badge expiry-badge-expired"><i class="fas fa-circle-xmark"></i> Expired</span>'
      : '<span class="expiry-badge expiry-badge-warn"><i class="fas fa-clock"></i> Expiring Soon</span>';
    var daysText = row.daysLeft < 0
      ? Math.abs(row.daysLeft) + 'd ago'
      : (row.daysLeft === 0 ? 'Today' : row.daysLeft + 'd left');
    var actionBtns =
      '<button class="btn btn-sm btn-primary" style="margin-right:4px;" onclick="moveExpiryToActive(\'' + esc(c.id) + '\')" title="Renew subscription and move to Top Up list"><i class="fas fa-rotate-right"></i> Active</button>' +
      '<button class="btn btn-sm btn-danger" onclick="moveExpiryToTerminate(\'' + esc(c.id) + '\')" title="Terminate service and move to Terminate list"><i class="fas fa-ban"></i> Terminate</button>';
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td>' + esc(c.name) + '</td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + statusBadge + '</td>' +
      '<td><strong>' + daysText + '</strong></td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(displayDate) + '</td>' +
      '<td style="white-space:nowrap;">' + actionBtns + '</td>' +
      '</tr>';
  }).join('');

  var body = g('expiry-notice-modal-body');
  if (body) {
    if (rows.length) {
      body.innerHTML =
        '<div class="table-responsive">' +
          '<table class="data-table">' +
            '<thead><tr>' +
              '<th>#</th><th>Customer</th><th>Phone</th><th>Status</th>' +
              '<th>Days</th><th>Agent</th><th>Branch</th><th>Expiry Date</th><th>Actions</th>' +
            '</tr></thead>' +
            '<tbody>' + tbodyHtml + '</tbody>' +
          '</table>' +
        '</div>';
    } else {
      body.innerHTML = '<p style="text-align:center;padding:32px;color:#999;">No customers found.</p>';
    }
  }
  openModal('modal-expiry-notice');
}

function moveExpiryToActive(id) {
  var item = topUpList.find(function(x) { return x.id === id; });
  if (!item) { showAlert('Record not found.', 'error'); return; }
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  // Close the expiry notice modal first, then open the TopUp edit form
  closeModal('modal-expiry-notice');
  // Pre-fill the TopUp form with existing data so user can update the end date / renew
  openCustomerModal('topup', item);
}

function moveExpiryToTerminate(id) {
  var item = topUpList.find(function(x) { return x.id === id; });
  if (!item) { showAlert('Record not found.', 'error'); return; }
  if (!canModifyRecord(item)) { showAlert('You do not have permission to modify this record.', 'error'); return; }
  showConfirm(
    'Are you sure you want to terminate the service for ' + esc(item.name) + '? This will move the customer to the Terminate list.',
    function() {
      // Add to termination list if not already there
      var existingTermRec = terminationList.find(function(t) {
        return (item.customerId && t.customerId === item.customerId) || (t.name === item.name && t.phone === item.phone);
      });
      if (!existingTermRec) {
        terminationList.push({
          id: uid(), customerId: item.customerId || '', name: item.name, phone: item.phone,
          reason: 'Service terminated (expired)',
          agent: item.agent, branch: item.branch, date: todayStr(), lat: item.lat || '', lng: item.lng || ''
        });
        syncSheet('Terminations', terminationList);
        renderTerminationTable();
      }
      // Remove from Top Up list
      topUpList = topUpList.filter(function(x) { return x.id !== id; });
      syncSheet('TopUp', topUpList);
      saveAllData();
      renderTopUpTable();
      addNotification((currentUser ? currentUser.name : 'User') + ' terminated an expired customer (' + item.name + ').');
      closeModal('modal-expiry-notice');
      showToast('Customer terminated and moved to Terminate list.', 'info');
    },
    'Terminate Service',
    'Terminate',
    true
  );
}

function submitTermination(e) {
  e.preventDefault();
  const editId = rv('term-edit-id');
  const obj = {
    id: editId || uid(),
    name: rv('term-name'), phone: rv('term-phone'), reason: rv('term-reason'),
    agent: rv('term-agent'), branch: rv('term-branch'), date: rv('term-date'),
    lat: rv('term-lat') || '', lng: rv('term-lng') || ''
  };
  if (!obj.name) { showAlert('Please enter customer name'); return; }
  if (!obj.phone) { showAlert('Please enter phone number'); return; }
  if (!/^\d{6,15}$/.test(obj.phone.replace(/[\s\-+()]/g, ''))) { showAlert('Please enter a valid phone number (6–15 digits, separators allowed)'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  if (editId) {
    const idx = terminationList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) terminationList[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a termination record.');
  } else {
    terminationList.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' submitted a termination.');
  }
  closeModal('modal-termination');
  renderTerminationTable();
  syncSheet('Terminations', terminationList);
  saveAllData();
  showToast(editId ? 'Termination record updated.' : 'Termination submitted successfully.', 'success');
}

function editTermination(id) {
  const item = terminationList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  openCustomerModal('termination', item);
}

function deleteTermination(id) {
  const item = terminationList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Are you sure you want to delete this termination record? This action cannot be undone.', function() {
    terminationList = terminationList.filter(function(x) { return x.id !== id; });
    renderTerminationTable();
    syncSheet('Terminations', terminationList);
    saveAllData();
    showToast('Termination record deleted.', 'success');
  }, 'Delete Termination Record', 'Delete');
}

function renderTerminationTable(explicit) {
  const tbody = g('termination-table');
  if (!tbody) return;
  const baseTermList = getBaseRecordsForRole(terminationList);
  if (!baseTermList.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-times-circle" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No termination records yet</td></tr>';
    return;
  }
  const termSearchVal = (rv('term-search') || '').toLowerCase().trim();
  const termList = termSearchVal
    ? baseTermList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(termSearchVal) ||
               (c.phone || '').toLowerCase().includes(termSearchVal);
      })
    : baseTermList;
  if (!termList.length) {
    if (explicit && termSearchVal) { showAlert('No results found for "' + termSearchVal + '"', 'warning'); }
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-times-circle" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No results found</td></tr>';
    return;
  }
  tbody.innerHTML = termList.map(function(c, i) {
    const avIdx = i % 8;
    const canEdit = canModifyRecord(c);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + esc(c.reason || '') + '</td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      '<td>' + (c.lat || c.lng ? [c.lat, c.lng].filter(Boolean).map(esc).join(', ') : '—') + '</td>' +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-edit" onclick="editTermination(\'' + esc(c.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteTermination(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// Out Coverage Functions
// ------------------------------------------------------------
function renderOutCoverageTable(explicit) {
  const tbody = g('out-coverage-table');
  if (!tbody) return;
  const baseList = getBaseRecordsForRole(outCoverageList);
  const searchVal = (rv('oc-search') || '').toLowerCase().trim();
  const list = searchVal
    ? baseList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(searchVal) ||
               (c.phone || '').toLowerCase().includes(searchVal);
      })
    : baseList;
  if (!list.length) {
    if (explicit && searchVal) { showAlert('No results found for "' + searchVal + '"', 'warning'); }
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-signal" style="font-size:2rem;display:block;margin-bottom:8px;"></i>' + (searchVal ? 'No results found' : 'No out of coverage records yet') + '</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(function(c, i) {
    const avIdx = i % 8;
    const canEdit = canModifyRecord(c);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + esc(c.idNum || '') + '</td>' +
      '<td>' + esc(c.tariff || '') + '</td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      '<td style="white-space:nowrap;">' + (c.lat && c.lng
        ? '<a href="https://www.openstreetmap.org/?mlat=' + esc(c.lat) + '&mlon=' + esc(c.lng) + '&zoom=15" target="_blank" title="' + esc(c.lat) + ', ' + esc(c.lng) + '" style="color:#1B7D3D;text-decoration:none;"><i class="fas fa-map-marker-alt"></i> ' + esc(c.lat) + ', ' + esc(c.lng) + '</a>'
        : '<span style="color:#ccc;font-size:0.8rem;">—</span>') + '</td>' +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-delete" onclick="deleteOutCoverage(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function deleteOutCoverage(id) {
  const item = outCoverageList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Delete this out coverage record?', function() {
    outCoverageList = outCoverageList.filter(function(x) { return x.id !== id; });
    syncSheet('OutCoverage', outCoverageList);
    saveAllData();
    renderOutCoverageTable();
    showToast('Out coverage record deleted.', 'success');
  });
}

// ------------------------------------------------------------
// Promotion Functions
// ------------------------------------------------------------
function isPromoExpired(p) {
  if (!p.endDate) return false;
  const today = new Date(todayStr());
  return new Date(p.endDate) < today;
}

function openNewPromotionModal(item) {
  const form = g('form-newPromotion');
  if (form) form.reset();
  const editEl = g('np-edit-id');
  if (editEl) editEl.value = '';
  const title = g('modal-newPromotion-title');
  const btn = g('np-submit-btn');
  if (item) {
    if (title) title.textContent = 'Edit Promotion';
    if (btn) btn.textContent = 'Update Promotion';
    if (editEl) editEl.value = item.id;
    const c = g('np-campaign'); if (c) c.value = item.campaign || '';
    const ch = g('np-channel'); if (ch) ch.value = item.channel || '';
    const s = g('np-start'); if (s) s.value = item.startDate || '';
    const e = g('np-end'); if (e) e.value = item.endDate || '';
    const t = g('np-terms'); if (t) t.value = item.terms || '';
  } else {
    if (title) title.textContent = 'New Promotion';
    if (btn) btn.textContent = 'Add Promotion';
    const s = g('np-start'); if (s) s.value = todayStr();
  }
  openModal('modal-newPromotion');
}

function submitNewPromotion(e) {
  e.preventDefault();
  const editId = rv('np-edit-id');
  const obj = {
    id: editId || uid(),
    campaign: rv('np-campaign'),
    channel: rv('np-channel'),
    startDate: rv('np-start'),
    endDate: rv('np-end'),
    terms: g('np-terms') ? g('np-terms').value : ''
  };
  if (!obj.campaign) { showAlert('Please enter campaign name'); return; }
  if (!obj.startDate || !obj.endDate) { showAlert('Please enter the promotion period'); return; }
  if (editId) {
    const idx = promotionList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) promotionList[idx] = obj;
  } else {
    promotionList.push(obj);
  }
  closeModal('modal-newPromotion');
  renderPromotionCards();
  renderPromoSettingTable();
  syncSheet('Promotions', promotionList);
  saveAllData();
  showToast(editId ? 'Promotion updated.' : 'Promotion created successfully.', 'success');
}

function editNewPromotion(id) {
  const item = promotionList.find(function(x) { return x.id === id; });
  if (item) openNewPromotionModal(item);
}

function deleteNewPromotion(id) {
  showConfirm('Are you sure you want to delete this promotion? This action cannot be undone.', function() {
    promotionList = promotionList.filter(function(x) { return x.id !== id; });
    renderPromotionCards();
    renderPromoSettingTable();
    syncSheet('Promotions', promotionList);
    saveAllData();
    showToast('Promotion deleted.', 'success');
  }, 'Delete Promotion', 'Delete');
}

function renderPromotionCards() {
  const available = promotionList.filter(function(p) { return !isPromoExpired(p); });
  const expired = promotionList.filter(function(p) { return isPromoExpired(p); });

  const avCount = g('promo-available-count');
  const exCount = g('promo-expired-count');
  if (avCount) avCount.textContent = available.length;
  if (exCount) exCount.textContent = expired.length;

  const avGrid = g('promo-available-grid');
  const avEmpty = g('promo-available-empty');
  if (avGrid) {
    if (!available.length) {
      avGrid.innerHTML = '';
      if (avEmpty) avEmpty.style.display = '';
    } else {
      if (avEmpty) avEmpty.style.display = 'none';
      avGrid.innerHTML = available.map(function(p) { return buildPromoCard(p, false); }).join('');
    }
  }

  const exGrid = g('promo-expired-grid');
  const exEmpty = g('promo-expired-empty');
  if (exGrid) {
    if (!expired.length) {
      exGrid.innerHTML = '';
      if (exEmpty) exEmpty.style.display = '';
    } else {
      if (exEmpty) exEmpty.style.display = 'none';
      exGrid.innerHTML = expired.map(function(p) { return buildPromoCard(p, true); }).join('');
    }
  }
  var newBtn = g('promo-new-btn');
  if (newBtn) newBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';
  // Apply current promo view to show/hide sections
  setPromoView(currentPromoView);
}

function buildPromoCard(p, isExpired) {
  var isAdmin = (currentRole === 'admin' || currentRole === 'cluster');
  var statusBadge = isExpired
    ? '<span class="promo-status-badge promo-status-expired">Expired</span>'
    : '<span class="promo-status-badge promo-status-active">Active</span>';
  var termsHtml = p.terms ? '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-file-lines"></i></span><span class="promo-info-value promo-terms-text">' + esc(p.terms) + '</span></div>' : '';
  var actionsHtml;
  if (isAdmin) {
    if (isExpired) {
      actionsHtml = '<button class="btn-restore" onclick="restorePromotion(\'' + esc(p.id) + '\')"><i class="fas fa-rotate-left"></i> Restore</button>' +
        '<button class="btn-delete" onclick="deleteNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-trash"></i></button>';
    } else {
      actionsHtml = '<button class="btn-edit" onclick="editNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-edit"></i> Edit</button>' +
        '<button class="btn-delete" onclick="deleteNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-trash"></i></button>';
    }
  } else {
    actionsHtml = '<button class="btn-view-promo" onclick="openPromoViewModal(\'' + esc(p.id) + '\')"><i class="fas fa-eye"></i> View</button>';
  }
  return '<div class="promo-card-v2' + (isExpired ? ' expired' : '') + '">' +
    '<div class="promo-card-v2-header">' +
      '<span class="promo-card-v2-title">' + esc(p.campaign) + '</span>' +
      statusBadge +
    '</div>' +
    '<div class="promo-card-v2-body">' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-broadcast-tower"></i></span><span><span class="promo-info-label">Channel: </span><span class="promo-info-value">' + esc(p.channel || '—') + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-calendar-range"></i></span><span><span class="promo-info-label">Period: </span><span class="promo-info-value">' + esc(p.startDate || '') + ' \u2192 ' + esc(p.endDate || '') + '</span></span></div>' +
      termsHtml +
    '</div>' +
    '<div class="promo-card-v2-footer">' + actionsHtml + '</div>' +
  '</div>';
}

function restorePromotion(id) {
  var idx = promotionList.findIndex(function(x) { return x.id === id; });
  if (idx < 0) return;
  var today = new Date();
  var future = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 30);
  promotionList[idx].endDate = future.getFullYear() + '-' + String(future.getMonth() + 1).padStart(2, '0') + '-' + String(future.getDate()).padStart(2, '0');
  renderPromotionCards();
  showToast('Promotion restored for 30 days.', 'success');
  syncSheet('Promotions', promotionList);
  saveAllData();
}

function openPromoViewModal(id) {
  var p = promotionList.find(function(x) { return x.id === id; });
  if (!p) return;
  var body = g('modal-viewPromo-body');
  if (body) {
    body.innerHTML =
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-tag"></i></span><span><span class="promo-info-label">Campaign: </span><span class="promo-info-value">' + esc(p.campaign) + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-broadcast-tower"></i></span><span><span class="promo-info-label">Channel: </span><span class="promo-info-value">' + esc(p.channel || '—') + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-calendar"></i></span><span><span class="promo-info-label">Start: </span><span class="promo-info-value">' + esc(p.startDate || '—') + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-calendar-xmark"></i></span><span><span class="promo-info-label">End: </span><span class="promo-info-value">' + esc(p.endDate || '—') + '</span></span></div>' +
      (p.terms ? '<div class="promo-info-row" style="margin-top:8px;"><span class="promo-info-icon"><i class="fas fa-file-lines"></i></span><span><span class="promo-info-label">Terms: </span><span class="promo-info-value" style="font-style:italic;color:#777;">' + esc(p.terms) + '</span></span></div>' : '');
  }
  openModal('modal-viewPromo');
}

function renderPromotionTable() {
  renderPromotionCards();
}

function renderPromoSettingTable() {
  const tbody = g('promo-setting-table') ? g('promo-setting-table').querySelector('tbody') : null;
  if (!tbody) return;
  if (!promotionList.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:30px;color:#999;">No promotions defined</td></tr>';
    return;
  }
  tbody.innerHTML = promotionList.map(function(p, i) {
    const expired = isPromoExpired(p);
    const statusPill = expired ? 'pill-gray' : 'pill-green';
    const statusLabel = expired ? 'Expired' : 'Active';
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><strong>' + esc(p.campaign || p.name || '') + '</strong></td>' +
      '<td>' + esc(p.channel || '') + '</td>' +
      '<td>' + esc(p.startDate || '') + ' – ' + esc(p.endDate || '') + '</td>' +
      '<td><span class="pill ' + statusPill + '">' + statusLabel + '</span></td>' +
      '<td style="white-space:nowrap;">' +
        (expired ? '' : '<button class="btn-edit" onclick="editNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-edit"></i></button> ') +
        '<button class="btn-delete" onclick="deleteNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-trash"></i></button>' +
      '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// Deposit Functions
// ------------------------------------------------------------
var USD_DENOMS = [100, 50, 20, 10, 5, 1];
var KHR_DENOMS = [100000, 50000, 20000, 10000, 5000, 2000, 1000, 500, 100];
var KHR_EXCHANGE_RATE = 4000; // 1 USD = 4000 KHR

function khrToUsd(riel) {
  return (riel || 0) / KHR_EXCHANGE_RATE;
}

function formatKHR(n) {
  return (n || 0).toLocaleString('en-US') + ' ៛';
}

function toggleDenomSection() {
  var section = g('denom-section');
  var btn = g('denom-toggle-btn');
  if (!section) return;
  var isOpen = section.style.display !== 'none';
  section.style.display = isOpen ? 'none' : 'block';
  if (btn) btn.classList.toggle('open', !isOpen);
}

function calcDenomTotals() {
  var usdTotal = 0;
  USD_DENOMS.forEach(function(d) {
    var el = g('usd-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    var sub = qty * d;
    usdTotal += sub;
    var subEl = g('usd-sub-' + d);
    if (subEl) {
      subEl.textContent = qty > 0 ? '$' + sub.toFixed(2) : '—';
      subEl.classList.toggle('has-value', qty > 0);
    }
  });
  var usdHeaderEl = g('usd-denom-total');
  if (usdHeaderEl) usdHeaderEl.textContent = '$' + usdTotal.toFixed(2);
  var usdRowEl = g('usd-denom-total-row');
  if (usdRowEl) usdRowEl.textContent = '$' + usdTotal.toFixed(2);
  var cashEl = g('dep-cash'); if (cashEl && usdTotal > 0) cashEl.value = usdTotal.toFixed(2);

  var khrTotal = 0;
  KHR_DENOMS.forEach(function(d) {
    var el = g('khr-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    var sub = qty * d;
    khrTotal += sub;
    var subEl = g('khr-sub-' + d);
    if (subEl) {
      subEl.textContent = qty > 0 ? formatKHR(sub) : '—';
      subEl.classList.toggle('has-value', qty > 0);
    }
  });
  var khrHeaderEl = g('khr-denom-total');
  if (khrHeaderEl) khrHeaderEl.textContent = formatKHR(khrTotal);
  var khrRowEl = g('khr-denom-total-row');
  if (khrRowEl) khrRowEl.textContent = formatKHR(khrTotal);
  var rielEl = g('dep-riel'); if (rielEl && khrTotal > 0) rielEl.value = khrTotal;
}

function _resetDenomSection() {
  var section = g('denom-section');
  var btn = g('denom-toggle-btn');
  if (section) section.style.display = 'none';
  if (btn) btn.classList.remove('open');
  USD_DENOMS.forEach(function(d) {
    var el = g('usd-qty-' + d); if (el) el.value = '';
    var subEl = g('usd-sub-' + d); if (subEl) { subEl.textContent = '—'; subEl.classList.remove('has-value'); }
  });
  KHR_DENOMS.forEach(function(d) {
    var el = g('khr-qty-' + d); if (el) el.value = '';
    var subEl = g('khr-sub-' + d); if (subEl) { subEl.textContent = '—'; subEl.classList.remove('has-value'); }
  });
  var usdH = g('usd-denom-total'); if (usdH) usdH.textContent = '$0.00';
  var usdR = g('usd-denom-total-row'); if (usdR) usdR.textContent = '$0.00';
  var khrH = g('khr-denom-total'); if (khrH) khrH.textContent = '0 ៛';
  var khrR = g('khr-denom-total-row'); if (khrR) khrR.textContent = '0 ៛';
}

// Returns the credit amount from a deposit record, handling both the new
// `creditAmount` field and the legacy `credit` field for backward compatibility.
function _depCreditAmt(deposit) {
  return parseFloat(deposit.creditAmount !== undefined ? deposit.creditAmount : deposit.credit) || 0;
}

// Normalizes a sale record after reading from Google Sheets or localStorage,
// ensuring object fields (items, dollarItems) are parsed back from the JSON
// strings that the Sheets API stores them as, and that numeric fields
// (kpiAmount, kpiPoints) are stored as numbers.
function normalizeSaleRecord(s) {
  if (!s || typeof s !== 'object') return s;
  var out = Object.assign({}, s);
  // Parse items and dollarItems back from JSON strings (when loaded from Google Sheets)
  ['items', 'dollarItems'].forEach(function(key) {
    if (typeof out[key] === 'string') {
      try { out[key] = JSON.parse(out[key]); } catch (e) { console.warn('normalizeSaleRecord: failed to parse ' + key + ':', out[key], e); out[key] = {}; }
    }
    if (!out[key] || typeof out[key] !== 'object' || Array.isArray(out[key])) {
      out[key] = {};
    }
  });
  if (out.kpiAmount !== undefined && out.kpiAmount !== '') {
    out.kpiAmount = parseFloat(out.kpiAmount) || 0;
  }
  if (out.kpiPoints !== undefined && out.kpiPoints !== '') {
    out.kpiPoints = parseFloat(out.kpiPoints) || 0;
  }
  return out;
}

// Normalizes a deposit record to ensure all required fields are present.
// Back-fills the `cash` field for legacy records loaded from Google Sheets or
// localStorage that were created before the cash column existed.
// Also ensures the status-related fields have sensible defaults so that legacy
// records without these columns still display and sync correctly.
function normalizeDeposit(d) {
  if (!d || typeof d !== 'object') {
    console.warn('normalizeDeposit: received non-object value', d);
    return d;
  }
  var out = Object.assign({}, d);
  // Convert numeric fields to numbers – GAS readSheetData returns all cell
  // values as strings (String(raw)), so these must be explicitly coerced after
  // a sync-down to avoid string concatenation in arithmetic operations.
  out.riel          = parseFloat(out.riel)          || 0;
  out.creditAmount  = _depCreditAmt(out);
  out.amount        = parseFloat(out.amount)        || 0;
  // Ensure cash is a number; infer from total when missing or empty string
  if (out.cash === undefined || out.cash === '' || out.cash === null) {
    var rielInUsd = parseFloat(khrToUsd(out.riel)) || 0;
    out.cash = Math.max(0, out.amount - out.creditAmount - rielInUsd);
  } else {
    out.cash = parseFloat(out.cash) || 0;
  }
  // Ensure status-related fields are present so legacy records without these
  // columns are handled consistently and synced with correct values to the sheet.
  if (out.status === undefined || out.status === null || out.status === '') out.status = 'pending';
  if (out.approvedBy === undefined || out.approvedBy === null) out.approvedBy = '';
  if (out.approvedAt === undefined || out.approvedAt === null) out.approvedAt = '';
  return out;
}

function _loadDenomSection(cashDetail) {
  if (!cashDetail) return;
  var section = g('denom-section');
  var btn = g('denom-toggle-btn');
  if (section) section.style.display = 'block';
  if (btn) btn.classList.add('open');
  (cashDetail.usd || []).forEach(function(entry) {
    var el = g('usd-qty-' + entry.denom); if (el) el.value = entry.qty;
  });
  (cashDetail.khr || []).forEach(function(entry) {
    var el = g('khr-qty-' + entry.denom); if (el) el.value = entry.qty;
  });
  calcDenomTotals();
}

function openAddDeposit(el) {
  navigateTo('deposit', null);
  setActiveSubItem(el);
  openDepositModal(null);
}

function openDepositModal(item) {
  const form = g('form-addDeposit');
  if (form) form.reset();
  const editEl = g('dep-edit-id');
  if (editEl) editEl.value = '';
  _resetDenomSection();

  // Always reset field lock states so they don't persist between modal opens
  const agEl = g('dep-agent');
  const brEl = g('dep-branch');
  const brTextEl = g('dep-branch-text');
  if (agEl) agEl.readOnly = false;
  if (brEl) { brEl.disabled = false; brEl.style.display = ''; brEl.required = true; }
  if (brTextEl) brTextEl.style.display = 'none';

  const title = g('modal-addDeposit-title');
  const btn = g('dep-submit-btn');
  populateBranchSelects();

  if (item) {
    if (title) title.textContent = 'Edit Deposit';
    if (btn) btn.textContent = 'Update Deposit';
    if (editEl) editEl.value = item.id;
    if (agEl) agEl.value = item.agent || '';
    if (brEl) brEl.value = item.branch || '';
    const cashEl = g('dep-cash'); if (cashEl) cashEl.value = item.cash || '';
    const creditEl = g('dep-credit'); if (creditEl) creditEl.value = _depCreditAmt(item) || '';
    const rielEl = g('dep-riel'); if (rielEl) rielEl.value = item.riel || '';
    const dtEl = g('dep-date'); if (dtEl) dtEl.value = item.date || '';
    const ntEl = g('dep-remark'); if (ntEl) ntEl.value = item.remark || item.note || '';
    if (item.cashDetail) _loadDenomSection(item.cashDetail);
    if (currentUser && (currentRole === 'agent' || currentRole === 'supervisor')) {
      if (agEl) agEl.readOnly = true;
      if (brEl && brTextEl) {
        brTextEl.textContent = item.branch || '';
        brEl.style.display = 'none';
        brEl.required = false;
        brTextEl.style.display = '';
      }
    }
  } else {
    if (title) title.textContent = 'Add Deposit';
    if (btn) btn.textContent = 'Add Deposit';
    const dtEl = g('dep-date'); if (dtEl) dtEl.value = todayStr();
    if (currentUser) {
      if (agEl) {
        agEl.value = currentUser.name || '';
        if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true;
      }
      if (brEl) {
        if (currentUser.branch) brEl.value = currentUser.branch;
        if (currentRole === 'agent' || currentRole === 'supervisor') {
          if (brTextEl) {
            brTextEl.textContent = currentUser.branch || '';
            brTextEl.style.display = '';
          }
          brEl.style.display = 'none';
          brEl.required = false;
        }
      }
    }
  }
  openModal('modal-addDeposit');
}

function submitDeposit(e) {
  e.preventDefault();
  const editId = rv('dep-edit-id');
  const cash = parseFloat(rv('dep-cash')) || 0;
  const credit = parseFloat(rv('dep-credit')) || 0;
  const riel = parseFloat(rv('dep-riel')) || 0;
  const cashDetailUsd = [];
  const cashDetailKhr = [];
  USD_DENOMS.forEach(function(d) {
    var el = g('usd-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    if (qty > 0) cashDetailUsd.push({ denom: d, qty: qty });
  });
  KHR_DENOMS.forEach(function(d) {
    var el = g('khr-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    if (qty > 0) cashDetailKhr.push({ denom: d, qty: qty });
  });
  var existingRecord = editId ? depositList.find(function(x){return x.id===editId;}) : null;
  const obj = {
    id: editId || uid(),
    agent: rv('dep-agent'),
    branch: rv('dep-branch'),
    cash: cash,
    creditAmount: credit,
    riel: riel,
    cashDetail: (cashDetailUsd.length || cashDetailKhr.length) ? { usd: cashDetailUsd, khr: cashDetailKhr } : null,
    amount: cash + credit + khrToUsd(riel),
    date: rv('dep-date'),
    submittedAt: (existingRecord && existingRecord.submittedAt) || new Date().toISOString(),
    remark: rv('dep-remark'),
    status: editId ? (existingRecord || {}).status || 'pending' : 'pending',
    approvedBy: editId ? (existingRecord || {}).approvedBy || '' : '',
    approvedAt: editId ? (existingRecord || {}).approvedAt || '' : ''
  };
  if (!obj.agent) { showAlert('Please enter agent name'); return; }
  if (obj.amount <= 0 && riel <= 0) { showAlert('Please enter a cash, credit, and/or KHR riel amount greater than zero'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  if (editId) {
    const idx = depositList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) depositList[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a deposit record.');
  } else {
    depositList.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' added a deposit.');
  }
  closeModal('modal-addDeposit');
  renderDepositTable();
  updateDepositKpis();
  syncSheet('Deposits', depositList);
  saveAllData();
  showToast(editId ? 'Deposit record updated.' : 'Deposit submitted successfully.', 'success');
}

function editDeposit(id) {
  const item = depositList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  if (item.status === 'approved' && currentRole !== 'admin') { showAlert('This deposit has already been approved and cannot be edited.', 'warning'); return; }
  openDepositModal(item);
}

function deleteDeposit(id) {
  const item = depositList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  if (item.status === 'approved' && currentRole !== 'admin') { showAlert('This deposit has already been approved and cannot be deleted.', 'warning'); return; }
  showConfirm('Are you sure you want to delete this deposit record? This action cannot be undone.', function() {
    depositList = depositList.filter(function(x) { return x.id !== id; });
    renderDepositTable();
    updateDepositKpis();
    syncSheet('Deposits', depositList);
    saveAllData();
    showToast('Deposit record deleted.', 'success');
  }, 'Delete Deposit Record', 'Delete');
}

function approveDeposit(id) {
  var canApprove = (currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster');
  if (!canApprove) { showAlert('Only supervisor, admin, or cluster can approve deposits.', 'warning'); return; }
  showConfirm('Approve this deposit record?', function() {
    var idx = depositList.findIndex(function(x) { return x.id === id; });
    if (idx >= 0) {
      depositList[idx].status = 'approved';
      depositList[idx].approvedBy = currentUser ? currentUser.name : 'Supervisor';
      depositList[idx].approvedAt = todayStr();
      renderDepositTable();
      updateDepositKpis();
      saveAllData();
      // Show sync indicator
      var ind = document.getElementById('gs-sync-indicator');
      var lbl = document.getElementById('gs-sync-status');
      if (ind) ind.className = 'syncing';
      if (lbl) lbl.textContent = 'Syncing\u2026';
      var approvedRecord = depositList[idx];
      // Use a targeted updateRow to write only the three status-related fields so
      // that other concurrent changes are not overwritten.  Fall back to a full
      // sync if the targeted update fails (e.g. sheet does not yet have the row).
      var statusPatch = {
        id:         approvedRecord.id,
        status:     approvedRecord.status,
        approvedBy: approvedRecord.approvedBy,
        approvedAt: approvedRecord.approvedAt
      };
      var doFullSync = function(reason) {
        console.warn('GS updateRow failed (' + reason + '), falling back to full sync.');
        return _gsPost({ sheet: 'Deposits', action: 'sync', data: normalizeArrayForSheet(depositList) })
          .then(function() {
            if (ind) ind.className = '';
            if (lbl) lbl.textContent = 'Synced \u2713';
            setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
            showToast('Deposit approved.', 'success');
          })
          .catch(function(syncErr) {
            console.warn('GS sync error on approve:', syncErr);
            if (ind) ind.className = 'error';
            if (lbl) lbl.textContent = 'Sync failed';
            showToast('Deposit approved locally but failed to sync to Google Sheets. Please use Sync Up to push the data.', 'warning');
          });
      };
      _gsPost({ sheet: 'Deposits', action: 'updateRow', data: statusPatch })
        .then(function(resp) {
          // _gsPost resolves even for GAS-level errors; check the returned status.
          if (resp && resp.status === 'error') {
            return doFullSync(resp.message || 'GAS returned error');
          }
          if (ind) ind.className = '';
          if (lbl) lbl.textContent = 'Synced \u2713';
          setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
          showToast('Deposit approved.', 'success');
        })
        .catch(function(err) {
          return doFullSync(err);
        });
      // Show the approval form modal immediately (before sync resolves) so the
      // user sees confirmation regardless of sync outcome, as required.
      showApprovalFormModal('deposit', approvedRecord);
    }
  }, 'Approve Deposit', 'Approve', false);
}

function viewDepositRevenue(id) {
  var record = depositList.find(function(x) { return x.id === id; });
  if (record) showApprovalFormModal('deposit', record);
}

function updateDepositKpis() {
  const baseDeposits = getBaseRecordsForRole(depositList);
  let totalCash = 0, totalCredit = 0, totalRiel = 0;
  const agents = new Set();
  baseDeposits.forEach(function(d) {
    totalCash += (d.cash || 0);
    totalCredit += _depCreditAmt(d);
    totalRiel += (d.riel || 0);
    agents.add(d.agent);
  });
  const total = totalCash + totalCredit + khrToUsd(totalRiel);
  const el1 = g('dep-kpi-total'); if (el1) el1.textContent = fmtMoney(total);
  const el2 = g('dep-kpi-count'); if (el2) el2.textContent = baseDeposits.length;
  const el3 = g('dep-kpi-agents'); if (el3) el3.textContent = agents.size;
  const el4 = g('dep-kpi-cash'); if (el4) el4.textContent = fmtMoney(totalCash + khrToUsd(totalRiel));
  const el5 = g('dep-kpi-credit'); if (el5) el5.textContent = fmtMoney(totalCredit);
  renderDepositChart();
}

function renderDepositChart() {
  _cDepositPerf = destroyChart(_cDepositPerf);
  const canvas = g('cDepositPerf');
  if (!canvas || typeof Chart === 'undefined') return;
  const periodEl = g('dep-chart-period');
  const period = periodEl ? periodEl.value : 'monthly';
  const baseDepositList = getBaseRecordsForRole(depositList);

  const now = new Date();
  let labels = [], cashData = [], creditData = [];

  if (period === 'weekly') {
    for (let i = 6; i >= 0; i--) {
      const d = new Date(now); d.setDate(d.getDate() - i);
      const key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
      labels.push(key.slice(5));
      let c = 0, cr = 0;
      baseDepositList.forEach(function(dep) { if (dep.date === key) { c += (dep.cash||0); cr += _depCreditAmt(dep); } });
      cashData.push(c); creditData.push(cr);
    }
  } else if (period === 'monthly') {
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      labels.push(ymLabel(key));
      let c = 0, cr = 0;
      baseDepositList.forEach(function(dep) { if (dep.date && dep.date.startsWith(key)) { c += (dep.cash||0); cr += _depCreditAmt(dep); } });
      cashData.push(c); creditData.push(cr);
    }
  } else {
    for (let i = 2; i >= 0; i--) {
      const yr = now.getFullYear() - i;
      labels.push(String(yr));
      let c = 0, cr = 0;
      baseDepositList.forEach(function(dep) { if (dep.date && dep.date.startsWith(String(yr))) { c += (dep.cash||0); cr += _depCreditAmt(dep); } });
      cashData.push(c); creditData.push(cr);
    }
  }

  _cDepositPerf = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Cash ($)', data: cashData, backgroundColor: 'rgba(27,125,61,0.8)', borderColor: '#1B7D3D', borderWidth: 1 },
        { label: 'Credit ($)', data: creditData, backgroundColor: 'rgba(21,101,192,0.8)', borderColor: '#1565C0', borderWidth: 1 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } } },
      scales: {
        x: { stacked: true, ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, stacked: true, ticks: { font: { size: 10 } } }
      }
    }
  });
}

function renderDepositTable() {
  const table = g('deposit-table');
  if (!table) return;
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  if (!tbody) return;

  const baseDepositList = getBaseRecordsForRole(depositList);

  // Update header to include cash, credit, status columns
  if (thead) {
    thead.innerHTML = '<tr><th>#</th><th>Agent</th><th>Branch</th><th>Cash ($)</th><th>Cash KHR (៛)</th><th>Credit ($)</th><th>Total</th><th>Date</th><th>Remark</th><th>Status</th><th>Actions</th></tr>';
  }

  if (!baseDepositList.length) {
    tbody.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-piggy-bank" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No deposit records yet</td></tr>';
    return;
  }
  const canApprove = (currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster');
  tbody.innerHTML = baseDepositList.map(function(d, i) {
    const avIdx = i % 8;
    const status = d.status || 'pending';
    const statusPill = status === 'approved' ? 'pill-green' : 'pill-orange';
    const statusLabel = status === 'approved' ? 'Approved' : 'Pending';
    const approveBtn = (canApprove && status !== 'approved') ? '<button class="btn-edit" onclick="approveDeposit(\'' + esc(d.id) + '\')" title="Approve"><i class="fas fa-check-circle"></i></button> ' : '';
    const viewBtn = (status === 'approved') ? '<button class="btn-edit" onclick="viewDepositRevenue(\'' + esc(d.id) + '\')" title="View Revenue Detail" style="color:#1565C0;"><i class="fas fa-eye"></i></button> ' : '';
    // Approved records may only be edited/deleted by admin
    const canEdit = canModifyRecord(d) && (status !== 'approved' || currentRole === 'admin');
    var dCreditAmt = _depCreditAmt(d);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(d.agent)) + '</span>' + esc(d.agent) + '</div></td>' +
      '<td>' + esc(d.branch || '') + '</td>' +
      '<td style="color:#1B7D3D;font-weight:600;">' + (d.cash ? '$' + Number(d.cash).toFixed(2) : '—') + '</td>' +
      '<td style="color:#8B6914;font-weight:600;">' + (d.riel ? formatKHR(d.riel) : '—') + '</td>' +
      '<td style="color:#1565C0;font-weight:600;">' + (dCreditAmt ? '$' + Number(dCreditAmt).toFixed(2) : '—') + '</td>' +
      '<td style="font-weight:700;color:#1B7D3D;">$' + Number((d.cash || 0) + khrToUsd(d.riel) + (dCreditAmt || 0)).toFixed(2) + '</td>' +
      '<td>' + esc(toLocalDateStr(d.date) || '') + '</td>' +
      '<td style="color:#888;font-size:0.8rem;">' + esc(d.remark || d.note || '') + '</td>' +
      '<td><span class="pill ' + statusPill + '">' + statusLabel + '</span></td>' +
      '<td style="white-space:nowrap;">' +
        approveBtn +
        viewBtn +
        (canEdit ? '<button class="btn-edit" onclick="editDeposit(\'' + esc(d.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteDeposit(\'' + esc(d.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function setDepositView(view) {
  var depBtns = document.querySelectorAll('#page-deposit .view-toggle-btn');
  depBtns.forEach(function(b) { b.classList.remove('active'); });
  var btn = g(view === 'table' ? 'dep-view-btn-table' : 'dep-view-btn-summary');
  if (btn) btn.classList.add('active');

  var tableCard = g('deposit-table-card');
  var summaryView = g('deposit-summary-view');

  if (view === 'table') {
    if (tableCard) tableCard.style.display = '';
    if (summaryView) summaryView.style.display = 'none';
  } else {
    if (tableCard) tableCard.style.display = 'none';
    if (summaryView) summaryView.style.display = '';
    renderDepositSummaryView();
  }
}

function renderDepositSummaryView() {
  var container = g('deposit-summary-view');
  if (!container) return;

  const baseDepositList = getBaseRecordsForRole(depositList);
  if (!baseDepositList.length) {
    container.innerHTML = '<div style="text-align:center;padding:40px;color:#999;"><i class="fas fa-inbox fa-3x" style="display:block;margin-bottom:12px;"></i>No deposit records found</div>';
    return;
  }

  var agentMap = {};
  baseDepositList.forEach(function(d) {
    if (!agentMap[d.agent]) agentMap[d.agent] = { cash: 0, riel: 0, credit: 0, total: 0, count: 0, branch: d.branch || '' };
    var ag = agentMap[d.agent];
    var dCreditAmt = _depCreditAmt(d);
    ag.cash += (d.cash || 0);
    ag.riel += (d.riel || 0);
    ag.credit += dCreditAmt;
    ag.total += (d.cash || 0) + khrToUsd(d.riel) + dCreditAmt;
    ag.count++;
  });

  var cards = Object.keys(agentMap).map(function(agent, i) {
    var ag = agentMap[agent];
    var avIdx = i % 8;
    return '<div class="summary-card">' +
      '<div class="summary-card-header">' +
        '<span class="avatar-circle av-' + avIdx + '" style="width:36px;height:36px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.8rem;font-weight:700;color:#fff;">' + esc(ini(agent)) + '</span>' +
        '<div><div class="sc-name">' + esc(agent) + '</div><div style="font-size:0.72rem;opacity:0.8;">' + esc(ag.branch) + ' &middot; ' + ag.count + ' deposit' + (ag.count !== 1 ? 's' : '') + '</div></div>' +
      '</div>' +
      '<div class="summary-card-body">' +
        '<div class="summary-row"><span>Cash</span><span style="color:#1B7D3D;font-weight:600;">$' + Number(ag.cash).toFixed(2) + '</span></div>' +
        '<div class="summary-row"><span>Cash KHR</span><span style="color:#8B6914;font-weight:600;">' + formatKHR(ag.riel) + '</span></div>' +
        '<div class="summary-row"><span>Credit</span><span style="color:#1565C0;font-weight:600;">$' + Number(ag.credit).toFixed(2) + '</span></div>' +
        '<div class="summary-row"><span>Total</span><span style="font-weight:700;color:#1A1A2E;">$' + Number(ag.total).toFixed(2) + '</span></div>' +
      '</div>' +
    '</div>';
  }).join('');

  container.innerHTML = '<div class="summary-grid">' + cards + '</div>';
}

// ------------------------------------------------------------
// Staff Functions
// ------------------------------------------------------------
function openUserModal(user) {
  const form = g('form-addUser');
  if (form) form.reset();
  g('user-edit-id').value = '';
  const title = g('modal-addUser-title');
  const btn = g('user-submit-btn');
  populateBranchSelects();

  if (user) {
    if (title) title.textContent = 'Edit User';
    if (btn) btn.textContent = 'Update User';
    g('user-edit-id').value = user.id;
    g('user-name').value = user.name || '';
    g('user-username').value = user.username || '';
    g('user-password').value = user.password || '';
    g('user-role').value = user.role || 'Agent';
    const bInput = g('user-branch'); if (bInput) bInput.value = user.branch || '';
    g('user-status').value = user.status || 'active';
    const emailInput = g('user-email'); if (emailInput) emailInput.value = user.email || '';
  } else {
    if (title) title.textContent = 'Add User';
    if (btn) btn.textContent = 'Add User';
  }
  openModal('modal-addUser');
}

function submitUser(e) {
  e.preventDefault();
  const editId = rv('user-edit-id');
  const obj = {
    id: editId || uid(),
    name: rv('user-name'), username: rv('user-username'), password: rv('user-password'),
    role: rv('user-role'), branch: rv('user-branch'), status: rv('user-status'),
    email: rv('user-email') || ''
  };
  if (!obj.name) { showAlert('Please enter user name'); return; }
  if (!obj.username) { showAlert('Please enter username'); return; }
  if (!editId && !obj.password) { showAlert('Please enter a password for the new user'); return; }
  const dupUser = staffList.find(function(x) { return x.username.toLowerCase() === obj.username.toLowerCase() && x.id !== editId; });
  if (dupUser) { showAlert('Username already exists. Please choose a different username.'); return; }
  if (editId) {
    const idx = staffList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) {
      // If password field was left blank on edit, preserve the existing password
      if (!obj.password) { obj.password = staffList[idx].password || ''; }
      staffList[idx] = obj;
    }
  } else {
    staffList.push(obj);
  }
  closeModal('modal-addUser');
  renderStaffTable();
  syncSheet('Staff', staffList);
  saveAllData();
  showToast(editId ? 'User updated successfully.' : 'User added successfully.', 'success');
}

function editUser(id) {
  const user = staffList.find(function(x) { return x.id === id; });
  if (user) openUserModal(user);
}

function deleteUser(id) {
  showConfirm('Are you sure you want to delete this user? This action cannot be undone.', function() {
    staffList = staffList.filter(function(x) { return x.id !== id; });
    renderStaffTable();
    syncSheet('Staff', staffList);
    saveAllData();
    showToast('User deleted.', 'success');
  }, 'Delete User', 'Delete');
}

function renderStaffTable() {
  const tbody = g('staff-table');
  if (!tbody) return;
  if (!staffList.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-users-cog" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No users yet</td></tr>';
    return;
  }
  tbody.innerHTML = staffList.map(function(u, i) {
    const rolePill = u.role === 'Admin' ? 'pill-green' : u.role === 'Cluster' ? 'pill-purple' : u.role === 'Supervisor' ? 'pill-blue' : u.role === 'Agent' ? 'pill-orange' : 'pill-gray';
    const statusPill = u.status === 'active' ? 'pill-green' : 'pill-red';
    const avIdx = i % 8;
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(u.name)) + '</span>' + esc(u.name) + '</div></td>' +
      '<td>' + esc(u.username) + '</td>' +
      '<td><span class="pill ' + rolePill + '">' + esc(u.role) + '</span></td>' +
      '<td>' + esc(u.branch || '') + '</td>' +
      '<td>' + (u.email
        ? '<a href="mailto:' + encodeURIComponent(u.email) + '" style="color:#1B7D3D;text-decoration:none;"><i class="fas fa-envelope" style="margin-right:4px;font-size:.8rem;"></i>' + esc(u.email) + '</a>'
        : '<span style="color:#ccc;">—</span>') + '</td>' +
      '<td><span class="pill ' + statusPill + '">' + esc(u.status) + '</span></td>' +
      '<td style="white-space:nowrap;">' +
        '<button class="btn-edit" onclick="editUser(\'' + esc(u.id) + '\')"><i class="fas fa-edit"></i></button> ' +
        '<button class="btn-delete" onclick="deleteUser(\'' + esc(u.id) + '\')"><i class="fas fa-trash"></i></button>' +
      '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// KPI Functions
// ------------------------------------------------------------
function onUserRoleChange() {
  // reserved for role-specific UI changes
}

function onKpiShopNameInput(val) {
  var hiddenEl = g('kpi-shop-assignee');
  if (!hiddenEl) return;
  var sups = staffList.filter(function(u) { return u.role === 'Supervisor'; });
  var searchVal = (val || '').trim().toLowerCase();
  var found = sups.find(function(u) { return u.name.toLowerCase() === searchVal; });
  hiddenEl.value = found ? found.id : '';
}

function selectKpiFor(type) {
  kpiForSelected = type;
  const shopBtn = g('kpi-for-shop');
  const agentBtn = g('kpi-for-agent');
  if (shopBtn) { shopBtn.classList.toggle('active', type === 'shop'); }
  if (agentBtn) { agentBtn.classList.toggle('active', type === 'agent'); }

  const shopGroup = g('kpi-shop-assignee-group');
  const agentBranchGroup = g('kpi-agent-branch-group');
  const agentAssigneeGroup = g('kpi-agent-assignee-group');

  if (type === 'shop') {
    if (shopGroup) shopGroup.style.display = '';
    if (agentBranchGroup) agentBranchGroup.style.display = 'none';
    if (agentAssigneeGroup) agentAssigneeGroup.style.display = 'none';
    populateKpiShopAssignee();
  } else {
    if (shopGroup) shopGroup.style.display = 'none';
    if (agentBranchGroup) agentBranchGroup.style.display = '';
    if (agentAssigneeGroup) agentAssigneeGroup.style.display = '';
    populateKpiAgentBranch();
  }
}

function populateKpiShopAssignee(preselectedId) {
  const textEl = g('kpi-shop-assignee-name');
  const hiddenEl = g('kpi-shop-assignee');
  const datalist = g('kpi-sup-list');
  if (!textEl) return;
  const sups = staffList.filter(function(u) { return u.role === 'Supervisor'; });

  if (currentRole === 'supervisor' && currentUser) {
    // Supervisor can only assign KPI to themselves — show readonly textbox
    textEl.value = currentUser.name;
    textEl.readOnly = true;
    textEl.style.background = '#f5f5f5';
    if (hiddenEl) hiddenEl.value = currentUser.id;
    if (datalist) datalist.innerHTML = '';
  } else {
    // Admin: populate datalist with all supervisors, allow typing to choose
    textEl.readOnly = false;
    textEl.style.background = '';
    if (datalist) {
      datalist.innerHTML = sups.map(function(u) {
        return '<option value="' + esc(u.name) + '"></option>';
      }).join('');
    }
    if (preselectedId) {
      var sup = sups.find(function(u) { return u.id === preselectedId; });
      textEl.value = sup ? sup.name : '';
      if (hiddenEl) hiddenEl.value = preselectedId;
    } else {
      textEl.value = '';
      if (hiddenEl) hiddenEl.value = '';
    }
  }
}

function populateKpiAgentBranch() {
  const branchSel = g('kpi-agent-branch');
  if (!branchSel) return;

  if (currentRole === 'supervisor' && currentUser) {
    // Lock branch dropdown to supervisor's own branch
    branchSel.innerHTML = '<option value="' + esc(currentUser.branch) + '">' + esc(currentUser.branch) + '</option>';
    branchSel.disabled = true;
    branchSel.value = currentUser.branch;
  } else {
    branchSel.disabled = false;
    branchSel.innerHTML = '<option value="">Select branch</option>' +
      getBranches().map(function(b) { return '<option value="' + esc(b) + '">' + esc(b) + '</option>'; }).join('');
  }

  const agentSel = g('kpi-agent-assignee');
  if (agentSel) agentSel.innerHTML = '<option value="">Select Agent</option>';

  // Auto-load agents when supervisor's branch is pre-set
  if (currentRole === 'supervisor' && currentUser) {
    onKpiBranchChange();
  }
}

function onKpiBranchChange() {
  const branch = (currentRole === 'supervisor' && currentUser) ? currentUser.branch : rv('kpi-agent-branch');
  const agentSel = g('kpi-agent-assignee');
  if (!agentSel) return;
  const agents = staffList.filter(function(u) { return u.role === 'Agent' && u.branch === branch; });
  agentSel.innerHTML = '<option value="">Select Agent</option>' +
    agents.map(function(u) { return '<option value="' + esc(u.id) + '">' + esc(u.name) + '</option>'; }).join('');
}

function openKpiModal(item) {
  const form = g('form-kpi');
  if (form) form.reset();
  g('kpi-edit-id').value = '';
  kpiTypeSelected = 'Sales';
  setValueMode('currency');
  kpiForSelected = (item && item.kpiFor) ? item.kpiFor : 'shop';
  selectKpiFor(kpiForSelected);

  const title = g('modal-kpi-title');
  const btn = g('kpi-submit-btn');

  $$('.kpi-type-chip').forEach(function(c) { c.classList.remove('active'); });
  const firstChip = g('kpi-chip-Sales');
  if (firstChip) firstChip.classList.add('active');

  if (item) {
    if (title) title.textContent = 'Edit KPI';
    if (btn) btn.textContent = 'Update KPI';
    g('kpi-edit-id').value = item.id;
    g('kpi-name').value = item.name || '';
    g('kpi-target').value = item.target || '';
    g('kpi-period').value = item.period || 'Monthly';

    // Populate point tier fields if editing a point-based KPI
    var pt = item.pointTiers || {};
    var minEl = g('kpi-point-min'); if (minEl) minEl.value = pt.min || '';
    var gateEl = g('kpi-point-gate'); if (gateEl) gateEl.value = pt.gate || '';
    var otbEl = g('kpi-point-otb'); if (otbEl) otbEl.value = pt.otb || '';
    var oabEl = g('kpi-point-oab'); if (oabEl) oabEl.value = pt.oab || '';

    kpiTypeSelected = item.type || 'Sales';
    $$('.kpi-type-chip').forEach(function(c) { c.classList.remove('active'); });
    const chip = g('kpi-chip-' + kpiTypeSelected);
    if (chip) chip.classList.add('active');

    switchKpiScheme(item.kpiMode === 'point' ? 'point' : 'volume');

    if (item.kpiMode !== 'point') {
      setValueMode(item.valueMode || 'unit');
      if (item.valueMode === 'unit') {
        const itemSel = g('kpi-item-sel'); if (itemSel && item.itemId) itemSel.value = item.itemId;
      } else {
        const csEl = g('kpi-currency-sel'); if (csEl) csEl.value = item.currency || 'USD';
        const ciEl = g('kpi-currency-item-sel'); if (ciEl && item.itemId) ciEl.value = item.itemId;
      }
    }
    if (item.kpiFor === 'shop') {
      populateKpiShopAssignee(item.assigneeId);
    } else if (item.kpiFor === 'agent') {
      populateKpiAgentBranch();
      const branchSel = g('kpi-agent-branch');
      if (branchSel && item.assigneeBranch) {
        branchSel.value = item.assigneeBranch;
        onKpiBranchChange();
        const agentSel = g('kpi-agent-assignee');
        if (agentSel && item.assigneeId) agentSel.value = item.assigneeId;
      }
    }
  } else {
    if (title) title.textContent = 'Add KPI';
    if (btn) btn.textContent = 'Add KPI';
    switchKpiScheme('volume');
  }
  openModal('modal-kpi');
}

function selectKpiType(el) {
  const type = el ? el.getAttribute('data-value') : 'Sales';
  kpiTypeSelected = type;
  $$('.kpi-type-chip').forEach(function(c) { c.classList.remove('active'); });
  if (el) el.classList.add('active');
  var autoMode = (type === 'Sales' || type === 'Revenue') ? 'currency' : 'unit';
  setValueMode(autoMode);
}

function switchKpiScheme(mode) {
  kpiModeSelected = mode;
  var volBtn = g('kpi-kpimode-volume');
  var ptBtn = g('kpi-kpimode-point');
  if (volBtn) volBtn.classList.toggle('active', mode === 'volume');
  if (ptBtn) ptBtn.classList.toggle('active', mode === 'point');
  var isPoint = mode === 'point';
  var typeGroup = g('kpi-type-group');
  var valuemodeGroup = g('kpi-valuemode-group');
  if (typeGroup) typeGroup.style.display = isPoint ? 'none' : '';
  if (valuemodeGroup) valuemodeGroup.style.display = isPoint ? 'none' : '';
  // Toggle standard target field vs. point tier targets
  var targetField = g('kpi-target-field');
  var pointTargetsField = g('kpi-point-targets-field');
  if (targetField) targetField.style.display = isPoint ? 'none' : '';
  if (pointTargetsField) pointTargetsField.style.display = isPoint ? '' : 'none';
  if (isPoint) {
    var unitField = g('kpi-unit-field');
    var curField = g('kpi-currency-field');
    if (unitField) unitField.style.display = 'none';
    if (curField) curField.style.display = 'none';
  } else {
    // When restoring volume mode, ensure a valid KPI type chip is active
    if (!g('kpi-chip-' + kpiTypeSelected) || kpiTypeSelected === 'Point') {
      kpiTypeSelected = 'Sales';
      $$('.kpi-type-chip').forEach(function(c) { c.classList.remove('active'); });
      var salesChip = g('kpi-chip-Sales');
      if (salesChip) salesChip.classList.add('active');
      setValueMode('currency');
    } else {
      setValueMode(kpiValueMode);
    }
  }
}

function setValueMode(mode) {
  kpiValueMode = mode;
  const unitField = g('kpi-unit-field');
  const curField = g('kpi-currency-field');
  const unitToggle = g('kpi-unit-toggle');
  const curToggle = g('kpi-currency-toggle');

  if (mode === 'unit') {
    if (unitField) unitField.style.display = '';
    if (curField) curField.style.display = 'none';
    if (unitToggle) unitToggle.classList.add('active');
    if (curToggle) curToggle.classList.remove('active');
    populateKpiItemSel();
  } else {
    if (unitField) unitField.style.display = 'none';
    if (curField) curField.style.display = '';
    if (unitToggle) unitToggle.classList.remove('active');
    if (curToggle) curToggle.classList.add('active');
    populateKpiCurrencyItemSel();
  }
}

function populateKpiItemSel() {
  const sel = g('kpi-item-sel');
  if (!sel) return;
  const allItems = itemCatalogue.filter(function(x) { return x.status === 'active' && x.id !== ITEM_ID_REVENUE; });
  sel.innerHTML = '<option value="">All Items</option>' +
    allItems.map(function(item) {
      return '<option value="' + esc(item.id) + '">' + esc(item.name) + '</option>';
    }).join('');
}

function populateKpiCurrencyItemSel() {
  const sel = g('kpi-currency-item-sel');
  if (!sel) return;
  const dollarItems = itemCatalogue.filter(function(x) { return x.status === 'active' && x.group === 'dollar' && x.id !== ITEM_ID_REVENUE; });
  sel.innerHTML = '<option value="">All Dollar Items</option>' +
    dollarItems.map(function(item) {
      return '<option value="' + esc(item.id) + '">' + esc(item.name) + '</option>';
    }).join('');
}

function submitKpi(e) {
  e.preventDefault();
  const editId = rv('kpi-edit-id');

  // Permission check: only admin, cluster, and supervisor can create/edit KPIs
  if (!canManageKpis()) { showAlert('You do not have permission to manage KPIs.', 'error'); return; }

  // Resolve supervisor id from text input (for admin role)
  if (kpiForSelected === 'shop') {
    var shopName = rv('kpi-shop-assignee-name');
    var hiddenEl = g('kpi-shop-assignee');
    var isSupervisorWithAutoFill = currentRole === 'supervisor' && currentUser;
    if (!isSupervisorWithAutoFill && shopName && hiddenEl && !hiddenEl.value) {
      var sups = staffList.filter(function(u) { return u.role === 'Supervisor'; });
      var found = sups.find(function(u) { return u.name.toLowerCase() === shopName.trim().toLowerCase(); });
      if (found) hiddenEl.value = found.id;
    }
  }

  const obj = {
    id: editId || uid(),
    name: rv('kpi-name'),
    kpiMode: kpiModeSelected,
    type: kpiModeSelected === 'point' ? 'Point' : kpiTypeSelected,
    kpiFor: kpiForSelected,
    assigneeId: kpiForSelected === 'shop' ? rv('kpi-shop-assignee') : rv('kpi-agent-assignee'),
    assigneeBranch: kpiForSelected === 'agent' ? rv('kpi-agent-branch') : '',
    target: kpiModeSelected === 'point' ? 0 : (parseFloat(rv('kpi-target')) || 0),
    valueMode: kpiModeSelected === 'point' ? 'unit' : kpiValueMode,
    itemId: (kpiModeSelected !== 'point' && kpiValueMode === 'unit') ? rv('kpi-item-sel') :
            (kpiModeSelected !== 'point' && kpiValueMode === 'currency') ? rv('kpi-currency-item-sel') : '',
    unit: (kpiModeSelected !== 'point' && kpiValueMode === 'unit') ? rv('kpi-unit-val') : '',
    currency: (kpiModeSelected !== 'point' && kpiValueMode === 'currency') ? rv('kpi-currency-sel') : '',
    pointTiers: kpiModeSelected === 'point' ? {
      min:  parseFloat(rv('kpi-point-min'))  || 0,
      gate: parseFloat(rv('kpi-point-gate')) || 0,
      otb:  parseFloat(rv('kpi-point-otb'))  || 0,
      oab:  parseFloat(rv('kpi-point-oab'))  || 0
    } : null,
    period: rv('kpi-period')
  };
  if (!obj.name) { showAlert('Please enter KPI name'); return; }
  if (kpiForSelected === 'shop' && !obj.assigneeId) { showAlert('Please select a supervisor'); return; }
  if (kpiForSelected === 'agent' && !obj.assigneeId) { showAlert('Please select an agent'); return; }
  if (kpiModeSelected !== 'point' && obj.target <= 0) { showAlert('Please enter a target value greater than 0.'); return; }
  if (kpiModeSelected === 'point' && (!obj.pointTiers || obj.pointTiers.gate <= 0)) { showAlert('Please enter at least a Gate target for the point-based KPI.'); return; }
  if (editId) {
    const idx = kpiList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) kpiList[idx] = obj;
  } else {
    kpiList.push(obj);
  }
  closeModal('modal-kpi');
  renderKpiTable();
  syncSheet('KPI', kpiList);
  saveAllData();
  showToast(editId ? 'KPI updated successfully.' : 'KPI added successfully.', 'success');
  // Refresh dashboard KPI section if on dashboard
  if (currentPage === 'dashboard') renderDashboardKpiSection();
}

function editKpi(id) {
  const item = kpiList.find(function(x) { return x.id === id; });
  if (!item) return;
  // Permission check: only admin, cluster, or supervisor (for their own KPIs) can edit
  if (!canManageKpis()) { showAlert('You do not have permission to edit KPIs.', 'error'); return; }
  if (currentRole === 'supervisor' && !canSupervisorModifyKpi(item)) {
    showAlert('You can only edit KPIs assigned to your shop or agents.', 'error'); return;
  }
  openKpiModal(item);
}

function deleteKpi(id) {
  const item = kpiList.find(function(x) { return x.id === id; });
  if (!item) return;
  // Permission check
  if (!canManageKpis()) { showAlert('You do not have permission to delete KPIs.', 'error'); return; }
  if (currentRole === 'supervisor' && !canSupervisorModifyKpi(item)) {
    showAlert('You can only delete KPIs assigned to your shop or agents.', 'error'); return;
  }
  showConfirm('Are you sure you want to delete this KPI? This action cannot be undone.', function() {
    kpiList = kpiList.filter(function(x) { return x.id !== id; });
    renderKpiTable();
    syncSheet('KPI', kpiList);
    saveAllData();
    showToast('KPI deleted.', 'success');
    if (currentPage === 'dashboard') renderDashboardKpiSection();
  }, 'Delete KPI', 'Delete');
}

function setKpiMonth(mode) {
  const picker = g('kpi-month-picker');
  if (mode === 'current') {
    kpiSelectedMonth = ymNow();
  } else if (mode === 'prev') {
    const d = new Date();
    d.setMonth(d.getMonth() - 1);
    kpiSelectedMonth = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  } else {
    kpiSelectedMonth = '';
  }
  if (picker) picker.value = kpiSelectedMonth;
  renderKpiTable();
}

function onKpiMonthChange() {
  const picker = g('kpi-month-picker');
  kpiSelectedMonth = picker ? picker.value : '';
  renderKpiTable();
}

function initKpiMonthPicker() {
  if (!kpiSelectedMonth) kpiSelectedMonth = ymNow();
  const picker = g('kpi-month-picker');
  if (picker) picker.value = kpiSelectedMonth;
}

// Returns the branch associated with a KPI entry.
// For agent KPIs the branch is stored directly; for shop KPIs it comes from the assignee supervisor.
function getKpiBranch(k) {
  if (k.kpiFor === 'agent') return k.assigneeBranch || '';
  if (k.kpiFor === 'shop') {
    var sup = staffList.find(function(u) { return u.id === k.assigneeId; });
    return sup ? (sup.branch || '') : '';
  }
  return '';
}

function renderKpiTable() {
  const tbody = g('kpi-table');
  if (!tbody) return;

  // Show/hide KPI add button based on role
  var kpiAddBtn = g('kpi-add-btn');
  if (kpiAddBtn) {
    kpiAddBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster' || currentRole === 'supervisor') ? '' : 'none';
  }

  // Only show KPIs visible to the current user's branch
  var visibleKpis = getVisibleKpis();

  if (!visibleKpis.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-chart-line" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No KPIs defined yet</td></tr>';
    return;
  }

  // Group KPIs by branch
  var branchMap = {};
  var branchOrder = [];
  visibleKpis.forEach(function(k) {
    var branch = getKpiBranch(k) || 'Unassigned';
    if (!branchMap[branch]) {
      branchMap[branch] = [];
      branchOrder.push(branch);
    }
    branchMap[branch].push(k);
  });
  branchOrder.sort();

  var html = '';
  branchOrder.forEach(function(branch) {
    var kpis = branchMap[branch];
    var rowNum = 0;
    html += '<tr><td colspan="10" class="kpi-branch-header">' +
      '<i class="fas fa-building" style="margin-right:6px;"></i>' + esc(branch) +
      ' <span class="kpi-branch-count">(' + kpis.length + ' KPI' + (kpis.length !== 1 ? 's' : '') + ')</span>' +
      '</td></tr>';
    kpis.forEach(function(k) {
      rowNum++;
      var isPoint = k.kpiMode === 'point';
      const typeLabel = k.type || (k.valueMode === 'currency' ? 'Amount' : 'Unit');
      const typePill = typeLabel === 'Sales' ? 'pill-green' : typeLabel === 'Revenue' || typeLabel === 'Amount' ? 'pill-orange' : typeLabel === 'Units' || typeLabel === 'Unit' ? 'pill-blue' : 'pill-purple';
      const unitLabel = getKpiUnitLabel(k);
      var valueDisplay;
      if (isPoint && k.pointTiers) {
        var pt = k.pointTiers;
        valueDisplay = '<small style="color:#555;">MIN:<b>' + (pt.min || 0) + '</b> Gate:<b>' + (pt.gate || 0) + '</b> OTB:<b>' + (pt.otb || 0) + '</b> OAB:<b>' + (pt.oab || 0) + '</b></small>';
      } else {
        valueDisplay = (!isPoint && k.valueMode === 'currency')
          ? fmtMoney(k.target, esc(k.currency) + ' ')
          : k.target + (unitLabel ? ' ' + esc(unitLabel) : '');
      }
      const assignee = staffList.find(function(u) { return u.id === k.assigneeId; });
      const forLabel = k.kpiFor === 'shop' ? '<span class="pill pill-blue"><i class="fas fa-store"></i> Shop</span>' : '<span class="pill pill-orange"><i class="fas fa-user"></i> Agent</span>';
      const assigneeName = assignee ? esc(assignee.name) : (k.assigneeBranch ? esc(k.assigneeBranch) : '—');
      // Compute actual & achievement for selected month
      const ym = kpiSelectedMonth || ymNow();
      const actual = Math.round(calcKpiActual(k, ym) * 100) / 100;
      var gateTarget = (isPoint && k.pointTiers) ? (k.pointTiers.gate || 0) : k.target;
      const pct = gateTarget > 0 ? Math.round(actual / gateTarget * 100) : 0;
      const pctClass = pct >= 100 ? 'pill-green' : pct >= 70 ? 'pill-orange' : 'pill-red';
      const actualDisplay = (!isPoint && k.valueMode === 'currency')
        ? fmtMoney(actual, esc(k.currency) + ' ')
        : actual + (unitLabel ? ' ' + esc(unitLabel) : '');
      const progressBar = '<div style="background:#eee;border-radius:4px;height:6px;width:80px;display:inline-block;vertical-align:middle;margin-right:4px;">' +
        '<div style="background:' + (pct >= 100 ? '#1B7D3D' : pct >= 70 ? '#FF9800' : '#E53935') + ';width:' + Math.min(pct, 100) + '%;height:100%;border-radius:4px;"></div></div>';
      // Determine if the current user can modify this KPI using centralised helpers
      var canModifyKpi = canManageKpis() &&
        (currentRole !== 'supervisor' || canSupervisorModifyKpi(k));
      var actionBtns = canModifyKpi
        ? '<button class="btn-edit" onclick="editKpi(\'' + esc(k.id) + '\')"><i class="fas fa-edit"></i></button> ' +
          '<button class="btn-delete" onclick="deleteKpi(\'' + esc(k.id) + '\')"><i class="fas fa-trash"></i></button>'
        : '<span style="color:#bbb;font-size:.75rem;">View only</span>';
      var modePill = isPoint
        ? '<span class="pill pill-purple"><i class="fas fa-star"></i> Point</span>'
        : '<span class="pill pill-teal"><i class="fas fa-chart-bar"></i> Volume</span>';
      html += '<tr>' +
        '<td>' + rowNum + '</td>' +
        '<td>' + esc(k.name) + '</td>' +
        '<td>' + modePill + '</td>' +
        '<td><span class="pill ' + typePill + '">' + esc(typeLabel) + '</span></td>' +
        '<td>' + forLabel + '<br><small style="color:#888;">' + assigneeName + '</small></td>' +
        '<td>' + valueDisplay + '</td>' +
        '<td>' + actualDisplay + '</td>' +
        '<td>' + progressBar + '<span class="pill ' + pctClass + '" style="font-size:.72rem;">' + pct + '%</span></td>' +
        '<td>' + esc(k.period || '') + '</td>' +
        '<td style="white-space:nowrap;">' + actionBtns + '</td>' +
        '</tr>';
    });
  });
  tbody.innerHTML = html;
}

// ------------------------------------------------------------
// Login / Logout
// ------------------------------------------------------------
function handleLogin(e) {
  e.preventDefault();
  var username = rv('login-username');
  var password = rv('login-password');
  var errEl = g('login-error');
  var btn = g('login-submit-btn');

  console.log('[AUTH] Login attempt started');
  console.log('[AUTH] Username entered:', username);

  if (!username || !password) {
    console.warn('[AUTH] Missing credentials');
    if (errEl) { errEl.textContent = 'Please enter username and password.'; errEl.style.display = ''; }
    return;
  }
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Signing in\u2026'; }
  function doAuth() {
    console.log('[AUTH] Starting authentication check');
    console.log('[AUTH] Total staff records loaded:', staffList.length);
    console.log('[AUTH] Staff usernames available:', staffList.map(function(u) { return u.username; }));

    var user = staffList.find(function(u) {
      var usernameMatch = (u.username || '').toLowerCase() === username.toLowerCase();
      var passwordMatch = u.password === password;
      var statusMatch = (u.status || '').toLowerCase() === 'active';

      if (usernameMatch) {
        console.log('[AUTH] Username match found:', u.username);
        console.log('[AUTH] Password match:', passwordMatch);
        console.log('[AUTH] Status:', u.status, '| Status active:', statusMatch);
      }

      return usernameMatch && passwordMatch && statusMatch;
    });
    if (user) {
      console.log('[AUTH] \u2713 Authentication successful for user:', user.username);
      var roleMap = { 'Admin': 'admin', 'Cluster': 'cluster', 'Supervisor': 'supervisor', 'Agent': 'agent' };
      currentUser = user;
      if (errEl) errEl.style.display = 'none';
      var ls = g('login-screen'); if (ls) ls.style.display = 'none';
      var as = g('app-shell'); if (as) { as.style.display = 'flex'; }
      switchRole(roleMap[user.role] || 'user');
      // switchRole overwrites currentUser with a representative user for demo switching;
      // restore it to the authenticated user so all data filtering uses the correct identity.
      currentUser = user;
      lsSave(LS_KEYS.session, Object.assign({}, user, { loginTime: Date.now() }));
      startSessionTimer();
      var nameEl = g('topbar-name'); if (nameEl) nameEl.textContent = user.name;
      var avatarEl = g('topbar-avatar'); if (avatarEl) avatarEl.textContent = ini(user.name);
      navigateTo('dashboard', null);
      // Auto-refresh all data from Google Sheets on login so users always see the latest records
      refreshAllData();
    } else {
      console.error('[AUTH] \u2717 Authentication failed');
      console.error('[AUTH] No matching user found for username:', username);
      console.error('[AUTH] Available users:', staffList.map(function(u) {
        return { username: u.username, status: u.status, hasPassword: !!u.password };
      }));
      if (errEl) { errEl.textContent = 'Invalid username or password, or account is inactive.'; errEl.style.display = ''; }
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-right-to-bracket"></i> Sign In'; }
    }
  }
  console.log('[AUTH] Fetching staff data from Google Sheets...');
  fetchStaffFromSheet()
    .then(function() {
      console.log('[AUTH] Staff data fetch completed');
      doAuth();
    })
    .catch(function(err) {
      console.error('[AUTH] Staff data fetch error:', err);
      doAuth(); // Still attempt auth with cached data
    });
}

function toggleLoginPwd() {
  var inp = g('login-password');
  var eye = g('login-pwd-eye');
  if (!inp) return;
  if (inp.type === 'password') { inp.type = 'text'; if (eye) eye.className = 'fas fa-eye-slash'; }
  else { inp.type = 'password'; if (eye) eye.className = 'fas fa-eye'; }
}

function handleLogout() {
  showConfirm('Are you sure you want to sign out of Smart 5G Dashboard?', function() {
    clearSessionTimer();
    currentUser = null;
    localStorage.removeItem(LS_KEYS.session);
    var as = g('app-shell'); if (as) as.style.display = 'none';
    var ls = g('login-screen'); if (ls) ls.style.display = 'flex';
    var lf = g('login-form'); if (lf) lf.reset();
    var errEl = g('login-error'); if (errEl) errEl.style.display = 'none';
    var btn = g('login-submit-btn'); if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-right-to-bracket"></i> Sign In'; }
  }, 'Sign Out', 'Sign Out', false);
}

function addNotification(msg) {
  notifications.push({ id: uid(), msg: msg, time: new Date().toLocaleString(), read: false });
  updateNotificationBadge();
}

function updateNotificationBadge() {
  var unread = notifications.filter(function(n) { return !n.read; }).length;
  var badge = g('notif-badge');
  if (badge) {
    badge.textContent = unread > 9 ? '9+' : (unread || '');
    badge.style.display = unread > 0 ? 'flex' : 'none';
  }
}

function toggleNotifications() {
  var panel = g('notif-panel');
  if (!panel) return;
  var isOpen = panel.style.display === 'block';
  panel.style.display = isOpen ? 'none' : 'block';
  if (!isOpen) {
    notifications.forEach(function(n) { n.read = true; });
    updateNotificationBadge();
    renderNotificationPanel();
  }
}

function renderNotificationPanel() {
  var list = g('notif-list');
  if (!list) return;
  if (!notifications.length) {
    list.innerHTML = '<div style="text-align:center;padding:24px;color:#999;font-size:.85rem;"><i class="fas fa-bell-slash" style="font-size:1.5rem;display:block;margin-bottom:8px;opacity:.3;"></i>No notifications yet</div>';
    return;
  }
  list.innerHTML = notifications.slice().reverse().map(function(n) {
    return '<div style="padding:10px 16px;border-bottom:1px solid #F5F5F5;font-size:.82rem;">' +
      '<div style="color:#1A1A2E;font-weight:500;">' + esc(n.msg) + '</div>' +
      '<div style="color:#999;font-size:.72rem;margin-top:2px;">' + esc(n.time) + '</div>' +
    '</div>';
  }).join('');
}

function loginContactSupport() {
  showAlert('Email: ' + SUPPORT_CONTACT.email + '\nPhone: ' + SUPPORT_CONTACT.phone, 'info', 'Contact Support');
}

// ------------------------------------------------------------
// Contact Support Modal
// ------------------------------------------------------------
function openContactSupportModal() {
  if (currentUser) {
    var nameEl = g('cs-name');
    var usernameEl = g('cs-username');
    var branchEl = g('cs-branch');
    if (nameEl) nameEl.value = currentUser.name || '';
    if (usernameEl) usernameEl.value = currentUser.username || '';
    if (branchEl) branchEl.value = currentUser.branch || '';
  }
  var msgEl = g('cs-message');
  if (msgEl) msgEl.value = '';
  openModal('modal-contactSupport');
}

function submitContactSupport() {
  var msg = g('cs-message') ? g('cs-message').value.trim() : '';
  if (!msg) { showToast('Please enter a message.', 'error'); return; }
  closeModal('modal-contactSupport');
  showToast('Support request sent successfully!', 'success');
}

// ------------------------------------------------------------
// Compatibility Aliases
// ------------------------------------------------------------
function applySaleFilters() { applyReportFilters(); }
function clearSaleFilters() { clearReportFilters(); }
function loadDashboard() { renderDashboard(); }
function selectKpiMode(mode) { setValueMode(mode); }

function toggleUnitTableView() {
  saleUnitShowAll = !saleUnitShowAll;
  renderSaleTable();
}
function togglePointTableView() {
  salePointShowAll = !salePointShowAll;
  renderSaleTable();
}
function submitKPI(e) { submitKpi(e); }
function switchSaleView(view) { setReportView(view); }
function togglePasswordVisibility(inputId, eyeId) { togglePwd(inputId, eyeId); }
function toggleSidebar() {
  const sidebar = g('sidebar');
  if (sidebar) sidebar.classList.toggle('sidebar-collapsed');
}

// ------------------------------------------------------------
// Social Login
// ------------------------------------------------------------
function loginWithSocial(provider) {
  if (provider === 'telegram') {
    window.open('https://t.me/saray2026123', '_blank', 'noopener,noreferrer');
    return;
  }
  if (provider === 'phone') {
    openModal('modal-phone-login');
    return;
  }
  var names = { facebook: 'Facebook', google: 'Google' };
  showAlert(
    'To register with ' + names[provider] + ', please contact the admin on Telegram:\n\n@saray2026123\n\nThey will help set up your account.',
    'info',
    'Register with ' + names[provider]
  );
}

function submitPhoneLogin() {
  var phone = g('phone-login-number') ? g('phone-login-number').value.trim() : '';
  var name = g('phone-login-name') ? g('phone-login-name').value.trim() : '';
  if (!phone || !name) { showToast('Please fill in all required fields.', 'error'); return; }
  closeModal('modal-phone-login');
  showToast('Request submitted! Please contact @saray2026123 on Telegram to complete setup.', 'success');
}

// ------------------------------------------------------------
// Inventory Management
// ------------------------------------------------------------
var LS_INV_STOCK = 'smart5g_inv_stock';
var LS_INV_REQUESTS = 'smart5g_inv_requests';
var LS_INV_HISTORY = 'smart5g_inv_history';

var invStock = [];       // { id, itemId, itemName, category, inStock, allocated, lastUpdated }
var invRequests = [];    // { id, date, requestedBy, itemId, itemName, qty, purpose, status, reviewedBy, reviewNote }
var invHistory = [];     // { id, date, itemId, itemName, type('in'/'out'), qty, before, after, by, note }
var _reviewAllocId = null;

function loadInvData() {
  try { invStock = JSON.parse(localStorage.getItem(LS_INV_STOCK)) || []; } catch(e) { invStock = []; }
  try { invRequests = JSON.parse(localStorage.getItem(LS_INV_REQUESTS)) || []; } catch(e) { invRequests = []; }
  try { invHistory = JSON.parse(localStorage.getItem(LS_INV_HISTORY)) || []; } catch(e) { invHistory = []; }
}

function saveInvData() {
  localStorage.setItem(LS_INV_STOCK, JSON.stringify(invStock));
  localStorage.setItem(LS_INV_REQUESTS, JSON.stringify(invRequests));
  localStorage.setItem(LS_INV_HISTORY, JSON.stringify(invHistory));
}

function getAvailableQty(itemId) {
  var entry = invStock.find(function(s) { return s.itemId === itemId; });
  if (!entry) return 0;
  return Math.max(0, (entry.inStock || 0) - (entry.allocated || 0));
}

function openInvTab(tab, btn) {
  $$('.inv-tab-content').forEach(function(c) { c.classList.remove('active'); });
  var tc = g('invtab-content-' + tab);
  if (tc) tc.classList.add('active');
  $$('.tab-btn[id^="invtab-"]').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  if (tab === 'requests') renderInvRequestsTable();
  if (tab === 'history') renderInvHistoryTable();
}

function renderInvSaleStock() {
  loadInvData();
  // Update KPI
  var totalIn = invStock.reduce(function(s, x) { return s + (x.inStock || 0); }, 0);
  var totalAlloc = invStock.reduce(function(s, x) { return s + (x.allocated || 0); }, 0);
  var pending = invRequests.filter(function(r) { return r.status === 'pending'; }).length;
  var kvTotal = g('inv-kv-total'); if (kvTotal) kvTotal.textContent = totalIn;
  var kvAlloc = g('inv-kv-allocated'); if (kvAlloc) kvAlloc.textContent = totalAlloc;
  var kvPending = g('inv-kv-pending'); if (kvPending) kvPending.textContent = pending;

  // Role-based UI
  var isSup = currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster';
  var supActions = g('inv-sale-sup-actions'); if (supActions) supActions.style.display = isSup ? '' : 'none';
  var allocBtn = g('inv-allocate-btn'); if (allocBtn) allocBtn.style.display = isSup ? '' : 'none';
  var newReqBtn = g('inv-new-request-btn'); if (newReqBtn) newReqBtn.style.display = isSup ? '' : 'none';
  var actionsCol = g('inv-stock-actions-col'); if (actionsCol) actionsCol.style.display = isSup ? '' : 'none';

  // Render stock table
  var tbody = g('inv-stock-table');
  if (!tbody) return;
  if (!invStock.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-inbox fa-2x" style="display:block;margin-bottom:10px;"></i>No stock items yet.' + (isSup ? ' Click <strong>Add Stock</strong> to get started.' : '') + '</td></tr>';
    return;
  }
  tbody.innerHTML = invStock.map(function(s, idx) {
    var avail = Math.max(0, (s.inStock || 0) - (s.allocated || 0));
    var statusBadge = avail <= 5
      ? '<span class="inv-badge-status inv-badge-low"><i class="fas fa-triangle-exclamation"></i> Low</span>'
      : '<span class="inv-badge-status inv-badge-ok"><i class="fas fa-check"></i> OK</span>';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td style="font-weight:600;">' + esc(s.itemName) + '</td>' +
      '<td>' + esc(s.category || '-') + '</td>' +
      '<td style="font-weight:700;color:#1B7D3D;">' + (s.inStock || 0) + '</td>' +
      '<td style="color:#1565C0;">' + (s.allocated || 0) + '</td>' +
      '<td style="font-weight:700;">' + avail + ' ' + statusBadge + '</td>' +
      '<td style="color:#888;font-size:.8rem;">' + esc(s.lastUpdated || '-') + '</td>' +
      (isSup ? '<td><button class="btn-edit" onclick="openAddStockModal(\'' + esc(s.id) + '\')"><i class="fas fa-edit"></i></button> <button class="btn-delete" onclick="deleteStockItem(\'' + esc(s.id) + '\')"><i class="fas fa-trash"></i></button></td>' : '<td></td>') +
      '</tr>';
  }).join('');
}

function renderInvRequestsTable() {
  loadInvData();
  var filter = g('inv-req-status-filter') ? g('inv-req-status-filter').value : '';
  var isSup = currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster';
  var tbody = g('inv-requests-table');
  if (!tbody) return;
  var data = invRequests.filter(function(r) { return !filter || r.status === filter; });
  if (!data.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-inbox fa-2x" style="display:block;margin-bottom:10px;"></i>No requests found.</td></tr>';
    return;
  }
  tbody.innerHTML = data.map(function(r, idx) {
    var statusClass = { pending: 'inv-badge-pending', approved: 'inv-badge-approved', rejected: 'inv-badge-rejected' }[r.status] || 'inv-badge-pending';
    var statusLabel = r.status ? (r.status.charAt(0).toUpperCase() + r.status.slice(1)) : 'Pending';
    var actions = '';
    if (isSup && r.status === 'pending') {
      actions = '<button class="btn btn-sm btn-primary" onclick="openReviewAllocModal(\'' + esc(r.id) + '\')"><i class="fas fa-clipboard-check"></i> Review</button>';
    }
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + esc(r.date || '-') + '</td>' +
      '<td>' + esc(r.requestedBy || '-') + '</td>' +
      '<td style="font-weight:600;">' + esc(r.itemName || '-') + '</td>' +
      '<td style="font-weight:700;">' + (r.qty || 0) + '</td>' +
      '<td style="color:#888;">' + esc(r.purpose || '-') + '</td>' +
      '<td><span class="inv-badge-status ' + statusClass + '">' + statusLabel + '</span></td>' +
      '<td>' + (actions || '<span style="color:#ccc;font-size:.8rem;">—</span>') + '</td>' +
      '</tr>';
  }).join('');
}

function renderInvHistoryTable() {
  loadInvData();
  var tbody = g('inv-history-table');
  if (!tbody) return;
  var data = invHistory.slice().reverse();
  if (!data.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-inbox fa-2x" style="display:block;margin-bottom:10px;"></i>No history yet.</td></tr>';
    return;
  }
  tbody.innerHTML = data.map(function(h, idx) {
    var typeLabel = h.type === 'in'
      ? '<span class="inv-type-in"><i class="fas fa-arrow-down"></i> IN</span>'
      : '<span class="inv-type-out"><i class="fas fa-arrow-up"></i> OUT</span>';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + esc(h.date || '-') + '</td>' +
      '<td style="font-weight:600;">' + esc(h.itemName || '-') + '</td>' +
      '<td>' + typeLabel + '</td>' +
      '<td style="font-weight:700;">' + (h.qty || 0) + '</td>' +
      '<td>' + (h.before || 0) + '</td>' +
      '<td>' + (h.after || 0) + '</td>' +
      '<td>' + esc(h.by || '-') + '</td>' +
      '<td style="color:#888;font-size:.8rem;">' + esc(h.note || '-') + '</td>' +
      '</tr>';
  }).join('');
}

function openAddStockModal(editId) {
  var form = g('form-addStock');
  if (form) form.reset();
  // Populate item select
  var sel = g('addstock-item');
  if (sel) {
    sel.innerHTML = '<option value="">Select item</option>' +
      itemCatalogue.filter(function(x) { return x.status === 'active'; }).map(function(it) {
        return '<option value="' + esc(it.id) + '">' + esc(it.name) + '</option>';
      }).join('');
  }
  var editIdEl = g('addstock-edit-id');
  if (editIdEl) editIdEl.value = '';
  if (editId) {
    loadInvData();
    var entry = invStock.find(function(s) { return s.id === editId; });
    if (entry) {
      if (editIdEl) editIdEl.value = editId;
      if (sel) sel.value = entry.itemId;
      var qtyEl = g('addstock-qty'); if (qtyEl) qtyEl.value = entry.inStock || '';
      var noteEl = g('addstock-note'); if (noteEl) noteEl.value = entry.note || '';
    }
  }
  openModal('modal-addStock');
}

function submitAddStock(e) {
  e.preventDefault();
  loadInvData();
  var itemId = rv('addstock-item');
  var qty = parseInt(rv('addstock-qty'), 10);
  var note = rv('addstock-note');
  if (!itemId) { showToast('Please select an item.', 'error'); return; }
  if (!qty || qty < 1) { showToast('Please enter a valid quantity.', 'error'); return; }
  var item = itemCatalogue.find(function(x) { return x.id === itemId; });
  if (!item) { showToast('Item not found.', 'error'); return; }
  var editId = rv('addstock-edit-id');
  var now = todayStr();
  var byUser = currentUser ? currentUser.name : currentRole;

  var existing = invStock.find(function(s) { return s.id === editId; }) ||
                 invStock.find(function(s) { return s.itemId === itemId; });

  if (existing && !editId) {
    // Add to existing stock
    var before = existing.inStock || 0;
    existing.inStock = before + qty;
    existing.lastUpdated = now;
    invHistory.push({ id: 'ih' + Date.now(), date: now, itemId: itemId, itemName: item.name, type: 'in', qty: qty, before: before, after: existing.inStock, by: byUser, note: note || 'Stock added' });
  } else if (existing && editId) {
    var before2 = existing.inStock || 0;
    existing.inStock = qty;
    existing.lastUpdated = now;
    invHistory.push({ id: 'ih' + Date.now(), date: now, itemId: itemId, itemName: item.name, type: 'in', qty: qty, before: before2, after: qty, by: byUser, note: note || 'Stock adjusted' });
  } else {
    invStock.push({ id: 'is' + Date.now(), itemId: itemId, itemName: item.name, category: item.category || '-', inStock: qty, allocated: 0, lastUpdated: now });
    invHistory.push({ id: 'ih' + Date.now(), date: now, itemId: itemId, itemName: item.name, type: 'in', qty: qty, before: 0, after: qty, by: byUser, note: note || 'Initial stock' });
  }
  saveInvData();
  closeModal('modal-addStock');
  showToast('Stock updated successfully!', 'success');
  renderInvSaleStock();
}

function deleteStockItem(id) {
  showConfirm('Are you sure you want to delete this stock entry?', function() {
    loadInvData();
    invStock = invStock.filter(function(s) { return s.id !== id; });
    saveInvData();
    showToast('Stock entry deleted.', 'success');
    renderInvSaleStock();
  }, 'Delete Stock Entry', 'Delete', true);
}

function openAllocateModal() {
  var form = g('form-allocate');
  if (form) form.reset();
  loadInvData();
  var sel = g('allocate-item');
  if (sel) {
    sel.innerHTML = '<option value="">Select item</option>' +
      invStock.map(function(s) {
        var avail = Math.max(0, (s.inStock || 0) - (s.allocated || 0));
        return '<option value="' + esc(s.itemId) + '">' + esc(s.itemName) + ' (Available: ' + avail + ')</option>';
      }).join('');
  }
  var infoEl = g('allocate-available-info');
  if (infoEl) infoEl.style.display = 'none';
  openModal('modal-allocate');
}

function updateAllocateAvailable() {
  var itemId = g('allocate-item') ? g('allocate-item').value : '';
  var infoEl = g('allocate-available-info');
  var qtyEl = g('allocate-avail-qty');
  if (!itemId) { if (infoEl) infoEl.style.display = 'none'; return; }
  var avail = getAvailableQty(itemId);
  if (infoEl) infoEl.style.display = '';
  if (qtyEl) qtyEl.textContent = avail;
}

function submitAllocationRequest(e) {
  e.preventDefault();
  loadInvData();
  var itemId = rv('allocate-item');
  var qty = parseInt(rv('allocate-qty'), 10);
  var purpose = rv('allocate-purpose');
  if (!itemId) { showToast('Please select an item.', 'error'); return; }
  if (!qty || qty < 1) { showToast('Please enter a valid quantity.', 'error'); return; }

  // Verify stock exists
  var stockEntry = invStock.find(function(s) { return s.itemId === itemId; });
  if (!stockEntry) { showToast('No stock found for this item.', 'error'); return; }

  // Verify available quantity
  var avail = getAvailableQty(itemId);
  if (qty > avail) {
    showToast('Requested quantity (' + qty + ') exceeds available stock (' + avail + ').', 'error');
    return;
  }

  var now = todayStr();
  var byUser = currentUser ? currentUser.name : currentRole;
  invRequests.push({
    id: 'ir' + Date.now(),
    date: now,
    requestedBy: byUser,
    itemId: itemId,
    itemName: stockEntry.itemName,
    qty: qty,
    purpose: purpose,
    status: 'pending'
  });
  saveInvData();
  closeModal('modal-allocate');
  showToast('Allocation request submitted! Awaiting supervisor approval.', 'success');
  renderInvRequestsTable();
  renderInvSaleStock();
}

function openReviewAllocModal(requestId) {
  loadInvData();
  var req = invRequests.find(function(r) { return r.id === requestId; });
  if (!req) return;
  _reviewAllocId = requestId;
  var avail = getAvailableQty(req.itemId);
  var body = g('review-alloc-body');
  if (body) {
    var match = req.itemName && req.qty;
    body.innerHTML =
      '<div class="form-group">' +
        '<div style="background:#F5F5F5;border-radius:10px;padding:16px;">' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Request ID</span>' +
            '<span style="font-size:.8125rem;font-weight:600;">' + esc(req.id) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Requested By</span>' +
            '<span style="font-size:.8125rem;font-weight:600;">' + esc(req.requestedBy) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Date</span>' +
            '<span style="font-size:.8125rem;">' + esc(req.date) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;align-items:center;">' +
            '<span style="font-size:.8125rem;color:#888;">Item Name</span>' +
            '<span style="font-size:.875rem;font-weight:700;color:#1A1A2E;">' + esc(req.itemName) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;align-items:center;">' +
            '<span style="font-size:.8125rem;color:#888;">Qty Requested</span>' +
            '<span style="font-size:1rem;font-weight:800;color:#1565C0;">' + req.qty + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Available in Stock</span>' +
            '<span style="font-size:.875rem;font-weight:700;color:' + (avail >= req.qty ? '#2E7D32' : '#C62828') + ';">' + avail + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:4px;">' +
            '<span style="font-size:.8125rem;color:#888;">Purpose</span>' +
            '<span style="font-size:.8125rem;max-width:55%;text-align:right;">' + esc(req.purpose) + '</span>' +
          '</div>' +
        '</div>' +
      '</div>' +
      (avail < req.qty ? '<div style="background:#FFEBEE;border-radius:8px;padding:10px 14px;font-size:.8125rem;color:#C62828;margin-top:8px;"><i class="fas fa-triangle-exclamation"></i> <strong>Warning:</strong> Available stock (' + avail + ') is less than requested (' + req.qty + '). Approval will not be possible.</div>' : '') +
      '<div class="form-group" style="margin-top:12px;">' +
        '<label class="form-label">Review Note (optional)</label>' +
        '<input type="text" class="form-input" id="review-note-input" placeholder="e.g., Approved for weekly allocation" />' +
      '</div>';
  }
  // Disable approve if not enough stock
  var approveBtn = g('review-approve-btn');
  if (approveBtn) approveBtn.disabled = avail < req.qty;
  openModal('modal-reviewAlloc');
}

function processAllocation(action) {
  loadInvData();
  var req = invRequests.find(function(r) { return r.id === _reviewAllocId; });
  if (!req) return;
  var reviewNote = g('review-note-input') ? g('review-note-input').value.trim() : '';
  var now = todayStr();
  var byUser = currentUser ? currentUser.name : currentRole;

  if (action === 'approved') {
    // Verify item name and amount one more time
    var stockEntry = invStock.find(function(s) { return s.itemId === req.itemId; });
    if (!stockEntry) { showToast('Stock entry not found. Cannot approve.', 'error'); return; }
    var avail = getAvailableQty(req.itemId);
    if (req.qty > avail) {
      showToast('Insufficient stock (' + avail + ' available, ' + req.qty + ' requested). Cannot approve.', 'error');
      return;
    }
    // Auto-calculate: deduct from available (add to allocated)
    var before = stockEntry.inStock || 0;
    stockEntry.allocated = (stockEntry.allocated || 0) + req.qty;
    invHistory.push({
      id: 'ih' + Date.now(),
      date: now,
      itemId: req.itemId,
      itemName: req.itemName,
      type: 'out',
      qty: req.qty,
      before: before,
      after: before,  // inStock doesn't change, allocated increases
      by: byUser,
      note: 'Allocated: ' + (reviewNote || req.purpose)
    });
    req.status = 'approved';
    req.reviewedBy = byUser;
    req.reviewNote = reviewNote;
    req.reviewedAt = now;
    saveInvData();
    closeModal('modal-reviewAlloc');
    showToast('Allocation approved! ' + req.qty + ' units of ' + req.itemName + ' allocated.', 'success');
    showApprovalFormModal('allocation', req);
  } else {
    req.status = 'rejected';
    req.reviewedBy = byUser;
    req.reviewNote = reviewNote;
    saveInvData();
    closeModal('modal-reviewAlloc');
    showToast('Allocation request rejected.', 'info');
  }
  _reviewAllocId = null;
  renderInvRequestsTable();
  renderInvSaleStock();
}

function renderShopStockTable() {
  loadInvData();
  var shopFilterEl = g('shop-stock-branch-filter');
  // Auto-apply branch filter for supervisor and agent roles
  if (shopFilterEl && currentUser && (currentRole === 'supervisor' || currentRole === 'agent') && !shopFilterEl.value) {
    shopFilterEl.value = currentUser.branch || '';
  }
  var branchFilter = shopFilterEl ? shopFilterEl.value : '';
  var tbody = g('shop-stock-table');
  if (!tbody) return;

  // Build shop stock from approved allocations
  var approved = invRequests.filter(function(r) {
    return r.status === 'approved';
  });

  if (!approved.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-store fa-2x" style="display:block;margin-bottom:10px;"></i>No allocated stock for shops yet.</td></tr>';
    return;
  }

  tbody.innerHTML = approved.filter(function(r) {
    return !branchFilter || (r.purpose && r.purpose.toLowerCase().indexOf(branchFilter.toLowerCase()) !== -1);
  }).map(function(r, idx) {
    var stockEntry = invStock.find(function(s) { return s.itemId === r.itemId; });
    var remaining = r.qty;  // simplified: show full allocated qty
    var statusBadge = '<span class="inv-badge-status inv-badge-approved">Active</span>';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + esc(r.purpose || '-') + '</td>' +
      '<td style="font-weight:600;">' + esc(r.itemName) + '</td>' +
      '<td style="font-weight:700;color:#1565C0;">' + (r.qty || 0) + '</td>' +
      '<td>0</td>' +
      '<td style="font-weight:700;color:#1B7D3D;">' + remaining + '</td>' +
      '<td>' + esc(r.date || '-') + '</td>' +
      '<td>' + statusBadge + '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// Approval Form — PDF & Email
// ------------------------------------------------------------

function initApprovalSignaturePad() {
  var oldCanvas = g('approval-sig-canvas');
  if (!oldCanvas || !oldCanvas.parentNode) return;
  // Replace canvas with a fresh clone to remove any stale event listeners
  var canvas = oldCanvas.cloneNode(false);
  oldCanvas.parentNode.replaceChild(canvas, oldCanvas);
  _sigCanvas = canvas;
  _sigCtx = canvas.getContext('2d');
  _sigCtx.strokeStyle = '#1A1A2E';
  _sigCtx.lineWidth = 2;
  _sigCtx.lineCap = 'round';
  _sigCtx.lineJoin = 'round';

  function getPos(e) {
    var rect = canvas.getBoundingClientRect();
    var scaleX = canvas.width / rect.width;
    var scaleY = canvas.height / rect.height;
    var src = e.touches ? e.touches[0] : e;
    return { x: (src.clientX - rect.left) * scaleX, y: (src.clientY - rect.top) * scaleY };
  }

  canvas.addEventListener('mousedown', function(e) { _sigDrawing = true; var p = getPos(e); _sigCtx.beginPath(); _sigCtx.moveTo(p.x, p.y); });
  canvas.addEventListener('mousemove', function(e) { if (!_sigDrawing) return; var p = getPos(e); _sigCtx.lineTo(p.x, p.y); _sigCtx.stroke(); });
  canvas.addEventListener('mouseup', function() { _sigDrawing = false; });
  canvas.addEventListener('mouseleave', function() { _sigDrawing = false; });
  canvas.addEventListener('touchstart', function(e) { e.preventDefault(); _sigDrawing = true; var p = getPos(e); _sigCtx.beginPath(); _sigCtx.moveTo(p.x, p.y); }, { passive: false });
  canvas.addEventListener('touchmove', function(e) { e.preventDefault(); if (!_sigDrawing) return; var p = getPos(e); _sigCtx.lineTo(p.x, p.y); _sigCtx.stroke(); }, { passive: false });
  canvas.addEventListener('touchend', function() { _sigDrawing = false; });
}

function clearApprovalSignature() {
  if (_sigCtx && _sigCanvas) {
    _sigCtx.clearRect(0, 0, _sigCanvas.width, _sigCanvas.height);
  }
}

function showApprovalFormModal(type, data) {
  _approvalFormData = { type: type, data: data };
  var titleEl = g('approval-form-title');
  var contentEl = g('approval-form-content');
  var submitterEl = g('approval-submitter-name');
  var approverEl = g('approval-approver-name');

  var v = _getApprovalFormVars(type, data);
  if (titleEl) titleEl.innerHTML = '<i class="fas fa-file-contract" style="color:#1B7D3D;margin-right:8px;"></i>' + v.formTitle;

  var detailRows = '';
  if (v.isDeposit) {
    var dCreditAmt = _depCreditAmt(data);
    var totalAmt = (Number(data.cash) || 0) + khrToUsd(data.riel || 0) + dCreditAmt;
    detailRows =
      '<tr><td colspan="2" style="padding:6px 12px;background:#e8f5e9;font-weight:700;font-size:.78rem;color:#1B7D3D;letter-spacing:.6px;border-bottom:1px solid #c8e6c9;text-transform:uppercase;">Revenue Detail</td></tr>' +
      '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;width:45%;">Cash ($)</td>' +
        '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:700;color:#1B7D3D;">$' + Number(data.cash || 0).toFixed(2) + '</td></tr>' +
      '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Cash KHR (៛)</td>' +
        '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:700;color:#8B6914;">' + formatKHR(data.riel || 0) + '</td></tr>' +
      '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Credit ($)</td>' +
        '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:700;color:#1565C0;">$' + Number(dCreditAmt || 0).toFixed(2) + '</td></tr>' +
      '<tr><td style="padding:8px 12px;font-weight:700;color:#fff;background:#1B7D3D;border-bottom:1px solid #4caf50;">Total Amount</td>' +
        '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:800;color:#E65100;font-size:1.05rem;">$' + totalAmt.toFixed(2) + '</td></tr>';
  } else {
    detailRows =
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;width:45%;">Item Name</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;">' + esc(data.itemName || '') + '</td></tr>' +
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Quantity Allocated</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;color:#1565C0;">' + (data.qty || 0) + '</td></tr>' +
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Purpose</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;">' + esc(data.purpose || '') + '</td></tr>';
  }

  if (contentEl) {
    contentEl.innerHTML =
      '<div style="border:2px solid #1B7D3D;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(27,125,61,.10);">' +
        '<div style="background:linear-gradient(135deg,#1B7D3D 0%,#2e9d58 100%);color:#fff;padding:16px 20px;text-align:center;">' +
          '<div style="font-size:1.15rem;font-weight:800;letter-spacing:.8px;text-transform:uppercase;">SMART 5G DASHBOARD</div>' +
          '<div style="font-size:.85rem;opacity:.9;margin-top:4px;font-weight:500;">' + v.formTitle.toUpperCase() + '</div>' +
          (v.isDeposit ? '<div style="margin-top:8px;"><span style="background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.5);border-radius:20px;padding:3px 14px;font-size:.78rem;font-weight:700;letter-spacing:.5px;">✓ APPROVED</span></div>' : '') +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;">' +
          '<tr><td colspan="2" style="padding:6px 12px;background:#f0f4f8;font-weight:700;font-size:.78rem;color:#555;letter-spacing:.6px;border-bottom:1px solid #dde3ea;text-transform:uppercase;">Submission Info</td></tr>' +
          '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;width:45%;">Submitter Name</td>' +
            '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:600;">' + esc(v.submitter) + '</td></tr>' +
          (v.branch ? '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Branch</td><td style="padding:8px 12px;border-bottom:1px solid #eee;">' + esc(v.branch) + '</td></tr>' : '') +
          '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Date Submitted</td>' +
            '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + esc(v.dateSubmitted) + '</td></tr>' +
          detailRows +
          (v.remark ? '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Remark / Note</td><td style="padding:8px 12px;border-bottom:1px solid #eee;font-style:italic;color:#666;">' + esc(v.remark) + '</td></tr>' : '') +
          '<tr><td colspan="2" style="padding:6px 12px;background:#f0f4f8;font-weight:700;font-size:.78rem;color:#555;letter-spacing:.6px;border-bottom:1px solid #dde3ea;border-top:1px solid #dde3ea;text-transform:uppercase;">Approval Info</td></tr>' +
          '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Approved By</td>' +
            '<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:700;color:#1B7D3D;">' + esc(v.approver) + '</td></tr>' +
          '<tr><td style="padding:8px 12px;font-weight:600;color:#555;background:#f5f7fa;">Date Approved</td>' +
            '<td style="padding:8px 12px;font-weight:600;color:#333;">' + esc(v.dateApproved) + '</td></tr>' +
        '</table>' +
      '</div>';
  }

  if (submitterEl) submitterEl.textContent = v.submitter;
  if (approverEl) approverEl.textContent = v.approver;

  // Clear previous signature
  clearApprovalSignature();
  openModal('modal-approvalForm');

  // Init signature pad (after modal is visible)
  setTimeout(initApprovalSignaturePad, 50);
}

function _escHtml(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function _getApprovalFormVars(type, data) {
  var isDeposit = (type === 'deposit');
  return {
    isDeposit: isDeposit,
    formTitle: isDeposit ? 'Revenue Deposit Approval' : 'Stock Allocation Approval',
    submitter: isDeposit ? (data.agent || '') : (data.requestedBy || ''),
    approver:  isDeposit ? (data.approvedBy || '') : (data.reviewedBy || ''),
    dateSubmitted: isDeposit ? toLocalDateStr(data.date || (data.submittedAt ? data.submittedAt.split('T')[0] : '')) : (data.date || ''),
    dateApproved:  isDeposit ? toLocalDateStr(data.approvedAt || '') : (data.reviewedAt || ''),
    branch: isDeposit ? (data.branch || '') : '',
    remark: isDeposit ? (data.remark || '') : (data.reviewNote || data.purpose || '')
  };
}

function printApprovalForm() {
  if (!_approvalFormData) return;
  var v = _getApprovalFormVars(_approvalFormData.type, _approvalFormData.data);
  var data = _approvalFormData.data;
  var sigDataUrl = (_sigCanvas && _sigCanvas.getContext) ? _sigCanvas.toDataURL('image/png') : '';

  var detailHtml = '';
  if (v.isDeposit) {
    var dCreditAmtP = _depCreditAmt(data);
    var totalAmtP = (Number(data.cash) || 0) + khrToUsd(data.riel || 0) + dCreditAmtP;
    detailHtml =
      '<tr><td colspan="2" class="lbl" style="background:#e8f5e9;color:#1B7D3D;font-size:.78rem;letter-spacing:.5px;text-transform:uppercase;">Revenue Detail</td></tr>' +
      '<tr><td class="lbl">Cash ($)</td><td class="val" style="color:#1B7D3D;">$' + Number(data.cash || 0).toFixed(2) + '</td></tr>' +
      '<tr><td class="lbl">Cash KHR (&#x17DB;)</td><td class="val" style="color:#8B6914;">' + formatKHR(data.riel || 0) + '</td></tr>' +
      '<tr><td class="lbl">Credit ($)</td><td class="val" style="color:#1565C0;">$' + Number(dCreditAmtP || 0).toFixed(2) + '</td></tr>' +
      '<tr><td class="lbl" style="background:#1B7D3D;color:#fff;font-weight:700;">Total Amount</td><td class="val" style="font-weight:800;color:#E65100;font-size:1.05rem;">$' + totalAmtP.toFixed(2) + '</td></tr>';
  } else {
    detailHtml =
      '<tr><td class="lbl">Item Name</td><td class="val">' + _escHtml(data.itemName) + '</td></tr>' +
      '<tr><td class="lbl">Quantity Allocated</td><td class="val" style="color:#1565C0;font-weight:700;">' + _escHtml(String(data.qty || 0)) + '</td></tr>' +
      '<tr><td class="lbl">Purpose</td><td class="val">' + _escHtml(data.purpose) + '</td></tr>';
  }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + _escHtml(v.formTitle) + '</title><style>' +
    'body{font-family:Arial,sans-serif;padding:32px;color:#1A1A2E;font-size:13px;max-width:700px;margin:0 auto;}' +
    '.header{background:linear-gradient(135deg,#1B7D3D 0%,#2e9d58 100%);color:#fff;padding:18px 20px;border-radius:8px 8px 0 0;text-align:center;}' +
    '.header h1{margin:0;font-size:1.15rem;font-weight:800;letter-spacing:.8px;text-transform:uppercase;}' +
    '.header p{margin:4px 0 0;font-size:.8rem;opacity:.9;font-weight:500;}' +
    '.status-badge{display:inline-block;margin-top:10px;background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.5);border-radius:20px;padding:3px 14px;font-size:.78rem;font-weight:700;letter-spacing:.5px;}' +
    'table{width:100%;border-collapse:collapse;border:2px solid #1B7D3D;border-top:none;border-radius:0 0 8px 8px;overflow:hidden;}' +
    'td{padding:9px 14px;border-bottom:1px solid #e8e8e8;}' +
    '.lbl{background:#f5f7fa;font-weight:600;color:#555;width:45%;}' +
    '.val{font-weight:500;}' +
    '.section-hdr td{background:#f0f4f8;font-weight:700;font-size:.78rem;color:#555;letter-spacing:.6px;text-transform:uppercase;padding:6px 14px;border-bottom:1px solid #dde3ea;}' +
    '.sig-section{display:flex;gap:32px;margin-top:28px;}' +
    '.sig-box{flex:1;}' +
    '.sig-box p{font-weight:600;font-size:.85rem;color:#555;margin:0 0 6px;}' +
    '.sig-line{border-bottom:2px solid #333;height:60px;margin-bottom:4px;}' +
    '.sig-name{font-size:.75rem;color:#333;font-style:italic;}' +
    '.sig-img{max-width:100%;height:60px;object-fit:contain;display:block;border:1px solid #ddd;border-radius:4px;background:#fafafa;}' +
    '.footer{margin-top:20px;text-align:center;font-size:.7rem;color:#999;border-top:1px solid #eee;padding-top:10px;}' +
    '@media print{body{padding:16px;} button{display:none!important;} .no-print{display:none!important;}}' +
    '</style></head><body>' +
    '<div class="no-print" style="margin-bottom:16px;text-align:right;">' +
      '<button onclick="window.print()" style="background:#1B7D3D;color:#fff;border:none;padding:9px 18px;border-radius:6px;cursor:pointer;font-size:.875rem;">\uD83D\uDDB6 Print / Save as PDF</button>' +
    '</div>' +
    '<div class="header"><h1>SMART 5G DASHBOARD</h1><p>' + _escHtml(v.formTitle).toUpperCase() + '</p>' +
      (v.isDeposit ? '<div class="status-badge">\u2713 APPROVED</div>' : '') +
    '</div>' +
    '<table>' +
      '<tr class="section-hdr"><td colspan="2">Submission Info</td></tr>' +
      '<tr><td class="lbl">Submitter Name</td><td class="val">' + _escHtml(v.submitter) + '</td></tr>' +
      (v.branch ? '<tr><td class="lbl">Branch</td><td class="val">' + _escHtml(v.branch) + '</td></tr>' : '') +
      '<tr><td class="lbl">Date Submitted</td><td class="val">' + _escHtml(v.dateSubmitted) + '</td></tr>' +
      detailHtml +
      (v.remark ? '<tr><td class="lbl">Remark / Note</td><td class="val" style="font-style:italic;color:#666;">' + _escHtml(v.remark) + '</td></tr>' : '') +
      '<tr class="section-hdr"><td colspan="2">Approval Info</td></tr>' +
      '<tr><td class="lbl">Approved By</td><td class="val" style="color:#1B7D3D;font-weight:700;">' + _escHtml(v.approver) + '</td></tr>' +
      '<tr><td class="lbl">Date Approved</td><td class="val" style="font-weight:600;">' + _escHtml(v.dateApproved) + '</td></tr>' +
    '</table>' +
    '<div class="sig-section">' +
      '<div class="sig-box"><p>Submitter Signature:</p><div class="sig-line"></div><div class="sig-name">' + _escHtml(v.submitter) + '</div></div>' +
      '<div class="sig-box"><p>Supervisor / Approver Signature:</p>' +
        (sigDataUrl ? '<img src="' + _escHtml(sigDataUrl) + '" class="sig-img" alt="Signature" />' : '<div class="sig-line"></div>') +
        '<div class="sig-name">' + _escHtml(v.approver) + '</div></div>' +
    '</div>' +
    '<div class="footer">Generated by Smart 5G Dashboard &bull; ' + new Date().toLocaleString() + '</div>' +
    '</body></html>';

  var win = window.open('', '_blank', 'width=780,height=700,scrollbars=yes');
  if (win) {
    win.document.open();
    win.document.write(html);
    win.document.close();
  }
}

function emailApprovalForm() {
  if (!_approvalFormData) return;
  var v = _getApprovalFormVars(_approvalFormData.type, _approvalFormData.data);
  var data = _approvalFormData.data;

  // Look up submitter's email from staffList by name (case-insensitive, trimmed)
  var submitterNameNorm = (v.submitter || '').trim().toLowerCase();
  var submitterUser = staffList.find(function(u) { return (u.name || '').trim().toLowerCase() === submitterNameNorm; });
  var submitterEmail = (submitterUser && submitterUser.email) ? submitterUser.email : '';
  if (!submitterEmail) {
    showToast('No email on file for ' + (v.submitter || 'submitter') + '. Add an email in Settings \u2192 Permission to auto-fill the recipient.', 'warning');
  }

  var subject = 'Smart 5G \u2014 ' + v.formTitle + ' \u2014 ' + v.dateApproved;

  var bodyLines = [
    'SMART 5G DASHBOARD',
    v.formTitle.toUpperCase(),
    '========================================',
    '',
    'Submitter Name  : ' + v.submitter,
  ];
  if (v.branch) bodyLines.push('Branch          : ' + v.branch);
  bodyLines.push('Date Submitted  : ' + v.dateSubmitted);

  if (v.isDeposit) {
    var dCreditAmtE = _depCreditAmt(data);
    var totalAmtE = (Number(data.cash) || 0) + khrToUsd(data.riel || 0) + dCreditAmtE;
    bodyLines.push('--- Revenue Detail ---');
    bodyLines.push('Cash ($)        : $' + Number(data.cash || 0).toFixed(2));
    bodyLines.push('Cash KHR (\u17DB)   : ' + formatKHR(data.riel || 0));
    bodyLines.push('Credit ($)      : $' + Number(dCreditAmtE || 0).toFixed(2));
    bodyLines.push('Total Amount    : $' + totalAmtE.toFixed(2));
  } else {
    bodyLines.push('Item Name       : ' + (data.itemName || ''));
    bodyLines.push('Quantity        : ' + (data.qty || 0));
    bodyLines.push('Purpose         : ' + (data.purpose || ''));
  }

  if (v.remark) bodyLines.push('Remark / Note   : ' + v.remark);
  bodyLines.push('Approved By     : ' + v.approver);
  bodyLines.push('Date Approved   : ' + v.dateApproved);
  bodyLines.push('');
  bodyLines.push('---');
  bodyLines.push('Note: Please print the approval form (use "Print / Save PDF" button), sign it, and attach to this email for records.');
  bodyLines.push('');
  bodyLines.push('Generated by Smart 5G Dashboard \u2014 ' + new Date().toLocaleString());

  var mailtoLink = 'mailto:' + encodeURIComponent(submitterEmail) + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(bodyLines.join('\r\n'));
  window.open(mailtoLink);
}

// ------------------------------------------------------------
// Init
// ------------------------------------------------------------
document.addEventListener('DOMContentLoaded', function() {
  loadAllData();
  loadInvData();
  var mStart = currentMonthStart();
  var mEnd = todayStr();
  var fromEl = g('sale-date-from'); if (fromEl) fromEl.value = mStart;
  var toEl = g('sale-date-to'); if (toEl) toEl.value = mEnd;
  filteredSales = saleRecords.filter(function(s) { return s.date >= mStart && s.date <= mEnd; });
  populateBranchSelects();
  renderItemChips();
  renderNewCustomerTable();
  renderTopUpTable();
  renderTerminationTable();
  renderStaffTable();
  renderKpiTable();
  renderPromotionCards();
  renderDepositTable();
  updateDepositKpis();
  renderSaleTable();
  updateSaleKpis();

  // Restore login session if user was previously authenticated
  var savedSession = lsLoad(LS_KEYS.session, null);
  if (savedSession && savedSession.username) {
    var sessionAge = savedSession.loginTime ? (Date.now() - savedSession.loginTime) : SESSION_TIMEOUT_MS;
    if (sessionAge >= SESSION_TIMEOUT_MS) {
      // Session has expired – clear it and stay on login screen
      localStorage.removeItem(LS_KEYS.session);
    } else {
      var roleMap = { 'Admin': 'admin', 'Cluster': 'cluster', 'Supervisor': 'supervisor', 'Agent': 'agent' };
      currentUser = savedSession;
      var ls = g('login-screen'); if (ls) ls.style.display = 'none';
      var as = g('app-shell'); if (as) as.style.display = 'flex';
      switchRole(roleMap[savedSession.role] || 'user');
      currentUser = savedSession;
      var nameEl = g('topbar-name'); if (nameEl) nameEl.textContent = savedSession.name;
      var avatarEl = g('topbar-avatar'); if (avatarEl) avatarEl.textContent = ini(savedSession.name);
      navigateTo('dashboard', null);
      startSessionTimer(SESSION_TIMEOUT_MS - sessionAge);
      // Refresh all data from Google Sheets so every role sees up-to-date records after a page reload.
      refreshAllData();
    }
  }

  // Populate shop stock branch filter
  var shopFilter = g('shop-stock-branch-filter');
  if (shopFilter) {
    BRANCHES.forEach(function(b) {
      var opt = document.createElement('option');
      opt.value = b; opt.textContent = b;
      shopFilter.appendChild(opt);
    });
  }
});

document.addEventListener('click', function(e) {
  var panel = g('notif-panel');
  var btn = g('notif-btn');
  if (panel && panel.style.display === 'block') {
    if (!panel.contains(e.target) && (!btn || !btn.contains(e.target))) {
      panel.style.display = 'none';
    }
  }
  resetSessionTimer();
});

// Reset session timer on keyboard and mouse activity so active users are not signed out
['mousemove', 'keydown'].forEach(function(evt) {
  document.addEventListener(evt, function() { resetSessionTimer(); }, { passive: true });
});

// ============================================================
// Coverage Feature
// ============================================================

var COV_TABS = [
  { id: 'smart-home',    field: 'smartHome',    color: '#1B7D3D', mapVar: '_covMapSmartHome' },
  { id: 'smart-home-5g', field: 'smartHome5G',  color: '#1565C0', mapVar: '_covMapSmartHome5G' },
  { id: 'smart-fiber',   field: 'smartFiber',   color: '#FB8C00', mapVar: '_covMapSmartFiber' }
];

var KH_CENTER = [12.5657, 104.9910]; // Cambodia center

function switchCoverageTab(tab) {
  currentCoverageTab = tab;
  $$('.cov-tab-content').forEach(function(c) { c.classList.remove('active'); });
  var tc = g('cov-content-' + tab);
  if (tc) tc.classList.add('active');

  // Invalidate the map so it renders correctly when becoming visible
  setTimeout(function() {
    if (typeof L !== 'undefined') {
      if (tab === 'smart-home' && _covMapSmartHome) { _covMapSmartHome.invalidateSize(); }
      if (tab === 'smart-home-5g' && _covMapSmartHome5G) { _covMapSmartHome5G.invalidateSize(); }
      if (tab === 'smart-fiber' && _covMapSmartFiber) { _covMapSmartFiber.invalidateSize(); }
    }
    renderCoverageMap(tab);
    renderCoverageTable(tab);
  }, 50);
}

// Called from sidebar nav — shows only the selected coverage type (tab header hidden)
function openCoverageFromNav(tab, navItem) {
  navigateTo('coverage', null);
  setActiveSubItem(navItem);
  var tabHeader = document.querySelector('#page-coverage .tab-header');
  if (tabHeader) tabHeader.style.display = 'none';
  $$('#page-coverage .tab-btn').forEach(function(b) { b.classList.remove('active'); });
  var activeBtn = g('cov-tab-' + tab);
  if (activeBtn) activeBtn.classList.add('active');
  switchCoverageTab(tab);
}

// Called from the tab buttons inside the page — shows the tab header
function openCoverageTab(tab, btn) {
  var tabHeader = document.querySelector('#page-coverage .tab-header');
  if (tabHeader) tabHeader.style.display = '';
  switchCoverageTab(tab);
  $$('#page-coverage .tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
}

function initCoveragePage() {
  // Show Add button for admin/cluster only
  var addBtn = g('cov-add-btn');
  if (addBtn) addBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';

  // Show/hide Actions column based on role
  var isAdmin = (currentRole === 'admin' || currentRole === 'cluster');
  $$('.cov-admin-col').forEach(function(el) { el.style.display = isAdmin ? '' : 'none'; });

  // Initialize maps if not already done
  initCoverageMap('smart-home');
  initCoverageMap('smart-home-5g');
  initCoverageMap('smart-fiber');

  // Render tables
  renderCoverageTable('smart-home');
  renderCoverageTable('smart-home-5g');
  renderCoverageTable('smart-fiber');

  // Activate current tab button
  $$('#page-coverage .tab-btn').forEach(function(b) { b.classList.remove('active'); });
  var activeBtn = g('cov-tab-' + currentCoverageTab);
  if (activeBtn) activeBtn.classList.add('active');

  // Show/hide correct tab content
  $$('.cov-tab-content').forEach(function(c) { c.classList.remove('active'); });
  var tc = g('cov-content-' + currentCoverageTab);
  if (tc) tc.classList.add('active');

  setTimeout(function() {
    if (typeof L !== 'undefined') {
      if (_covMapSmartHome) _covMapSmartHome.invalidateSize();
      if (_covMapSmartHome5G) _covMapSmartHome5G.invalidateSize();
      if (_covMapSmartFiber) _covMapSmartFiber.invalidateSize();
    }
    renderCoverageMap('smart-home');
    renderCoverageMap('smart-home-5g');
    renderCoverageMap('smart-fiber');
  }, 100);
}

function initCoverageMap(tab) {
  if (typeof L === 'undefined') return; // Leaflet not loaded
  var mapId = 'cov-map-' + tab;
  var el = g(mapId);
  if (!el) return;

  var tabInfo = COV_TABS.find(function(t) { return t.id === tab; });
  if (!tabInfo) return;

  // Destroy existing map instance if any
  if (tab === 'smart-home' && _covMapSmartHome) { _covMapSmartHome.remove(); _covMapSmartHome = null; }
  if (tab === 'smart-home-5g' && _covMapSmartHome5G) { _covMapSmartHome5G.remove(); _covMapSmartHome5G = null; }
  if (tab === 'smart-fiber' && _covMapSmartFiber) { _covMapSmartFiber.remove(); _covMapSmartFiber = null; }

  var map = L.map(mapId).setView(KH_CENTER, 7);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 18,
    attribution: '© OpenStreetMap contributors'
  }).addTo(map);

  if (tab === 'smart-home') _covMapSmartHome = map;
  if (tab === 'smart-home-5g') _covMapSmartHome5G = map;
  if (tab === 'smart-fiber') _covMapSmartFiber = map;
}

function getMapForTab(tab) {
  if (tab === 'smart-home') return _covMapSmartHome;
  if (tab === 'smart-home-5g') return _covMapSmartHome5G;
  if (tab === 'smart-fiber') return _covMapSmartFiber;
  return null;
}

function renderCoverageMap(tab) {
  var tabInfo = COV_TABS.find(function(t) { return t.id === tab; });
  if (!tabInfo) return;

  // Update count badge regardless of Leaflet availability
  var locs = coverageLocations.filter(function(loc) { return loc[tabInfo.field]; });
  var badge = g('cov-count-' + tab);
  if (badge) badge.textContent = locs.length + ' location' + (locs.length !== 1 ? 's' : '');

  if (typeof L === 'undefined') return; // Leaflet not loaded
  var map = getMapForTab(tab);
  if (!map) return;

  // Remove existing layers (except tile layer)
  map.eachLayer(function(layer) {
    if (layer instanceof L.CircleMarker || layer instanceof L.Marker) {
      map.removeLayer(layer);
    }
  });

  locs.forEach(function(loc) {
    if (!loc.lat || !loc.lng) return;
    var circle = L.circleMarker([loc.lat, loc.lng], {
      radius: 10,
      fillColor: tabInfo.color,
      color: '#fff',
      weight: 2,
      opacity: 1,
      fillOpacity: 0.85
    }).addTo(map);
    circle.bindPopup(
      '<strong>' + esc(loc.commune) + '</strong><br/>' +
      '<small>' + esc(loc.district) + ', ' + esc(loc.province) + '</small>'
    );
  });

}

function renderCoverageTable(tab) {
  var tbody = g('cov-table-' + tab);
  if (!tbody) return;

  var tabInfo = COV_TABS.find(function(t) { return t.id === tab; });
  if (!tabInfo) return;

  var searchEl = g('cov-search-' + tab);
  var q = searchEl ? searchEl.value.trim().toLowerCase() : '';

  var locs = coverageLocations.filter(function(loc) { return loc[tabInfo.field]; });
  if (q) {
    locs = locs.filter(function(loc) {
      return (loc.commune || '').toLowerCase().includes(q) ||
             (loc.district || '').toLowerCase().includes(q) ||
             (loc.province || '').toLowerCase().includes(q);
    });
  }

  var isAdmin = (currentRole === 'admin' || currentRole === 'cluster');

  if (!locs.length) {
    tbody.innerHTML = '<tr><td colspan="' + (isAdmin ? 6 : 5) + '" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-map" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No coverage locations found</td></tr>';
    return;
  }

  tbody.innerHTML = locs.map(function(loc, i) {
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td>' + esc(loc.commune) + '</td>' +
      '<td>' + esc(loc.district) + '</td>' +
      '<td>' + esc(loc.province) + '</td>' +
      '<td>' + esc(loc.date || '') + '</td>' +
      (isAdmin ? '<td><button class="btn-edit" onclick="editCoverageLocation(\'' + loc.id + '\')" title="Edit"><i class="fas fa-pen"></i></button> <button class="btn-delete" onclick="deleteCoverageLocation(\'' + loc.id + '\')" title="Delete"><i class="fas fa-trash-can"></i></button></td>' : '<td style="display:none;"></td>') +
      '</tr>';
  }).join('');
}

function openAddCoverageModal() {
  var f = g('form-coverage');
  if (f) f.reset();
  var titleEl = g('modal-coverage-title');
  if (titleEl) titleEl.innerHTML = '<i class="fas fa-map-location-dot" style="color:#1B7D3D;margin-right:8px;"></i>Add Coverage Location';
  if (f) f.removeAttribute('data-edit-id');

  openModal('modal-coverage');

  setTimeout(function() {
    if (typeof L === 'undefined') return; // Leaflet not loaded
    if (_covPickerMap) { _covPickerMap.remove(); _covPickerMap = null; _covPickerMarker = null; _covPickerHighlight = null; }
    var pickerEl = g('cov-picker-map');
    if (!pickerEl) return;
    _covPickerMap = L.map('cov-picker-map').setView(KH_CENTER, 7);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 18,
      attribution: '© OpenStreetMap contributors'
    }).addTo(_covPickerMap);
    _covPickerMap.on('click', function(e) {
      var lat = parseFloat(e.latlng.lat.toFixed(6));
      var lng = parseFloat(e.latlng.lng.toFixed(6));
      var latEl = g('cov-lat'); var lngEl = g('cov-lng');
      if (latEl) latEl.value = lat;
      if (lngEl) lngEl.value = lng;
      if (_covPickerMarker) _covPickerMap.removeLayer(_covPickerMarker);
      _covPickerMarker = L.marker([lat, lng]).addTo(_covPickerMap);
    });
  }, 200);
}

function editCoverageLocation(id) {
  var loc = coverageLocations.find(function(l) { return l.id === id; });
  if (!loc) return;

  var titleEl = g('modal-coverage-title');
  if (titleEl) titleEl.innerHTML = '<i class="fas fa-pen" style="color:#1B7D3D;margin-right:8px;"></i>Edit Coverage Location';

  var f = g('form-coverage');
  if (f) f.setAttribute('data-edit-id', id);

  var prov = g('cov-province'); if (prov) prov.value = loc.province || '';
  var dist = g('cov-district'); if (dist) dist.value = loc.district || '';
  var comm = g('cov-commune'); if (comm) comm.value = loc.commune || '';
  var latEl = g('cov-lat'); if (latEl) latEl.value = loc.lat || '';
  var lngEl = g('cov-lng'); if (lngEl) lngEl.value = loc.lng || '';
  var sh = g('cov-prod-smart-home'); if (sh) sh.checked = !!loc.smartHome;
  var sh5g = g('cov-prod-smart-home-5g'); if (sh5g) sh5g.checked = !!loc.smartHome5G;
  var sf = g('cov-prod-smart-fiber'); if (sf) sf.checked = !!loc.smartFiber;

  openModal('modal-coverage');

  setTimeout(function() {
    if (typeof L === 'undefined') return; // Leaflet not loaded
    if (_covPickerMap) { _covPickerMap.remove(); _covPickerMap = null; _covPickerMarker = null; _covPickerHighlight = null; }
    var pickerEl = g('cov-picker-map');
    if (!pickerEl) return;
    var center = (loc.lat && loc.lng) ? [loc.lat, loc.lng] : KH_CENTER;
    var zoom = (loc.lat && loc.lng) ? 14 : 7;
    _covPickerMap = L.map('cov-picker-map').setView(center, zoom);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 18,
      attribution: '© OpenStreetMap contributors'
    }).addTo(_covPickerMap);
    if (loc.lat && loc.lng) {
      _covPickerMarker = L.marker([loc.lat, loc.lng]).addTo(_covPickerMap);
    }
    _covPickerMap.on('click', function(e) {
      var lat = parseFloat(e.latlng.lat.toFixed(6));
      var lng = parseFloat(e.latlng.lng.toFixed(6));
      var latEl = g('cov-lat'); var lngEl = g('cov-lng');
      if (latEl) latEl.value = lat;
      if (lngEl) lngEl.value = lng;
      if (_covPickerMarker) _covPickerMap.removeLayer(_covPickerMarker);
      _covPickerMarker = L.marker([lat, lng]).addTo(_covPickerMap);
    });
  }, 200);
}

function submitCoverageLocation(e) {
  e.preventDefault();
  var province = (g('cov-province').value || '').trim();
  var district = (g('cov-district').value || '').trim();
  var commune  = (g('cov-commune').value || '').trim();
  var lat      = parseFloat(g('cov-lat').value) || null;
  var lng      = parseFloat(g('cov-lng').value) || null;
  var smartHome    = g('cov-prod-smart-home').checked;
  var smartHome5G  = g('cov-prod-smart-home-5g').checked;
  var smartFiber   = g('cov-prod-smart-fiber').checked;

  if (!province || !district || !commune) {
    showToast('Please fill in Province, District, and Commune.', 'error'); return;
  }
  if (!smartHome && !smartHome5G && !smartFiber) {
    showToast('Please select at least one product.', 'error'); return;
  }

  var today = todayStr();
  var f = g('form-coverage');
  var editId = f ? f.getAttribute('data-edit-id') : null;

  if (editId) {
    var idx = coverageLocations.findIndex(function(l) { return l.id === editId; });
    if (idx !== -1) {
      coverageLocations[idx] = Object.assign(coverageLocations[idx], {
        province: province, district: district, commune: commune,
        lat: lat, lng: lng,
        smartHome: smartHome, smartHome5G: smartHome5G, smartFiber: smartFiber,
        date: today
      });
    }
  } else {
    coverageLocations.push({
      id: 'cv_' + Date.now(),
      province: province, district: district, commune: commune,
      lat: lat, lng: lng,
      smartHome: smartHome, smartHome5G: smartHome5G, smartFiber: smartFiber,
      date: today
    });
  }

  lsSave(LS_KEYS.coverage, coverageLocations);
  closeModal('modal-coverage');
  showToast('Coverage location saved.', 'success');

  // Refresh all maps and tables
  renderCoverageMap('smart-home');
  renderCoverageMap('smart-home-5g');
  renderCoverageMap('smart-fiber');
  renderCoverageTable('smart-home');
  renderCoverageTable('smart-home-5g');
  renderCoverageTable('smart-fiber');
}

function deleteCoverageLocation(id) {
  showConfirm('Delete this coverage location?', function() {
    coverageLocations = coverageLocations.filter(function(l) { return l.id !== id; });
    lsSave(LS_KEYS.coverage, coverageLocations);
    showToast('Coverage location deleted.', 'success');
    renderCoverageMap('smart-home');
    renderCoverageMap('smart-home-5g');
    renderCoverageMap('smart-fiber');
    renderCoverageTable('smart-home');
    renderCoverageTable('smart-home-5g');
    renderCoverageTable('smart-fiber');
  });
}

// ------------------------------------------------------------
// Commune Autocomplete (Nominatim)
// ------------------------------------------------------------
function onCommuneInput(input) {
  var q = input.value.trim();
  var dropdown = g('cov-commune-suggestions');
  if (!dropdown) return;

  clearTimeout(_covCommuneDebounce);
  if (q.length < 2) { dropdown.style.display = 'none'; return; }

  _covCommuneDebounce = setTimeout(function() {
    var url = 'https://nominatim.openstreetmap.org/search?q=' +
      encodeURIComponent(q + ' Cambodia') +
      '&countrycodes=kh&format=json&addressdetails=1&limit=6';
    fetch(url)
      .then(function(r) { return r.json(); })
      .then(function(results) {
        if (!results || !results.length) { dropdown.style.display = 'none'; return; }
        dropdown.innerHTML = results.map(function(r, i) {
          var displayName = r.display_name || '';
          return '<div class="cov-suggestion-item" onmousedown="selectCommuneSuggestion(' + i + ')" data-idx="' + i + '">' +
            '<i class="fas fa-location-dot" style="color:#1B7D3D;margin-right:6px;flex-shrink:0;"></i>' +
            '<span>' + esc(displayName) + '</span>' +
            '</div>';
        }).join('');
        dropdown._results = results;
        dropdown.style.display = 'block';
      })
      .catch(function() { dropdown.style.display = 'none'; });
  }, 320);
}

function selectCommuneSuggestion(idx) {
  var dropdown = g('cov-commune-suggestions');
  if (!dropdown || !dropdown._results) return;
  var r = dropdown._results[idx];
  if (!r) return;
  dropdown.style.display = 'none';

  var addr = r.address || {};

  // Fill commune field
  var communeVal = addr.village || addr.commune || addr.suburb || addr.neighbourhood || addr.quarter || r.name || '';
  var communeEl = g('cov-commune');
  if (communeEl) communeEl.value = communeVal;

  // Fill district
  var districtVal = addr.county || addr.city_district || addr.district || addr.town || addr.city || '';
  var districtEl = g('cov-district');
  if (districtEl) districtEl.value = districtVal;

  // Fill province
  var provinceVal = addr.state || addr.province || '';
  var provinceEl = g('cov-province');
  if (provinceEl) provinceEl.value = provinceVal;

  // Fill lat/lng with centre
  var lat = parseFloat(r.lat);
  var lng = parseFloat(r.lon);
  if (!isNaN(lat) && !isNaN(lng)) {
    var latEl = g('cov-lat'); var lngEl = g('cov-lng');
    if (latEl) latEl.value = lat.toFixed(6);
    if (lngEl) lngEl.value = lng.toFixed(6);
  }

  // Highlight on picker map
  if (typeof L !== 'undefined' && _covPickerMap) {
    // Remove old highlight and marker
    if (_covPickerHighlight) { _covPickerMap.removeLayer(_covPickerHighlight); _covPickerHighlight = null; }
    if (_covPickerMarker)    { _covPickerMap.removeLayer(_covPickerMarker);    _covPickerMarker = null; }

    // Draw bounding box highlight
    if (r.boundingbox && r.boundingbox.length === 4) {
      var s = parseFloat(r.boundingbox[0]), n = parseFloat(r.boundingbox[1]);
      var w = parseFloat(r.boundingbox[2]), e = parseFloat(r.boundingbox[3]);
      if (!isNaN(s) && !isNaN(n) && !isNaN(w) && !isNaN(e)) {
        _covPickerHighlight = L.rectangle([[s, w], [n, e]], {
          color: '#1B7D3D', weight: 2, fillColor: '#1B7D3D', fillOpacity: 0.2
        }).addTo(_covPickerMap);
        _covPickerMap.fitBounds([[s, w], [n, e]], { padding: [30, 30] });
      } else if (!isNaN(lat) && !isNaN(lng)) {
        _covPickerMap.setView([lat, lng], 14);
      }
    } else if (!isNaN(lat) && !isNaN(lng)) {
      _covPickerMap.setView([lat, lng], 14);
    }

    // Add centre marker
    if (!isNaN(lat) && !isNaN(lng)) {
      _covPickerMarker = L.marker([lat, lng]).addTo(_covPickerMap);
    }
  }
}

function closeCommuneSuggestions() {
  var dropdown = g('cov-commune-suggestions');
  if (dropdown) dropdown.style.display = 'none';
}

// ------------------------------------------------------------
// Sale Tracking
// ------------------------------------------------------------
function initSaleTracking() {
  if (!stSelectedMonth) stSelectedMonth = ymNow();
  var picker = g('st-month-picker');
  if (picker) picker.value = stSelectedMonth;
  populateStAgentSel();
  renderSaleTracking();
}

function setStMonth(mode) {
  if (mode === 'current') {
    stSelectedMonth = ymNow();
  } else if (mode === 'prev') {
    var base = stSelectedMonth || ymNow();
    var parts = base.split('-');
    var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 2, 1);
    stSelectedMonth = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  }
  var picker = g('st-month-picker');
  if (picker) picker.value = stSelectedMonth;
  renderSaleTracking();
}

function onStMonthChange() {
  var picker = g('st-month-picker');
  stSelectedMonth = picker ? picker.value : ymNow();
  populateStKpiSel();
  renderSaleTracking();
}

function populateStAgentSel() {
  var wrap = g('st-agent-filter-wrap');
  var sel = g('st-agent-sel');
  if (!sel) return;

  if (wrap) wrap.style.display = '';

  // Agents see only themselves — auto-fill their name as read-only
  if (currentRole === 'agent' && currentUser) {
    sel.innerHTML = '<option value="' + esc(currentUser.id) + '">' + esc(currentUser.name) + '</option>';
    sel.value = currentUser.id;
    sel.disabled = true;
    populateStKpiSel();
    return;
  }
  sel.disabled = false;

  var agents;
  if (currentRole === 'supervisor' && currentUser) {
    agents = staffList.filter(function(u) {
      return u.role === 'Agent' && u.branch === currentUser.branch && u.status === 'active';
    });
  } else {
    agents = staffList.filter(function(u) { return u.role === 'Agent' && u.status === 'active'; });
  }

  var cur = sel.value;
  var shopOption = '';
  if (currentRole === 'supervisor' && currentUser) {
    var shopVal = 'shop_' + currentUser.id;
    shopOption = '<option value="' + esc(shopVal) + '"' + (cur === shopVal ? ' selected' : '') + '>\uD83C\uDFEA Shop: ' + esc(currentUser.branch) + '</option>';
  }
  sel.innerHTML = '<option value="">— Select Agent —</option>' + shopOption +
    agents.map(function(a) {
      return '<option value="' + esc(a.id) + '"' + (cur === a.id ? ' selected' : '') + '>' + esc(a.name) + '</option>';
    }).join('');
  if (!cur && agents.length === 1 && !shopOption) sel.value = agents[0].id;
  populateStKpiSel();
}

function onStAgentChange() {
  populateStKpiSel();
  renderSaleTracking();
}

function populateStKpiSel() {
  var sel = g('st-kpi-sel');
  if (!sel) return;
  var agentId = getStAgentId();
  if (!agentId) {
    sel.innerHTML = '<option value="">— Select Agent first —</option>';
    return;
  }
  var isShopView = agentId.indexOf('shop_') === 0;
  var supervisorId = isShopView ? agentId.slice(5) : '';
  // Items shown in the KPI tracking dropdown (must match itemCatalogue IDs):
  // Smart Fiber+ (i3), Smart@Home (i2), SmartNas (i4),
  // Gross Ads (i1), Monthly Upsell (i5), Recharge (i7)
  var trackItemIds = ['i3', 'i2', 'i4', 'i1', 'i5', 'i7'];
  var cur = sel.value;
  sel.innerHTML = '<option value="">— Select KPI —</option>' +
    trackItemIds.map(function(itemId) {
      var item = itemCatalogue.find(function(x) { return x.id === itemId; });
      if (!item) return ''; // skip if item not in catalogue (should not happen)
      var kpi;
      if (isShopView) {
        kpi = kpiList.find(function(k) {
          return k.kpiFor === 'shop' && k.assigneeId === supervisorId && k.itemId === itemId;
        });
      } else {
        kpi = kpiList.find(function(k) {
          return k.kpiFor === 'agent' && k.assigneeId === agentId && k.itemId === itemId;
        });
        // Fallback: use shop-level KPI for this agent's branch
        if (!kpi) {
          var agentUser = staffList.find(function(u) { return u.id === agentId; });
          if (agentUser && agentUser.branch) {
            var branchSup = staffList.find(function(u) { return u.role === 'Supervisor' && u.branch === agentUser.branch; });
            if (branchSup) {
              kpi = kpiList.find(function(k) {
                return k.kpiFor === 'shop' && k.assigneeId === branchSup.id && k.itemId === itemId;
              });
            }
          }
        }
      }
      var lbl = esc(item.name) + (kpi ? ' (' + kpi.target + '/month)' : '');
      return '<option value="' + esc(itemId) + '"' + (cur === itemId ? ' selected' : '') + '>' + lbl + '</option>';
    }).filter(function(s) { return s !== ''; }).join('');
}

function getStAgentId() {
  if (currentRole === 'agent' && currentUser) return currentUser.id;
  var sel = g('st-agent-sel');
  return sel ? sel.value : '';
}

function renderSaleTracking() {
  var tbody = g('st-table');
  var cards = g('st-summary-cards');
  if (!tbody) return;

  var ym = stSelectedMonth || ymNow();
  var agentId = getStAgentId();
  var selectedItemId = rv('st-kpi-sel');

  if (!agentId) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;">' +
      '<i class="fas fa-user" style="font-size:2rem;display:block;margin-bottom:8px;"></i>' +
      'Please select an agent or shop to view tracking.</td></tr>';
    if (cards) cards.innerHTML = '';
    return;
  }

  if (!selectedItemId) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;">' +
      '<i class="fas fa-crosshairs" style="font-size:2rem;display:block;margin-bottom:8px;"></i>' +
      'Please select an item from the KPI dropdown.</td></tr>';
    if (cards) cards.innerHTML = '';
    return;
  }

  var isShopView = agentId.indexOf('shop_') === 0;
  var supervisorId = isShopView ? agentId.slice(5) : '';
  var kpi = null;
  var agentName = '';
  var filterBranch = '';

  if (isShopView) {
    kpi = kpiList.find(function(k) {
      return k.kpiFor === 'shop' && k.assigneeId === supervisorId && k.itemId === selectedItemId;
    });
    var supUser = staffList.find(function(u) { return u.id === supervisorId; });
    filterBranch = supUser ? supUser.branch : '';
  } else {
    kpi = kpiList.find(function(k) {
      return k.kpiFor === 'agent' && k.assigneeId === agentId && k.itemId === selectedItemId;
    });
    // Fallback: use shop-level KPI for this agent's branch
    if (!kpi) {
      var agentUser = staffList.find(function(u) { return u.id === agentId; });
      if (agentUser && agentUser.branch) {
        var branchSup = staffList.find(function(u) { return u.role === 'Supervisor' && u.branch === agentUser.branch; });
        if (branchSup) {
          kpi = kpiList.find(function(k) {
            return k.kpiFor === 'shop' && k.assigneeId === branchSup.id && k.itemId === selectedItemId;
          });
        }
      }
    }
    var agentRec = staffList.find(function(u) { return u.id === agentId; });
    agentName = agentRec ? agentRec.name : '';
  }

  if (!kpi) {
    var itemDef = itemCatalogue.find(function(x) { return x.id === selectedItemId; });
    var itemName = itemDef ? itemDef.name : selectedItemId;
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;">' +
      '<i class="fas fa-crosshairs" style="font-size:2rem;display:block;margin-bottom:8px;"></i>' +
      'No KPI target set for <b>' + esc(itemName) + '</b>. Please add one in <b>KPI Setting</b>.</td></tr>';
    if (cards) cards.innerHTML = '';
    return;
  }

  var unitLabel = getKpiUnitLabel(kpi) || 'Units';
  var rows = calcSaleTrackingRows(agentName, kpi, ym, filterBranch);
  var today = todayStr();

  // Summary stats
  var totalActual = 0;
  rows.forEach(function(r) { if (!r.isFuture) totalActual += (r.actual || 0); });
  var remaining = Math.max(0, kpi.target - totalActual);
  var daysLeft = rows.filter(function(r) { return r.isFuture; }).length;
  var pct = kpi.target > 0 ? Math.round(totalActual / kpi.target * 100) : 0;
  var dailyRequired = daysLeft > 0 ? Math.ceil(remaining / daysLeft) : 0;
  var pctColor = pct >= 100 ? '#1B7D3D' : pct >= 70 ? '#FF9800' : '#E53935';

  // Render summary cards
  if (cards) {
    cards.innerHTML =
      '<div class="kpi-card">' +
        '<div class="kpi-card-icon kpi-icon-purple"><i class="fas fa-crosshairs"></i></div>' +
        '<div class="kpi-card-body"><p class="kpi-label">Monthly Target</p>' +
        '<p class="kpi-value">' + kpi.target + ' <small style="font-size:.75rem;color:#888;">' + esc(unitLabel) + '</small></p></div>' +
      '</div>' +
      '<div class="kpi-card">' +
        '<div class="kpi-card-icon kpi-icon-green"><i class="fas fa-check-circle"></i></div>' +
        '<div class="kpi-card-body"><p class="kpi-label">Achieved</p>' +
        '<p class="kpi-value" style="color:' + pctColor + ';">' + totalActual +
        ' <small style="font-size:.75rem;">(' + pct + '%)</small></p></div>' +
      '</div>' +
      '<div class="kpi-card">' +
        '<div class="kpi-card-icon kpi-icon-orange"><i class="fas fa-hourglass-half"></i></div>' +
        '<div class="kpi-card-body"><p class="kpi-label">Remaining</p>' +
        '<p class="kpi-value">' + remaining + ' <small style="font-size:.75rem;color:#888;">' + esc(unitLabel) + '</small></p></div>' +
      '</div>' +
      '<div class="kpi-card">' +
        '<div class="kpi-card-icon kpi-icon-blue"><i class="fas fa-calendar-days"></i></div>' +
        '<div class="kpi-card-body"><p class="kpi-label">Days Left · Required/Day</p>' +
        '<p class="kpi-value">' + daysLeft + ' days · ' + dailyRequired + '</p></div>' +
      '</div>';
  }

  // Render daily rows
  var weekdays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var html = '';
  rows.forEach(function(r, idx) {
    var isToday = r.date === today;
    var rowStyle = isToday ? ' style="background:#f0f9f4;font-weight:600;"' : '';

    // Carry-over column (carryIn — the carry coming into this day)
    var carryStr;
    if (r.carryIn === 0) {
      carryStr = '<span style="color:#aaa;">0</span>';
    } else if (r.carryIn > 0) {
      carryStr = '<span style="color:#E53935;">+' + r.carryIn + '</span>';
    } else {
      carryStr = '<span style="color:#1B7D3D;">' + r.carryIn + '</span>';
    }

    // Actual column
    var actualStr = r.isFuture ? '<span style="color:#bbb;">—</span>' : String(r.actual);

    // Balance column (carry-out — shortfall or surplus after this day)
    var balStr;
    if (r.isFuture) {
      balStr = '<span style="color:#bbb;">—</span>';
    } else if (r.balance === 0) {
      balStr = '<span style="color:#888;">0</span>';
    } else if (r.balance > 0) {
      balStr = '<span style="color:#E53935;">+' + r.balance + '</span>';
    } else {
      balStr = '<span style="color:#1B7D3D;">' + r.balance + '</span>';
    }

    // Cumulative column
    var cumStr = r.isFuture ? '<span style="color:#bbb;">—</span>' : String(r.cumulative);

    // Status badge
    var statusBadge;
    if (r.isFuture) {
      statusBadge = '<span class="pill pill-gray">Projected</span>';
    } else if (r.actual > r.adjustedTarget) {
      statusBadge = '<span class="pill pill-green">Ahead</span>';
    } else if (r.actual === r.adjustedTarget) {
      statusBadge = '<span class="pill pill-green">On Track</span>';
    } else {
      statusBadge = '<span class="pill pill-red">Behind</span>';
    }

    var dayName = weekdays[new Date(r.date + 'T00:00:00').getDay()];
    var todayBadge = isToday ? ' <span class="pill pill-blue" style="font-size:.65rem;padding:1px 5px;">Today</span>' : '';

    html += '<tr' + rowStyle + '>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + esc(r.date) + ' <small style="color:#888;">' + dayName + '</small>' + todayBadge + '</td>' +
      '<td style="color:#555;">' + r.baseDisplay + '</td>' +
      '<td>' + carryStr + '</td>' +
      '<td><strong>' + r.adjustedTarget + '</strong></td>' +
      '<td>' + actualStr + '</td>' +
      '<td>' + balStr + '</td>' +
      '<td>' + cumStr + '</td>' +
      '<td>' + statusBadge + '</td>' +
      '</tr>';
  });

  if (!html) {
    html = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;">No days found for this month.</td></tr>';
  }
  tbody.innerHTML = html;
}

// Computes day-by-day carry-over tracking rows for a given agent, KPI, and year-month.
// Each row contains: date, baseDisplay, carryIn, adjustedTarget, actual, balance, cumulative, isFuture.
// carry-over logic: if actual < adjustedTarget, the shortfall carries forward to the next day.
// If actual > adjustedTarget, the surplus deducts from the next day's target.
function calcSaleTrackingRows(agentName, kpi, ym, branchFilter) {
  var parts = ym.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();
  var monthlyTarget = kpi.target || 0;
  var baseDaily = monthlyTarget / daysInMonth; // exact float

  var today = todayStr();
  var rows = [];
  var carryIn = 0; // exact carry coming into the current day
  var cumulative = 0;
  var kpiItem = kpi.itemId ? itemCatalogue.find(function(x) { return x.id === kpi.itemId; }) : null;
  // kpiItem is resolved once here (not inside the loop) for efficiency.

  for (var d = 1; d <= daysInMonth; d++) {
    var dateStr = year + '-' + String(month).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    var isFuture = dateStr > today;

    // Adjusted target (exact for carry math; ceil for display so agent always hits an integer)
    var adjustedExact = baseDaily + carryIn;
    var adjustedTarget = adjustedExact > 0 ? Math.ceil(adjustedExact) : 0;

    var actual = 0;
    var carryOut = carryIn; // future days don't update carry

    if (!isFuture) {
      // Sum actual sales for this agent/shop on this date
      saleRecords.forEach(function(s) {
        if (toLocalDateStr(s.date) !== dateStr) return;
        var shouldSkip = branchFilter ? s.branch !== branchFilter : s.agent !== agentName;
        if (shouldSkip) return;
        if (kpi.itemId) {
          // Dollar-group items (e.g. Recharge) are stored in s.dollarItems; unit-group items in s.items.
          if (kpiItem && kpiItem.group === 'dollar') {
            actual += (s.dollarItems && s.dollarItems[kpi.itemId]) || 0;
          } else {
            actual += (s.items && s.items[kpi.itemId]) || 0;
          }
        } else {
          Object.keys(s.items || {}).forEach(function(iid) { actual += s.items[iid]; });
        }
      });
      cumulative += actual;
      // Use exact adjusted target for carry so rounding doesn't drift over the month
      carryOut = adjustedExact - actual;
    }

    rows.push({
      date: dateStr,
      baseDisplay: Math.round(baseDaily * 10) / 10,
      carryIn: Math.round(carryIn * 10) / 10,
      adjustedTarget: adjustedTarget,
      actual: isFuture ? null : actual,
      balance: isFuture ? null : Math.round((adjustedExact - actual) * 10) / 10,
      cumulative: isFuture ? null : cumulative,
      isFuture: isFuture
    });

    carryIn = carryOut;
  }
  return rows;
}

// ============================================================
// Performance Tracking Page
// ============================================================

function getAgentPointsThisMonth(agentName) {
  var ym = ymNow();
  var pts = 0;
  saleRecords.forEach(function(s) {
    if (agentName && s.agent !== agentName) return;
    if (ymOf(s.date) !== ym) return;
    if (s.transactionType === 'point' && s.kpiPoints) {
      pts += parseFloat(s.kpiPoints) || 0;
    }
  });
  return pts;
}

function renderPerformancePage() {
  var isAdmin   = (currentRole === 'admin' || currentRole === 'cluster');
  var isSupervisor = (currentRole === 'supervisor');
  var agentName = (currentUser && !isAdmin && !isSupervisor) ? (currentUser.name || '') : '';

  // Resolve tier thresholds from active point-based KPI setting; blank (null) if not yet configured
  var role = isSupervisor ? 'supervisor' : 'agent';
  var tiers = { min: null, gate: null, otb: null, oab: null }; // default: blank until KPI is set

  // Find a point-based KPI assigned to the current user's shop/branch
  var pointKpi = null;
  if (isAdmin || isSupervisor) {
    // Supervisors/admins: look for a shop-level point KPI for this branch
    if (currentUser) {
      pointKpi = kpiList.find(function(k) {
        if (k.kpiMode !== 'point' || !k.pointTiers) return false;
        if (k.kpiFor === 'shop') {
          var sup = staffList.find(function(u) { return u.id === k.assigneeId; });
          return sup && sup.branch === currentUser.branch;
        }
        return false;
      });
    }
    if (!pointKpi) {
      // Try any point KPI for admins
      pointKpi = kpiList.find(function(k) { return k.kpiMode === 'point' && k.pointTiers; });
    }
  } else if (currentUser) {
    // Agent: look for agent-level point KPI first, then shop-level for their branch
    pointKpi = kpiList.find(function(k) {
      return k.kpiMode === 'point' && k.pointTiers && k.kpiFor === 'agent' && k.assigneeId === currentUser.id;
    });
    if (!pointKpi && currentUser.branch) {
      var branchSup = staffList.find(function(u) { return u.role === 'Supervisor' && u.branch === currentUser.branch; });
      if (branchSup) {
        pointKpi = kpiList.find(function(k) {
          return k.kpiMode === 'point' && k.pointTiers && k.kpiFor === 'shop' && k.assigneeId === branchSup.id;
        });
      }
    }
  }

  if (pointKpi && pointKpi.pointTiers) {
    tiers = {
      min:  pointKpi.pointTiers.min  > 0 ? pointKpi.pointTiers.min  : null,
      gate: pointKpi.pointTiers.gate > 0 ? pointKpi.pointTiers.gate : null,
      otb:  pointKpi.pointTiers.otb  > 0 ? pointKpi.pointTiers.otb  : null,
      oab:  pointKpi.pointTiers.oab  > 0 ? pointKpi.pointTiers.oab  : null
    };
  }

  var totalPts = getAgentPointsThisMonth(agentName);
  var today = new Date();
  var daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  var dayOfMonth = today.getDate();
  var daysLeft = daysInMonth - dayOfMonth;
  var remaining = tiers.gate != null ? Math.max(0, tiers.gate - totalPts) : null;
  var dailyReq = remaining != null ? (daysLeft > 0 ? Math.ceil(remaining / daysLeft) : remaining) : null;
  var tierInfo = getKpiTierLabel(totalPts, role);

  // Update KPI cards
  var totalEl = g('perf-total-points');
  var minEl = g('perf-min-target');
  var gateEl = g('perf-gate-target');
  var otbEl = g('perf-otb-target');
  var oabEl = g('perf-oab-target');
  var dailyEl = g('perf-daily-required');
  var tierEl = g('perf-tier-badge');
  if (totalEl) totalEl.textContent = totalPts % 1 === 0 ? totalPts.toString() : totalPts.toFixed(2);
  if (minEl) minEl.textContent = tiers.min != null ? tiers.min.toLocaleString() : '-';
  if (gateEl) gateEl.textContent = tiers.gate != null ? tiers.gate.toLocaleString() : '-';
  if (otbEl) otbEl.textContent = tiers.otb != null ? tiers.otb.toLocaleString() : '-';
  if (oabEl) oabEl.textContent = tiers.oab != null ? tiers.oab.toLocaleString() : '-';
  if (dailyEl) dailyEl.textContent = dailyReq != null ? (dailyReq > 0 ? dailyReq.toString() : '0') : '-';
  if (tierEl) {
    tierEl.innerHTML = '<span class="tier-badge" style="background:' + tierInfo.color + ';color:#fff;padding:2px 10px;border-radius:12px;font-size:0.85rem;">' +
      tierInfo.label + '</span>';
  }

  // Render tier achievement table
  var tbody = g('perf-tier-tbody');
  if (tbody) {
    var tierRows = [
      { name: 'Min',  pts: tiers.min  },
      { name: 'Gate', pts: tiers.gate },
      { name: 'OTB',  pts: tiers.otb  },
      { name: 'OAB',  pts: tiers.oab  }
    ];
    tbody.innerHTML = tierRows.map(function(t) {
      var pct = t.pts != null && t.pts > 0 ? Math.min(100, totalPts > 0 ? Math.round((totalPts / t.pts) * 100) : 0) : 0;
      var rem = t.pts != null ? Math.max(0, t.pts - totalPts) : null;
      var reached = t.pts != null && t.pts > 0 && totalPts >= t.pts;
      return '<tr>' +
        '<td><strong>' + t.name + '</strong></td>' +
        '<td>' + (t.pts != null ? t.pts.toLocaleString() : '') + '</td>' +
        '<td>' + (totalPts % 1 === 0 ? totalPts.toString() : totalPts.toFixed(2)) + '</td>' +
        '<td>' + (t.pts != null ? '<div class="progress-bar-wrap"><div class="progress-bar-fill" style="width:' + pct + '%;background:' + (reached ? '#4CAF50' : '#1B7D3D') + ';"></div></div> ' + pct + '%' : '') + '</td>' +
        '<td>' + (t.pts != null ? (reached ? '<span style="color:#4CAF50;">✓ Reached</span>' : rem.toFixed(1)) : '') + '</td>' +
        '<td><span class="status-badge ' + (reached ? 'status-active' : 'status-pending') + '">' + (reached ? 'Done' : 'In Progress') + '</span></td>' +
        '</tr>';
    }).join('');
  }

  // Agent Performance Table (supervisors and admins)
  var agentWrap = g('perf-agent-table-wrap');
  if (agentWrap) {
    if (isAdmin || isSupervisor) {
      agentWrap.style.display = '';
      var agentTbody = g('perf-agent-tbody');
      if (agentTbody) {
        // Collect all agents from sale records this month
        var ym = ymNow();
        var agentMap = {};
        saleRecords.forEach(function(s) {
          if (!s.agent) return;
          if (ymOf(s.date) !== ym) return;
          if (isSupervisor && currentUser && s.branch !== currentUser.branch) return;
          if (!agentMap[s.agent]) agentMap[s.agent] = { name: s.agent, branch: s.branch || '', points: 0 };
          if (s.transactionType === 'point' && s.kpiPoints) {
            agentMap[s.agent].points += parseFloat(s.kpiPoints) || 0;
          }
        });
        var agentRows = Object.values(agentMap).sort(function(a, b) { return b.points - a.points; });
        if (agentRows.length === 0) {
          agentTbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#999;">No point-based sales this month</td></tr>';
        } else {
          agentTbody.innerHTML = agentRows.map(function(a) {
            var pct = tiers.gate > 0 ? Math.min(100, a.points > 0 ? Math.round((a.points / tiers.gate) * 100) : 0) : 0;
            var tier = getKpiTierLabel(a.points, 'agent');
            return '<tr>' +
              '<td>' + esc(a.name) + '</td>' +
              '<td>' + esc(a.branch) + '</td>' +
              '<td>' + (a.points % 1 === 0 ? a.points.toString() : a.points.toFixed(2)) + '</td>' +
              '<td>' + tiers.gate.toLocaleString() + '</td>' +
              '<td>' + pct + '%</td>' +
              '<td><span class="tier-badge" style="background:' + tier.color + ';color:#fff;padding:2px 8px;border-radius:10px;font-size:0.78rem;">' + tier.label + '</span></td>' +
              '</tr>';
          }).join('');
        }
      }
    } else {
      agentWrap.style.display = 'none';
    }
  }
}
