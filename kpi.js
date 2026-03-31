// ============================================================
// KPI Tracker — kpi.js
// Point-Based KPI tracking for Agents, Interns,
// Assist Supervisors and Supervisors.
// ============================================================

(function () {
  'use strict';

  // ----------------------------------------------------------
  // LocalStorage Keys
  // ----------------------------------------------------------
  var LS_KPI_SETTINGS = 'kpi_settings';
  var LS_KPI_ENTRIES  = 'kpi_daily_entries';

  // ----------------------------------------------------------
  // Default Activity Definitions
  // ----------------------------------------------------------
  var DEFAULT_ACTIVITIES = [
    // MBB / Pre-Paid (1 pt per USD)
    { id: 'a01', group: 'MBB / Pre-Paid',         name: 'Gross Add',                       pointRate: 1,   countBasis: 'USD amount',                      enabled: true },
    { id: 'a02', group: 'MBB / Pre-Paid',         name: 'Change SIM',                      pointRate: 1,   countBasis: 'USD amount',                      enabled: true },
    { id: 'a03', group: 'MBB / Pre-Paid',         name: 'Pre-Paid Sub Recharge',            pointRate: 1,   countBasis: 'USD amount',                      enabled: true },
    { id: 'a04', group: 'MBB / Pre-Paid',         name: 'SC Selling',                      pointRate: 1,   countBasis: 'USD amount',                      enabled: true },
    // FBB / Home (2 pts per USD)
    { id: 'a05', group: 'FBB / Home',             name: 'Home Internet Gross Add',          pointRate: 2,   countBasis: 'USD amount',                      enabled: true },
    { id: 'a06', group: 'FBB / Home',             name: 'FWBB/FTTx Deposit & Fees',         pointRate: 2,   countBasis: 'USD amount (deposit/signup/install)', enabled: true },
    { id: 'a07', group: 'FBB / Home',             name: 'Refund FWBB & FTTx',              pointRate: 2,   countBasis: 'USD amount',                      enabled: true },
    { id: 'a08', group: 'FBB / Home',             name: 'Home Internet Migration',          pointRate: 2,   countBasis: 'USD amount',                      enabled: true },
    { id: 'a09', group: 'FBB / Home',             name: 'Home Internet Recharge',           pointRate: 2,   countBasis: 'USD amount',                      enabled: true },
    // FBB / SME Connectivity (2 pts per USD)
    { id: 'a10', group: 'FBB / SME Connectivity', name: 'SME New Sub Gross Add',            pointRate: 2,   countBasis: 'USD amount',                      enabled: true },
    { id: 'a11', group: 'FBB / SME Connectivity', name: 'SME Existing Recharge',            pointRate: 2,   countBasis: 'USD amount',                      enabled: true },
    // MBB / ICT Solution (2 pts per USD)
    { id: 'a12', group: 'MBB / ICT Solution',     name: 'ICT Solution',                    pointRate: 2,   countBasis: 'USD amount',                      enabled: true },
    // Other
    { id: 'a13', group: 'Other',                  name: 'Device Handset / Accessory',       pointRate: 0.5, countBasis: 'USD price',                       enabled: true },
    { id: 'a14', group: 'Other',                  name: 'eSIM',                             pointRate: 2,   countBasis: 'Units (count new eSIM)',           enabled: true },
    { id: 'a15', group: 'Other',                  name: 'Smart Nas Download',               pointRate: 2,   countBasis: 'Unique subscribers',              enabled: true }
  ];

  // ----------------------------------------------------------
  // Default Position Targets
  // ----------------------------------------------------------
  var DEFAULT_POSITIONS = [
    {
      id: 'pos1',
      name: 'Agent / Intern',
      minPts:  800,  gatePts: 1000, otbPts: 1100, oabPts: 1300,
      minPay:   25,  gatePay:   50, otbPay:   75, oabPay:  100
    },
    {
      id: 'pos2',
      name: 'Assist Sup / Sup',
      minPts:  5600, gatePts: 7000, otbPts: 7500, oabPts: 8000,
      minPay:    38, gatePay:   75, otbPay:  150, oabPay:  200
    }
  ];

  // ----------------------------------------------------------
  // State
  // ----------------------------------------------------------
  var _settings  = null;  // { activities, positions }
  var _entries   = null;  // array of entry objects
  var _cKpiReport = null; // Chart.js instance for report page

  // Current sub-page: 'entry' | 'report' | 'config'
  var _currentSubPage = 'entry';

  // ----------------------------------------------------------
  // Persistence Helpers
  // ----------------------------------------------------------
  function loadSettings() {
    try {
      var raw = localStorage.getItem(LS_KPI_SETTINGS);
      if (raw) {
        var parsed = JSON.parse(raw);
        // Merge any new default activities added after first save
        var existing = parsed.activities || [];
        DEFAULT_ACTIVITIES.forEach(function (def) {
          if (!existing.find(function (a) { return a.id === def.id; })) {
            existing.push(JSON.parse(JSON.stringify(def)));
          }
        });
        parsed.activities = existing;
        if (!parsed.positions || !parsed.positions.length) {
          parsed.positions = JSON.parse(JSON.stringify(DEFAULT_POSITIONS));
        }
        _settings = parsed;
      }
    } catch (e) {}
    if (!_settings) {
      _settings = {
        activities: JSON.parse(JSON.stringify(DEFAULT_ACTIVITIES)),
        positions:  JSON.parse(JSON.stringify(DEFAULT_POSITIONS))
      };
    }
  }

  function saveSettings() {
    localStorage.setItem(LS_KPI_SETTINGS, JSON.stringify(_settings));
  }

  function loadEntries() {
    try {
      var raw = localStorage.getItem(LS_KPI_ENTRIES);
      if (raw) { _entries = JSON.parse(raw); }
    } catch (e) {}
    if (!Array.isArray(_entries)) { _entries = []; }
  }

  function saveEntries() {
    localStorage.setItem(LS_KPI_ENTRIES, JSON.stringify(_entries));
  }

  // ----------------------------------------------------------
  // Utility
  // ----------------------------------------------------------
  function emptyState(icon, msg) {
    return '<div class="kpit-empty"><i class="fas fa-' + icon + '"></i><p>' + esc(msg) + '</p></div>';
  }

  function todayStr() {
    var d = new Date();
    return d.getFullYear() + '-' +
           String(d.getMonth() + 1).padStart(2, '0') + '-' +
           String(d.getDate()).padStart(2, '0');
  }

  function monthOf(dateStr) { return dateStr ? dateStr.slice(0, 7) : ''; }

  function fmtDate(dateStr) {
    if (!dateStr) return '';
    var parts = dateStr.split('-');
    if (parts.length < 3) return dateStr;
    return parts[2] + '/' + parts[1] + '/' + parts[0];
  }

  function daysInMonth(ym) { // ym = 'YYYY-MM'
    var parts = ym.split('-');
    return new Date(parseInt(parts[0]), parseInt(parts[1]), 0).getDate();
  }

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Escape for HTML attribute values delimited by single quotes
  function escAttr(s) {
    return esc(s).replace(/'/g, '&#39;');
  }

  function g(id) { return document.getElementById(id); }

  // ----------------------------------------------------------
  // Point / Payment Calculation
  // ----------------------------------------------------------
  function calcEntryPoints(activityValues) {
    // activityValues: { activityId: numericAmount, ... }
    var total = 0;
    (_settings.activities || []).forEach(function (act) {
      if (!act.enabled) return;
      var val = parseFloat(activityValues[act.id]) || 0;
      total += val * act.pointRate;
    });
    return Math.round(total * 100) / 100;
  }

  function calcPayment(posId, totalPoints) {
    var pos = (_settings.positions || []).find(function (p) { return p.id === posId; });
    if (!pos) return { payment: 0, tier: 'Below Min', pct: 0 };

    var tiers = [
      { name: 'Min',  pts: pos.minPts,  pay: pos.minPay  },
      { name: 'Gate', pts: pos.gatePts, pay: pos.gatePay },
      { name: 'OTB',  pts: pos.otbPts,  pay: pos.otbPay  },
      { name: 'OAB',  pts: pos.oabPts,  pay: pos.oabPay  }
    ];

    if (totalPoints < pos.minPts) {
      var pct = pos.minPts > 0 ? Math.round(totalPoints / pos.minPts * 100) : 0;
      return { payment: 0, tier: 'Below Min', pct: pct, pos: pos, tiers: tiers };
    }

    // Pro-Rata interpolation between tiers
    var achieved = tiers[0]; // at least Min
    for (var i = 0; i < tiers.length; i++) {
      if (totalPoints >= tiers[i].pts) { achieved = tiers[i]; }
    }

    var nextTier = null;
    for (var j = 0; j < tiers.length; j++) {
      if (tiers[j].pts > totalPoints) { nextTier = tiers[j]; break; }
    }

    var payment = achieved.pay;
    if (nextTier) {
      // Pro-Rata: interpolate between achieved and next tier
      var rangeRatio = (totalPoints - achieved.pts) / (nextTier.pts - achieved.pts);
      payment = achieved.pay + rangeRatio * (nextTier.pay - achieved.pay);
      payment = Math.round(payment * 100) / 100;
    }

    var topPts = pos.oabPts;
    var pct = topPts > 0 ? Math.round(totalPoints / topPts * 100) : 0;

    return { payment: payment, tier: achieved.name, pct: pct, pos: pos, tiers: tiers, nextTier: nextTier, achieved: achieved };
  }

  // ----------------------------------------------------------
  // Init — called by navigateTo
  // ----------------------------------------------------------
  window.initKpiTracker = function () {
    loadSettings();
    loadEntries();
    renderKpiTrackerPage();
  };

  // ----------------------------------------------------------
  // Tab switching
  // ----------------------------------------------------------
  window.kpitSwitchTab = function (tab) {
    _currentSubPage = tab;
    ['entry', 'report', 'config'].forEach(function (t) {
      var btn = g('kpit-tab-' + t);
      var pg  = g('kpit-sub-' + t);
      if (btn) btn.classList.toggle('active', t === tab);
      if (pg)  pg.classList.toggle('active', t === tab);
    });
    if (tab === 'report')  renderKpiReport();
    if (tab === 'config')  renderKpiConfig();
  };

  // ----------------------------------------------------------
  // Render full page scaffold (once per init)
  // ----------------------------------------------------------
  function renderKpiTrackerPage() {
    var wrap = g('kpit-page-wrap');
    if (!wrap) return;

    wrap.innerHTML =
      '<div class="kpit-tabs">' +
        '<button class="kpit-tab active" id="kpit-tab-entry"  onclick="kpitSwitchTab(\'entry\')">' +
          '<i class="fas fa-pen-to-square"></i> Daily Entry</button>' +
        '<button class="kpit-tab" id="kpit-tab-report" onclick="kpitSwitchTab(\'report\')">' +
          '<i class="fas fa-chart-line"></i> My Report</button>' +
        '<button class="kpit-tab" id="kpit-tab-config" onclick="kpitSwitchTab(\'config\')">' +
          '<i class="fas fa-sliders"></i> KPI Settings</button>' +
      '</div>' +
      '<div class="kpit-subpage active" id="kpit-sub-entry"></div>' +
      '<div class="kpit-subpage" id="kpit-sub-report"></div>' +
      '<div class="kpit-subpage" id="kpit-sub-config"></div>';

    renderKpiEntry();
    // Report and config rendered lazily on tab click
  }

  // ----------------------------------------------------------
  // Daily Entry Sub-Page
  // ----------------------------------------------------------
  function renderKpiEntry() {
    var wrap = g('kpit-sub-entry');
    if (!wrap) return;

    var posOptions = (_settings.positions || []).map(function (p) {
      return '<option value="' + esc(p.id) + '">' + esc(p.name) + '</option>';
    }).join('');

    wrap.innerHTML =
      '<div class="kpit-card">' +
        '<div class="kpit-card-header">' +
          '<span class="kpit-card-title"><i class="fas fa-pen-to-square" style="color:#1B7D3D"></i> Daily Sales Entry</span>' +
        '</div>' +

        '<div class="kpit-form-row">' +
          '<div class="kpit-form-group">' +
            '<label for="kpit-date"><i class="fas fa-calendar"></i> Date</label>' +
            '<input type="date" id="kpit-date" value="' + todayStr() + '" />' +
          '</div>' +
          '<div class="kpit-form-group">' +
            '<label for="kpit-position"><i class="fas fa-id-badge"></i> Position</label>' +
            '<select id="kpit-position" onchange="kpitOnPositionChange()">' +
              posOptions +
            '</select>' +
          '</div>' +
          '<div class="kpit-form-group" style="flex:2 1 240px;">' +
            '<label for="kpit-staff-input"><i class="fas fa-user"></i> Staff Name</label>' +
            '<div class="kpit-staff-search">' +
              '<input type="text" id="kpit-staff-input" placeholder="Type or select staff…" autocomplete="off" ' +
                'oninput="kpitFilterStaff(this.value)" onfocus="kpitOpenStaffDrop()" />' +
              '<div class="kpit-staff-dropdown" id="kpit-staff-drop"></div>' +
            '</div>' +
          '</div>' +
        '</div>' +

        '<div class="table-responsive">' +
          '<table class="kpit-activity-table" id="kpit-activity-table"></table>' +
        '</div>' +

        '<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-top:14px;">' +
          '<div style="font-size:.88rem;color:#555;">' +
            'Daily Total: <strong id="kpit-daily-total" style="font-size:1.1rem;color:#1B7D3D;">0 pts</strong>' +
          '</div>' +
          '<button class="btn btn-primary" onclick="kpitSaveEntry()">' +
            '<i class="fas fa-floppy-disk"></i> Save Entry' +
          '</button>' +
        '</div>' +
      '</div>' +

      // Recent entries for same staff
      '<div class="kpit-card" id="kpit-recent-card">' +
        '<div class="kpit-card-header">' +
          '<span class="kpit-card-title"><i class="fas fa-clock-rotate-left" style="color:#888"></i> Recent Entries</span>' +
        '</div>' +
        '<div id="kpit-recent-body">' + emptyState('inbox', 'No entries yet.') + '</div>' +
      '</div>';

    renderActivityInputs();
    renderRecentEntries();
  }

  function renderActivityInputs() {
    var tbl = g('kpit-activity-table');
    if (!tbl) return;
    var acts = (_settings.activities || []).filter(function (a) { return a.enabled; });

    var groups = [];
    acts.forEach(function (a) {
      if (groups.indexOf(a.group) === -1) groups.push(a.group);
    });

    var html = '<thead><tr>' +
      '<th>#</th><th>Activity</th><th>Count Basis</th>' +
      '<th style="text-align:right">Rate</th>' +
      '<th style="text-align:right">Amount / Qty</th>' +
      '<th style="text-align:right">Points</th>' +
    '</tr></thead><tbody>';

    var rowNo = 1;
    groups.forEach(function (grp) {
      html += '<tr class="kpit-group-header"><td colspan="6">' + esc(grp) + '</td></tr>';
      acts.filter(function (a) { return a.group === grp; }).forEach(function (act) {
        html +=
          '<tr>' +
          '<td style="color:#aaa;font-size:.78rem;">' + rowNo++ + '</td>' +
          '<td>' + esc(act.name) + '</td>' +
          '<td style="color:#888;font-size:.78rem;">' + esc(act.countBasis) + '</td>' +
          '<td style="text-align:right;color:#1B7D3D;font-weight:700;">' + act.pointRate + ' <span style="font-size:.72rem;color:#aaa;font-weight:400;">pt/$</span></td>' +
          '<td style="text-align:right;">' +
            '<input type="number" min="0" step="0.01" class="kpit-qty-input" id="kpit-qty-' + act.id + '" ' +
              'placeholder="0" oninput="kpitUpdateDailyTotal()" />' +
          '</td>' +
          '<td class="kpit-pts-cell" id="kpit-pts-' + act.id + '">0' +
            '<span class="kpit-pts-rate">pts</span>' +
          '</td>' +
          '</tr>';
      });
    });

    html += '<tr class="kpit-total-row">' +
      '<td colspan="4" style="font-weight:700;">Daily Total</td>' +
      '<td></td>' +
      '<td class="kpit-pts-cell" id="kpit-daily-total-cell">0 pts</td>' +
    '</tr>';

    html += '</tbody>';
    tbl.innerHTML = html;
  }

  window.kpitUpdateDailyTotal = function () {
    var vals = collectActivityValues();
    var total = calcEntryPoints(vals);
    var el = g('kpit-daily-total');
    var cell = g('kpit-daily-total-cell');
    if (el) el.textContent = total.toLocaleString() + ' pts';
    if (cell) cell.textContent = total.toLocaleString() + ' pts';

    // Update per-activity points
    (_settings.activities || []).filter(function (a) { return a.enabled; }).forEach(function (act) {
      var val = parseFloat(vals[act.id]) || 0;
      var pts = Math.round(val * act.pointRate * 100) / 100;
      var el2 = g('kpit-pts-' + act.id);
      if (el2) el2.innerHTML = pts.toLocaleString() + '<span class="kpit-pts-rate">pts</span>';
    });
  };

  function collectActivityValues() {
    var vals = {};
    (_settings.activities || []).filter(function (a) { return a.enabled; }).forEach(function (act) {
      var inp = g('kpit-qty-' + act.id);
      vals[act.id] = inp ? parseFloat(inp.value) || 0 : 0;
    });
    return vals;
  }

  window.kpitOnPositionChange = function () {
    // nothing extra needed yet; hook for future use
  };

  // Staff search helpers
  window.kpitOpenStaffDrop = function () {
    var inp = g('kpit-staff-input');
    kpitFilterStaff(inp ? inp.value : '');
  };

  window.kpitFilterStaff = function (query) {
    var drop = g('kpit-staff-drop');
    if (!drop) return;

    // Build staff list from existing entries + app staffList if available
    var names = [];
    (_entries || []).forEach(function (e) {
      if (e.staffName && names.indexOf(e.staffName) === -1) names.push(e.staffName);
    });
    if (typeof staffList !== 'undefined' && Array.isArray(staffList)) {
      staffList.forEach(function (s) {
        var n = s.name || s.staffName;
        if (n && names.indexOf(n) === -1) names.push(n);
      });
    }

    var q = (query || '').toLowerCase();
    var filtered = names.filter(function (n) { return n.toLowerCase().indexOf(q) !== -1; });

    if (!filtered.length) { drop.classList.remove('open'); return; }
    drop.innerHTML = filtered.map(function (n) {
      return '<div class="kpit-staff-opt" onclick="kpitSelectStaff(\'' + escAttr(n) + '\')">' +
        '<div class="kpit-staff-avatar">' + esc(n.charAt(0).toUpperCase()) + '</div>' +
        esc(n) +
      '</div>';
    }).join('');
    drop.classList.add('open');
  };

  window.kpitSelectStaff = function (name) {
    var inp = g('kpit-staff-input');
    if (inp) inp.value = name;
    var drop = g('kpit-staff-drop');
    if (drop) drop.classList.remove('open');
    renderRecentEntries();
  };

  // Close dropdown on outside click
  document.addEventListener('click', function (e) {
    var drop = g('kpit-staff-drop');
    if (!drop) return;
    if (!drop.contains(e.target) && e.target.id !== 'kpit-staff-input') {
      drop.classList.remove('open');
    }
  });

  // ----------------------------------------------------------
  // Save Entry
  // ----------------------------------------------------------
  window.kpitSaveEntry = function () {
    var dateEl = g('kpit-date');
    var posEl  = g('kpit-position');
    var staffEl = g('kpit-staff-input');

    var date      = dateEl  ? dateEl.value  : todayStr();
    var posId     = posEl   ? posEl.value   : (_settings.positions[0] || {}).id;
    var staffName = (staffEl ? staffEl.value : '').trim();

    if (!staffName) {
      alert('Please enter a staff name before saving.');
      if (staffEl) staffEl.focus();
      return;
    }

    var vals  = collectActivityValues();
    var total = calcEntryPoints(vals);

    // Check if entry for same date + staff already exists → update
    var existing = _entries.findIndex(function (e) {
      return e.date === date && e.staffName === staffName && e.positionId === posId;
    });

    var entry = {
      id:         (existing >= 0 ? _entries[existing].id : 'ke_' + Date.now()),
      date:       date,
      positionId: posId,
      staffName:  staffName,
      values:     vals,
      totalPts:   total,
      savedAt:    new Date().toISOString()
    };

    if (existing >= 0) {
      _entries[existing] = entry;
    } else {
      _entries.push(entry);
    }
    saveEntries();

    // Show quick feedback
    var saveBtn = document.querySelector('[onclick="kpitSaveEntry()"]');
    if (saveBtn) {
      var orig = saveBtn.innerHTML;
      saveBtn.innerHTML = '<i class="fas fa-check"></i> Saved!';
      saveBtn.style.background = '#2E7D32';
      setTimeout(function () {
        saveBtn.innerHTML = orig;
        saveBtn.style.background = '';
      }, 1500);
    }

    renderRecentEntries();
  };

  // ----------------------------------------------------------
  // Recent Entries (for current staff in entry form)
  // ----------------------------------------------------------
  function renderRecentEntries() {
    var wrap = g('kpit-recent-body');
    if (!wrap) return;

    var staffEl = g('kpit-staff-input');
    var staffName = staffEl ? staffEl.value.trim() : '';

    var relevant = (_entries || [])
      .filter(function (e) { return !staffName || e.staffName === staffName; })
      .sort(function (a, b) { return b.date.localeCompare(a.date); })
      .slice(0, 10);

    if (!relevant.length) {
      wrap.innerHTML = emptyState('inbox', 'No entries yet.');
      return;
    }

    var rows = relevant.map(function (e) {
      var pos = (_settings.positions || []).find(function (p) { return p.id === e.positionId; });
      return '<tr>' +
        '<td class="date-cell">' + esc(fmtDate(e.date)) + '</td>' +
        '<td>' + esc(e.staffName) + '</td>' +
        '<td>' + esc((pos ? pos.name : e.positionId)) + '</td>' +
        '<td class="pts-cell" style="text-align:right;">' + (e.totalPts || 0).toLocaleString() + ' pts</td>' +
        '<td style="text-align:center;">' +
          '<button onclick="kpitDeleteEntry(\'' + escAttr(e.id) + '\')" title="Delete" ' +
            'style="background:none;border:none;color:#e53935;cursor:pointer;font-size:.9rem;">' +
            '<i class="fas fa-trash-can"></i></button>' +
        '</td>' +
      '</tr>';
    }).join('');

    wrap.innerHTML =
      '<div class="table-responsive">' +
      '<table class="kpit-report-table">' +
        '<thead><tr><th>Date</th><th>Staff</th><th>Position</th><th style="text-align:right">Points</th><th></th></tr></thead>' +
        '<tbody>' + rows + '</tbody>' +
      '</table></div>';
  }

  window.kpitDeleteEntry = function (id) {
    if (!confirm('Delete this entry?')) return;
    _entries = _entries.filter(function (e) { return e.id !== id; });
    saveEntries();
    renderRecentEntries();
    if (_currentSubPage === 'report') renderKpiReport();
  };

  // ----------------------------------------------------------
  // Report Sub-Page
  // ----------------------------------------------------------
  function renderKpiReport() {
    var wrap = g('kpit-sub-report');
    if (!wrap) return;

    // Build staff list from entries
    var staffNames = [];
    (_entries || []).forEach(function (e) {
      if (e.staffName && staffNames.indexOf(e.staffName) === -1) staffNames.push(e.staffName);
    });
    if (typeof staffList !== 'undefined' && Array.isArray(staffList)) {
      staffList.forEach(function (s) {
        var n = s.name || s.staffName;
        if (n && staffNames.indexOf(n) === -1) staffNames.push(n);
      });
    }

    var posOptions = (_settings.positions || []).map(function (p) {
      return '<option value="' + esc(p.id) + '">' + esc(p.name) + '</option>';
    }).join('');

    var staffOptions = staffNames.map(function (n) {
      return '<option value="' + esc(n) + '">' + esc(n) + '</option>';
    }).join('');

    // Current YYYY-MM
    var today = todayStr();
    var curYM = today.slice(0, 7);

    wrap.innerHTML =
      '<div class="kpit-card">' +
        '<div class="kpit-card-header">' +
          '<span class="kpit-card-title"><i class="fas fa-chart-line" style="color:#1B7D3D"></i> KPI Report</span>' +
        '</div>' +
        '<div class="kpit-form-row">' +
          '<div class="kpit-form-group">' +
            '<label><i class="fas fa-calendar"></i> Month</label>' +
            '<input type="month" id="kpit-report-month" value="' + curYM + '" onchange="kpitRefreshReport()" />' +
          '</div>' +
          '<div class="kpit-form-group">' +
            '<label><i class="fas fa-user"></i> Staff</label>' +
            '<select id="kpit-report-staff" onchange="kpitRefreshReport()">' +
              '<option value="">-- Select Staff --</option>' +
              staffOptions +
            '</select>' +
          '</div>' +
          '<div class="kpit-form-group">' +
            '<label><i class="fas fa-id-badge"></i> Position</label>' +
            '<select id="kpit-report-pos" onchange="kpitRefreshReport()">' +
              posOptions +
            '</select>' +
          '</div>' +
        '</div>' +
      '</div>' +

      '<div id="kpit-report-content">' +
        emptyState('chart-bar', 'Select a staff member to view their report.') +
      '</div>';
  }

  window.kpitRefreshReport = function () {
    var staffEl = g('kpit-report-staff');
    var monthEl = g('kpit-report-month');
    var posEl   = g('kpit-report-pos');
    if (!staffEl || !monthEl || !posEl) return;

    var staffName = staffEl.value;
    var ym        = monthEl.value;
    var posId     = posEl.value;

    if (!staffName) {
      var rc = g('kpit-report-content');
      if (rc) rc.innerHTML = emptyState('chart-bar', 'Select a staff member to view their report.');
      return;
    }

    buildReportContent(staffName, ym, posId);
  };

  function buildReportContent(staffName, ym, posId) {
    var wrap = g('kpit-report-content');
    if (!wrap) return;

    var entries = (_entries || [])
      .filter(function (e) { return e.staffName === staffName && monthOf(e.date) === ym; })
      .sort(function (a, b) { return a.date.localeCompare(b.date); });

    // Cumulative points per day
    var cumTotal = 0;
    var dailyData = [];
    entries.forEach(function (e) {
      cumTotal += (e.totalPts || 0);
      dailyData.push({ date: e.date, pts: e.totalPts || 0, cum: cumTotal });
    });

    var pos = (_settings.positions || []).find(function (p) { return p.id === posId; });
    var result = pos ? calcPayment(posId, cumTotal) : null;

    var today = todayStr();
    var dim = daysInMonth(ym);
    var todayDay = today.slice(0, 7) === ym ? parseInt(today.slice(8)) : dim;
    var remaining = dim - todayDay;
    if (remaining < 0) remaining = 0;

    // ── Payment highlight ──
    var payHtml = '';
    if (result) {
      var nextInfo = '';
      if (result.nextTier) {
        var needed = result.nextTier.pts - cumTotal;
        nextInfo = needed > 0
          ? needed.toLocaleString() + ' pts needed to reach ' + result.nextTier.name
          : 'Reached ' + result.nextTier.name + '!';
      } else {
        nextInfo = 'Maximum tier reached!';
      }
      payHtml =
        '<div class="kpit-payment-box">' +
          '<div>' +
            '<div class="pay-label">Estimated Incentive Payment</div>' +
            '<div class="pay-amount">$' + (result.payment).toFixed(2) + '</div>' +
            '<div class="pay-tier">Tier: <strong>' + esc(result.tier) + '</strong></div>' +
          '</div>' +
          '<div class="pay-right">' +
            '<div class="pay-pts">' + cumTotal.toLocaleString() + ' / ' + (pos.oabPts).toLocaleString() + ' pts</div>' +
            '<div class="pay-next">' + esc(nextInfo) + '</div>' +
            '<div style="margin-top:6px;font-size:.78rem;opacity:.75;">' + remaining + ' day(s) remaining in month</div>' +
          '</div>' +
        '</div>';
    }

    // ── Tier badges ──
    var tierBadgesHtml = '';
    if (pos) {
      var tierDefs = [
        { key: 'kpit-min',  name: 'Min',  pts: pos.minPts,  pay: pos.minPay  },
        { key: 'kpit-gate', name: 'Gate', pts: pos.gatePts, pay: pos.gatePay },
        { key: 'kpit-otb',  name: 'OTB',  pts: pos.otbPts,  pay: pos.otbPay  },
        { key: 'kpit-oab',  name: 'OAB',  pts: pos.oabPts,  pay: pos.oabPay  }
      ];
      tierBadgesHtml = '<div class="kpit-tier-badges">';
      tierDefs.forEach(function (td) {
        var achieved = cumTotal >= td.pts;
        tierBadgesHtml +=
          '<div class="kpit-tier-badge ' + td.key + (achieved ? ' achieved' : '') + '">' +
            '<span class="tier-name">' + td.name + '</span>' +
            '<span class="tier-pts">' + td.pts.toLocaleString() + ' pts</span>' +
            '<span class="tier-pay">$' + td.pay + ' USD</span>' +
            (achieved ? '<span class="tier-check"><i class="fas fa-circle-check"></i> Achieved</span>' : '') +
          '</div>';
      });
      tierBadgesHtml += '</div>';
    }

    // ── Progress bar ──
    var barHtml = '';
    if (pos) {
      var barPct = pos.oabPts > 0 ? Math.min(cumTotal / pos.oabPts * 100, 100) : 0;
      var markers = [
        { pts: pos.minPts,  label: 'Min',  pct: Math.min(pos.minPts  / pos.oabPts * 100, 100) },
        { pts: pos.gatePts, label: 'Gate', pct: Math.min(pos.gatePts / pos.oabPts * 100, 100) },
        { pts: pos.otbPts,  label: 'OTB',  pct: Math.min(pos.otbPts  / pos.oabPts * 100, 100) },
        { pts: pos.oabPts,  label: 'OAB',  pct: 100 }
      ];
      var markerHtml = markers.map(function (m) {
        return '<div class="kpit-tier-marker" style="left:' + m.pct.toFixed(1) + '%;">' +
          '<span class="kpit-tier-marker-label">' + m.label + '<br>' + m.pts.toLocaleString() + '</span>' +
        '</div>';
      }).join('');

      barHtml =
        '<div class="kpit-tier-bar-wrap">' +
          '<div class="kpit-tier-bar-fill" style="width:' + barPct.toFixed(1) + '%;"></div>' +
          markerHtml +
        '</div>';
    }

    // ── Summary stats ──
    var statsHtml = '';
    if (pos) {
      var dailyAvg = dailyData.length > 0 ? Math.round(cumTotal / dailyData.length) : 0;
      var projEnd  = dailyAvg * dim;
      statsHtml =
        '<div class="kpit-stats-row">' +
          '<div class="kpit-stat-card">' +
            '<div class="stat-label">Total Points</div>' +
            '<div class="stat-value">' + cumTotal.toLocaleString() + '</div>' +
            '<div class="stat-sub">this month</div>' +
          '</div>' +
          '<div class="kpit-stat-card">' +
            '<div class="stat-label">Days Entered</div>' +
            '<div class="stat-value">' + dailyData.length + '</div>' +
            '<div class="stat-sub">of ' + dim + ' days</div>' +
          '</div>' +
          '<div class="kpit-stat-card">' +
            '<div class="stat-label">Daily Average</div>' +
            '<div class="stat-value">' + dailyAvg.toLocaleString() + '</div>' +
            '<div class="stat-sub">pts/day</div>' +
          '</div>' +
          '<div class="kpit-stat-card">' +
            '<div class="stat-label">Projected EOMonth</div>' +
            '<div class="stat-value">' + projEnd.toLocaleString() + '</div>' +
            '<div class="stat-sub">pts (at current pace)</div>' +
          '</div>' +
        '</div>';
    }

    // ── Chart ──
    var chartHtml =
      '<div class="kpit-card">' +
        '<div class="kpit-card-header">' +
          '<span class="kpit-card-title"><i class="fas fa-chart-area" style="color:#1B7D3D"></i> Cumulative Points</span>' +
        '</div>' +
        '<div class="kpit-chart-wrap"><canvas id="kpit-report-chart"></canvas></div>' +
      '</div>';

    // ── Daily table ──
    var tableHtml = '';
    if (entries.length) {
      var acts = (_settings.activities || []).filter(function (a) { return a.enabled; });
      var tableRows = entries.map(function (e, idx) {
        var breakdown = acts.map(function (act) {
          var val = (e.values && e.values[act.id]) ? e.values[act.id] : 0;
          if (!val) return '';
          var pts = Math.round(val * act.pointRate * 100) / 100;
          return '<span style="font-size:.76rem;display:inline-block;margin:1px 3px;">' +
            esc(act.name) + ': <strong>' + pts + ' pts</strong></span>';
        }).filter(Boolean).join('');

        return '<tr>' +
          '<td class="date-cell">' + esc(fmtDate(e.date)) + '</td>' +
          '<td>' + (breakdown || '<span style="color:#aaa;font-size:.78rem;">—</span>') + '</td>' +
          '<td class="pts-cell" style="text-align:right;white-space:nowrap;">' + (e.totalPts || 0).toLocaleString() + ' pts</td>' +
          '<td class="cum-cell">' + (dailyData[idx] ? dailyData[idx].cum.toLocaleString() : 0) + '</td>' +
        '</tr>';
      }).join('');

      tableHtml =
        '<div class="kpit-card">' +
          '<div class="kpit-card-header">' +
            '<span class="kpit-card-title"><i class="fas fa-table-list" style="color:#888"></i> Daily Breakdown</span>' +
          '</div>' +
          '<div class="table-responsive">' +
          '<table class="kpit-report-table">' +
            '<thead><tr><th>Date</th><th>Activity Points</th><th style="text-align:right">Daily Pts</th><th style="text-align:right">Cumulative</th></tr></thead>' +
            '<tbody>' + tableRows + '</tbody>' +
          '</table></div>' +
        '</div>';
    } else {
      tableHtml =
        '<div class="kpit-card">' +
          emptyState('inbox', 'No entries found for this month.') +
        '</div>';
    }

    wrap.innerHTML =
      payHtml +
      '<div class="kpit-card">' +
        '<div class="kpit-card-header">' +
          '<span class="kpit-card-title"><i class="fas fa-trophy" style="color:#FFA000"></i> KPI Tiers — ' +
          esc(staffName) + '</span>' +
        '</div>' +
        tierBadgesHtml +
        barHtml +
        statsHtml +
      '</div>' +
      chartHtml +
      tableHtml;

    // Draw chart after DOM settles
    setTimeout(function () { drawReportChart(dailyData, pos); }, 0);
  }

  function drawReportChart(dailyData, pos) {
    var canvas = g('kpit-report-chart');
    if (!canvas || typeof Chart === 'undefined') return;

    if (_cKpiReport) { try { _cKpiReport.destroy(); } catch (e) {} _cKpiReport = null; }

    if (!dailyData.length) return;

    var labels = dailyData.map(function (d) { return fmtDate(d.date); });
    var cumPts  = dailyData.map(function (d) { return d.cum; });
    var dayPts  = dailyData.map(function (d) { return d.pts; });

    var annotations = {};
    if (pos) {
      [
        { name: 'Min',  pts: pos.minPts,  color: '#FFC107' },
        { name: 'Gate', pts: pos.gatePts, color: '#4CAF50' },
        { name: 'OTB',  pts: pos.otbPts,  color: '#2196F3' },
        { name: 'OAB',  pts: pos.oabPts,  color: '#9C27B0' }
      ].forEach(function (t, i) {
        annotations['tier' + i] = {
          type: 'line',
          yMin: t.pts, yMax: t.pts,
          borderColor: t.color,
          borderWidth: 1.5,
          borderDash: [5, 4],
          label: { content: t.name + ' (' + t.pts.toLocaleString() + ')', enabled: true, position: 'end', font: { size: 10 } }
        };
      });
    }

    _cKpiReport = new Chart(canvas, {
      data: {
        labels: labels,
        datasets: [
          {
            type: 'line',
            label: 'Cumulative Pts',
            data: cumPts,
            borderColor: '#1B7D3D',
            backgroundColor: 'rgba(27,125,61,0.08)',
            fill: true,
            tension: 0.35,
            pointRadius: 4,
            pointBackgroundColor: '#1B7D3D',
            yAxisID: 'y'
          },
          {
            type: 'bar',
            label: 'Daily Pts',
            data: dayPts,
            backgroundColor: 'rgba(27,125,61,0.25)',
            borderColor: 'rgba(27,125,61,0.5)',
            borderWidth: 1,
            yAxisID: 'y'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { display: true, position: 'top', labels: { boxWidth: 12, font: { size: 11 } } },
          annotation: Object.keys(annotations).length ? { annotations: annotations } : {}
        },
        scales: {
          y: { beginAtZero: true, ticks: { font: { size: 10 } }, grid: { color: 'rgba(0,0,0,.06)' } },
          x: { ticks: { font: { size: 10 } }, grid: { display: false } }
        }
      }
    });
  }

  // ----------------------------------------------------------
  // Config / Settings Sub-Page
  // ----------------------------------------------------------
  function renderKpiConfig() {
    var wrap = g('kpit-sub-config');
    if (!wrap) return;

    var actRows = (_settings.activities || []).map(function (act, i) {
      return '<tr>' +
        '<td>' + esc(act.group) + '</td>' +
        '<td>' + esc(act.name) + '</td>' +
        '<td>' + esc(act.countBasis) + '</td>' +
        '<td style="text-align:right;">' +
          '<input type="number" min="0" step="0.01" value="' + act.pointRate + '" ' +
            'onchange="kpitUpdateActivity(' + i + ',\'pointRate\',this.value)" style="width:80px;" />' +
        '</td>' +
        '<td style="text-align:center;">' +
          '<label class="kpit-toggle">' +
            '<input type="checkbox" ' + (act.enabled ? 'checked' : '') + ' ' +
              'onchange="kpitUpdateActivity(' + i + ',\'enabled\',this.checked)" />' +
            '<span class="kpit-toggle-slider"></span>' +
          '</label>' +
        '</td>' +
      '</tr>';
    }).join('');

    var posHtml = (_settings.positions || []).map(function (pos, pi) {
      return '<div class="kpit-card">' +
        '<div class="kpit-card-header"><span class="kpit-card-title">' + esc(pos.name) + '</span></div>' +
        '<div class="kpit-incentive-grid">' +
          '<div>' +
            '<h4 style="margin-bottom:8px;font-size:.82rem;color:#1B7D3D;">Point Targets</h4>' +
            '<table class="kpit-settings-table">' +
              '<thead><tr><th>Tier</th><th style="text-align:right">Points</th></tr></thead><tbody>' +
              '<tr><td>Min</td><td style="text-align:right;">' +
                '<input type="number" min="0" value="' + pos.minPts + '" onchange="kpitUpdatePosition(' + pi + ',\'minPts\',this.value)" /></td></tr>' +
              '<tr><td>Gate</td><td style="text-align:right;">' +
                '<input type="number" min="0" value="' + pos.gatePts + '" onchange="kpitUpdatePosition(' + pi + ',\'gatePts\',this.value)" /></td></tr>' +
              '<tr><td>OTB</td><td style="text-align:right;">' +
                '<input type="number" min="0" value="' + pos.otbPts + '" onchange="kpitUpdatePosition(' + pi + ',\'otbPts\',this.value)" /></td></tr>' +
              '<tr><td>OAB</td><td style="text-align:right;">' +
                '<input type="number" min="0" value="' + pos.oabPts + '" onchange="kpitUpdatePosition(' + pi + ',\'oabPts\',this.value)" /></td></tr>' +
            '</tbody></table>' +
          '</div>' +
          '<div>' +
            '<h4 style="margin-bottom:8px;font-size:.82rem;color:#1B7D3D;">Incentive Payments (USD)</h4>' +
            '<table class="kpit-settings-table">' +
              '<thead><tr><th>Tier</th><th style="text-align:right">Payment ($)</th></tr></thead><tbody>' +
              '<tr><td>Min</td><td style="text-align:right;">' +
                '<input type="number" min="0" step="0.01" value="' + pos.minPay + '" onchange="kpitUpdatePosition(' + pi + ',\'minPay\',this.value)" /></td></tr>' +
              '<tr><td>Gate</td><td style="text-align:right;">' +
                '<input type="number" min="0" step="0.01" value="' + pos.gatePay + '" onchange="kpitUpdatePosition(' + pi + ',\'gatePay\',this.value)" /></td></tr>' +
              '<tr><td>OTB</td><td style="text-align:right;">' +
                '<input type="number" min="0" step="0.01" value="' + pos.otbPay + '" onchange="kpitUpdatePosition(' + pi + ',\'otbPay\',this.value)" /></td></tr>' +
              '<tr><td>OAB</td><td style="text-align:right;">' +
                '<input type="number" min="0" step="0.01" value="' + pos.oabPay + '" onchange="kpitUpdatePosition(' + pi + ',\'oabPay\',this.value)" /></td></tr>' +
            '</tbody></table>' +
          '</div>' +
        '</div>' +
      '</div>';
    }).join('');

    wrap.innerHTML =
      '<div class="kpit-card">' +
        '<div class="kpit-card-header">' +
          '<span class="kpit-card-title"><i class="fas fa-list-check" style="color:#1B7D3D"></i> Activities &amp; Point Rates</span>' +
          '<button class="btn btn-outline btn-sm" onclick="kpitResetSettings()" style="font-size:.78rem;">' +
            '<i class="fas fa-rotate-left"></i> Reset to Default</button>' +
        '</div>' +
        '<div class="table-responsive">' +
          '<table class="kpit-settings-table">' +
            '<thead><tr><th>Group</th><th>Activity</th><th>Count Basis</th>' +
              '<th style="text-align:right">Pt Rate</th><th style="text-align:center">Enabled</th></tr></thead>' +
            '<tbody>' + actRows + '</tbody>' +
          '</table></div>' +
      '</div>' +

      '<div style="margin-bottom:14px;">' +
        '<div class="kpit-card-title" style="margin-bottom:14px;">' +
          '<i class="fas fa-trophy" style="color:#FFA000"></i> Position Targets &amp; Incentives' +
        '</div>' +
        posHtml +
      '</div>';
  }

  window.kpitUpdateActivity = function (idx, field, value) {
    if (!_settings.activities[idx]) return;
    if (field === 'enabled') {
      _settings.activities[idx].enabled = !!value;
    } else if (field === 'pointRate') {
      _settings.activities[idx].pointRate = parseFloat(value) || 0;
    }
    saveSettings();
  };

  window.kpitUpdatePosition = function (idx, field, value) {
    if (!_settings.positions[idx]) return;
    _settings.positions[idx][field] = parseFloat(value) || 0;
    saveSettings();
  };

  window.kpitResetSettings = function () {
    if (!confirm('Reset all KPI settings to default values? (Saved entries will NOT be affected.)')) return;
    _settings = {
      activities: JSON.parse(JSON.stringify(DEFAULT_ACTIVITIES)),
      positions:  JSON.parse(JSON.stringify(DEFAULT_POSITIONS))
    };
    saveSettings();
    renderKpiConfig();
  };

})();
