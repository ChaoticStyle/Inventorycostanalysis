/* app — Inventory Cost Analysis dashboard (New + Used). */
(function () {
  "use strict";
  var Core = window.InventoryCore;

  // NEW state (unchanged shape) + USED sub-state (U) + active mode.
  var state = { mode: "new", snap: null, prior: null, seg: "All", chart: null,
    allWeeks: [], trendMetric: "availUnits", trendChart: null };
  var U = { snap: null, prior: null, allWeeks: [], trendMetric: "uTotalUnits",
    trendChart: null, agingChart: null, bootstrapped: false };

  var METRICS = {
    availUnits: { label: "Available units (A)", get: function (s) { return s.countA; }, money: false },
    availCost: { label: "$ in Available", get: function (s) { return s.statCost.A; }, money: true },
    aoUnits: { label: "Units (A + O)", get: function (s) { return s.countAO; }, money: false },
    onOrder: { label: "$ in On Order", get: function (s) { return s.statCost.O; }, money: true },
    soldAvail: { label: "$ in Sold Available", get: function (s) { return s.statCost.SA; }, money: true },
    totalUnits: { label: "Total units (A+SA+O)", get: function (s) { return s.count; }, money: false },
  };
  var UMETRICS = {
    uTotalUnits: { label: "Total units", get: function (s) { return s.total; }, money: false },
    uTotalCost: { label: "Total cost", get: function (s) { return s.grand; }, money: true },
  };

  // ---------- helpers ----------
  var $ = function (id) { return document.getElementById(id); };
  function money(x) { return "$" + (Math.round(x * 100) / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function money0(x) { return "$" + Math.round(x).toLocaleString("en-US"); }
  function intc(x) { return Number(x).toLocaleString("en-US"); }
  function pct(part, whole) { return whole ? (part / whole * 100).toFixed(1) + "%" : "0.0%"; }
  function signedInt(n) { return (n > 0 ? "+" : n < 0 ? "-" : "") + intc(Math.abs(n)); }
  function signedMoney(x) { return (x >= 0 ? "+" : "-") + money(Math.abs(x)); }
  function signedPct(cur, prev) {
    if (prev === 0) return cur === 0 ? "0%" : "new";
    var p = (cur - prev) / prev * 100;
    return (p >= 0 ? "+" : "-") + Math.abs(p).toFixed(0) + "%";
  }
  function cls(n) { return n > 0 ? "pos" : n < 0 ? "neg" : "flat"; }
  function sumVals(o) { var s = 0; Object.keys(o).forEach(function (k) { s += o[k]; }); return s; }
  function toast(msg) {
    var t = $("toast"); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2600);
  }
  async function api(path, opts) {
    opts = opts || {}; opts.headers = { "content-type": "application/json" };
    var res = await fetch(path, opts); return res.json();
  }
  // append ?kind=used for used-mode calls; new mode uses the default store
  function ep(base, used) {
    if (!used) return base;
    return base + (base.indexOf("?") >= 0 ? "&" : "?") + "kind=used";
  }
  function deltaCell(n, isMoney) {
    if (n === null) return "<td class='cell-delta flat'>—</td>";
    return "<td class='cell-delta " + cls(n) + "'>" + (isMoney ? signedMoney(n) : signedInt(n)) + "</td>";
  }
  function dollar(l, v) { return "<div class='dollar'><div class='l'>" + l + "</div><div class='v'>" + money(v) + "</div></div>"; }

  function sheetGrid(wb) {
    // Pick the FIRST data sheet whose header matches; skip a "Template" tab (the export
    // ships one with the same headers but placeholder rows).
    var chosen = null;
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var nm = wb.SheetNames[i];
      if (String(nm).trim().toLowerCase() === "template") continue;
      var r0 = XLSX.utils.sheet_to_json(wb.Sheets[nm], { header: 1, blankrows: false })[0] || [];
      var up = r0.map(function (c) { return String(c).trim().toUpperCase(); });
      if (up.indexOf("TOTAL COST") >= 0 || up.indexOf("STOCK") >= 0 || up.indexOf("STK.#") >= 0) { chosen = nm; break; }
    }
    if (!chosen) chosen = wb.SheetNames[0];
    return XLSX.utils.sheet_to_json(wb.Sheets[chosen], { header: 1, blankrows: false });
  }

  // =========================================================================
  // NEW inventory
  // =========================================================================
  function readRows(grid) {
    if (!grid.length) return [];
    var hdr = grid[0].map(function (c) { return String(c).trim().toUpperCase(); });
    function col(name, fb) { var i = hdr.indexOf(name); return i >= 0 ? i : fb; }
    var c = { stock: col("STOCK", 0), year: col("YEAR", 1), type: col("TYPE", 4),
      cost: col("TOTAL COST", 6), days: col("DAYS", 9), stat: col("STAT", 11) };
    var out = [];
    for (var i = 1; i < grid.length; i++) {
      var r = grid[i];
      out.push({ stock: r[c.stock], year: r[c.year], type: r[c.type], cost: r[c.cost], days: r[c.days], stat: r[c.stat] });
    }
    return out;
  }

  async function handleFile(file) {
    var wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
    var built = Core.buildSegments(readRows(sheetGrid(wb)));
    var d = Core.parseDateFromName(file.name);
    if (!d) toast("Could not read a date from the filename — expected …_DDMMYYYY…");
    var snap = {
      date: d ? d.iso : new Date().toISOString().slice(0, 10),
      dateDisplay: d ? d.display : "(today)", fileName: file.name,
      meta: { kept: built.kept, excludedSO: built.excludedSO, unclassifiedTypes: built.unclassifiedTypes },
      segments: built.segments,
    };
    state.snap = snap;
    try { await api("/api/snapshots", { method: "POST", body: JSON.stringify(snap) }); toast("Saved New week " + snap.dateDisplay); } catch (e) {}
    state.prior = null;
    try { var res = await api("/api/snapshots?before=" + encodeURIComponent(snap.date)); state.prior = res.snapshot || null; } catch (e) {}
    await loadAllWeeks();
    $("dropzone").hidden = true; $("dash").hidden = false;
    render(); if (state.mode === "new") loadWeeks();
  }

  async function loadAllWeeks() {
    try { var res = await api("/api/snapshots?all=1"); state.allWeeks = res.snapshots || []; }
    catch (e) { state.allWeeks = state.snap ? [state.snap] : []; }
  }

  async function bootstrapNew() {
    await loadAllWeeks();
    if (state.allWeeks.length) {
      state.snap = state.allWeeks[state.allWeeks.length - 1];
      state.prior = state.allWeeks.length > 1 ? state.allWeeks[state.allWeeks.length - 2] : null;
      $("dropzone").hidden = true; $("dash").hidden = false; render();
    }
  }

  function render() {
    if (!state.snap) return;
    var seg = state.snap.segments[state.seg];
    var prev = state.prior ? state.prior.segments[state.seg] : null;
    var s = state.snap, p = state.prior;
    if (state.mode === "new") setDateline(s, p);
    var m = s.meta || { excludedSO: 0, unclassifiedTypes: [] };
    $("metaNote").textContent = (m.excludedSO || 0) + " Sold Order unit(s) excluded"
      + (m.unclassifiedTypes && m.unclassifiedTypes.length ? "  ·  unclassified TYPE: " + m.unclassifiedTypes.join(", ") : "");
    renderKpis(seg, prev); renderComparison(seg, prev, s, p); renderTrend();
    renderYearTable(seg, prev); renderStatusTable(seg); renderAging(seg);
  }

  function setDateline(s, p) {
    $("dateline").textContent = p
      ? "Inventory as of " + s.dateDisplay + "  ·  compared to " + p.dateDisplay
      : "Inventory as of " + s.dateDisplay + "  ·  no prior week yet";
  }

  function renderTrend() {
    var weeks = (state.allWeeks || []).filter(function (w) { return w.segments && w.segments[state.seg]; });
    var metric = METRICS[state.trendMetric];
    lineTrend($("trendChart"), $("trendNote"), weeks.map(function (w) { return w.dateDisplay || w.date; }),
      weeks.map(function (w) { return metric.get(w.segments[state.seg]) || 0; }), metric, "state");
  }

  function renderKpis(seg, prev) {
    var dA = prev ? seg.countA - prev.countA : null, dAvail = prev ? seg.statCost.A - prev.statCost.A : null;
    $("kpis").innerHTML =
      kpiCard("Available units (A)", intc(seg.countA), dA, false) +
      kpiCard("$ in Available", money0(seg.statCost.A), dAvail, true) +
      kpiCard("$ in Sold Available", money0(seg.statCost.SA), null) +
      kpiCard("$ in On Order", money0(seg.statCost.O), null) +
      kpiCard("Total units (A+SA+O)", intc(seg.count), null);
  }
  function kpiCard(label, value, delta, isMoney) {
    var d = "";
    if (delta !== null && delta !== undefined)
      d = "<div class='delta " + cls(delta) + "'>" + (isMoney ? signedMoney(delta) : signedInt(delta)) + " vs prior</div>";
    return "<div class='kpi'><div class='label'>" + label + "</div><div class='value'>" + value + "</div>" + d + "</div>";
  }

  function renderComparison(seg, prev, s, p) {
    $("cmpDates").textContent = p ? (s.dateDisplay + " vs " + p.dateDisplay) : s.dateDisplay + " (no prior)";
    var years = Core.sortedYears(Object.assign({}, seg.yrCountA, prev ? prev.yrCountA : {}));
    var rows = "<thead><tr><th>Model Year</th><th>Avail. Units</th><th>% of Avail.</th><th>Δ Units</th><th>Total Cost</th><th>% Δ</th><th>$ Δ</th></tr></thead><tbody>";
    years.forEach(function (y) {
      var cu = seg.yrCountA[y] || 0, pu = prev ? (prev.yrCountA[y] || 0) : null;
      var cc = seg.yrCostA[y] || 0, pc = prev ? (prev.yrCostA[y] || 0) : 0;
      rows += "<tr><td>" + y + "</td><td>" + intc(cu) + "</td><td>" + pct(cu, seg.countA) + "</td>"
        + deltaCell(prev ? cu - pu : null, false) + "<td>" + money(cc) + "</td>"
        + "<td class='cell-delta " + (prev ? cls(cc - pc) : "flat") + "'>" + (prev ? signedPct(cc, pc) : "—") + "</td>"
        + deltaCell(prev ? cc - pc : null, true) + "</tr>";
    });
    var tc = seg.statCost.A, tcp = prev ? prev.statCost.A : 0;
    rows += "<tr class='total'><td>Total (Available)</td><td>" + intc(seg.countA) + "</td><td>100.0%</td>"
      + deltaCell(prev ? seg.countA - prev.countA : null, false) + "<td>" + money(tc) + "</td>"
      + "<td class='cell-delta " + (prev ? cls(tc - tcp) : "flat") + "'>" + (prev ? signedPct(tc, tcp) : "—") + "</td>"
      + deltaCell(prev ? tc - tcp : null, true) + "</tr></tbody>";
    $("cmpTable").innerHTML = rows;
    $("cmpStatus").innerHTML = dollar("$ in Available", seg.statCost.A) + dollar("$ in Sold Available", seg.statCost.SA) + dollar("$ in On Order", seg.statCost.O);
  }

  function renderYearTable(seg, prev) {
    var years = Core.sortedYears(Object.assign({}, seg.yrCountAO, prev ? prev.yrCountAO : {}));
    var totCost = sumVals(seg.yrCostAO);
    var rows = "<thead><tr><th>Model Year</th><th>Units</th><th>%</th><th>Δ Units</th><th>Total Cost</th><th>%</th><th>Δ Cost</th></tr></thead><tbody>";
    years.forEach(function (y) {
      var cu = seg.yrCountAO[y] || 0, pu = prev ? (prev.yrCountAO[y] || 0) : null;
      var cc = seg.yrCostAO[y] || 0, pc = prev ? (prev.yrCostAO[y] || 0) : null;
      rows += "<tr><td>" + y + "</td><td>" + intc(cu) + "</td><td>" + pct(cu, seg.countAO) + "</td>"
        + deltaCell(prev ? cu - pu : null, false) + "<td>" + money(cc) + "</td><td>" + pct(cc, totCost) + "</td>"
        + deltaCell(prev ? cc - pc : null, true) + "</tr>";
    });
    rows += "<tr class='total'><td>Total</td><td>" + intc(seg.countAO) + "</td><td>100.0%</td>"
      + deltaCell(prev ? seg.countAO - prev.countAO : null, false) + "<td>" + money(totCost) + "</td><td>100.0%</td>"
      + deltaCell(prev ? totCost - sumVals(prev.yrCostAO) : null, true) + "</tr></tbody>";
    $("yearTable").innerHTML = rows;
  }

  function renderStatusTable(seg) {
    var labels = [["A", "Available"], ["SA", "Sold Available"], ["O", "On Order"]];
    var rows = "<thead><tr><th>Status</th><th>Description</th><th>Units</th><th>Total Cost</th></tr></thead><tbody>";
    labels.forEach(function (l) {
      rows += "<tr><td>" + l[0] + "</td><td>" + l[1] + "</td><td>" + intc(seg.statCount[l[0]] || 0) + "</td><td>" + money(seg.statCost[l[0]] || 0) + "</td></tr>";
    });
    rows += "<tr class='total'><td colspan='2'>Total Units (A + SA + O)</td><td>" + intc(seg.count) + "</td><td>" + money(seg.statCost.A + seg.statCost.SA + seg.statCost.O) + "</td></tr></tbody>";
    $("statusTable").innerHTML = rows;
  }

  function renderAging(seg) {
    var names = Core.BUCKETS.map(function (b) { return b[0]; });
    var cost = names.map(function (n) { return seg.bktCostAO[n] || 0; });
    var units = names.map(function (n) { return seg.bktUnitsAO[n] || 0; });
    if (state.chart) state.chart.destroy();
    state.chart = agingBar($("agingChart"), names, cost, units, "Age (Days)");
  }

  // =========================================================================
  // USED / PREOWNED inventory
  // =========================================================================
  async function handleUsedFile(file) {
    var wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
    var u = Core.buildUsed(sheetGrid(wb));
    var d = Core.parseDateFromName(file.name);
    if (!d) toast("Could not read a date from the filename — expected …_DDMMYYYY…");
    var snap = {
      kind: "used", date: d ? d.iso : new Date().toISOString().slice(0, 10),
      dateDisplay: d ? d.display : "(today)", fileName: file.name,
      total: u.total, grand: u.grand, yrCount: u.yrCount, yrCost: u.yrCost, ageBands: u.ageBands,
    };
    U.snap = snap;
    try { await api(ep("/api/snapshots", true), { method: "POST", body: JSON.stringify(snap) }); toast("Saved Used week " + snap.dateDisplay); } catch (e) {}
    U.prior = null;
    try { var res = await api(ep("/api/snapshots?before=" + encodeURIComponent(snap.date), true)); U.prior = res.snapshot || null; } catch (e) {}
    await loadUsedWeeks();
    $("uDropzone").hidden = true; $("uDash").hidden = false;
    renderUsed(); if (state.mode === "used") loadWeeks();
  }

  async function loadUsedWeeks() {
    try { var res = await api(ep("/api/snapshots?all=1", true)); U.allWeeks = res.snapshots || []; }
    catch (e) { U.allWeeks = U.snap ? [U.snap] : []; }
  }

  async function bootstrapUsed() {
    U.bootstrapped = true;
    await loadUsedWeeks();
    if (U.allWeeks.length) {
      U.snap = U.allWeeks[U.allWeeks.length - 1];
      U.prior = U.allWeeks.length > 1 ? U.allWeeks[U.allWeeks.length - 2] : null;
      $("uDropzone").hidden = true; $("uDash").hidden = false; renderUsed();
    }
  }

  function renderUsed() {
    if (!U.snap) return;
    var s = U.snap, p = U.prior;
    if (state.mode === "used") setDateline(s, p);
    $("uMeta").textContent = intc(s.total) + " units  ·  grand total " + money(s.grand)
      + (p ? "  ·  vs " + p.dateDisplay : "");
    // KPIs
    var dU = p ? s.total - p.total : null, dC = p ? s.grand - p.grand : null;
    $("uKpis").innerHTML = kpiCard("Total units", intc(s.total), dU, false) + kpiCard("Total cost", money0(s.grand), dC, true);
    // Year table
    $("uCmpDates").textContent = p ? (s.dateDisplay + " vs " + p.dateDisplay) : s.dateDisplay + " (no prior)";
    var years = Core.sortedYears(Object.assign({}, s.yrCount, p ? p.yrCount : {}));
    var rows = "<thead><tr><th>Model Year</th><th>Units</th><th>%</th><th>Δ Units</th><th>Total Cost</th><th>% Δ</th><th>$ Δ</th></tr></thead><tbody>";
    years.forEach(function (y) {
      var cu = s.yrCount[y] || 0, pu = p ? (p.yrCount[y] || 0) : null;
      var cc = s.yrCost[y] || 0, pc = p ? (p.yrCost[y] || 0) : 0;
      rows += "<tr><td>" + y + "</td><td>" + intc(cu) + "</td><td>" + pct(cu, s.total) + "</td>"
        + deltaCell(p ? cu - pu : null, false) + "<td>" + money(cc) + "</td>"
        + "<td class='cell-delta " + (p ? cls(cc - pc) : "flat") + "'>" + (p ? signedPct(cc, pc) : "—") + "</td>"
        + deltaCell(p ? cc - pc : null, true) + "</tr>";
    });
    rows += "<tr class='total'><td>Total</td><td>" + intc(s.total) + "</td><td>100.0%</td>"
      + deltaCell(p ? s.total - p.total : null, false) + "<td>" + money(s.grand) + "</td><td>100.0%</td>"
      + deltaCell(p ? s.grand - p.grand : null, true) + "</tr></tbody>";
    $("uYearTable").innerHTML = rows;
    // Age table
    var at = "<thead><tr><th>Age (Days)</th><th>Units</th><th>Total Cost</th><th>% of Cost</th></tr></thead><tbody>";
    s.ageBands.forEach(function (b) {
      at += "<tr><td>" + b.label + "</td><td>" + intc(b.units) + "</td><td>" + money(b.cost) + "</td><td>" + pct(b.cost, s.grand) + "</td></tr>";
    });
    at += "<tr class='total'><td>Total</td><td>" + intc(s.total) + "</td><td>" + money(s.grand) + "</td><td>100.0%</td></tr></tbody>";
    $("uAgeTable").innerHTML = at;
    // Aging chart (cost by age band)
    if (U.agingChart) U.agingChart.destroy();
    U.agingChart = agingBar($("uAgingChart"),
      s.ageBands.map(function (b) { return b.label; }),
      s.ageBands.map(function (b) { return b.cost; }),
      s.ageBands.map(function (b) { return b.units; }), "Age (Days)");
    renderUsedTrend();
  }

  function renderUsedTrend() {
    var weeks = U.allWeeks || [];
    var metric = UMETRICS[U.trendMetric];
    lineTrend($("uTrendChart"), $("uTrendNote"), weeks.map(function (w) { return w.dateDisplay || w.date; }),
      weeks.map(function (w) { return metric.get(w) || 0; }), metric, "U");
  }

  // =========================================================================
  // Shared chart builders
  // =========================================================================
  function agingBar(canvas, names, cost, units, xlabel) {
    var total = cost.reduce(function (a, b) { return a + b; }, 0) || 1;
    var base = ["#0a1f5c", "#123fa8", "#1f6fe0", "#2f7ff0", "#5aa8f7", "#8cc2fb", "#b9dcfd"];
    var colors = names.map(function (_, i) { return base[Math.min(i, base.length - 1)]; });
    var labels = names.map(function (n, i) { return [n, (cost[i] / total * 100).toFixed(0) + "%"]; });
    return new Chart(canvas, {
      type: "bar",
      data: { labels: labels, datasets: [{ data: cost, backgroundColor: colors, borderRadius: 6, maxBarThickness: 120 }] },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 28 } },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: function (c) { return intc(units[c.dataIndex]) + " units · " + money(c.parsed.y); } } },
          datalabels: { anchor: "end", align: "end", offset: 2, color: "#14203a", font: { weight: "700", size: 11 },
            formatter: function (v, c) { return intc(units[c.dataIndex]) + "\n" + money0(v); } },
        },
        scales: { y: { beginAtZero: true, ticks: { callback: function (v) { return money0(v); } }, grid: { color: "#eef1f6" } }, x: { grid: { display: false } } },
      },
      plugins: [ChartDataLabels],
    });
  }

  function lineTrend(canvas, noteEl, labels, vals, metric, which) {
    var fmt = metric.money ? money0 : intc;
    if (vals.length >= 2) {
      var d = vals[vals.length - 1] - vals[0];
      noteEl.innerHTML = metric.label + " — latest <b>" + fmt(vals[vals.length - 1]) + "</b> (" + labels[labels.length - 1]
        + "), net <span class='cell-delta " + cls(d) + "'>" + (metric.money ? signedMoney(d) : signedInt(d)) + "</span> since " + labels[0] + " across " + vals.length + " weeks.";
    } else if (vals.length === 1) {
      noteEl.innerHTML = metric.label + " — <b>" + fmt(vals[0]) + "</b> for " + labels[0] + ". Save at least two weeks to see the trend line.";
    } else { noteEl.textContent = "No saved weeks yet."; }
    var holder = which === "U" ? U : state;
    if (holder.trendChart) holder.trendChart.destroy();
    holder.trendChart = new Chart(canvas, {
      type: "line",
      data: { labels: labels, datasets: [{ data: vals, borderColor: "#1f6fe0", backgroundColor: "rgba(31,111,224,.12)", borderWidth: 2.5, fill: true, tension: 0.25, pointRadius: 4, pointBackgroundColor: "#0a1f5c" }] },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 24 } },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: function (c) {
            var i = c.dataIndex, dd = i > 0 ? vals[i] - vals[i - 1] : null;
            return dd === null ? fmt(vals[i]) : fmt(vals[i]) + "  (" + (metric.money ? signedMoney(dd) : signedInt(dd)) + " vs prior)";
          } } },
          datalabels: { align: "top", offset: 4, color: "#14203a", font: { weight: "700", size: 10.5 },
            formatter: function (v) { return metric.money ? money0(v) : intc(v); }, display: function (c) { return c.chart.data.labels.length <= 16; } },
        },
        scales: { y: { beginAtZero: false, ticks: { callback: function (v) { return metric.money ? money0(v) : intc(v); } }, grid: { color: "#eef1f6" } }, x: { grid: { display: false } } },
      },
      plugins: [ChartDataLabels],
    });
  }

  // =========================================================================
  // Saved weeks drawer (mode-aware)
  // =========================================================================
  async function loadWeeks() {
    var used = state.mode === "used";
    var list = $("weeksList");
    $("drawerTitle").textContent = used ? "Saved Used weeks" : "Saved New weeks";
    try {
      var res = await api(ep("/api/snapshots", used));
      var items = res.snapshots || [];
      if (!items.length) { list.innerHTML = "<p class='note'>No weeks saved yet.</p>"; return; }
      var cur = used ? (U.snap && U.snap.date) : (state.snap && state.snap.date);
      list.innerHTML = items.map(function (it) {
        var isCur = it.date === cur;
        return "<div class='week-item " + (isCur ? "current" : "") + "'>"
          + "<div><div class='w-date'>" + (it.dateDisplay || it.date) + (isCur ? " · current" : "") + "</div>"
          + "<div class='w-sub'>" + (it.fileName || "") + "</div></div>"
          + "<button class='del' data-date='" + it.date + "'>Delete</button></div>";
      }).join("");
      Array.prototype.forEach.call(list.querySelectorAll(".del"), function (b) {
        b.addEventListener("click", async function () {
          if (!confirm("Delete the saved week " + b.dataset.date + "?")) return;
          await api(ep("/api/snapshots?date=" + encodeURIComponent(b.dataset.date), used), { method: "DELETE" });
          toast("Deleted " + b.dataset.date);
          if (used) {
            await loadUsedWeeks();
            if (U.snap) { var r = await api(ep("/api/snapshots?before=" + encodeURIComponent(U.snap.date), true)); U.prior = r.snapshot || null; renderUsed(); }
          } else {
            await loadAllWeeks();
            if (state.snap) { var r2 = await api("/api/snapshots?before=" + encodeURIComponent(state.snap.date)); state.prior = r2.snapshot || null; render(); }
          }
          loadWeeks();
        });
      });
    } catch (e) { list.innerHTML = "<p class='note'>Could not load saved weeks.</p>"; }
  }

  // =========================================================================
  // Mode switching + wiring
  // =========================================================================
  function setMode(mode) {
    state.mode = mode;
    $("newView").hidden = mode !== "new";
    $("usedView").hidden = mode !== "used";
    Array.prototype.forEach.call(document.querySelectorAll(".toptab"), function (t) {
      t.classList.toggle("active", t.dataset.mode === mode);
    });
    if (mode === "used") {
      if (!U.bootstrapped) bootstrapUsed();
      else if (U.snap) { setDateline(U.snap, U.prior); } else $("dateline").textContent = "Drop the Used export to begin";
    } else {
      if (state.snap) setDateline(state.snap, state.prior); else $("dateline").textContent = "Drop this week's export to begin";
    }
  }

  Array.prototype.forEach.call(document.querySelectorAll(".toptab"), function (t) {
    t.addEventListener("click", function () { setMode(t.dataset.mode); });
  });

  // top-bar upload button routes to active mode
  $("uploadBtn").addEventListener("click", function () { (state.mode === "used" ? $("uFileInput") : $("fileInput")).click(); });

  // New view inputs
  $("dzPick").addEventListener("click", function () { $("fileInput").click(); });
  $("fileInput").addEventListener("change", function (e) { if (e.target.files[0]) handleFile(e.target.files[0]); e.target.value = ""; });
  // Used view inputs
  $("uPick").addEventListener("click", function () { $("uFileInput").click(); });
  $("uFileInput").addEventListener("change", function (e) { if (e.target.files[0]) handleUsedFile(e.target.files[0]); e.target.value = ""; });

  function wireDrop(zoneId, handler) {
    var dz = $(zoneId);
    ["dragenter", "dragover"].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add("drag"); }); });
    ["dragleave", "drop"].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove("drag"); }); });
    dz.addEventListener("drop", function (e) { if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]); });
  }
  wireDrop("dropzone", handleFile);
  wireDrop("uDropzone", handleUsedFile);
  window.addEventListener("dragover", function (e) { e.preventDefault(); });
  window.addEventListener("drop", function (e) {
    e.preventDefault(); if (!e.dataTransfer.files[0]) return;
    if (state.mode === "used") handleUsedFile(e.dataTransfer.files[0]); else handleFile(e.dataTransfer.files[0]);
  });

  // New sub-tabs (segments)
  Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (t) {
    t.addEventListener("click", function () {
      document.querySelectorAll(".tab").forEach(function (x) { x.classList.remove("active"); });
      t.classList.add("active"); state.seg = t.dataset.seg; render();
    });
  });

  $("trendMetric").addEventListener("change", function (e) { state.trendMetric = e.target.value; renderTrend(); });
  $("uTrendMetric").addEventListener("change", function (e) { U.trendMetric = e.target.value; renderUsedTrend(); });

  $("weeksBtn").addEventListener("click", function () { $("drawer").hidden = false; loadWeeks(); });
  $("drawerClose").addEventListener("click", function () { $("drawer").hidden = true; });
  $("drawerScrim").addEventListener("click", function () { $("drawer").hidden = true; });

  // open in New mode
  setMode("new");
  bootstrapNew();
})();
