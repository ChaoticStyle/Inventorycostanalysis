/* app.js — New Inventory weekly dashboard front-end. */
(function () {
  "use strict";
  var Core = window.InventoryCore;

  var state = { snap: null, prior: null, seg: "All", chart: null,
    allWeeks: [], trendMetric: "availUnits", trendChart: null };

  // Metrics available in the weekly trend chart (value pulled from a segment object).
  var METRICS = {
    availUnits: { label: "Available units (A)", get: function (s) { return s.countA; }, money: false },
    availCost: { label: "$ in Available", get: function (s) { return s.statCost.A; }, money: true },
    aoUnits: { label: "Units (A + O)", get: function (s) { return s.countAO; }, money: false },
    onOrder: { label: "$ in On Order", get: function (s) { return s.statCost.O; }, money: true },
    soldAvail: { label: "$ in Sold Available", get: function (s) { return s.statCost.SA; }, money: true },
    totalUnits: { label: "Total units (A+SA+O)", get: function (s) { return s.count; }, money: false },
  };

  // ---------- helpers ----------
  var $ = function (id) { return document.getElementById(id); };
  function money(x) { return "$" + (Math.round(x * 100) / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function money0(x) { return "$" + Math.round(x).toLocaleString("en-US"); }
  function intc(x) { return Number(x).toLocaleString("en-US"); }
  function pct(part, whole) { return whole ? (part / whole * 100).toFixed(1) + "%" : "0.0%"; }
  function signedInt(n) { return (n > 0 ? "+" : n < 0 ? "−" : "") + intc(Math.abs(n)); }
  function signedMoney(x) { return (x >= 0 ? "+" : "−") + money(Math.abs(x)); }
  function signedPct(cur, prev) {
    if (prev === 0) return cur === 0 ? "0%" : "new";
    var p = (cur - prev) / prev * 100;
    return (p >= 0 ? "+" : "−") + Math.abs(p).toFixed(0) + "%";
  }
  function cls(n) { return n > 0 ? "pos" : n < 0 ? "neg" : "flat"; }
  function toast(msg) {
    var t = $("toast"); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2600);
  }
  async function api(path, opts) {
    opts = opts || {};
    opts.headers = { "content-type": "application/json" };
    var res = await fetch(path, opts);
    return res.json();
  }

  // ---------- file handling ----------
  function readRows(wb) {
    var ws = wb.Sheets[wb.SheetNames[0]];
    wb.SheetNames.forEach(function (nm) {
      var r0 = XLSX.utils.sheet_to_json(wb.Sheets[nm], { header: 1, blankrows: false })[0] || [];
      var up = r0.map(function (c) { return String(c).trim().toUpperCase(); });
      if (up.indexOf("TOTAL COST") >= 0 || up.indexOf("STOCK") >= 0) ws = wb.Sheets[nm];
    });
    var grid = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
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
    var buf = await file.arrayBuffer();
    var wb = XLSX.read(buf, { type: "array" });
    var built = Core.buildSegments(readRows(wb));
    var d = Core.parseDateFromName(file.name);
    if (!d) { toast("Could not read a date from the filename — expected …_DDMMYYYY…"); }
    var snap = {
      date: d ? d.iso : new Date().toISOString().slice(0, 10),
      dateDisplay: d ? d.display : "(today)",
      fileName: file.name,
      meta: { kept: built.kept, excludedSO: built.excludedSO, unclassifiedTypes: built.unclassifiedTypes },
      segments: built.segments,
    };
    state.snap = snap;

    // auto-save this week's final numbers
    try {
      await api("/api/snapshots", { method: "POST", body: JSON.stringify(snap) });
      toast("Saved snapshot for " + snap.dateDisplay);
    } catch (e) { /* ignore save errors */ }

    // fetch prior week (latest before this date) for deltas
    state.prior = null;
    try {
      var res = await api("/api/snapshots?before=" + encodeURIComponent(snap.date));
      state.prior = res.snapshot || null;
    } catch (e) {}

    await loadAllWeeks();
    $("dropzone").hidden = true;
    $("dash").hidden = false;
    render();
    loadWeeks();
  }

  async function loadAllWeeks() {
    try {
      var res = await api("/api/snapshots?all=1");
      state.allWeeks = res.snapshots || [];
    } catch (e) { state.allWeeks = state.snap ? [state.snap] : []; }
  }

  // On open, show the latest saved week (and its prior) without needing a re-drop.
  async function bootstrap() {
    await loadAllWeeks();
    if (state.allWeeks.length) {
      state.snap = state.allWeeks[state.allWeeks.length - 1];
      state.prior = state.allWeeks.length > 1 ? state.allWeeks[state.allWeeks.length - 2] : null;
      $("dropzone").hidden = true;
      $("dash").hidden = false;
      render();
    }
    loadWeeks();
  }

  // ---------- rendering ----------
  function render() {
    if (!state.snap) return;
    var seg = state.snap.segments[state.seg];
    var prev = state.prior ? state.prior.segments[state.seg] : null;
    var s = state.snap, p = state.prior;

    $("dateline").textContent = p
      ? "Inventory as of " + s.dateDisplay + "  ·  compared to " + p.dateDisplay
      : "Inventory as of " + s.dateDisplay + "  ·  no prior week yet";

    var m = s.meta || { excludedSO: 0, unclassifiedTypes: [] };
    $("metaNote").textContent = (m.excludedSO || 0) + " Sold Order unit(s) excluded"
      + (m.unclassifiedTypes && m.unclassifiedTypes.length ? "  ·  unclassified TYPE: " + m.unclassifiedTypes.join(", ") : "");

    renderKpis(seg, prev);
    renderComparison(seg, prev, s, p);
    renderTrend();
    renderYearTable(seg, prev);
    renderStatusTable(seg);
    renderAging(seg);
  }

  // Weekly trend line across every saved week, for the active segment + chosen metric.
  function renderTrend() {
    var weeks = (state.allWeeks || []).filter(function (w) { return w.segments && w.segments[state.seg]; });
    var metric = METRICS[state.trendMetric];
    var labels = weeks.map(function (w) { return w.dateDisplay || w.date; });
    var vals = weeks.map(function (w) { return metric.get(w.segments[state.seg]) || 0; });
    var fmt = metric.money ? money0 : intc;

    // net-change readout
    if (weeks.length >= 2) {
      var first = vals[0], last = vals[vals.length - 1], d = last - first;
      $("trendNote").innerHTML = metric.label + " — latest <b>" + fmt(last) + "</b> ("
        + (labels[labels.length - 1]) + "), net <span class='cell-delta " + cls(d) + "'>"
        + (metric.money ? signedMoney(d) : signedInt(d)) + "</span> since " + labels[0]
        + " across " + weeks.length + " weeks.";
    } else if (weeks.length === 1) {
      $("trendNote").innerHTML = metric.label + " — <b>" + fmt(vals[0]) + "</b> for " + labels[0]
        + ". Save at least two weeks to see the trend line.";
    } else {
      $("trendNote").textContent = "No saved weeks yet.";
    }

    if (state.trendChart) state.trendChart.destroy();
    state.trendChart = new Chart($("trendChart"), {
      type: "line",
      data: {
        labels: labels,
        datasets: [{
          data: vals, label: metric.label, borderColor: "#1f6fe0", backgroundColor: "rgba(31,111,224,.12)",
          borderWidth: 2.5, fill: true, tension: 0.25, pointRadius: 4, pointBackgroundColor: "#0a1f5c",
        }],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { top: 24 } },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: function (c) {
            var i = c.dataIndex, prevv = i > 0 ? vals[i - 1] : null, dd = prevv === null ? null : vals[i] - prevv;
            var base = fmt(vals[i]);
            return dd === null ? base : base + "  (" + (metric.money ? signedMoney(dd) : signedInt(dd)) + " vs prior)";
          } } },
          datalabels: {
            align: "top", offset: 4, color: "#14203a", font: { weight: "700", size: 10.5 },
            formatter: function (v) { return metric.money ? money0(v) : intc(v); },
            display: function (c) { return c.chart.data.labels.length <= 16; },
          },
        },
        scales: {
          y: { beginAtZero: false, ticks: { callback: function (v) { return metric.money ? money0(v) : intc(v); } }, grid: { color: "#eef1f6" } },
          x: { grid: { display: false } },
        },
      },
      plugins: [ChartDataLabels],
    });
  }

  function deltaCell(n, isMoney) {
    if (n === null) return "<td class='cell-delta flat'>—</td>";
    var txt = isMoney ? signedMoney(n) : signedInt(n);
    return "<td class='cell-delta " + cls(n) + "'>" + txt + "</td>";
  }

  function renderKpis(seg, prev) {
    var dA = prev ? seg.countA - prev.countA : null;
    var dAvail = prev ? seg.statCost.A - prev.statCost.A : null;
    function kpi(label, value, delta, isMoney) {
      var d = "";
      if (delta !== null && delta !== undefined)
        d = "<div class='delta " + cls(delta) + "'>" + (isMoney ? signedMoney(delta) : signedInt(delta)) + " vs prior</div>";
      return "<div class='kpi'><div class='label'>" + label + "</div><div class='value'>" + value + "</div>" + d + "</div>";
    }
    $("kpis").innerHTML =
      kpi("Available units (A)", intc(seg.countA), dA, false) +
      kpi("$ in Available", money0(seg.statCost.A), dAvail, true) +
      kpi("$ in Sold Available", money0(seg.statCost.SA), null) +
      kpi("$ in On Order", money0(seg.statCost.O), null) +
      kpi("Total units (A+SA+O)", intc(seg.count), null);
  }

  // Week-to-week comparison — Available (STAT A) only, matches the Word docs.
  function renderComparison(seg, prev, s, p) {
    $("cmpDates").textContent = p ? (s.dateDisplay + " vs " + p.dateDisplay) : s.dateDisplay + " (no prior)";
    var years = Core.sortedYears(Object.assign({}, seg.yrCountA, prev ? prev.yrCountA : {}));
    var rows = "<thead><tr><th>Model Year</th><th>Avail. Units</th><th>% of Avail.</th><th>Δ Units</th><th>Total Cost</th><th>% Δ</th><th>$ Δ</th></tr></thead><tbody>";
    years.forEach(function (y) {
      var cu = seg.yrCountA[y] || 0, pu = prev ? (prev.yrCountA[y] || 0) : null;
      var cc = seg.yrCostA[y] || 0, pc = prev ? (prev.yrCostA[y] || 0) : 0;
      rows += "<tr><td>" + y + "</td><td>" + intc(cu) + "</td><td>" + pct(cu, seg.countA) + "</td>"
        + deltaCell(prev ? cu - pu : null, false)
        + "<td>" + money(cc) + "</td>"
        + "<td class='cell-delta " + (prev ? cls(cc - pc) : "flat") + "'>" + (prev ? signedPct(cc, pc) : "—") + "</td>"
        + deltaCell(prev ? cc - pc : null, true) + "</tr>";
    });
    var tc = seg.statCost.A, tcp = prev ? prev.statCost.A : 0;
    rows += "<tr class='total'><td>Total (Available)</td><td>" + intc(seg.countA) + "</td><td>100.0%</td>"
      + deltaCell(prev ? seg.countA - prev.countA : null, false)
      + "<td>" + money(tc) + "</td>"
      + "<td class='cell-delta " + (prev ? cls(tc - tcp) : "flat") + "'>" + (prev ? signedPct(tc, tcp) : "—") + "</td>"
      + deltaCell(prev ? tc - tcp : null, true) + "</tr></tbody>";
    $("cmpTable").innerHTML = rows;

    $("cmpStatus").innerHTML =
      dollar("$ in Available", seg.statCost.A) +
      dollar("$ in Sold Available", seg.statCost.SA) +
      dollar("$ in On Order", seg.statCost.O);
  }
  function dollar(l, v) { return "<div class='dollar'><div class='l'>" + l + "</div><div class='v'>" + money(v) + "</div></div>"; }

  // Units & cost by model year — A + O.
  function renderYearTable(seg, prev) {
    var years = Core.sortedYears(Object.assign({}, seg.yrCountAO, prev ? prev.yrCountAO : {}));
    var totCost = 0; Object.keys(seg.yrCostAO).forEach(function (y) { totCost += seg.yrCostAO[y]; });
    var rows = "<thead><tr><th>Model Year</th><th>Units</th><th>%</th><th>Δ Units</th><th>Total Cost</th><th>%</th><th>Δ Cost</th></tr></thead><tbody>";
    years.forEach(function (y) {
      var cu = seg.yrCountAO[y] || 0, pu = prev ? (prev.yrCountAO[y] || 0) : null;
      var cc = seg.yrCostAO[y] || 0, pc = prev ? (prev.yrCostAO[y] || 0) : null;
      rows += "<tr><td>" + y + "</td><td>" + intc(cu) + "</td><td>" + pct(cu, seg.countAO) + "</td>"
        + deltaCell(prev ? cu - pu : null, false)
        + "<td>" + money(cc) + "</td><td>" + pct(cc, totCost) + "</td>"
        + deltaCell(prev ? cc - pc : null, true) + "</tr>";
    });
    rows += "<tr class='total'><td>Total</td><td>" + intc(seg.countAO) + "</td><td>100.0%</td>"
      + deltaCell(prev ? seg.countAO - prev.countAO : null, false)
      + "<td>" + money(totCost) + "</td><td>100.0%</td>"
      + deltaCell(prev ? totCost - sumVals(prev.yrCostAO) : null, true) + "</tr></tbody>";
    $("yearTable").innerHTML = rows;
  }
  function sumVals(o) { var s = 0; Object.keys(o).forEach(function (k) { s += o[k]; }); return s; }

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
    var total = cost.reduce(function (a, b) { return a + b; }, 0) || 1;
    var colors = ["#0a1f5c", "#123fa8", "#1f6fe0", "#2f7ff0", "#5aa8f7"];
    var labels = names.map(function (n, i) { return [n, (cost[i] / total * 100).toFixed(0) + "%"]; });
    if (state.chart) state.chart.destroy();
    state.chart = new Chart($("agingChart"), {
      type: "bar",
      data: { labels: labels, datasets: [{ data: cost, backgroundColor: colors, borderRadius: 6, maxBarThickness: 120 }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { top: 28 } },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: function (c) { return intc(units[c.dataIndex]) + " units · " + money(c.parsed.y); } } },
          datalabels: {
            anchor: "end", align: "end", offset: 2, color: "#14203a", font: { weight: "700", size: 11 },
            formatter: function (v, c) { return intc(units[c.dataIndex]) + "\n" + money0(v); },
          },
        },
        scales: {
          y: { beginAtZero: true, ticks: { callback: function (v) { return money0(v); } }, grid: { color: "#eef1f6" } },
          x: { grid: { display: false } },
        },
      },
      plugins: [ChartDataLabels],
    });
  }

  // ---------- saved weeks drawer ----------
  async function loadWeeks() {
    var list = $("weeksList");
    try {
      var res = await api("/api/snapshots");
      var items = res.snapshots || [];
      if (!items.length) { list.innerHTML = "<p class='note'>No weeks saved yet.</p>"; return; }
      var curDate = state.snap ? state.snap.date : null;
      list.innerHTML = items.map(function (it) {
        var current = it.date === curDate;
        return "<div class='week-item " + (current ? "current" : "") + "'>"
          + "<div><div class='w-date'>" + (it.dateDisplay || it.date) + (current ? " · current" : "") + "</div>"
          + "<div class='w-sub'>" + (it.fileName || "") + "</div></div>"
          + "<button class='del' data-date='" + it.date + "'>Delete</button></div>";
      }).join("");
      Array.prototype.forEach.call(list.querySelectorAll(".del"), function (b) {
        b.addEventListener("click", async function () {
          if (!confirm("Delete the saved week " + b.dataset.date + "?")) return;
          await api("/api/snapshots?date=" + encodeURIComponent(b.dataset.date), { method: "DELETE" });
          toast("Deleted " + b.dataset.date);
          await loadAllWeeks();
          // if we deleted the current prior, recompute comparison
          if (state.snap) {
            var r = await api("/api/snapshots?before=" + encodeURIComponent(state.snap.date));
            state.prior = r.snapshot || null; render();
          }
          loadWeeks();
        });
      });
    } catch (e) { list.innerHTML = "<p class='note'>Could not load saved weeks.</p>"; }
  }

  // ---------- wiring ----------
  $("uploadBtn").addEventListener("click", function () { $("fileInput").click(); });
  $("dzPick").addEventListener("click", function () { $("fileInput").click(); });
  $("fileInput").addEventListener("change", function (e) { if (e.target.files[0]) handleFile(e.target.files[0]); e.target.value = ""; });

  var dz = $("dropzone");
  ["dragenter", "dragover"].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add("drag"); }); });
  ["dragleave", "drop"].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove("drag"); }); });
  dz.addEventListener("drop", function (e) { if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  // allow dropping anywhere once a dashboard is shown
  window.addEventListener("dragover", function (e) { e.preventDefault(); });
  window.addEventListener("drop", function (e) { e.preventDefault(); if (!$("dash").hidden && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });

  Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (t) {
    t.addEventListener("click", function () {
      document.querySelectorAll(".tab").forEach(function (x) { x.classList.remove("active"); });
      t.classList.add("active"); state.seg = t.dataset.seg; render();
    });
  });

  $("trendMetric").addEventListener("change", function (e) { state.trendMetric = e.target.value; renderTrend(); });

  $("weeksBtn").addEventListener("click", function () { $("drawer").hidden = false; loadWeeks(); });
  $("drawerClose").addEventListener("click", function () { $("drawer").hidden = true; });
  $("drawerScrim").addEventListener("click", function () { $("drawer").hidden = true; });

  // No auth gate — open straight into the app.
  bootstrap();
})();
