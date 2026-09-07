/*
 * inventory-core.js — shared rules engine for the New Inventory dashboard.
 * Mirrors scripts/inventory_analysis.py so the dashboard and the Word docs agree.
 * Runs in the browser (window.InventoryCore) and in Node (module.exports) so the
 * logic can be unit-tested outside a browser.
 *
 * STATUS MAP (must match the docs):
 *   Total Units (headline)        = A + SA + O   (full segment, SO already dropped)
 *   Cost by Status table          = A, SA, O     (each listed)
 *   Units/Cost by Model Year      = A + O        (SA excluded)
 *   Aging by Cost chart           = A + O        (SA excluded)
 *   Comparison "available units"  = A only       (its cost-by-year sums to $ in Available)
 *   Comparison "$ in Sold Available" = SA
 *   SO (Sold Order)               = excluded everywhere
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else if (root) root.InventoryCore = factory();
})(typeof self !== "undefined" ? self : (typeof globalThis !== "undefined" ? globalThis : this), function () {
  "use strict";

  var MOTOR_TYPES = ["A", "B", "C", "D"];
  var TOW_TYPES = ["5W", "TT", "TH", "5WTH"];
  var YEAR_TABLE_STATS = ["A", "O"]; // A+O drives year tables + aging chart
  var BUCKETS = [
    ["0-90", 0, 90],
    ["91-180", 91, 180],
    ["181-270", 181, 270],
    ["271-364", 271, 364],
    ["365+", 365, Infinity],
  ];

  function num(x) {
    var n = parseFloat(x);
    return isFinite(n) ? n : 0;
  }

  function bucketName(days) {
    for (var i = 0; i < BUCKETS.length; i++) {
      if (days >= BUCKETS[i][1] && days <= BUCKETS[i][2]) return BUCKETS[i][0];
    }
    return null;
  }

  function emptyBuckets() {
    var o = {};
    BUCKETS.forEach(function (b) { o[b[0]] = 0; });
    return o;
  }

  // rows: [{year, type, cost, days, stat}] — already SO-excluded.
  function analyze(rows) {
    var out = {
      count: rows.length,
      statCount: { A: 0, SA: 0, O: 0 },
      statCost: { A: 0, SA: 0, O: 0 },
      yrCountAO: {}, yrCostAO: {},
      yrCountA: {}, yrCostA: {}, countA: 0,
      bktUnitsAO: emptyBuckets(), bktCostAO: emptyBuckets(),
    };
    rows.forEach(function (v) {
      var y = String(v.year).trim();
      var s = v.stat;
      if (out.statCount[s] === undefined) { out.statCount[s] = 0; out.statCost[s] = 0; }
      out.statCount[s] += 1;
      out.statCost[s] += v.cost;
      if (YEAR_TABLE_STATS.indexOf(s) >= 0) {
        out.yrCountAO[y] = (out.yrCountAO[y] || 0) + 1;
        out.yrCostAO[y] = (out.yrCostAO[y] || 0) + v.cost;
        var bn = bucketName(v.days);
        if (bn) { out.bktUnitsAO[bn] += 1; out.bktCostAO[bn] += v.cost; }
      }
      if (s === "A") {
        out.yrCountA[y] = (out.yrCountA[y] || 0) + 1;
        out.yrCostA[y] = (out.yrCostA[y] || 0) + v.cost;
        out.countA += 1;
      }
    });
    out.countAO = 0;
    Object.keys(out.yrCountAO).forEach(function (y) { out.countAO += out.yrCountAO[y]; });
    return out;
  }

  function inSet(t, set) { return set.indexOf(String(t).trim().toUpperCase()) >= 0; }

  // rawRows: [{stock, year, type, cost, days, stat}] straight from the sheet.
  // Returns {kept, excludedSO, segments:{All,Motors,Towable}, unclassifiedTypes}.
  function buildSegments(rawRows) {
    var kept = [], excludedSO = 0, other = {};
    rawRows.forEach(function (r) {
      var stock = r.stock, yr = r.year;
      if ((stock === undefined || stock === null || stock === "") &&
          (yr === undefined || yr === null || yr === "")) return;
      var stat = (r.stat === undefined || r.stat === null) ? "" : String(r.stat).trim().toUpperCase();
      if (stat === "SO") { excludedSO += 1; return; }
      var row = {
        year: String(yr).trim(),
        type: (r.type === undefined || r.type === null) ? "" : String(r.type).trim(),
        cost: num(r.cost),
        days: num(r.days),
        stat: stat,
      };
      kept.push(row);
      var tu = row.type.toUpperCase();
      if (!inSet(tu, MOTOR_TYPES) && !inSet(tu, TOW_TYPES)) other[row.type] = true;
    });
    var motors = kept.filter(function (v) { return inSet(v.type, MOTOR_TYPES); });
    var towable = kept.filter(function (v) { return inSet(v.type, TOW_TYPES); });
    // "All New" = Motors + Towable combined (unclassified TYPEs are flagged but not counted).
    return {
      kept: kept.length,
      excludedSO: excludedSO,
      unclassifiedTypes: Object.keys(other),
      segments: {
        All: analyze(motors.concat(towable)),
        Motors: analyze(motors),
        Towable: analyze(towable),
      },
    };
  }

  // "..._DDMMYYYYHHMMSS" -> {iso:'YYYY-MM-DD', display:'M/D/YYYY'}
  function parseDateFromName(name) {
    var m = /COSTS[_-]?(\d{2})(\d{2})(\d{4})/i.exec(String(name || ""));
    if (!m) return null;
    var dd = +m[1], mm = +m[2], yyyy = +m[3];
    if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
    var iso = yyyy + "-" + String(mm).padStart(2, "0") + "-" + String(dd).padStart(2, "0");
    return { iso: iso, display: mm + "/" + dd + "/" + yyyy };
  }

  function sortedYears(map) {
    return Object.keys(map).sort(function (a, b) { return (+a || 9999) - (+b || 9999); });
  }

  // ---------------- USED / PREOWNED ----------------
  // Different export layout: STK.# | YR | MAKE & MODEL | TYPE | VIN # | LOC | DAYS | TOTAL COST
  // All units, no Motors/Towable split, no status. Grouped by model year (2-digit, as
  // shown) and by 100-day age bands. A DAYS value of -1 is normalized to 1.
  var USED_HEADER_MAP = {
    "STK.#": "stock", "STK #": "stock", "STOCK": "stock",
    "YR": "year", "YEAR": "year", "TYPE": "type", "DAYS": "days", "TOTAL COST": "cost",
  };
  var USED_FALLBACK = { stock: 0, year: 1, type: 3, days: 6, cost: 7 };

  function usedAgeBand(days, width) {
    width = width || 100;
    var d = num(days); if (d < 0) d = 1;         // -1 (just arrived) counts as 1
    var lo = Math.floor(d / width) * width;
    return { lo: lo, label: lo + "-" + (lo + width - 1) };
  }

  function buildUsed(rawRows) {
    var header = null, cols = {};
    // rawRows: array of arrays (grid). First non-empty row = header.
    if (rawRows.length && Array.isArray(rawRows[0])) {
      header = rawRows[0].map(function (c) { return c == null ? "" : String(c).trim().toUpperCase(); });
      Object.keys(USED_HEADER_MAP).forEach(function (h) {
        var i = header.indexOf(h); if (i >= 0 && cols[USED_HEADER_MAP[h]] === undefined) cols[USED_HEADER_MAP[h]] = i;
      });
    }
    Object.keys(USED_FALLBACK).forEach(function (f) { if (cols[f] === undefined) cols[f] = USED_FALLBACK[f]; });

    var yrCount = {}, yrCost = {}, ageCount = {}, ageCost = {}, ageLabel = {}, total = 0, grand = 0, blank = 0;
    for (var i = 1; i < rawRows.length; i++) {
      var r = rawRows[i];
      var stock = r[cols.stock], yr = r[cols.year];
      if ((stock == null || stock === "") && (yr == null || yr === "")) { if (++blank > 5) break; continue; }
      blank = 0;
      var cost = num(r[cols.cost]);
      var y = String(yr).trim();
      yrCount[y] = (yrCount[y] || 0) + 1;
      yrCost[y] = (yrCost[y] || 0) + cost;
      var b = usedAgeBand(r[cols.days]);
      ageCount[b.lo] = (ageCount[b.lo] || 0) + 1;
      ageCost[b.lo] = (ageCost[b.lo] || 0) + cost;
      ageLabel[b.lo] = b.label;
      total += 1; grand += cost;
    }
    var maxLo = 0; Object.keys(ageLabel).forEach(function (k) { maxLo = Math.max(maxLo, +k); });
    for (var lo = 0; lo <= maxLo; lo += 100) if (!ageLabel[lo]) ageLabel[lo] = lo + "-" + (lo + 99);
    var ageBands = Object.keys(ageLabel).map(Number).sort(function (a, b2) { return a - b2; })
      .map(function (lo2) { return { lo: lo2, label: ageLabel[lo2], units: ageCount[lo2] || 0, cost: ageCost[lo2] || 0 }; });

    return { total: total, grand: grand, yrCount: yrCount, yrCost: yrCost, ageBands: ageBands };
  }

  return {
    MOTOR_TYPES: MOTOR_TYPES, TOW_TYPES: TOW_TYPES, BUCKETS: BUCKETS,
    analyze: analyze, buildSegments: buildSegments, buildUsed: buildUsed,
    parseDateFromName: parseDateFromName, sortedYears: sortedYears,
  };
});
