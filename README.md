# New Inventory — Weekly Dashboard

A Netlify-hosted dashboard where leadership drops the weekly **"INVENTORY WITH GL COSTS"**
export and instantly sees a week-over-week comparison — the same numbers, rules, and
chart as the Word reports, in a cleaner view. Only the **final aggregated numbers** for
each week are stored server-side (Netlify Blobs); the raw Excel file never leaves the
browser.

## What it shows (three tabs from one file)

**All New**, **All Motors** (TYPE A/B/C/D), **All Towables** (TYPE 5W/TT/TH/5WTH) — all
computed from a single dropped file. Each tab shows:

- **KPI cards** — Available units (A) with Δ, $ in Available with Δ, $ in Sold Available,
  $ in On Order, Total units (A+SA+O).
- **Week-to-Week Comparison** (Available = STAT A) — model-year units + % of available,
  cost by year with % change and $ change vs the prior saved week. Sums to $ in Available.
- **Weekly Trend** — a line chart across every saved week with a metric selector
  (Available units, $ in Available, Units A+O, $ On Order, $ Sold Available, Total units)
  and a net-change readout. This is the running over/under tracker.
- **Units & Cost by Model Year** (A + O) with week-over-week Δ.
- **Inventory Cost by Status** (A / SA / O) and Total Units (A+SA+O).
- **Aging Inventory by Cost** chart (A + O), dark-to-light blue, unit + cost labels and
  cost-share % under each bar.

Opening the link shows the **latest saved week** automatically (with the trend), so you
don't have to re-drop a file to review history. Dropping a new file adds/updates that week.

## The rules (identical to the Word reports)

| View | Statuses counted |
|------|------------------|
| Total Units (headline / status table) | A + SA + O |
| Units / Cost by Model Year, Aging chart | A + O |
| Comparison "available units" + its year breakdown | A only |
| Comparison "$ in Sold Available" | SA |
| Sold Order (SO) | excluded everywhere |

Dates come from the filename (`…_DDMMYYYYHHMMSS`). Dropping a file **auto-saves** that
week's final numbers (keyed by date; re-dropping the same date overwrites) and compares
to the most recent earlier saved week. Use **Saved weeks** to review or delete stored weeks.

## Deploy to Netlify

**Option A — drag & drop (fastest):**
1. Go to https://app.netlify.com → *Add new site* → *Deploy manually*.
2. Drag this `dashboard/` folder onto the page.
3. Netlify installs `@netlify/blobs`, publishes the site, and wires up the function at
   `/api/*`. Blobs storage is enabled automatically — no database to configure.

**Option B — Git + CLI:**
```bash
cd dashboard
npm install
netlify deploy --build --prod      # or connect the repo in the Netlify UI
```

### Set the access password
Netlify → Site configuration → **Environment variables** → add:

```
DASHBOARD_PASSWORD = <your shared password>
```

Redeploy after setting it. The page then prompts for this password and every data call
(save / list / delete) requires it. If the variable is left unset, the site runs open
(no password) — useful for a first test, but set it before sharing the link.

## Local development
```bash
cd dashboard
npm install
netlify dev      # serves the site + functions + Blobs locally at http://localhost:8888
```
(Plain `python -m http.server` serves the UI but not the `/api` functions, so save/compare
won't work — the tables still render from a dropped file.)

## Files
- `index.html`, `styles.css` — UI
- `inventory-core.js` — the shared rules engine (mirrors `scripts/inventory_analysis.py`)
- `app.js` — Excel parsing (SheetJS), rendering, charts, save/delete
- `netlify/functions/api.js` — `/api/auth` + `/api/snapshots` (Netlify Blobs store)
- `netlify.toml`, `package.json`
