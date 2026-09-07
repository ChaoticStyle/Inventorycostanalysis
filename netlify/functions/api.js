// Netlify Function (v2) — snapshot store + password check for the inventory dashboard.
// Paths: /api/auth and /api/snapshots. Server-side store = Netlify Blobs.
// Only the FINAL AGGREGATED NUMBERS are stored (one blob per week, keyed by date).
import { getStore } from "@netlify/blobs";

// Password gate. Set DASHBOARD_PASSWORD in Netlify env to require it; if unset the site
// runs open (prevents lockout). Data calls must carry the password in the x-dash-key header.
const password = () => process.env.DASHBOARD_PASSWORD || "";
function authorized(req) {
  const pw = password();
  if (!pw) return true;
  return req.headers.get("x-dash-key") === pw;
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json", "cache-control": "no-store" },
  });
}

export default async (req) => {
  const url = new URL(req.url);

  // ---- /api/auth ----
  if (url.pathname.endsWith("/auth")) {
    const pw = password();
    if (req.method === "GET") return json({ required: !!pw });
    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      const ok = !pw || body.password === pw;
      return json({ ok, required: !!pw }, ok ? 200 : 401);
    }
    return json({ error: "method not allowed" }, 405);
  }

  // ---- /api/snapshots ----
  if (!authorized(req)) return json({ error: "unauthorized" }, 401);
  // New and Used inventory are stored in separate blob stores (?kind=used for Used).
  const used = url.searchParams.get("kind") === "used";
  const store = getStore({ name: used ? "inventory-snapshots-used" : "inventory-snapshots", consistency: "strong" });

  if (req.method === "GET") {
    const date = url.searchParams.get("date");
    const before = url.searchParams.get("before");
    if (url.searchParams.get("all")) {
      // full snapshots (all segments), oldest first — powers the weekly trend chart
      const { blobs } = await store.list();
      const keys = blobs.map((b) => b.key).sort();
      const snaps = [];
      for (const k of keys) { const s = await store.get(k, { type: "json" }); if (s) snaps.push(s); }
      return json({ snapshots: snaps });
    }
    if (date) {
      const snap = await store.get(date, { type: "json" });
      return json({ snapshot: snap || null });
    }
    if (before) {
      const { blobs } = await store.list();
      const keys = blobs.map((b) => b.key).filter((k) => k < before).sort();
      const latest = keys[keys.length - 1];
      const snap = latest ? await store.get(latest, { type: "json" }) : null;
      return json({ snapshot: snap });
    }
    // list metadata, newest first
    const { blobs } = await store.list();
    const items = [];
    for (const b of blobs) {
      const s = await store.get(b.key, { type: "json" });
      if (s) items.push({ date: s.date, dateDisplay: s.dateDisplay, fileName: s.fileName, savedAt: s.savedAt });
    }
    items.sort((a, b) => (a.date < b.date ? 1 : -1));
    return json({ snapshots: items });
  }

  if (req.method === "POST") {
    const snap = await req.json().catch(() => null);
    if (!snap || !snap.date) return json({ error: "missing date" }, 400);
    snap.savedAt = new Date().toISOString();
    await store.setJSON(snap.date, snap);
    return json({ ok: true, date: snap.date });
  }

  if (req.method === "DELETE") {
    const date = url.searchParams.get("date");
    if (!date) return json({ error: "missing date" }, 400);
    await store.delete(date);
    return json({ ok: true, date });
  }

  return json({ error: "method not allowed" }, 405);
};

export const config = { path: ["/api/auth", "/api/snapshots"] };
