/* Temuan Lapangan QSHE — logika aplikasi. Untuk kata kunci & klausul, edit js/rapikan.js dan js/data-klausul.js */
"use strict";
/* =========================================================
   Temuan Lapangan QSHE — Hexindo
   ========================================================= */


const LIBS = {
  excel: "lib/exceljs.min.js",
  pptx: "lib/pptxgen.bundle.js",
  chart: "lib/chart.umd.js"
};

/* ---------------- utilities ---------------- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const uid = (p = "") => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const clone = o => JSON.parse(JSON.stringify(o ?? null));
const SEVS = ["Critical", "Major", "Minor"];
const HIER = ["Eliminasi", "Substitusi", "Rekayasa", "Administratif", "APD"];
const RC_CAT = ["Manusia", "Metode", "Mesin/Peralatan", "Material", "Lingkungan", "Manajemen"];
const BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const HARI = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

function localISO(d = new Date()) { const z = d.getTimezoneOffset() * 6e4; return new Date(d - z).toISOString().slice(0, 10); }
function parseISO(iso) { if (!iso) return null; const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); }
function addDays(iso, n) { const d = parseISO(iso) || new Date(); d.setDate(d.getDate() + Number(n || 0)); return localISO(d); }
function fmtTgl(iso, withDay) { const d = parseISO(iso); if (!d) return "-"; return (withDay ? HARI[d.getDay()] + ", " : "") + d.getDate() + " " + BULAN[d.getMonth()] + " " + d.getFullYear(); }
function fmtShort(iso) { const d = parseISO(iso); if (!d) return "-"; return d.getDate() + " " + BULAN[d.getMonth()].slice(0, 3) + " " + String(d.getFullYear()).slice(2); }
function daysBetween(aIso, bIso) { const a = parseISO(aIso), b = parseISO(bIso); if (!a || !b) return 0; return Math.round((b - a) / 864e5); }
function lines(s) { return String(s || "").split(/\n+/).map(x => x.trim()).filter(Boolean); }

function toast(msg, kind) {
  const host = $("#toastHost"); host.innerHTML = "";
  const t = document.createElement("div"); t.className = "toast" + (kind === "err" ? " err" : ""); t.setAttribute("role", "status"); t.textContent = msg;
  host.appendChild(t); clearTimeout(toast._t); toast._t = setTimeout(() => t.remove(), kind === "err" ? 6500 : 3200);
}
const LS = {
  get(k) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : null; } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
  del(k) { try { localStorage.removeItem(k); } catch (e) {} }
};
function loadScript(src) {
  loadScript.cache = loadScript.cache || {};
  if (!loadScript.cache[src]) loadScript.cache[src] = new Promise((res, rej) => {
    if (src.startsWith("embed:")) {
      try { const code = document.getElementById(src.slice(6)).textContent; const el = document.createElement("script"); el.textContent = code; document.head.appendChild(el); res(); }
      catch (e) { delete loadScript.cache[src]; rej(new Error("Gagal memuat pustaka bawaan.")); }
      return;
    }
    const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => { delete loadScript.cache[src]; rej(new Error("Gagal memuat pustaka: " + src)); };
    document.head.appendChild(s);
  });
  return loadScript.cache[src];
}
function useCap(name) {
  try { if (window.claude && typeof window.claude.use === "function") return window.claude.use(name).catch(() => null); } catch (e) {}
  return Promise.resolve(null);
}
function dataUrlToBlob(u) {
  const [head, b64] = u.split(","); const mime = (head.match(/data:([^;]+)/) || [])[1] || "image/jpeg";
  const bin = atob(b64); const arr = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return new Blob([arr], { type: mime });
}
function loadImg(src) { return new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = src; }); }
async function compressImage(file, maxSide = 1280, maxLen = 240000) {
  const url = URL.createObjectURL(file);
  try {
    const im = await loadImg(url);
    let scale = Math.min(1, maxSide / Math.max(im.naturalWidth, im.naturalHeight));
    for (let pass = 0; pass < 5; pass++) {
      const w = Math.max(1, Math.round(im.naturalWidth * scale)), h = Math.max(1, Math.round(im.naturalHeight * scale));
      const c = document.createElement("canvas"); c.width = w; c.height = h;
      const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, w, h); g.drawImage(im, 0, 0, w, h);
      for (let q = 0.8; q >= 0.46; q -= 0.08) {
        const d = c.toDataURL("image/jpeg", q);
        if (d.length <= maxLen) return { dataUrl: d, w, h };
      }
      scale *= 0.78;
    }
    throw new Error("Foto terlalu besar untuk diproses");
  } finally { URL.revokeObjectURL(url); }
}

/* ---------------- state ---------------- */
const DEFAULT_SETTINGS = {
  due: { Critical: 7, Major: 14, Minor: 30 },
  pics: ["GA", "Branch Head", "Workshop / Service", "Warehouse / Part", "Mobile Workshop / Field Service", "HRGA", "IT", "Procurement", "QSHE", "Remanufacturing", "Project Site"],
  areas: [["Head Office Jakarta", "HO"], ["Cabang Jakarta", "JKT"], ["Cabang Balikpapan", "BPN"], ["Cabang Surabaya", "SBY"], ["Cabang Palembang", "PLG"], ["Cabang Pontianak", "PNK"], ["Site KPC Sangatta", "KPC"]],
  agenda: "Inspeksi Internal QSHE Dept",
  supervisors: [], qsheName: "Rosiana Rizky Amelia", onedrive: "", qsheDevice: false, qshePin: ""
};
const S = {
  inspections: [], findings: [], settings: clone(DEFAULT_SETTINGS),
  tab: "catat", currentInsp: LS.get("qf:currentInsp") || "",
  photoCache: new Map(), me: { id: null, name: "", owner: false }
};
function mergeSettings(x) {
  const s = clone(DEFAULT_SETTINGS); x = x || {};
  if (x.due) s.due = { ...s.due, ...x.due };
  if (Array.isArray(x.pics) && x.pics.length) s.pics = x.pics;
  if (Array.isArray(x.areas) && x.areas.length) s.areas = x.areas;
  if (x.agenda) s.agenda = x.agenda;
  if (Array.isArray(x.supervisors)) s.supervisors = x.supervisors;
  if (x.qsheName) s.qsheName = x.qsheName;
  if (typeof x.onedrive === "string") s.onedrive = x.onedrive;
  if (typeof x.qsheDevice === "boolean") s.qsheDevice = x.qsheDevice;
  if (typeof x.qshePin === "string") s.qshePin = x.qshePin;
  return s;
}

/* ---------------- storage adapter (db capability, local fallback) ---------------- */
function migrateRec(f) {
  if (f && Array.isArray(f.klausul_iso)) f.klausul_iso.forEach(x => { if (x.standar === "ISO 9001:2015") { x.standar = "ISO 9001:2026"; if (x.klausul === "10.3") x.klausul = "10.1"; if (x.klausul === "6.1") x.klausul = "6.1.1"; if (x.klausul === "5.1") x.klausul = "5.1.1"; } });
  return f;
}
const NOWRITE = new Set(["invalid_argument", "not_granted", "revoked", "not_writer", "permission_denied", "forbidden"]);
/* ---------------- penyimpanan perangkat (IndexedDB) ---------------- */
const IDB = {
  _db: null,
  open() {
    if (this._db) return this._db;
    this._db = new Promise((res, rej) => {
      const r = indexedDB.open("temuan-qshe", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("kv");
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
    return this._db;
  },
  async tx(mode, fn) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction("kv", mode); const st = t.objectStore("kv"); let out; const r = fn(st); if (r) r.onsuccess = () => { out = r.result; }; t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error || new Error("Penyimpanan penuh")); }); },
  get(k) { return this.tx("readonly", st => st.get(k)); },
  set(k, v) { return this.tx("readwrite", st => st.put(v, k)); },
  del(k) { return this.tx("readwrite", st => st.delete(k)); }
};
const Store = {
  mode: "local", readOnly: false, ready: null,
  init() {
    this.ready = (async () => {
      try { if (navigator.storage && navigator.storage.persist) await navigator.storage.persist(); } catch (e) {}
      let ins = await IDB.get("inspections"), fin = await IDB.get("findings"), set = await IDB.get("settings");
      if (!ins && !fin && (LS.get("qf:inspections") || LS.get("qf:findings"))) {        /* pindahkan data lama dari localStorage */
        ins = LS.get("qf:inspections") || []; fin = LS.get("qf:findings") || []; set = LS.get("qf:settings");
        await IDB.set("inspections", ins); await IDB.set("findings", fin); if (set) await IDB.set("settings", set);
        for (const f of fin) for (const p of [...(f.photos || []), ...(f.closePhotos || [])]) { const d = LS.get("qf:ph:" + p.id); if (d) await IDB.set("ph:" + p.id, d); }
      }
      S.inspections = ins || []; S.findings = (fin || []).map(migrateRec); S.settings = mergeSettings(set);
      S.deleted = (await IDB.get("deleted")) || [];
      S.me.owner = !!S.settings.qsheDevice;
      await Folder.init();
      onData(); updateStoreStatus();
    })();
    return this.ready;
  },
  async _save(col) {
    try { await IDB.set(col, clone(col === "inspections" ? S.inspections : S.findings)); }
    catch (e) { throw new Error("Penyimpanan perangkat penuh. Hapus foto/temuan lama atau pindahkan ke folder."); }
    Folder.schedule();
  },
  async put(col, obj) {
    await this.ready;
    const data = clone(obj); data.updatedAt = Date.now();
    const arr = col === "inspections" ? S.inspections : S.findings;
    const i = arr.findIndex(x => x.id === data.id); if (i >= 0) arr[i] = data; else arr.push(data);
    await this._save(col); onData();
  },
  async del(col, id) {
    await this.ready;
    const key = col === "inspections" ? "inspections" : "findings";
    S[key] = S[key].filter(x => x.id !== id);
    S.deleted = (S.deleted || []).filter(x => x.id !== id).concat([{ id, at: Date.now() }]);
    await IDB.set("deleted", S.deleted).catch(() => {});
    await this._save(col); onData();
  },
  async putSettings(obj) { await this.ready; await IDB.set("settings", clone(obj)); S.settings = mergeSettings(obj); S.me.owner = !!S.settings.qsheDevice; onData(); Folder.schedule(); },
  async putPhoto(id, dataUrl) {
    await this.ready;
    try { await IDB.set("ph:" + id, dataUrl); } catch (e) { throw new Error("Penyimpanan perangkat penuh — foto tidak tersimpan."); }
    S.photoCache.set(id, dataUrl); Folder.queuePhoto(id);
  },
  async getPhoto(id) {
    if (!id) return null;
    if (S.photoCache.has(id)) return S.photoCache.get(id);
    await this.ready; const d = await IDB.get("ph:" + id).catch(() => null);
    if (d) S.photoCache.set(id, d); return d || null;
  },
  async delPhoto(id) { await this.ready; S.photoCache.delete(id); await IDB.del("ph:" + id).catch(() => {}); }
};

/* ---------------- folder penyimpanan (Drive lokal / folder OneDrive di laptop) ---------------- */
const Folder = {
  supported: typeof window.showDirectoryPicker === "function",
  handle: null, state: "none", _t: null, _photos: new Set(),
  async init() {
    if (!this.supported) { this.state = "unsupported"; return; }
    try { this.handle = (await IDB.get("folder")) || null; } catch (e) {}
    await this.check(false);
    if (this.state === "ok") this.backupAll(false).catch(() => {});
    window.addEventListener("focus", () => { if (this.state === "ok") this.backupAll(false).catch(() => {}); });
  },
  async check(ask) {
    if (!this.handle) { this.state = "none"; updateStoreStatus(); return false; }
    const o = { mode: "readwrite" };
    try {
      let p = await this.handle.queryPermission(o);
      if (p !== "granted" && ask) p = await this.handle.requestPermission(o);
      this.state = p === "granted" ? "ok" : "ask";
    } catch (e) { this.state = "ask"; }
    updateStoreStatus(); return this.state === "ok";
  },
  async syncNow(silent) {
    if (!(await this.check(!silent))) return false;
    try { await this.backupAll(false); if (!silent) toast("Data tersinkron dengan folder " + this.handle.name + "."); return true; }
    catch (e) { if (!silent) toast("Sinkron gagal: " + (e.message || e), "err"); return false; }
  },
  async pick() {
    const h = await window.showDirectoryPicker({ id: "temuan-qshe", mode: "readwrite" });
    this.handle = h; await IDB.set("folder", h); await this.check(true);
    if (this.state === "ok") { await this.backupAll(true); toast(`Folder "${h.name}" terhubung. Database dan foto disalin ke folder ini.`); }
  },
  async forget() { this.handle = null; this.state = "none"; await IDB.del("folder"); updateStoreStatus(); },
  async dir(path, create = true) { let d = this.handle; for (const part of path.split("/").filter(Boolean)) d = await d.getDirectoryHandle(part, { create }); return d; },
  async write(path, blob) {
    const parts = path.split("/"); const name = parts.pop();
    const d = await this.dir(parts.join("/"));
    const fh = await d.getFileHandle(name, { create: true }); const w = await fh.createWritable(); await w.write(blob); await w.close();
  },
  async exists(path) { try { const parts = path.split("/"); const name = parts.pop(); const d = await this.dir(parts.join("/"), false); await d.getFileHandle(name); return true; } catch (e) { return false; } },
  schedule() { if (this.state !== "ok") return; clearTimeout(this._t); this._t = setTimeout(() => this.backupAll(false).catch(e => { this.state = "ask"; updateStoreStatus(); }), 2500); },
  queuePhoto(id) { this._photos.add(id); this.schedule(); },
  dbJSON() { return JSON.stringify({ format: "hexindo-temuan-qshe", versi: 2, jenis: "database", dibuat: new Date().toISOString(), inspections: S.inspections, findings: S.findings, deleted: S.deleted || [], settings: (({ qshePin, qsheDevice, ...x }) => x)(S.settings) }); },
  /* Gabungkan isi folder (bisa diisi banyak laptop lewat OneDrive/SharePoint) dengan data perangkat ini */
  async pull() {
    let fh; try { fh = await (await this.dir("Database", false)).getFileHandle("database_temuan.json"); } catch (e) { return 0; }
    let p; try { p = JSON.parse(await (await fh.getFile()).text()); } catch (e) { return 0; }
    if (!p || p.format !== "hexindo-temuan-qshe") return 0;
    const tomb = new Map(); [...(S.deleted || []), ...(p.deleted || [])].forEach(t => { if (!tomb.has(t.id) || tomb.get(t.id) < t.at) tomb.set(t.id, t.at); });
    S.deleted = [...tomb].map(([id, at]) => ({ id, at }));
    const merge = (a, b) => { const m = new Map(a.map(x => [x.id, x])); b.forEach(x => { const o = m.get(x.id); if (!o || (x.updatedAt || 0) > (o.updatedAt || 0)) m.set(x.id, x); });
      return [...m.values()].filter(x => !(tomb.has(x.id) && tomb.get(x.id) >= (x.updatedAt || 0))); };
    const before = JSON.stringify([S.inspections.length, S.findings.length, S.findings.map(f => f.updatedAt)]);
    S.inspections = merge(S.inspections, p.inspections || []); S.findings = merge(S.findings, (p.findings || []).map(migrateRec));
    await IDB.set("inspections", clone(S.inspections)); await IDB.set("findings", clone(S.findings)); await IDB.set("deleted", S.deleted);
    let fotoDir = null; try { fotoDir = await this.dir("Database/foto", false); } catch (e) {}
    if (fotoDir) for (const f of S.findings) for (const ph of [...(f.photos || []), ...(f.closePhotos || [])]) {
      if (await IDB.get("ph:" + ph.id)) continue;
      try { const file = await (await fotoDir.getFileHandle(ph.id + ".jpg")).getFile(); const d = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(file); }); await IDB.set("ph:" + ph.id, d); } catch (e) {}
    }
    const changed = before !== JSON.stringify([S.inspections.length, S.findings.length, S.findings.map(f => f.updatedAt)]);
    if (changed) onData();
    return changed;
  },
  async backupAll(full) {
    if (this.state !== "ok") return;
    await this.pull();
    await this.write("Database/database_temuan.json", new Blob([this.dbJSON()], { type: "application/json" }));
    const ids = full ? S.findings.flatMap(f => [...(f.photos || []), ...(f.closePhotos || [])].map(p => p.id)) : [...this._photos];
    this._photos.clear();
    for (const id of ids) {
      const path = `Database/foto/${id}.jpg`;
      if (full && await this.exists(path)) continue;
      const d = await Store.getPhoto(id); if (d) await this.write(path, dataUrlToBlob(d));
    }
    this.last = Date.now(); updateStoreStatus();
  },
  async restore() {
    if (!(await this.check(true))) throw new Error("Izin folder belum diberikan.");
    let fh; try { fh = await (await this.dir("Database", false)).getFileHandle("database_temuan.json"); } catch (e) { throw new Error("File Database/database_temuan.json tidak ditemukan di folder ini."); }
    const p = JSON.parse(await (await fh.getFile()).text());
    if (!p || p.format !== "hexindo-temuan-qshe") throw new Error("Isi database tidak dikenali.");
    let fotoDir = null; try { fotoDir = await this.dir("Database/foto", false); } catch (e) {}
    let nPh = 0;
    for (const f of p.findings || []) for (const ph of [...(f.photos || []), ...(f.closePhotos || [])]) {
      if (await IDB.get("ph:" + ph.id)) continue;
      if (!fotoDir) continue;
      try { const file = await (await fotoDir.getFileHandle(ph.id + ".jpg")).getFile(); const d = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(file); }); await IDB.set("ph:" + ph.id, d); nPh++; } catch (e) {}
    }
    const byId = (a, b) => { const m = new Map(a.map(x => [x.id, x])); b.forEach(x => { const o = m.get(x.id); if (!o || (x.updatedAt || 0) >= (o.updatedAt || 0)) m.set(x.id, x); }); return [...m.values()]; };
    S.inspections = byId(S.inspections, p.inspections || []); S.findings = byId(S.findings, (p.findings || []).map(migrateRec));
    await IDB.set("inspections", clone(S.inspections)); await IDB.set("findings", clone(S.findings));
    if (p.settings) { const keep = { qshePin: S.settings.qshePin, qsheDevice: S.settings.qsheDevice }; await Store.putSettings({ ...p.settings, ...keep }); }
    onData(); return { nI: (p.inspections || []).length, nF: (p.findings || []).length, nPh };
  }
};
function dbMsg(e) {
  const c = e && e.code;
  if (c === "quota_exceeded" || c === "resource_exhausted") return "Kuota penyimpanan penuh. Hapus foto/temuan lama yang tidak diperlukan.";
  if (c === "unavailable") return "Koneksi penyimpanan terputus. Coba lagi sebentar.";
  if (c === "not_granted" || c === "revoked") return "Akses penyimpanan tidak diizinkan untuk akun ini.";
  return (e && e.message) || "Gagal menyimpan.";
}
function dbErr(e) { toast(dbMsg(e), "err"); }
document.addEventListener("click", e => {
  if (!e.target.closest("#storeStatus")) return;
  if (Folder.state === "ask") Folder.check(true).then(ok => { if (ok) { Folder.backupAll(false); toast("Izin folder diberikan — penyalinan otomatis aktif lagi."); } });
  else if (Folder.state !== "ok") $("#btnSettings").click();
});
function updateStoreStatus() {
  const el = $("#storeStatus"); if (!el) return; const dot = el.querySelector(".dot"); const tx = el.querySelector("span");
  dot.className = "dot"; el.style.cursor = "pointer";
  if (Folder.state === "ok") { dot.classList.add("ok"); tx.textContent = `Folder: ${Folder.handle.name}`; el.title = "Data tersimpan di perangkat dan disalin otomatis ke folder " + Folder.handle.name + (Folder.last ? " (terakhir " + new Date(Folder.last).toLocaleTimeString("id-ID") + ")" : ""); }
  else if (Folder.state === "ask") { dot.classList.add("warn"); tx.textContent = "Ketuk untuk izinkan folder"; el.title = "Browser meminta izin ulang untuk menulis ke folder " + (Folder.handle ? Folder.handle.name : ""); }
  else { dot.classList.add("warn"); tx.textContent = "Tersimpan di perangkat ini"; el.title = Folder.supported ? "Belum ada folder penyimpanan. Atur di Pengaturan → Folder penyimpanan." : "Data tersimpan di perangkat ini. File laporan tersimpan ke folder Download."; }
}

/* ---------------- AI (sample capability) ---------------- */
const AI = { fn: null, img: null, ready: null };
AI.ready = (async () => {
  AI.fn = null; /* Fitur AI dimatikan permanen: aplikasi 100% gratis, tidak memakai kuota/akun Claude siapa pun */
  if (AI.fn) { try { const l = await AI.fn.limits(); AI.img = l && l.images ? l.images : null; } catch (e) { AI.img = null; } }
  updateAINote();
})();
function updateAINote() {
  const n = $("#aiNote"); if (!n) return;
  const bA = $("#btnAI"), bR = $("#btnRapi"), bC = $("#btnConc");
  if (AI.fn) {
    n.textContent = (AI.img ? "Rapikan jadi laporan (AI) membaca catatan dan foto, lalu menyusun deskripsi baku, klausul, regulasi, tingkat temuan, dan tindak lanjut." : "Rapikan jadi laporan (AI) membaca catatan Anda.") + " Rapikan tanpa AI bekerja di perangkat, tanpa kuota, dan tetap jalan saat sinyal hilang. Semua hasil bisa diedit.";
    bA.hidden = false; bA.disabled = false; bR.className = "btn"; if (bC) bC.textContent = "Susun dengan AI";
  } else {
    n.textContent = "Gratis tanpa AI: bahasa catatan dirapikan dan klausul ISO, kriteria SMK3, serta regulasi dipilih otomatis dari kata kunci. Berjalan di perangkat, tanpa internet. Periksa hasilnya sebelum dikirim.";
    bA.hidden = true; bR.className = "btn brand"; if (bC) bC.textContent = "Susun otomatis";
  }
}
function aiErrMsg(e) {
  const c = e && e.code;
  return ({
    not_granted: "Izin AI ditolak. Muat ulang halaman lalu izinkan saat diminta, atau gunakan Isi manual.",
    sampling_disabled: "AI tidak aktif untuk akun ini. Gunakan Isi manual.",
    rate_limited: "Batas penggunaan AI tercapai. Coba lagi beberapa saat lagi.",
    refused: "AI menolak memproses entri ini. Ubah catatan lalu coba lagi.",
    invalid_json: "Hasil AI tidak terbaca utuh. Tekan tombol lagi, atau persingkat catatan.",
    image_rejected: "Foto tidak bisa dibaca AI. Coba foto lain.",
    session_expired: "Sesi berakhir. Masuk ulang ke claude.ai.",
    prompt_too_large: "Catatan terlalu panjang. Persingkat lalu coba lagi.",
    cancelled: "Dibatalkan."
  })[c] || "AI sedang bermasalah. Coba lagi, atau gunakan Isi manual.";
}

let _kbText = null;
function kbText() {
  if (_kbText) return _kbText;
  const iso = (arr, name) => name + "\n" + arr.map(([k, t]) => "  " + k + " " + t).join("\n");
  let smk3 = "KRITERIA AUDIT SMK3 — Lampiran II PP 50/2012\n";
  let lastE = "", lastS = "";
  for (const [k, t] of KB.crit) {
    const [e, s] = k.split(".");
    if (e !== lastE) { smk3 += "Elemen " + e + " " + (KB.elem[e] || "") + "\n"; lastE = e; }
    const sk = e + "." + s; if (sk !== lastS) { smk3 += " " + sk + " " + (KB.subs[sk] || "") + "\n"; lastS = sk; }
    smk3 += "  " + k + " " + t + "\n";
  }
  const regs = "DAFTAR REGULASI (tulis nama persis)\n" + KB.regs.map(([n, c]) => "- " + n + " — cakupan: " + c).join("\n");
  _kbText = [iso(KB.iso45, "ISO 45001:2018 (K3)"), iso(KB.iso14, "ISO 14001:2026 (Lingkungan)"), iso(KB.iso9, "ISO 9001:2026 (Mutu)"), smk3, regs].join("\n\n");
  return _kbText;
}

function findingPrompt(inp) {
  const good = inp.jenis === "good";
  return `Anda adalah Lead Auditor QSHE senior di PT Hexindo Adiperkasa Tbk (distributor dan layanan purna jual alat berat Hitachi: workshop cabang, site tambang customer, gudang part, kantor, kendaraan service/mobile workshop di jalan hauling). Tugas Anda: mengubah catatan kasar inspeksi lapangan menjadi satu entri laporan inspeksi yang baku, objektif, dan siap dipresentasikan ke manajemen.

DATA INPUT
- Jenis entri: ${good ? "GOOD PRACTICE (praktik K3L yang baik)" : "TEMUAN (ketidaksesuaian / kondisi atau tindakan tidak aman)"}
- Area inspeksi: ${inp.area || "-"}
- Lokasi spesifik: ${inp.lokasi || "-"}
- Tanggal inspeksi: ${inp.tanggal || "-"}
- Catatan kasar inspector: """${inp.raw}"""
- Foto: ${inp.nImg ? inp.nImg + " foto terlampir — amati kondisi yang terlihat jelas" : "tidak ada foto yang dikirim"}

ATURAN PENULISAN
1. Bahasa Indonesia baku sesuai EYD, kalimat efektif, sudut pandang orang ketiga. Jangan menyalahkan individu dan jangan menyebut nama orang.
2. Susunan deskripsi: lokasi → kondisi aktual yang diamati (fakta) → persyaratan yang tidak terpenuhi → potensi risiko atau dampak. 3–5 kalimat, maksimal 90 kata. Contoh gaya: "Pada area workshop Jakarta, tepatnya di rak penyimpanan safety shoes, ditemukan beberapa sepatu dengan sol yang telah getas dan robek. APD yang tidak layak pakai tersebut masih disimpan bercampur dengan stok yang layak. Kondisi ini berpotensi menyebabkan pekerja tergelincir atau mengalami cedera kaki akibat menurunnya fungsi perlindungan APD."
3. Jangan mengarang fakta. Gunakan hanya informasi dari catatan dan hal yang jelas terlihat di foto. Jangan menambahkan angka, nomor unit, atau durasi yang tidak disebut. Hal penting yang belum diketahui tuliskan di "catatan_konfirmasi".
4. Istilah teknis yang lazim (APD, APAR, LOTO, housekeeping, TPS LB3, jack stand, sling, spill kit) boleh dipakai.
5. judul: ringkas, maksimal 10 kata, menggambarkan kondisi (bukan rekomendasi).

PEMETAAN STANDAR
- klausul_iso: 1–3 klausul HANYA dari daftar di bawah. ISO 45001:2018 untuk aspek K3, ISO 14001:2026 untuk aspek lingkungan, ISO 9001:2026 untuk mutu/proses/dokumen. Pilih klausul yang persyaratannya benar-benar tidak terpenuhi, bukan sekadar mirip topik.
- kriteria_smk3: 1–3 kode kriteria audit SMK3 (3 tingkat, contoh "6.5.3") HANYA dari daftar. Untuk temuan murni lingkungan/mutu tanpa aspek K3 boleh kosong.
- regulasi: 0–3 regulasi HANYA dari daftar regulasi, nama ditulis persis. Isi "pasal" hanya jika Anda yakin nomor pasal/ayatnya; jika ragu kosongkan. Jika regulasi paling relevan tidak ada di daftar, boleh ditambahkan dengan "perlu_verifikasi": true.

TINGKAT TEMUAN (kriteria perusahaan)
- Critical: potensi cedera berat/fatal atau LTI; ATAU pelanggaran regulasi pemerintah/perizinan yang berisiko sanksi; ATAU potensi kerusakan/pencemaran lingkungan yang signifikan. Pekerjaan/area wajib dihentikan atau diisolasi segera (stop_work = true).
- Major: potensi atau aktual cedera ringan; ATAU temuan/peringatan regulator yang tidak fatal; ATAU ketidaksesuaian proses yang berdampak pada operasional.
- Minor: selain Critical dan Major (dampak minimal, observasi, celah kecil dokumentasi).
Nilai berdasarkan konsekuensi terburuk yang wajar (credible worst case), bukan skenario ekstrem. alasan_severity: 1 kalimat yang menyebut kriteria yang terpenuhi.

TINDAK LANJUT
- tindakan_segera: koreksi/containment yang bisa dilakukan hari itu juga (maks. 30 kata).
- tindakan_korektif: 1–3 tindakan untuk menghilangkan akar masalah. Tiap tindakan diberi level hierarki pengendalian: Eliminasi, Substitusi, Rekayasa, Administratif, atau APD. Utamakan level tertinggi yang realistis. Tiap tindakan maks. 25 kata, diawali kata kerja, spesifik dan dapat diverifikasi.
- tindakan_pencegahan: perbaikan sistemik agar tidak terulang di area/cabang lain (maks. 30 kata).
- akar_masalah: dugaan awal (kategori salah satu dari: Manusia, Metode, Mesin/Peralatan, Material, Lingkungan, Manajemen), diawali "Dugaan awal:".
- bukti_penutupan: bukti objektif yang diminta untuk menutup temuan.
- pic_saran: pilih satu fungsi penanggung jawab dari daftar: ${S.settings.pics.join(", ")}. Jika tidak ada yang cocok, tulis fungsi yang paling tepat.
- kategori_aspek: satu atau lebih dari "K3", "Lingkungan", "Mutu".
- kategori_temuan: satu dari "Compliance" (terkait kepatuhan regulasi/perizinan), "Operational" (kondisi/praktik di lapangan), "Administration" (dokumen/catatan/prosedur).
${good ? `
KHUSUS GOOD PRACTICE: deskripsi apresiatif namun objektif (apa yang baik, mengapa efektif, catatan kecil peningkatan bila ada). severity diisi "-", stop_work false, tindakan_segera dan tindakan_korektif dikosongkan, saran peningkatan (bila ada) di tindakan_pencegahan, klausul/kriteria = persyaratan yang telah dipenuhi dengan baik.` : ""}

FORMAT JAWABAN — balas HANYA satu objek JSON, tanpa teks lain:
{"judul":"","deskripsi":"","bahaya":"","potensi_risiko":"","kategori_aspek":["K3"],"kategori_temuan":"Operational","klausul_iso":[{"standar":"ISO 45001:2018","klausul":"8.1.2","judul":"","alasan":""}],"kriteria_smk3":[{"kode":"6.1.1","ringkas":""}],"regulasi":[{"peraturan":"","pasal":"","relevansi":"","perlu_verifikasi":false}],"severity":"Major","alasan_severity":"","stop_work":false,"akar_masalah":{"kategori":"Metode","uraian":"Dugaan awal: "},"tindakan_segera":"","tindakan_korektif":[{"hierarki":"Administratif","tindakan":""}],"tindakan_pencegahan":"","bukti_penutupan":"","pic_saran":"","catatan_konfirmasi":""}

==== REFERENSI ====
${kbText()}`;
}

function normalizeAI(r, jenis) {
  const o = r && typeof r === "object" ? r : {};
  const arr = v => Array.isArray(v) ? v : [];
  const str = v => (v == null ? "" : String(v)).trim();
  const smk3Map = new Map(KB.crit);
  const out = {
    judul: str(o.judul), deskripsi: str(o.deskripsi), bahaya: str(o.bahaya), potensi_risiko: str(o.potensi_risiko),
    kategori_aspek: arr(o.kategori_aspek).map(str).filter(x => ["K3", "Lingkungan", "Mutu"].includes(x)),
    kategori_temuan: ["Compliance", "Operational", "Administration"].includes(o.kategori_temuan) ? o.kategori_temuan : "Operational",
    klausul_iso: arr(o.klausul_iso).map(x => ({ standar: str(x.standar), klausul: str(x.klausul).replace(/^klausul\s*/i, ""), judul: str(x.judul), alasan: str(x.alasan) })).filter(x => x.klausul),
    kriteria_smk3: arr(o.kriteria_smk3).map(x => { const k = str(x.kode); return { kode: k, ringkas: smk3Map.get(k) ? smk3Map.get(k) : str(x.ringkas) }; }).filter(x => x.kode),
    regulasi: arr(o.regulasi).map(x => ({ peraturan: str(x.peraturan), pasal: str(x.pasal), relevansi: str(x.relevansi), perlu_verifikasi: !!x.perlu_verifikasi || !KB.regs.some(([n]) => n === str(x.peraturan)) })).filter(x => x.peraturan),
    severity: jenis === "good" ? "" : (SEVS.includes(o.severity) ? o.severity : "Minor"),
    alasan_severity: str(o.alasan_severity), stop_work: jenis === "good" ? false : !!o.stop_work,
    akar_masalah: { kategori: RC_CAT.includes(o.akar_masalah && o.akar_masalah.kategori) ? o.akar_masalah.kategori : "Metode", uraian: str(o.akar_masalah && o.akar_masalah.uraian) },
    tindakan_segera: str(o.tindakan_segera),
    tindakan_korektif: arr(o.tindakan_korektif).map(x => ({ hierarki: HIER.includes(x.hierarki) ? x.hierarki : "Administratif", tindakan: str(x.tindakan) })).filter(x => x.tindakan),
    tindakan_pencegahan: str(o.tindakan_pencegahan), bukti_penutupan: str(o.bukti_penutupan),
    pic_saran: str(o.pic_saran), catatan_konfirmasi: str(o.catatan_konfirmasi)
  };
  if (!out.kategori_aspek.length) out.kategori_aspek = ["K3"];
  return out;
}

async function runFindingAI(inp, photos, signal, onProgress) {
  await AI.ready;
  if (!AI.fn) throw { code: "not_granted" };
  const opts = { modelTier: "default", signal, onText: ({ text }) => onProgress && onProgress(text.length) };
  let n = 0;
  if (AI.img && photos.length) {
    const max = Math.max(1, AI.img.maxCount || 1);
    const blobs = [];
    for (const p of photos.slice(0, max)) { const d = p.dataUrl || await Store.getPhoto(p.id); if (d) blobs.push(dataUrlToBlob(d)); }
    if (blobs.length) { opts.images = blobs; n = blobs.length; }
  }
  const res = await AI.fn.json(findingPrompt({ ...inp, nImg: n }), opts);
  return normalizeAI(res, inp.jenis);
}

function conclusionPrompt(insp, items, goods, tracked) {
  const f = items.map((x, i) => `${i + 1}. [${x.severity}] ${x.judul} — ${x.deskripsi} (PIC: ${x.pic || "-"}; tindakan segera: ${x.tindakan_segera || "-"})`).join("\n");
  const g = goods.map((x, i) => `${i + 1}. ${x.judul} — ${x.deskripsi}`).join("\n");
  const t = tracked.map((x, i) => `${i + 1}. ${x.judul} — status ${x.status}`).join("\n");
  return `Anda adalah Lead Auditor QSHE PT Hexindo Adiperkasa Tbk. Susun bagian KESIMPULAN laporan inspeksi untuk manajemen dalam Bahasa Indonesia baku.

Inspeksi: ${insp.agenda || "Inspeksi QSHE"} — ${insp.area} — ${fmtTgl(insp.tanggal, true)}
Temuan periode ini:
${f || "(tidak ada temuan)"}
Good practice:
${g || "(tidak ada)"}
Status temuan inspeksi sebelumnya di area yang sama:
${t || "(tidak ada data)"}

Ketentuan:
- "kesimpulan": 2 paragraf (total 90–150 kata). Paragraf 1: gambaran umum penerapan K3L termasuk aspek yang sudah baik. Paragraf 2: aspek kritikal yang memerlukan tindakan korektif segera, sebutkan tindakan paling mendesak. Pisahkan paragraf dengan "\\n\\n". Jangan menambah fakta di luar data.
- "rekomendasi": 3–5 rekomendasi tindakan perbaikan yang sistemik, masing-masing {"judul": frasa 2–5 kata, "uraian": maks. 25 kata}.
Balas HANYA JSON: {"kesimpulan":"","rekomendasi":[{"judul":"","uraian":""}]}`;
}

/* ---------------- aplikasi terpasang (PWA) ---------------- */
if ("serviceWorker" in navigator && /^https?:$/.test(location.protocol)) navigator.serviceWorker.register("sw.js").catch(() => {});
let installEvt = null;
window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); installEvt = e; const b = $("#btnInstall"); if (b) b.hidden = false; });
document.addEventListener("click", async e => {
  if (!e.target.closest("#btnInstall") || !installEvt) return;
  installEvt.prompt(); await installEvt.userChoice.catch(() => {}); installEvt = null; $("#btnInstall").hidden = true;
});
/* ======================= UI ======================= */
const REVIEW_LABEL = { Draft: "Draft", Review: "Menunggu Supervisor", QSHE: "Menunggu QSHE", Final: "Final" };

function areaCode(area) {
  const hit = S.settings.areas.find(a => a[0].toLowerCase() === String(area || "").toLowerCase());
  if (hit && hit[1]) return hit[1];
  const w = String(area || "").replace(/\b(cabang|site|project|kantor|area)\b/gi, "").trim().split(/\s+/).filter(Boolean);
  const base = (w[w.length - 1] || "INS").replace(/[aeiou]/gi, "").toUpperCase();
  return (base.length >= 3 ? base : (w[w.length - 1] || "INS").toUpperCase()).slice(0, 3);
}
function inspById(id) { return S.inspections.find(i => i.id === id); }
function inspLabel(i) { return i ? `${i.area} — ${fmtShort(i.tanggal)}` : "-"; }
function sortedInsps() { return [...S.inspections].sort((a, b) => (b.tanggal || "").localeCompare(a.tanggal || "") || (b.createdAt || 0) - (a.createdAt || 0)); }
function isOpen(f) { return f.jenis !== "good" && f.status !== "Closed"; }
function isLate(f) { return isOpen(f) && f.due && f.due < localISO(); }
function statusFlag(f) {
  if (f.jenis === "good") return "";
  if (isLate(f)) return `<span class="flag late">Lewat ${daysBetween(f.due, localISO())} hari</span>`;
  const c = f.status === "Closed" ? "closed" : f.status === "On Progress" ? "prog" : "open";
  return `<span class="flag ${c}">${esc(f.status || "Open")}</span>`;
}
function reviewFlag(f) { const r = f.review || "Draft"; return `<span class="flag ${r === "Final" ? "final" : r === "Review" || r === "QSHE" ? "rev" : ""}">${REVIEW_LABEL[r] || r}</span>`; }
function sevTag(f) { return f.jenis === "good" ? `<span class="tag Good">Good practice</span>` : `<span class="tag ${esc(f.severity || "None")}">${esc(f.severity || "-")}</span>`; }

function fcardHTML(f) {
  const ph = f.photos && f.photos[0] ? `<img alt="" data-ph="${esc(f.photos[0].id)}">` : "Tanpa foto";
  return `<button class="fcard" data-open="${esc(f.id)}">
    <div class="thumb">${ph}</div>
    <div>
      <div class="row" style="gap:8px"><span class="id">${esc(f.no || "")}</span>${sevTag(f)}${reviewFlag(f)}${statusFlag(f)}</div>
      <div class="t">${esc(f.judul || "(tanpa judul)")}</div>
      <div class="meta"><span>${esc(f.area || "")}${f.lokasi ? ", " + esc(f.lokasi) : ""}</span>${f.jenis !== "good" ? (f.pic || f.due ? `<span>PIC ${esc(f.pic || "-")}</span><span>Batas ${f.due ? fmtShort(f.due) : "-"}</span>` : `<span>Tindak lanjut belum diisi</span>`) : ""}</div>
    </div></button>`;
}
function hydrateThumbs(root) {
  $$("img[data-ph]", root).forEach(img => { if (img.src) return; Store.getPhoto(img.dataset.ph).then(d => { if (d) img.src = d; }); });
}

/* ---------------- editor (report sheet) ---------------- */
const isoSrc = std => String(std).startsWith("ISO 45001") ? KB.iso45 : String(std).startsWith("ISO 14001") ? KB.iso14 : KB.iso9;
const isoList = std => String(std).startsWith("ISO 45001") ? "dlIso45" : String(std).startsWith("ISO 14001") ? "dlIso14" : "dlIso9";
const STD_OPTS = ["ISO 45001:2018", "ISO 14001:2026", "ISO 9001:2026"];
function optList(arr, v) { return arr.map(x => `<option ${x === v ? "selected" : ""}>${esc(x)}</option>`).join(""); }
function getPath(o, k) { return k.split(".").reduce((a, p) => a == null ? a : a[p], o); }
function setPath(o, k, v) { const ps = k.split("."); let a = o; ps.slice(0, -1).forEach(p => { a[p] = a[p] || {}; a = a[p]; }); a[ps[ps.length - 1]] = v; }

function editorHTML(w, o) {
  const good = w.jenis === "good";
  const sevSeg = SEVS.map(v => `<button type="button" data-kseg="severity" data-v="${v}" aria-pressed="${w.severity === v}">${v}</button>`).join("");
  const aspek = ["K3", "Lingkungan", "Mutu"].map(v => `<button type="button" class="chip" data-aspek="${v}" aria-pressed="${(w.kategori_aspek || []).includes(v)}">${v}</button>`).join("");
  const kl = (w.klausul_iso || []).map((x, i) => `<div class="ref" style="grid-template-columns:158px 78px 1fr auto">
      <select class="in" data-arr="klausul_iso" data-i="${i}" data-f="standar">${optList(STD_OPTS, x.standar)}</select>
      <input class="in" list="${isoList(x.standar)}" data-arr="klausul_iso" data-i="${i}" data-f="klausul" value="${esc(x.klausul)}" placeholder="8.1.2">
      <input class="in" data-arr="klausul_iso" data-i="${i}" data-f="judul" value="${esc(x.judul)}" placeholder="Judul klausul">
      <button type="button" class="x" data-del="klausul_iso:${i}" aria-label="Hapus">×</button></div>`).join("");
  const sm = (w.kriteria_smk3 || []).map((x, i) => `<div class="ref" style="grid-template-columns:90px 1fr auto">
      <input class="in" list="dlSmk3" data-arr="kriteria_smk3" data-i="${i}" data-f="kode" value="${esc(x.kode)}" placeholder="6.5.3">
      <input class="in" data-arr="kriteria_smk3" data-i="${i}" data-f="ringkas" value="${esc(x.ringkas)}" placeholder="Bunyi kriteria">
      <button type="button" class="x" data-del="kriteria_smk3:${i}" aria-label="Hapus">×</button></div>`).join("");
  const rg = (w.regulasi || []).map((x, i) => `<div class="ref" style="grid-template-columns:1fr 120px auto auto">
      <input class="in" list="dlReg" data-arr="regulasi" data-i="${i}" data-f="peraturan" value="${esc(x.peraturan)}" placeholder="Nama peraturan">
      <input class="in" data-arr="regulasi" data-i="${i}" data-f="pasal" value="${esc(x.pasal)}" placeholder="Pasal (opsional)">
      ${x.perlu_verifikasi ? `<span class="verify" title="Tidak ada di daftar regulasi acuan — cek ulang nama & nomornya">Cek ulang</span>` : `<span></span>`}
      <button type="button" class="x" data-del="regulasi:${i}" aria-label="Hapus">×</button></div>`).join("");
  const ca = (w.tindakan_korektif || []).map((x, i) => `<div class="ref" style="grid-template-columns:140px 1fr auto">
      <select class="in" data-arr="tindakan_korektif" data-i="${i}" data-f="hierarki">${optList(HIER, x.hierarki)}</select>
      <input class="in" data-arr="tindakan_korektif" data-i="${i}" data-f="tindakan" value="${esc(x.tindakan)}" placeholder="Tindakan korektif">
      <button type="button" class="x" data-del="tindakan_korektif:${i}" aria-label="Hapus">×</button></div>`).join("");
  const phStrip = key => `<div class="photos">${(w[key] || []).map((p, i) => `<div class="ph"><img alt="" data-ph="${esc(p.id)}"><button type="button" data-phdel="${key}:${i}" aria-label="Hapus foto">×</button></div>`).join("")}
      ${(w[key] || []).length < 4 ? `<label class="ph-add"><input type="file" accept="image/*" multiple class="sr" data-phadd="${key}"><span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>Tambah</span></label>` : ""}</div>`;

  return `<article class="sheet">
    <div class="sheet-head">
      <div class="sheet-no"><small>${good ? "Good practice" : "Nomor temuan"}</small><span>${esc(w.no || "Baru")}</span></div>
      <div class="sheet-title"><input data-k="judul" value="${esc(w.judul)}" placeholder="Judul temuan" aria-label="Judul temuan"></div>
    </div>
    <div class="sheet-body">
      ${w.catatan_konfirmasi ? `<div class="note"><b>Perlu dikonfirmasi:</b> ${esc(w.catatan_konfirmasi)}</div>` : ""}
      ${good ? "" : `<section class="sec"><h4>Tingkat temuan</h4>
        <div class="row"><div class="seg sev">${sevSeg}</div>
        <label class="row small" style="font-weight:600;color:var(--crit)"><input type="checkbox" data-k="stop_work" ${w.stop_work ? "checked" : ""}> Hentikan pekerjaan / isolasi area segera</label></div>
        <input class="in" data-k="alasan_severity" value="${esc(w.alasan_severity)}" placeholder="Alasan penentuan tingkat"></section>`}
      ${o.mode === "edit" ? `<section class="sec"><h4>Foto temuan</h4>${phStrip("photos")}</section>` : ""}
      <section class="sec"><h4>${good ? "Uraian good practice" : "Deskripsi temuan"}</h4>
        <label class="f">Lokasi spesifik<input class="in" data-k="lokasi" value="${esc(w.lokasi)}"></label>
        <textarea class="in" data-k="deskripsi" rows="5" aria-label="Deskripsi">${esc(w.deskripsi)}</textarea>
        <div class="row" style="gap:16px"><button type="button" class="addlink" data-tool="bahasa">Rapikan bahasa (tanpa AI)</button><button type="button" class="addlink" data-tool="klausul">Cari klausul & regulasi dari deskripsi</button></div>
        ${good ? "" : `<label class="f">Potensi risiko / dampak<textarea class="in" data-k="potensi_risiko" rows="2">${esc(w.potensi_risiko)}</textarea></label>`}
        <div class="two"><div class="f"><span>Aspek</span><div class="chips">${aspek}</div></div>
        <label class="f">Kategori temuan<select class="in" data-k="kategori_temuan">${optList(["Compliance", "Operational", "Administration"], w.kategori_temuan)}</select></label></div>
      </section>
      <section class="sec"><h4>Standar, klausul & regulasi</h4>
        <div class="small muted">Klausul ISO</div><div class="refs">${kl}</div><button type="button" class="addlink" data-add="klausul_iso">+ Tambah klausul ISO</button>
        <div class="small muted">Kriteria audit SMK3 — PP 50/2012</div><div class="refs">${sm}</div><button type="button" class="addlink" data-add="kriteria_smk3">+ Tambah kriteria SMK3</button>
        <div class="small muted">Regulasi terkait</div><div class="refs">${rg}</div><button type="button" class="addlink" data-add="regulasi">+ Tambah regulasi</button>
      </section>
      ${o.mode === "new" ? "" : good ? `<section class="sec"><h4>Saran peningkatan</h4><textarea class="in" data-k="tindakan_pencegahan" rows="2">${esc(w.tindakan_pencegahan)}</textarea></section>` : `
      <section class="sec"><h4>Akar masalah (dugaan awal)</h4>
        <div class="ref" style="grid-template-columns:160px 1fr"><select class="in" data-k="akar_masalah.kategori"><option value="">— pilih —</option>${optList(RC_CAT, (w.akar_masalah || {}).kategori)}</select>
        <input class="in" data-k="akar_masalah.uraian" value="${esc((w.akar_masalah || {}).uraian)}"></div></section>
      <section class="sec"><h4>Tindak lanjut</h4>
        <label class="f">Tindakan segera (koreksi hari ini)<textarea class="in" data-k="tindakan_segera" rows="2">${esc(w.tindakan_segera)}</textarea></label>
        <div class="small muted">Tindakan korektif — urut hierarki pengendalian</div><div class="refs">${ca}</div><button type="button" class="addlink" data-add="tindakan_korektif">+ Tambah tindakan korektif</button>
        <label class="f">Tindakan pencegahan (sistemik)<textarea class="in" data-k="tindakan_pencegahan" rows="2">${esc(w.tindakan_pencegahan)}</textarea></label>
        <label class="f">Bukti yang diminta untuk penutupan<input class="in" data-k="bukti_penutupan" value="${esc(w.bukti_penutupan)}"></label>
      </section>
      <section class="sec"><h4>Penanggung jawab <span class="small muted" style="font-family:var(--font);font-weight:400">— diisi saat tindak lanjut</span></h4>
        <div class="two"><label class="f">PIC<input class="in" list="dlPic" data-k="pic" value="${esc(w.pic)}"></label>
        <label class="f">Batas waktu<input class="in" type="date" data-k="due" value="${esc(w.due)}"></label></div>
      </section>`}
      ${o.extraHTML ? o.extraHTML(w) : ""}
    </div>
    <div class="sheet-foot">${o.footHTML(w)}</div>
  </article>`;
}

function renderEditor(host, rec, o) {
  const w = clone(rec);
  const draw = () => { host.innerHTML = editorHTML(w, o); bind(); hydrateThumbs(host); };
  function bind() {
    $$("[data-k]", host).forEach(el => {
      const ev = el.tagName === "SELECT" || el.type === "checkbox" || el.type === "date" ? "change" : "input";
      el.addEventListener(ev, () => {
        setPath(w, el.dataset.k, el.type === "checkbox" ? el.checked : el.value);
        if (el.dataset.k === "due") w.dueManual = true;
        if (el.dataset.k === "status" && el.value === "Closed" && !w.closedDate) { w.closedDate = localISO(); draw(); }
      });
    });
    $$("[data-arr]", host).forEach(el => {
      const ev = el.tagName === "SELECT" ? "change" : "input";
      el.addEventListener(ev, () => {
        const row = w[el.dataset.arr][+el.dataset.i]; row[el.dataset.f] = el.value;
        if (el.dataset.arr === "kriteria_smk3" && el.dataset.f === "kode") {
          const hit = KB.crit.find(c => c[0] === el.value.trim());
          if (hit) { row.ringkas = hit[1]; const sib = el.parentElement.querySelector('[data-f="ringkas"]'); if (sib) sib.value = hit[1]; }
        }
        if (el.dataset.arr === "klausul_iso" && el.dataset.f === "klausul") {
          const hit = isoSrc(row.standar).find(c => c[0] === el.value.trim());
          if (hit) { row.judul = hit[1]; const sib = el.parentElement.querySelector('[data-f="judul"]'); if (sib) sib.value = hit[1]; }
        }
        if (el.dataset.arr === "klausul_iso" && el.dataset.f === "standar") { row.klausul = ""; row.judul = ""; draw(); return; }
        if (el.dataset.arr === "regulasi" && el.dataset.f === "peraturan") row.perlu_verifikasi = !KB.regs.some(([n]) => n === el.value.trim());
      });
    });
    $$("[data-kseg]", host).forEach(b => b.addEventListener("click", () => {
      const k = b.dataset.kseg; w[k] = b.dataset.v;
      if (k === "severity") w.stop_work = w.severity === "Critical";
      if (k === "status" && w.status === "Closed" && !w.closedDate) w.closedDate = localISO();
      if (k === "status" && w.status !== "Closed") w.closedDate = "";
      draw();
    }));
    $$("[data-aspek]", host).forEach(b => b.addEventListener("click", () => {
      const v = b.dataset.aspek; const a = new Set(w.kategori_aspek || []); a.has(v) ? a.delete(v) : a.add(v); w.kategori_aspek = [...a]; b.setAttribute("aria-pressed", a.has(v));
    }));
    $$("[data-del]", host).forEach(b => b.addEventListener("click", () => { const [k, i] = b.dataset.del.split(":"); w[k].splice(+i, 1); draw(); }));
    $$("[data-add]", host).forEach(b => b.addEventListener("click", () => {
      const k = b.dataset.add; w[k] = w[k] || [];
      w[k].push({ klausul_iso: { standar: "ISO 45001:2018", klausul: "", judul: "", alasan: "" }, kriteria_smk3: { kode: "", ringkas: "" }, regulasi: { peraturan: "", pasal: "", relevansi: "", perlu_verifikasi: false }, tindakan_korektif: { hierarki: "Administratif", tindakan: "" } }[k]);
      draw(); const last = $$(`[data-arr="${k}"]`, host).filter(e => e.tagName === "INPUT").slice(-1)[0]; if (last) last.focus();
    }));
    $$("[data-phdel]", host).forEach(b => b.addEventListener("click", () => { const [k, i] = b.dataset.phdel.split(":"); w[k].splice(+i, 1); draw(); }));
    $$("[data-phadd]", host).forEach(inp => inp.addEventListener("change", async () => {
      const k = inp.dataset.phadd; const files = [...inp.files].slice(0, 4 - (w[k] || []).length);
      for (const f of files) {
        try { const c = await compressImage(f); const id = uid("ph_"); await Store.putPhoto(id, c.dataUrl); (w[k] = w[k] || []).push({ id, w: c.w, h: c.h }); }
        catch (e) { toast(e.message || "Foto gagal diproses", "err"); }
      }
      draw();
    }));
    $$("[data-tool]", host).forEach(b => b.addEventListener("click", () => {
      if (b.dataset.tool === "bahasa") {
        let n = 0; const fix = v => { const r = RAPI.bahasa(v); n += r.changes.length; return r.text; };
        ["deskripsi", "potensi_risiko", "tindakan_segera", "tindakan_pencegahan", "bukti_penutupan", "remarks"].forEach(k => { if (String(w[k] || "").trim()) w[k] = fix(w[k]); });
        if (w.akar_masalah && String(w.akar_masalah.uraian || "").trim()) w.akar_masalah.uraian = fix(w.akar_masalah.uraian);
        (w.tindakan_korektif || []).forEach(c => { if (String(c.tindakan || "").trim()) c.tindakan = fix(c.tindakan).replace(/\.$/, ""); });
        draw(); toast(n ? `${n} kata/penulisan dirapikan.` : "Tidak ada kata tidak baku yang ditemukan.");
      } else {
        const n = RAPI.tambahRujukan(w); draw();
        toast(n ? `${n} rujukan ditambahkan dari kata kunci deskripsi — cek kesesuaiannya.` : "Tidak ada kata kunci yang dikenali. Isi rujukan secara manual.", n ? "" : "err");
      }
    }));
    if (o.onBind) o.onBind(host, w, draw);
  }
  draw();
  return { get: () => w, draw };
}

/* ---------------- capture view: unggah banyak foto -> isi keterangan -> simpan semua ---------------- */
function baseRecord(insp, o) {
  o = o || {};
  const jenis = o.jenis || "temuan";
  return {
    id: uid("f_"), inspId: insp.id, no: "", jenis, area: insp.area, tanggal: insp.tanggal,
    lokasi: o.lokasi || "", raw: o.raw || "", photos: [], closePhotos: [],
    judul: "", deskripsi: "", bahaya: "", potensi_risiko: "", kategori_aspek: ["K3"], kategori_temuan: "Operational",
    klausul_iso: [], kriteria_smk3: [], regulasi: [], severity: jenis === "good" ? "" : "Minor", alasan_severity: "", stop_work: false,
    akar_masalah: { kategori: "", uraian: "" }, tindakan_segera: "", tindakan_korektif: [], tindakan_pencegahan: "", bukti_penutupan: "",
    pic: "", due: "", dueManual: false, status: "Open", closedDate: "", remarks: "", review: "Draft", reviewNote: "", catatan_konfirmasi: "",
    createdAt: Date.now(), createdBy: S.me.id || null
  };
}
function nextNo(insp, jenis) {
  const same = S.findings.filter(f => f.inspId === insp.id && f.jenis === jenis);
  let max = 0; same.forEach(f => { const m = String(f.no || "").match(/(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  const n = Math.max(max, same.length) + 1;
  const d = (insp.tanggal || localISO()).replace(/-/g, "");
  return `${insp.kode || areaCode(insp.area)}-${d}-${jenis === "good" ? "GP" : ""}${String(n).padStart(2, "0")}`;
}

const MAX_BATCH = 80, MIN_NOTE = 8;
/* Cap.items: satu elemen per foto yang sudah diunggah. Disimpan di memori sampai ditekan "Simpan semua". */
const Cap = { items: [], busy: false, saving: false };
function currentInsp() { return inspById(S.currentInsp); }
function requireInsp() {
  const i = currentInsp(); if (i) return i;
  toast("Buat atau pilih inspeksi dulu, supaya temuan tercatat di area & tanggal yang benar.", "err"); openInspModal(); return null;
}

function setUploadProgress(done, total, failed) {
  const box = $("#upProgress");
  if (total == null) { box.innerHTML = ""; return; }
  const pct = total ? Math.round(done / total * 100) : 0;
  box.innerHTML = `<div class="up-prog"><div class="row between"><span>${done < total ? `Mengunggah foto ${done} dari ${total}…` : `${total - (failed || 0)} foto siap`}${failed ? ` <span style="color:var(--crit)">(${failed} gagal)</span>` : ""}</span><span class="muted">${pct}%</span></div><div class="track"><i style="width:${pct}%"></i></div></div>`;
}
async function ingestFiles(fileList) {
  if (Cap.busy || Cap.saving) return;
  if (!requireInsp()) return;
  const imgs = [...fileList].filter(f => /^image\//.test(f.type) || /\.(jpe?g|png|webp|heic|heif)$/i.test(f.name));
  if (!imgs.length) { toast("Tidak ada file foto yang dipilih.", "err"); return; }
  const room = MAX_BATCH - Cap.items.length;
  if (room <= 0) { toast(`Maksimal ${MAX_BATCH} foto per kali simpan. Simpan dulu foto yang sudah ada.`, "err"); return; }
  const files = imgs.slice(0, room);
  if (imgs.length > room) toast(`Hanya ${room} foto pertama yang dimasukkan (maksimal ${MAX_BATCH} per kali simpan).`, "err");
  Cap.busy = true; $("#dropZone").classList.add("busy");
  let done = 0, failed = 0; setUploadProgress(0, files.length);
  for (const f of files) {
    try { const c = await compressImage(f); Cap.items.push({ key: uid("n_"), photo: c, jenis: "temuan", lokasi: "", raw: "", merge: false }); }
    catch (err) { failed++; console.error(err); }
    done++; setUploadProgress(done, files.length, failed);
  }
  Cap.busy = false; $("#dropZone").classList.remove("busy");
  if (failed) toast(`${failed} foto gagal diproses (format tidak didukung atau terlalu besar).`, "err");
  renderNotes(); /* form keterangan baru muncul SETELAH semua foto selesai diunggah */
  const first = $("#stepNotes"); if (!first.hidden) first.scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth", block: "start" });
}
$("#filePhoto").addEventListener("change", e => { const fl = [...e.target.files]; e.target.value = ""; ingestFiles(fl); });
(() => {
  const dz = $("#dropZone");
  ["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("over"); }));
  dz.addEventListener("drop", e => { if (e.dataTransfer && e.dataTransfer.files) ingestFiles(e.dataTransfer.files); });
})();

function noteCardHTML(it, i) {
  const good = it.jenis === "good";
  const canMerge = i > 0;
  return `<article class="ncard ${it.merge ? "merged" : ""}" data-key="${esc(it.key)}">
    <div class="ncard-ph"><img alt="Foto ${i + 1}" src="${it.photo.dataUrl}"><span class="num">${i + 1}</span><button type="button" data-rm aria-label="Hapus foto ${i + 1}">×</button></div>
    <div class="ncard-body">
      ${it.merge ? `<div class="mnote">Foto ini dipakai sebagai foto tambahan pada entri di atasnya.</div>` : `
      <div class="ntop"><div class="seg sm" data-jenis><button type="button" data-v="temuan" aria-pressed="${!good}">Temuan</button><button type="button" data-v="good" aria-pressed="${good}">Good practice</button></div></div>
      <input class="in" data-f="lokasi" value="${esc(it.lokasi)}" placeholder="Lokasi spesifik — contoh: rak safety shoes, workshop" autocomplete="off" aria-label="Lokasi foto ${i + 1}">
      <textarea class="in" data-f="raw" rows="3" placeholder="${good ? "Catatan singkat — contoh: toolbox di lemari no. 1 sudah dikelompokkan per jenis & terkunci" : "Catatan singkat — contoh: sepatu safety di rak solnya banyak yg getas & robek, masih dicampur sama yg bagus"}" aria-label="Catatan foto ${i + 1}">${esc(it.raw)}</textarea>`}
      ${canMerge ? `<label class="mrg"><input type="checkbox" data-merge ${it.merge ? "checked" : ""}> Foto tambahan untuk entri di atasnya</label>` : ""}
    </div></article>`;
}
function updateNotesHint() {
  const act = Cap.items.filter(x => !x.merge);
  const filled = act.filter(x => String(x.raw).trim().length >= MIN_NOTE).length;
  $("#notesCount").textContent = `— ${Cap.items.length} foto`;
  $("#notesHint").textContent = act.length === filled ? `${act.length} entri siap dirapikan.` : `${filled} dari ${act.length} entri sudah ada keterangannya.`;
}
function renderNotes() {
  const box = $("#stepNotes");
  box.hidden = !Cap.items.length;
  if (!Cap.items.length) { $("#noteList").innerHTML = ""; return; }
  $("#noteList").innerHTML = Cap.items.map(noteCardHTML).join("");
  updateNotesHint();
}
$("#noteList").addEventListener("input", e => {
  const card = e.target.closest(".ncard"); if (!card) return;
  const it = Cap.items.find(x => x.key === card.dataset.key); if (!it) return;
  const f = e.target.dataset.f; if (f) { it[f] = e.target.value; card.classList.remove("bad"); updateNotesHint(); }
});
$("#noteList").addEventListener("change", e => {
  const card = e.target.closest(".ncard"); if (!card) return;
  const it = Cap.items.find(x => x.key === card.dataset.key); if (!it) return;
  if (e.target.matches("[data-merge]")) { it.merge = e.target.checked; renderNotes(); }
});
$("#noteList").addEventListener("click", e => {
  const card = e.target.closest(".ncard"); if (!card) return;
  const idx = Cap.items.findIndex(x => x.key === card.dataset.key); if (idx < 0) return;
  const jb = e.target.closest("[data-jenis] button");
  if (jb) { Cap.items[idx].jenis = jb.dataset.v; renderNotes(); return; }
  if (e.target.closest("[data-rm]")) {
    Cap.items.splice(idx, 1);
    if (Cap.items[0]) Cap.items[0].merge = false; /* foto pertama tidak bisa jadi foto tambahan */
    renderNotes(); setUploadProgress(); 
  }
});
$("#btnClearAll").onclick = () => { if (Cap.items.length && confirm(`Buang ${Cap.items.length} foto yang belum disimpan?`)) { Cap.items = []; renderNotes(); setUploadProgress(); } };
window.addEventListener("beforeunload", e => { if (Cap.items.length && !Cap.saving) { e.preventDefault(); e.returnValue = ""; } });

/* kelompokkan foto: foto bertanda "tambahan" ikut ke entri di atasnya (maks. 4 foto per entri) */
function groupItems() {
  const groups = [];
  Cap.items.forEach(it => {
    const g = groups[groups.length - 1];
    if (it.merge && g && g.photos.length < 4) g.photos.push(it.photo); else groups.push({ ...it, photos: [it.photo] });
  });
  return groups;
}
async function saveAll(review, btn) {
  if (Cap.saving || Cap.busy) return;
  const insp = requireInsp(); if (!insp) return;
  const groups = groupItems();
  const bad = groups.filter(g => String(g.raw).trim().length < MIN_NOTE);
  if (bad.length) {
    $$(".ncard.bad").forEach(c => c.classList.remove("bad"));
    bad.forEach(g => { const c = $(`.ncard[data-key="${g.key}"]`); if (c) c.classList.add("bad"); });
    const c0 = $(`.ncard[data-key="${bad[0].key}"]`); if (c0) { c0.scrollIntoView({ block: "center" }); const t = c0.querySelector("textarea"); if (t) t.focus({ preventScroll: true }); }
    toast(`${bad.length} foto belum ada keterangannya (minimal satu kalimat singkat).`, "err"); return;
  }
  Cap.saving = true; const all = $$("#stepNotes button, #stepNotes input, #stepNotes textarea"); all.forEach(x => x.disabled = true);
  const lbl = btn.innerHTML; const lokAll = $("#inLokasiAll").value.trim();
  let n = 0, fail = 0, last = ""; const okG = [];
  try {
    for (const g of groups) {
      btn.textContent = `Menyimpan ${n + fail + 1} dari ${groups.length}…`;
      try {
        const rec = baseRecord(insp, { jenis: g.jenis, lokasi: String(g.lokasi).trim() || lokAll, raw: String(g.raw).trim() });
        Object.assign(rec, RAPI.susun({ jenis: rec.jenis, area: insp.area, lokasi: rec.lokasi, raw: rec.raw }).fields);
        const photos = [];
        for (const p of g.photos) { const id = uid("ph_"); await Store.putPhoto(id, p.dataUrl); photos.push({ id, w: p.w, h: p.h }); }
        rec.photos = photos; rec.review = review; rec.no = nextNo(insp, rec.jenis);
        if (review === "Review") rec.sentAt = Date.now();
        await Store.put("findings", rec); n++; okG.push(g); last = rec.no;
      } catch (err) { fail++; console.error(err); last = err.message || last; }
    }
  } finally { Cap.saving = false; }
  if (fail) {
    /* yang gagal tetap di daftar supaya bisa dicoba lagi: buang hanya yang sudah tersimpan */
    toast(`${n} entri tersimpan, ${fail} gagal${last ? " — " + last : ""}. Yang gagal masih ada di daftar; coba simpan lagi.`, "err");
    const saved = okG.flatMap(g => g.photos);
    Cap.items = Cap.items.filter(it => !saved.includes(it.photo));
    if (Cap.items[0]) Cap.items[0].merge = false;
    renderNotes(); btn.innerHTML = lbl; all.forEach(x => x.disabled = false);
  } else {
    Cap.items = []; renderNotes(); setUploadProgress(); $("#inLokasiAll").value = "";
    btn.innerHTML = lbl; all.forEach(x => x.disabled = false);
    toast(review === "Review" ? `${n} entri dirapikan dan dikirim ke Supervisor.` : `${n} entri dirapikan dan tersimpan sebagai draft. Ketuk entri untuk memeriksa atau mengubahnya.`);
    const sp = $("#sessionPanel"); if (sp) sp.scrollIntoView({ block: "start" });
  }
}
$("#btnSaveAll").onclick = e => saveAll("Draft", e.currentTarget);
$("#btnSaveSpv").onclick = e => saveAll("Review", e.currentTarget);

function renderInspSelect() {
  const sel = $("#selInsp"); const list = sortedInsps();
  sel.innerHTML = `<option value="">— pilih inspeksi —</option>` + list.map(i => `<option value="${esc(i.id)}" ${i.id === S.currentInsp ? "selected" : ""}>${esc(inspLabel(i))}</option>`).join("");
  const i = currentInsp();
  $("#inspInfo").innerHTML = i ? `<div class="insp-card"><div class="row between"><b>${esc(i.area)}</b><button class="btn ghost small" id="btnEditInsp" style="min-height:30px;padding:4px 8px">Ubah</button></div>
    <div class="small">${esc(fmtTgl(i.tanggal, true))}</div><div class="small muted">${esc(i.agenda || "")}</div>
    <div class="small muted">Inspector: ${esc((i.inspectors || []).join(", ") || "-")}</div></div>`
    : `<div class="empty small">Belum ada inspeksi dipilih. Buat inspeksi baru untuk mulai mencatat.</div>`;
  const be = $("#btnEditInsp"); if (be) be.onclick = () => openInspModal(i);
}
$("#selInsp").onchange = e => { S.currentInsp = e.target.value; LS.set("qf:currentInsp", S.currentInsp); renderInspSelect(); renderSession(); };

function renderSession() {
  const i = currentInsp(); const box = $("#sessList");
  if (!i) { $("#sessTitle").textContent = "Entri di inspeksi ini"; $("#sessCount").textContent = ""; box.innerHTML = `<div class="empty">Pilih inspeksi untuk melihat entri yang sudah dicatat.</div>`; return; }
  const items = S.findings.filter(f => f.inspId === i.id).sort((a, b) => (a.jenis === "good") - (b.jenis === "good") || String(a.no).localeCompare(String(b.no)));
  $("#sessTitle").textContent = `Entri di ${i.area}`;
  const nT = items.filter(f => f.jenis !== "good").length;
  $("#sessCount").textContent = `${nT} temuan, ${items.length - nT} good practice`;
  box.innerHTML = items.length ? items.map(fcardHTML).join("") : `<div class="empty">Belum ada entri. Unggah semua foto kondisi di lapangan, isi keterangan singkat tiap foto, lalu tekan <b>Rapikan &amp; simpan semua</b>.</div>`;
  hydrateThumbs(box);
}

/* ---------------- inspection modal ---------------- */
let editingInsp = null;
function openModal(id) { const m = $(id); m.hidden = false; const f = m.querySelector("input,select,textarea,button"); if (f) setTimeout(() => f.focus(), 30); }
function closeModal(m) { m.hidden = true; }
$$(".modal").forEach(m => {
  m.addEventListener("click", e => { if (e.target === m || e.target.closest("[data-close]")) closeModal(m); });
});
document.addEventListener("keydown", e => { if (e.key === "Escape") $$(".modal").forEach(m => { if (!m.hidden) closeModal(m); }); });

function openInspModal(i) {
  editingInsp = i || null;
  $("#mInspTitle").textContent = i ? "Ubah inspeksi" : "Inspeksi baru";
  $("#dlArea").innerHTML = S.settings.areas.map(a => `<option value="${esc(a[0])}">`).join("");
  $("#iArea").value = i ? i.area : ""; $("#iKode").value = i ? (i.kode || "") : "";
  $("#iTgl").value = i ? i.tanggal : localISO(); $("#iAgenda").value = i ? (i.agenda || "") : S.settings.agenda;
  $("#iInspector").value = i ? (i.inspectors || []).join(", ") : (S.me.name || "");
  $("#iInspectee").value = i ? (i.inspectees || []).join(", ") : "";
  openModal("#mInsp");
}
$("#iArea").addEventListener("input", () => { if (!editingInsp || !$("#iKode").value) $("#iKode").value = areaCode($("#iArea").value); });
$("#btnNewInsp").onclick = () => openInspModal();
$("#btnSaveInsp").onclick = async () => {
  const area = $("#iArea").value.trim(), tgl = $("#iTgl").value;
  if (!area || !tgl) { toast("Area dan tanggal wajib diisi.", "err"); return; }
  const split = s => s.split(",").map(x => x.trim()).filter(Boolean);
  const rec = { ...(editingInsp || { id: uid("i_"), createdAt: Date.now(), createdBy: S.me.id || null }),
    area, kode: ($("#iKode").value.trim() || areaCode(area)).toUpperCase(), tanggal: tgl, agenda: $("#iAgenda").value.trim(),
    inspectors: split($("#iInspector").value), inspectees: split($("#iInspectee").value) };
  try {
    await Store.put("inspections", rec);
    if (editingInsp) { for (const f of S.findings.filter(x => x.inspId === rec.id && (x.area !== rec.area || x.tanggal !== rec.tanggal))) await Store.put("findings", { ...f, area: rec.area, tanggal: rec.tanggal }); }
    S.currentInsp = rec.id; LS.set("qf:currentInsp", rec.id);
    closeModal($("#mInsp")); renderInspSelect(); renderSession(); toast(editingInsp ? "Inspeksi diperbarui." : "Inspeksi dibuat. Silakan mulai mencatat.");
  } catch (e) { toast(e.message, "err"); }
};

/* ---------------- register ---------------- */
function renderRegister() {
  const fi = $("#fInsp"); const cur = fi.value;
  fi.innerHTML = `<option value="">Semua inspeksi</option>` + sortedInsps().map(i => `<option value="${esc(i.id)}" ${i.id === cur ? "selected" : ""}>${esc(inspLabel(i))}</option>`).join("");
  const q = $("#fQ").value.trim().toLowerCase(), sev = $("#fSev").value, st = $("#fStatus").value, rv = $("#fReview").value;
  let list = S.findings.filter(f => {
    if (fi.value && f.inspId !== fi.value) return false;
    if (sev === "good" && f.jenis !== "good") return false;
    if (sev && sev !== "good" && f.severity !== sev) return false;
    if (st === "late" && !isLate(f)) return false;
    if (st && st !== "late" && (f.jenis === "good" || (f.status || "Open") !== st)) return false;
    if (rv && (f.review || "Draft") !== rv) return false;
    if (q) { const hay = [f.no, f.judul, f.deskripsi, f.lokasi, f.area, f.pic, ...(f.klausul_iso || []).map(k => k.standar + " " + k.klausul), ...(f.regulasi || []).map(r => r.peraturan)].join(" ").toLowerCase(); if (!hay.includes(q)) return false; }
    return true;
  });
  const sevRank = { Critical: 0, Major: 1, Minor: 2 };
  list.sort((a, b) => (isOpen(b) - isOpen(a)) || ((sevRank[a.severity] ?? 3) - (sevRank[b.severity] ?? 3)) || (b.tanggal || "").localeCompare(a.tanggal || ""));
  $("#regList").innerHTML = list.length ? list.map(fcardHTML).join("") : `<div class="empty">${S.findings.length ? "Tidak ada temuan yang cocok dengan filter." : "Register masih kosong. Catat temuan pertama di tab Catat."}</div>`;
  hydrateThumbs($("#regList"));
}
["#fQ", "#fInsp", "#fSev", "#fStatus", "#fReview"].forEach(s => $(s).addEventListener(s === "#fQ" ? "input" : "change", renderRegister));
function updateReviewCount() {
  const nS = S.findings.filter(f => f.review === "Review").length, nQ = S.findings.filter(f => f.review === "QSHE").length;
  const n = S.me.owner ? nQ : nS + nQ; const c = $("#cntReview");
  c.hidden = !n; c.textContent = n; c.title = `${nS} menunggu Supervisor, ${nQ} menunggu QSHE`;
}

/* ---------------- detail modal ---------------- */
let Detail = null;
document.addEventListener("click", e => { const b = e.target.closest("[data-open]"); if (b) openDetail(b.dataset.open); });
function openDetail(id) {
  const rec = S.findings.find(f => f.id === id); if (!rec) return;
  const host = $("#detailHost");
  host.innerHTML = `<button class="icon-btn modal-close" data-close aria-label="Tutup">×</button><div id="dEd"></div>`;
  const orig = clone(rec);
  Detail = renderEditor($("#dEd"), rec, {
    mode: "edit",
    extraHTML: w => w.jenis === "good" ? reviewHTML(w) : `
      <section class="sec" style="border-top:4px solid var(--ink);padding-top:14px"><h4>Status tindak lanjut</h4>
        <div class="row"><div class="seg">${["Open", "On Progress", "Closed"].map(v => `<button type="button" data-kseg="status" data-v="${v}" aria-pressed="${(w.status || "Open") === v}">${v}</button>`).join("")}</div>
        ${w.status === "Closed" ? `<label class="f" style="flex-direction:row">Tanggal ditutup <input class="in" type="date" data-k="closedDate" value="${esc(w.closedDate)}" style="width:auto"></label>` : ""}</div>
        <label class="f">Progres / keterangan<textarea class="in" data-k="remarks" rows="2">${esc(w.remarks)}</textarea></label>
        <div class="small muted">Foto bukti penutupan</div>
        <div class="photos">${(w.closePhotos || []).map((p, i) => `<div class="ph"><img alt="" data-ph="${esc(p.id)}"><button type="button" data-phdel="closePhotos:${i}" aria-label="Hapus foto">×</button></div>`).join("")}
        ${(w.closePhotos || []).length < 4 ? `<label class="ph-add"><input type="file" accept="image/*" multiple class="sr" data-phadd="closePhotos"><span>Tambah bukti</span></label>` : ""}</div>
      </section>${reviewHTML(w)}`,
    footHTML: w => `<button class="btn danger" data-act="del">Hapus</button><span style="flex:1"></span><button class="btn ghost" data-close>Tutup</button><button class="btn primary" data-act="save">Simpan perubahan</button>`,
    onBind: (h, w, draw) => {
      $$("[data-act]", h).forEach(b => b.onclick = async () => {
        const a = b.dataset.act;
        if (a === "del") { if (!confirm(`Hapus ${w.no}? Foto terkait juga dihapus.`)) return; try { await Store.del("findings", w.id); for (const p of [...(orig.photos || []), ...(orig.closePhotos || [])]) await Store.delPhoto(p.id); closeModal($("#mDetail")); toast("Temuan dihapus."); } catch (e) { toast(e.message, "err"); } return; }
        if (a === "toReview") { w.review = "Review"; w.sentAt = Date.now(); }
        if (a === "spvOk") {
          if (!String(w.spvName || "").trim()) { toast("Isi nama supervisor yang menyetujui.", "err"); return; }
          w.review = "QSHE"; w.spvAt = Date.now();
        }
        if (a === "final") {
          if (!S.me.owner) { toast("Persetujuan Final hanya dari perangkat QSHE (aktifkan di Pengaturan dengan PIN).", "err"); return; }
          w.review = "Final"; w.reviewedBy = S.me.id || null; w.reviewedAt = Date.now(); w.qsheName = S.settings.qsheName || "QSHE";
        }
        if (a === "reopen") { if (!S.me.owner) return; w.review = "QSHE"; w.reviewedAt = null; }
        if (a === "back" || a === "backSpv") {
          if (!String(w.reviewNote || "").trim()) { toast("Tulis catatan reviewer agar yang bersangkutan tahu apa yang perlu diperbaiki.", "err"); return; }
          if (a === "back") { w.review = "Draft"; w.spvName = w.spvName || ""; w.spvAt = null; } else { w.review = "Review"; w.spvAt = null; }
        }
        if (w.status === "Closed" && !w.closedDate) w.closedDate = localISO();
        try {
          b.disabled = true; await Store.put("findings", w);
          const keep = new Set([...(w.photos || []), ...(w.closePhotos || [])].map(p => p.id));
          for (const p of [...(orig.photos || []), ...(orig.closePhotos || [])]) if (!keep.has(p.id)) await Store.delPhoto(p.id);
          closeModal($("#mDetail"));
          toast({ final: `${w.no} disetujui QSHE — status Final.`, spvOk: `${w.no} disetujui Supervisor, diteruskan ke QSHE.`, back: `${w.no} dikembalikan ke pencatat.`, backSpv: `${w.no} dikembalikan ke Supervisor.`, reopen: `${w.no} dibuka kembali untuk review QSHE.`, toReview: `${w.no} dikirim ke Supervisor.` }[a] || "Perubahan tersimpan.");
        } catch (e) { b.disabled = false; toast(e.message, "err"); }
      });
    }
  });
  openModal("#mDetail");
}
function reviewHTML(w) {
  const r = w.review || "Draft";
  const dt = t => t ? fmtShort(localISO(new Date(t))) : "";
  const trail = `<div class="small" style="display:grid;gap:2px">
      <div>1. Supervisor: ${w.spvAt ? `<b>${esc(w.spvName)}</b> — disetujui ${dt(w.spvAt)}` : r === "Review" ? "menunggu" : "-"}</div>
      <div>2. QSHE: ${r === "Final" ? `<b>${esc(w.qsheName || S.settings.qsheName || "QSHE")}</b> — Final ${dt(w.reviewedAt)}` : r === "QSHE" ? "menunggu" : "-"}</div></div>`;
  let act = "";
  if (r === "Draft") act = `<button type="button" class="btn" data-act="toReview">Kirim ke Supervisor</button>`;
  if (r === "Review") act = `<label class="f" style="flex:1;min-width:200px">Nama supervisor yang menyetujui<input class="in" list="dlSpv" data-k="spvName" value="${esc(w.spvName)}" placeholder="Nama supervisor"></label>
      <button type="button" class="btn brand" data-act="spvOk">Setujui (Supervisor)</button><button type="button" class="btn" data-act="back">Kembalikan ke pencatat</button>`;
  if (r === "QSHE") act = S.me.owner
    ? `<button type="button" class="btn brand" data-act="final">Setujui jadi Final (QSHE)</button><button type="button" class="btn" data-act="backSpv">Kembalikan ke Supervisor</button><button type="button" class="btn" data-act="back">Kembalikan ke pencatat</button>`
    : `<span class="small muted">Menunggu persetujuan akhir QSHE — kirim paket data ke QSHE dari tab Laporan.</span>`;
  if (r === "Final" && S.me.owner) act = `<button type="button" class="btn" data-act="reopen">Buka kembali untuk review</button>`;
  return `<section class="sec" style="border-top:1px solid var(--line);padding-top:14px"><h4>Review: Supervisor → QSHE</h4>
    <div class="row">${reviewFlag(w)}</div>${trail}
    <label class="f">Catatan reviewer<textarea class="in" data-k="reviewNote" rows="2" placeholder="Masukan untuk pencatat, mis. klausul kurang tepat, foto kurang jelas">${esc(w.reviewNote)}</textarea></label>
    <div class="row" style="align-items:flex-end">${act}</div></section>`;
}

/* ---------------- settings ---------------- */
$("#btnSettings").onclick = () => {
  const s = S.settings;
  $("#sCrit").value = s.due.Critical; $("#sMajor").value = s.due.Major; $("#sMinor").value = s.due.Minor;
  $("#sPics").value = s.pics.join("\n"); $("#sAreas").value = s.areas.map(a => a[0] + (a[1] ? " | " + a[1] : "")).join("\n"); $("#sAgenda").value = s.agenda;
  $("#sSpv").value = (s.supervisors || []).join("\n"); $("#sQshe").value = s.qsheName || ""; $("#sOneDrive").value = s.onedrive || "";
  const th = LS.get("qf:theme") || "auto"; $$("#segTheme button").forEach(b => b.setAttribute("aria-pressed", b.dataset.v === th));
  renderFolderSettings(); openModal("#mSettings");
};
$$("#segTheme button").forEach(b => b.onclick = () => { LS.set("qf:theme", b.dataset.v); applyTheme(); $$("#segTheme button").forEach(x => x.setAttribute("aria-pressed", x === b)); });
function applyTheme() { const t = LS.get("qf:theme") || "auto"; if (t === "auto") document.documentElement.removeAttribute("data-theme"); else document.documentElement.setAttribute("data-theme", t); }
$("#btnSaveSettings").onclick = async () => {
  const n = (id, d) => Math.max(1, parseInt($(id).value, 10) || d);
  const obj = {
    due: { Critical: n("#sCrit", 7), Major: n("#sMajor", 14), Minor: n("#sMinor", 30) },
    pics: lines($("#sPics").value), areas: lines($("#sAreas").value).map(l => { const [a, k] = l.split("|").map(x => x.trim()); return [a, (k || "").toUpperCase()]; }),
    agenda: $("#sAgenda").value.trim() || DEFAULT_SETTINGS.agenda,
    supervisors: lines($("#sSpv").value), qsheName: $("#sQshe").value.trim() || DEFAULT_SETTINGS.qsheName,
    onedrive: /^https:\/\//i.test($("#sOneDrive").value.trim()) ? $("#sOneDrive").value.trim() : "",
    qsheDevice: !!S.settings.qsheDevice, qshePin: S.settings.qshePin || ""
  };
  try { await Store.putSettings(obj); S.settings = mergeSettings(obj); closeModal($("#mSettings")); renderDatalists(); toast("Pengaturan tersimpan."); } catch (e) { toast(e.message, "err"); }
};
function renderFolderSettings() {
  const i = $("#fldInfo");
  if (!Folder.supported) {
    i.innerHTML = "Browser ini belum bisa menghubungkan folder (umumnya HP Android/iPhone). File laporan tersimpan ke folder <b>Download</b>. Untuk menyatukan data, pakai <b>Unduh paket data</b> lalu impor di perangkat QSHE.";
    $("#fldBtns").hidden = true;
  } else {
    $("#fldBtns").hidden = false;
    i.innerHTML = Folder.handle
      ? `Terhubung ke folder <b>${esc(Folder.handle.name)}</b>${Folder.state === "ask" ? " — perlu izin ulang (ketuk status di kanan atas)" : ""}. Database dan foto disinkronkan otomatis ke subfolder <b>Database</b>; PowerPoint/Excel tersimpan di subfolder <b>Laporan</b>. Beberapa laptop yang memilih folder OneDrive/SharePoint yang sama akan saling berbagi data.`
      : "Belum ada folder. Pilih folder (mis. folder OneDrive di laptop) agar database, foto, dan laporan otomatis tersimpan di sana.";
    $("#btnForgetFolder").hidden = !Folder.handle; $("#btnRestoreFolder").hidden = !Folder.handle;
    $("#btnPickFolder").textContent = Folder.handle ? "Ganti folder" : "Pilih folder";
  }
  $("#qsheInfo").textContent = S.me.owner
    ? "Perangkat ini aktif sebagai perangkat QSHE: tombol persetujuan Final tersedia."
    : "Persetujuan Final hanya bisa dilakukan di perangkat QSHE. Aktifkan dengan PIN di perangkat milik QSHE saja.";
  $("#btnQsheMode").textContent = S.me.owner ? "Nonaktifkan mode QSHE" : "Aktifkan mode QSHE";
}
const pinHash = p => { let h = 5381; const t = "hexindo-qshe:" + p; for (let i = 0; i < t.length; i++) h = ((h << 5) + h + t.charCodeAt(i)) >>> 0; return "h" + h.toString(36); };
$("#btnPickFolder").onclick = async () => { try { await Folder.pick(); renderFolderSettings(); } catch (e) { if (e && e.name !== "AbortError") toast(e.message || "Folder gagal dipilih.", "err"); } };
$("#btnForgetFolder").onclick = async () => { if (!confirm("Lepas folder? Data di perangkat tetap ada, hanya penyalinan otomatis yang berhenti.")) return; await Folder.forget(); renderFolderSettings(); };
$("#btnRestoreFolder").onclick = async () => { await Folder.syncNow(false); renderFolderSettings(); };
$("#btnShareApp").onclick = async () => {
  const url = location.href.split("#")[0].replace(/index\.html$/, "");
  if (!/^https:/.test(url)) return toast("Aplikasi belum di-hosting. Unggah ke GitHub Pages dulu agar punya link untuk dibagikan.", "err");
  try { if (navigator.share) { await navigator.share({ title: "Temuan Lapangan QSHE", text: "Aplikasi Temuan Lapangan QSHE — buka lalu pasang di HP:", url }); return; } } catch (e) { if (e && e.name === "AbortError") return; }
  let ok = false; try { await navigator.clipboard.writeText(url); ok = true; } catch (e) {}
  toast(ok ? "Link aplikasi tersalin: " + url : "Link aplikasi: " + url, ok ? "" : "err");
};
$("#btnQsheMode").onclick = async () => {
  const st = S.settings;
  if (S.me.owner) { await Store.putSettings({ ...st, qsheDevice: false }); renderFolderSettings(); updateReviewCount(); return toast("Mode QSHE dinonaktifkan di perangkat ini."); }
  if (!st.qshePin) {
    const a = prompt("Buat PIN QSHE (minimal 4 angka):"); if (!a) return;
    if (!/^\d{4,}$/.test(a)) return toast("PIN minimal 4 angka.", "err");
    if (prompt("Ulangi PIN:") !== a) return toast("PIN tidak sama.", "err");
    await Store.putSettings({ ...st, qshePin: pinHash(a), qsheDevice: true });
  } else {
    const a = prompt("Masukkan PIN QSHE:"); if (!a) return;
    if (pinHash(a) !== st.qshePin) return toast("PIN salah.", "err");
    await Store.putSettings({ ...st, qsheDevice: true });
  }
  renderFolderSettings(); updateReviewCount(); toast("Mode QSHE aktif di perangkat ini.");
};
function renderDatalists() {
  $("#dlPic").innerHTML = S.settings.pics.map(p => `<option value="${esc(p)}">`).join("");
  $("#dlSpv").innerHTML = (S.settings.supervisors || []).map(p => `<option value="${esc(p)}">`).join("");
  if (!$("#dlSmk3")) {
    const d1 = document.createElement("datalist"); d1.id = "dlSmk3"; d1.innerHTML = KB.crit.map(([k, t]) => `<option value="${k}">${esc(t.slice(0, 90))}</option>`).join("");
    const d2 = document.createElement("datalist"); d2.id = "dlReg"; d2.innerHTML = KB.regs.map(([n]) => `<option value="${esc(n)}">`).join("");
    const mk = (id, arr) => { const d = document.createElement("datalist"); d.id = id; d.innerHTML = arr.map(([k, t]) => `<option value="${k}">${esc(t)}</option>`).join(""); return d; };
    document.body.append(d1, d2, mk("dlIso45", KB.iso45), mk("dlIso14", KB.iso14), mk("dlIso9", KB.iso9));
  }
}

/* ======================= DASHBOARD ======================= */
const Charts = {};
function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
function dashData() {
  const area = $("#dArea").value, aspek = $("#dAspek").value, per = $("#dPeriode").value;
  const from = per === "all" ? "" : addDays(localISO(), -Number(per));
  return S.findings.filter(f => f.jenis !== "good" && (!area || f.area === area) && (!aspek || (f.kategori_aspek || []).includes(aspek)) && (!from || (f.tanggal || "") >= from));
}
async function renderDashboard() {
  const areas = [...new Set(S.findings.map(f => f.area).filter(Boolean))].sort();
  const da = $("#dArea"), cur = da.value;
  da.innerHTML = `<option value="">Semua area</option>` + areas.map(a => `<option ${a === cur ? "selected" : ""}>${esc(a)}</option>`).join("");
  const data = dashData();
  const open = data.filter(isOpen), late = data.filter(isLate), closed = data.filter(f => f.status === "Closed");
  const critOpen = open.filter(f => f.severity === "Critical");
  const pct = data.length ? Math.round(closed.length / data.length * 100) : 0;
  $("#dashSub").textContent = `${data.length} temuan dalam filter ini, diperbarui ${new Date().toLocaleTimeString("id-ID", { hour: "2-digit", minute: "2-digit" })}`;
  $("#kpis").innerHTML = `
    <div class="kpi"><b>${data.length}</b><span>Total temuan</span></div>
    <div class="kpi"><b>${open.length}</b><span>Masih terbuka</span></div>
    <div class="kpi ${late.length ? "alert" : ""}"><b>${late.length}</b><span>Lewat batas waktu</span></div>
    <div class="kpi ${critOpen.length ? "alert" : ""}"><b>${critOpen.length}</b><span>Critical belum ditutup</span></div>
    <div class="kpi"><b>${pct}%</b><span>Tingkat penutupan</span><div class="bar-mini"><i style="width:${pct}%"></i></div></div>`;

  const lateRows = [...open].sort((a, b) => (a.due || "9").localeCompare(b.due || "9")).slice(0, 12);
  $("#tblLate").innerHTML = lateRows.length ? `<thead><tr><th>ID</th><th>Temuan</th><th>Tingkat</th><th>PIC</th><th>Batas waktu</th><th>Status</th></tr></thead><tbody>` +
    lateRows.map(f => `<tr class="click" data-open="${esc(f.id)}"><td>${esc(f.no)}</td><td>${esc(f.judul)}<div class="small muted">${esc(f.area)}</div></td><td>${sevTag(f)}</td><td>${esc(f.pic || "-")}</td><td>${fmtShort(f.due)}</td><td>${statusFlag(f)}</td></tr>`).join("") + "</tbody>"
    : `<tbody><tr><td class="muted">Tidak ada temuan terbuka. Semua tindak lanjut sudah ditutup.</td></tr></tbody>`;

  const byArea = {};
  data.forEach(f => { const a = byArea[f.area] = byArea[f.area] || { n: 0, C: 0, M: 0, m: 0, open: 0, late: 0, closed: 0, insp: new Set() }; a.n++; a.insp.add(f.inspId); if (f.severity === "Critical") a.C++; if (f.severity === "Major") a.M++; if (f.severity === "Minor") a.m++; if (isOpen(f)) a.open++; if (isLate(f)) a.late++; if (f.status === "Closed") a.closed++; });
  const ar = Object.entries(byArea).sort((a, b) => b[1].open - a[1].open);
  $("#tblArea").innerHTML = ar.length ? `<thead><tr><th>Area</th><th>Inspeksi</th><th>Temuan</th><th>Critical</th><th>Major</th><th>Minor</th><th>Terbuka</th><th>Lewat batas</th><th>% ditutup</th></tr></thead><tbody>` +
    ar.map(([k, a]) => `<tr><td>${esc(k)}</td><td>${a.insp.size}</td><td>${a.n}</td><td>${a.C}</td><td>${a.M}</td><td>${a.m}</td><td>${a.open}</td><td>${a.late ? `<b style="color:var(--crit)">${a.late}</b>` : 0}</td><td>${Math.round(a.closed / a.n * 100)}%</td></tr>`).join("") + "</tbody>"
    : `<tbody><tr><td class="muted">Belum ada data.</td></tr></tbody>`;

  try { await loadScript(LIBS.chart); } catch (e) { $$(".chart-box").forEach(b => b.innerHTML = `<p class="small muted">Grafik gagal dimuat. Periksa koneksi internet.</p>`); return; }
  const ink = cssVar("--ink"), line = cssVar("--line"), C = { Critical: cssVar("--crit"), Major: cssVar("--major"), Minor: cssVar("--minor") };
  Chart.defaults.color = cssVar("--steel"); Chart.defaults.font.family = cssVar("--font"); Chart.defaults.borderColor = line;
  const mk = (id, cfg) => { if (Charts[id]) Charts[id].destroy(); const cv = $("#" + id); if (!cv) return; Charts[id] = new Chart(cv, cfg); };
  const base = { responsive: true, maintainAspectRatio: false, animation: false, plugins: { legend: { position: "bottom", labels: { boxWidth: 12 } } } };

  const stCol = { Open: C.Critical, "On Progress": C.Major, Closed: cssVar("--good") };
  mk("cSev", { type: "bar", data: { labels: SEVS, datasets: ["Open", "On Progress", "Closed"].map(st => ({ label: st, backgroundColor: stCol[st], data: SEVS.map(s => data.filter(f => f.severity === s && (f.status || "Open") === st).length) })) },
    options: { ...base, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } } } } });

  const pics = {}; open.forEach(f => { const p = f.pic || "(kosong)"; pics[p] = pics[p] || { Critical: 0, Major: 0, Minor: 0, t: 0 }; pics[p][f.severity] = (pics[p][f.severity] || 0) + 1; pics[p].t++; });
  const pk = Object.keys(pics).sort((a, b) => pics[b].t - pics[a].t).slice(0, 10);
  mk("cPic", { type: "bar", data: { labels: pk, datasets: SEVS.map(s => ({ label: s, backgroundColor: C[s], data: pk.map(p => pics[p][s] || 0) })) },
    options: { ...base, indexAxis: "y", scales: { x: { stacked: true, beginAtZero: true, ticks: { precision: 0 } }, y: { stacked: true, grid: { display: false }, ticks: { autoSkip: false } } } } });

  const months = [...new Set(data.map(f => (f.tanggal || "").slice(0, 7)).filter(Boolean))].sort().slice(-12);
  mk("cTrend", { type: "bar", data: { labels: months.map(m => BULAN[+m.slice(5) - 1].slice(0, 3) + " " + m.slice(2, 4)), datasets: SEVS.map(s => ({ label: s, backgroundColor: C[s], data: months.map(m => data.filter(f => f.severity === s && (f.tanggal || "").startsWith(m)).length) })) },
    options: { ...base, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } } } } });

  const cl = {}; data.forEach(f => (f.klausul_iso || []).forEach(k => { const key = (k.standar || "").replace(":2018", "").replace(":2026", "").replace(":2015", "") + " " + k.klausul; cl[key] = (cl[key] || 0) + 1; }));
  const ck = Object.keys(cl).sort((a, b) => cl[b] - cl[a]).slice(0, 8);
  mk("cClause", { type: "bar", data: { labels: ck, datasets: [{ label: "Jumlah temuan", backgroundColor: cssVar("--brand"), data: ck.map(k => cl[k]) }] },
    options: { ...base, indexAxis: "y", plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, ticks: { precision: 0 } }, y: { grid: { display: false }, ticks: { autoSkip: false } } } } });
}
["#dArea", "#dAspek", "#dPeriode"].forEach(s => $(s).addEventListener("change", renderDashboard));

/* ======================= LAPORAN ======================= */
let rLoadedFor = null;
function reportSets(insp) {
  const finalOnly = $("#rFinalOnly").checked;
  const rank = { Critical: 0, Major: 1, Minor: 2 };
  const all = S.findings.filter(f => f.inspId === insp.id);
  const items = all.filter(f => f.jenis !== "good" && (!finalOnly || f.review === "Final")).sort((a, b) => (rank[a.severity] ?? 3) - (rank[b.severity] ?? 3) || String(a.no).localeCompare(String(b.no)));
  const goods = all.filter(f => f.jenis === "good" && (!finalOnly || f.review === "Final")).sort((a, b) => String(a.no).localeCompare(String(b.no)));
  const prevInsp = S.inspections.filter(i => i.area === insp.area && i.id !== insp.id && (i.tanggal || "") < (insp.tanggal || "")).sort((a, b) => b.tanggal.localeCompare(a.tanggal));
  const lastPrev = prevInsp[0];
  const tracked = S.findings.filter(f => f.jenis !== "good" && prevInsp.some(i => i.id === f.inspId) && (f.inspId === (lastPrev && lastPrev.id) || isOpen(f)))
    .sort((a, b) => (a.tanggal || "").localeCompare(b.tanggal || "") || String(a.no).localeCompare(String(b.no)));
  const isDraft = [...items, ...goods].some(f => f.review !== "Final");
  return { items, goods, tracked, isDraft };
}
function renderReportSide() {
  const sel = $("#rInsp"); const cur = sel.value || S.currentInsp;
  sel.innerHTML = `<option value="">— pilih inspeksi —</option>` + sortedInsps().map(i => `<option value="${esc(i.id)}" ${i.id === cur ? "selected" : ""}>${esc(inspLabel(i))}</option>`).join("");
  const insp = inspById(sel.value);
  if (!insp) { $("#rSummary").innerHTML = `<div class="empty small">Pilih inspeksi yang akan dilaporkan.</div>`; return; }
  const { items, goods, tracked, isDraft } = reportSets(insp);
  const all = S.findings.filter(f => f.inspId === insp.id);
  const nFinal = all.filter(f => f.review === "Final").length;
  const c = s => items.filter(f => f.severity === s).length;
  $("#rSummary").innerHTML = `<div class="insp-card"><b>${esc(insp.area)}</b><div class="small">${esc(fmtTgl(insp.tanggal, true))}</div>
    <div class="row small" style="margin-top:6px;gap:6px"><span class="tag Critical">${c("Critical")} Critical</span><span class="tag Major">${c("Major")} Major</span><span class="tag Minor">${c("Minor")} Minor</span><span class="tag Good">${goods.length} Good</span></div>
    <div class="small muted" style="margin-top:6px">${nFinal} dari ${all.length} entri sudah Final${tracked.length ? `, ${tracked.length} temuan sebelumnya akan ditampilkan di slide tracking` : ""}.</div></div>
    ${isDraft ? `<div class="note small">Masih ada entri yang belum Final — laporan akan diberi tanda <b>DRAFT</b>. Final = sudah disetujui Supervisor dan QSHE.</div>` : ""}`;
  if (rLoadedFor !== insp.id) {
    rLoadedFor = insp.id;
    $("#rConc").value = insp.kesimpulan || "";
    $("#rRec").value = (insp.rekomendasi || []).map(r => r.judul + ": " + r.uraian).join("\n");
  }
}
$("#rInsp").onchange = () => { rLoadedFor = null; renderReportSide(); };
["#rFinalOnly"].forEach(s => $(s).addEventListener("change", renderReportSide));
function parseRec() { return lines($("#rRec").value).map(l => { const i = l.indexOf(":"); return i > 0 ? { judul: l.slice(0, i).trim(), uraian: l.slice(i + 1).trim() } : { judul: l, uraian: "" }; }); }
$("#btnSaveConc").onclick = async () => {
  const insp = inspById($("#rInsp").value); if (!insp) return toast("Pilih inspeksi dulu.", "err");
  try { await Store.put("inspections", { ...insp, kesimpulan: $("#rConc").value.trim(), rekomendasi: parseRec() }); toast("Kesimpulan tersimpan."); } catch (e) { toast(e.message, "err"); }
};
$("#btnConc").onclick = async () => {
  const insp = inspById($("#rInsp").value); if (!insp) return toast("Pilih inspeksi dulu.", "err");
  await AI.ready;
  const { items, goods, tracked } = reportSets(insp);
  if (!items.length && !goods.length) return toast("Belum ada entri di inspeksi ini.", "err");
  if (!AI.fn) {
    const sets = reportSets(insp);
    $("#rConc").value = autoConclusion(insp, sets);
    $("#rRec").value = autoRecs(sets).map(x => `${x.judul}: ${x.uraian}`).join("\n");
    await Store.put("inspections", { ...insp, kesimpulan: $("#rConc").value.trim(), rekomendasi: parseRec() }).catch(() => {});
    return toast("Kesimpulan disusun otomatis dari data temuan. Silakan lengkapi.");
  }
  const ctl = new AbortController(); $("#btnConc").disabled = true;
  $("#concBusy").innerHTML = `<div class="ai-busy"><div style="flex:1"><div class="small">Menyusun kesimpulan…</div><div class="bar" style="margin-top:8px"></div></div><button class="btn" id="btnStopConc">Batalkan</button></div>`;
  $("#btnStopConc").onclick = () => ctl.abort();
  try {
    const r = await AI.fn.json(conclusionPrompt(insp, items, goods, tracked), { modelTier: "default", signal: ctl.signal, cache: false });
    $("#rConc").value = String(r.kesimpulan || "").trim();
    $("#rRec").value = (Array.isArray(r.rekomendasi) ? r.rekomendasi : []).map(x => `${String(x.judul || "").trim()}: ${String(x.uraian || "").trim()}`).join("\n");
    await Store.put("inspections", { ...insp, kesimpulan: $("#rConc").value.trim(), rekomendasi: parseRec() }).catch(() => {});
    $("#concBusy").innerHTML = "";
  } catch (e) { $("#concBusy").innerHTML = e && e.code === "cancelled" ? "" : `<div class="note">${esc(aiErrMsg(e))}</div>`; }
  finally { $("#btnConc").disabled = false; }
};

/* ---------------- kirim tanpa akun: teks & paket data ---------------- */
function reportText(insp) {
  const { items, goods } = reportSets(insp);
  const c = s => items.filter(f => f.severity === s).length;
  const ref = f => [...(f.klausul_iso || []).map(x => `${x.standar} kl. ${x.klausul}`), ...(f.kriteria_smk3 || []).map(x => `SMK3 ${x.kode}`), ...(f.regulasi || []).map(x => x.peraturan.replace(/ tentang .*/, "") + (x.pasal ? " " + x.pasal : ""))].join("; ");
  const out = [`*LAPORAN INSPEKSI QSHE — ${insp.area}*`, `${insp.agenda || ""}`.trim(), fmtTgl(insp.tanggal, true), `Inspector: ${(insp.inspectors || []).join(", ") || "-"}`,
    `Temuan: ${items.length} (${c("Critical")} Critical, ${c("Major")} Major, ${c("Minor")} Minor) · Good practice: ${goods.length}`, ""];
  items.forEach((f, i) => {
    out.push(`*${i + 1}. ${f.no || ""} [${f.severity || "-"}] ${f.judul || ""}*${f.stop_work ? " — HENTIKAN PEKERJAAN" : ""}`);
    if (f.deskripsi) out.push(f.deskripsi);
    const r = ref(f); if (r) out.push(`Rujukan: ${r}`);
    if (f.tindakan_segera) out.push(`Tindakan segera: ${f.tindakan_segera}`);
    (f.tindakan_korektif || []).forEach(x => out.push(`- [${x.hierarki}] ${x.tindakan}`));
    if (f.pic || f.due) out.push(`PIC: ${f.pic || "-"} · Batas: ${f.due ? fmtTgl(f.due) : "-"} · Status: ${f.status || "Open"}`);
    out.push(`Review: ${REVIEW_LABEL[f.review || "Draft"] || f.review}${f.spvAt ? ` (Spv: ${f.spvName})` : ""}`, "");
  });
  if (goods.length) { out.push("*GOOD PRACTICE*"); goods.forEach((f, i) => out.push(`${i + 1}. ${f.judul}${f.deskripsi ? " — " + f.deskripsi : ""}`)); out.push(""); }
  if (insp.kesimpulan) out.push("*KESIMPULAN*", insp.kesimpulan, "");
  out.push(Store.mode === "local" ? "(Dikirim dari perangkat lapangan — foto menyusul terpisah atau lewat paket data.)" : "");
  return out.join("\n").trim();
}
$("#btnText").onclick = () => {
  const insp = inspById($("#rInsp").value); if (!insp) return toast("Pilih inspeksi dulu.", "err");
  $("#txtOut").value = reportText(insp); openModal("#mText");
};
$("#btnCopyTxt").onclick = async () => {
  const ta = $("#txtOut"); let ok = false;
  try { await navigator.clipboard.writeText(ta.value); ok = true; } catch (e) {}
  if (!ok) { try { ta.focus(); ta.select(); ok = document.execCommand("copy"); } catch (e) {} }
  toast(ok ? "Teks tersalin — tempel di WhatsApp/email." : "Salin otomatis diblokir perangkat. Tekan lama teks lalu pilih Salin semua.", ok ? "" : "err");
};
$("#btnPack").onclick = async e => {
  const insp = inspById($("#rInsp").value); if (!insp) return toast("Pilih inspeksi dulu.", "err");
  const btn = e.currentTarget; busyBtn(btn, true, "Menyiapkan paket…");
  try {
    const fs = S.findings.filter(f => f.inspId === insp.id);
    const photos = {};
    for (const f of fs) for (const p of [...(f.photos || []), ...(f.closePhotos || [])]) { const d = await Store.getPhoto(p.id); if (d) photos[p.id] = d; }
    const pack = { format: "hexindo-temuan-qshe", versi: 1, dibuat: new Date().toISOString(), inspections: [insp], findings: fs, photos };
    const blob = new Blob([JSON.stringify(pack)], { type: "application/json" });
    await saveFile(`Paket_Temuan_${slug(insp.kode || insp.area)}_${(insp.tanggal || "").replace(/-/g, "")}.json`, blob);
  } catch (err) { toast(err.message || "Gagal menyiapkan paket.", "err"); }
  finally { busyBtn(btn, false); }
};
$("#fImport").addEventListener("change", async e => {
  const file = e.target.files[0]; e.target.value = ""; if (!file) return;
  try {
    const p = JSON.parse(await file.text());
    if (!p || p.format !== "hexindo-temuan-qshe" || !Array.isArray(p.findings)) throw new Error("File bukan paket data aplikasi ini.");
    const nI = (p.inspections || []).length, nF = p.findings.length;
    if (!confirm(`Impor ${nI} inspeksi dan ${nF} entri ke register ini? Entri dengan ID yang sama akan diperbarui.`)) return;
    for (const [id, d] of Object.entries(p.photos || {})) await Store.putPhoto(id, d);
    for (const i of (p.inspections || [])) await Store.put("inspections", i);
    for (const f of p.findings) await Store.put("findings", migrateRec({ ...f, review: f.review === "QSHE" || f.review === "Final" ? "QSHE" : "Review" }));
    toast(`${nF} entri diimpor — yang sudah disetujui Supervisor langsung masuk antrean QSHE.`);
  } catch (err) { toast(err.message || "Impor gagal.", "err"); }
});

function oneDriveNote(name, inFolder) {
  const el = $("#expNote"); const u = S.settings.onedrive;
  if (inFolder) { el.innerHTML = `<span style="display:block;margin-top:4px"><b>${esc(name)}</b> tersimpan di folder <b>${esc(Folder.handle.name)}/Laporan</b>.</span>`; return; }
  el.innerHTML = u
    ? `<span style="display:block;margin-top:4px"><b>${esc(name)}</b> sudah diunduh. Unggah ke OneDrive QSHE: <a href="${esc(u)}" target="_blank" rel="noopener" style="color:var(--brand-deep);font-weight:600">buka folder OneDrive</a> · <button type="button" class="addlink" id="btnCopyOD">salin link</button></span>`
    : `<span style="display:block;margin-top:4px"><b>${esc(name)}</b> sudah diunduh ke folder Download. Pindahkan ke folder OneDrive QSHE, atau hubungkan folder di Pengaturan (laptop).</span>`;
  const b = $("#btnCopyOD"); if (b) b.onclick = async () => { let ok = false; try { await navigator.clipboard.writeText(u); ok = true; } catch (e) {} toast(ok ? "Link OneDrive tersalin — tempel di browser." : "Salin manual: " + u, ok ? "" : "err"); };
}
async function saveFile(name, blob) {
  if (Folder.handle && (Folder.state === "ok" || await Folder.check(true))) {
    try { await Folder.write("Laporan/" + name, blob); toast(`Tersimpan di folder ${Folder.handle.name}/Laporan`); oneDriveNote(name, true); return true; }
    catch (e) { toast("Gagal menulis ke folder, file diunduh biasa.", "err"); }
  }
  const dl = await useCap("downloads");
  if (dl) {
    try { await dl.save({ filename: name, data: blob }); toast("File siap: " + name); oneDriveNote(name); return true; }
    catch (e) {
      const c = e && e.code;
      if (c === "declined") toast("Unduhan dibatalkan.");
      else if (c === "rate_limited") toast("Masih ada unduhan yang menunggu konfirmasi.", "err");
      else toast("Unduhan tidak tersedia di tampilan ini (" + (c || "error") + ").", "err");
      return false;
    }
  }
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000); oneDriveNote(name); return true;
}
function busyBtn(btn, on, label) { btn.disabled = on; if (on) { btn.dataset.l = btn.innerHTML; btn.textContent = label; } else if (btn.dataset.l) btn.innerHTML = btn.dataset.l; }
function slug(s) { return String(s || "").replace(/[^\w]+/g, "_").replace(/^_|_$/g, ""); }

$("#btnPPT").onclick = async e => {
  const insp = inspById($("#rInsp").value); if (!insp) return toast("Pilih inspeksi dulu.", "err");
  const btn = e.currentTarget; busyBtn(btn, true, "Menyusun slide…");
  try {
    const sets = reportSets(insp);
    if (!sets.items.length && !sets.goods.length) throw new Error($("#rFinalOnly").checked ? "Belum ada entri Final. Matikan opsi 'Hanya temuan Final' untuk laporan draft." : "Belum ada entri di inspeksi ini.");
    const blob = await buildPPT(insp, sets, { kesimpulan: $("#rConc").value.trim(), rekomendasi: parseRec(), track: $("#rTrack").checked, good: $("#rGood").checked });
    await saveFile(`QSHE_Inspection_${slug(insp.kode || insp.area)}_${insp.tanggal.replace(/-/g, "")}${sets.isDraft ? "_DRAFT" : ""}.pptx`, blob);
  } catch (err) { toast(err.message || "Gagal menyusun PowerPoint.", "err"); console.error(err); }
  finally { busyBtn(btn, false); }
};
$("#btnXLSInsp").onclick = async e => {
  const insp = inspById($("#rInsp").value); if (!insp) return toast("Pilih inspeksi dulu.", "err");
  const btn = e.currentTarget; busyBtn(btn, true, "Menyusun Excel…");
  try {
    const sets = reportSets(insp);
    const blob = await buildExcel([...sets.items], [...sets.goods], { title: `${insp.area} — ${fmtTgl(insp.tanggal, true)}`, insp, isDraft: sets.isDraft });
    await saveFile(`Register_Temuan_${slug(insp.kode || insp.area)}_${insp.tanggal.replace(/-/g, "")}${sets.isDraft ? "_DRAFT" : ""}.xlsx`, blob);
  } catch (err) { toast(err.message || "Gagal menyusun Excel.", "err"); console.error(err); }
  finally { busyBtn(btn, false); }
};
$("#btnXLSAll").onclick = async e => {
  const btn = e.currentTarget; busyBtn(btn, true, "Menyusun Excel…");
  try {
    const rank = { Critical: 0, Major: 1, Minor: 2 };
    const items = S.findings.filter(f => f.jenis !== "good").sort((a, b) => (a.tanggal || "").localeCompare(b.tanggal || "") || String(a.no).localeCompare(String(b.no)));
    const goods = S.findings.filter(f => f.jenis === "good").sort((a, b) => (a.tanggal || "").localeCompare(b.tanggal || ""));
    if (!items.length && !goods.length) throw new Error("Register masih kosong.");
    const blob = await buildExcel(items, goods, { title: "Seluruh register temuan", isDraft: items.some(f => f.review !== "Final") });
    await saveFile(`Register_Temuan_QSHE_${localISO().replace(/-/g, "")}.xlsx`, blob);
  } catch (err) { toast(err.message || "Gagal menyusun Excel.", "err"); console.error(err); }
  finally { busyBtn(btn, false); }
};

/* ---------------- shared text builders ---------------- */
function stdLines(f, long) {
  const a = [];
  (f.klausul_iso || []).forEach(k => a.push(`${k.standar} Klausul ${k.klausul}${k.judul ? " – " + k.judul : ""}`));
  (f.kriteria_smk3 || []).forEach(k => a.push(`PP 50/2012 Kriteria ${k.kode}${long && k.ringkas ? " – " + k.ringkas : ""}`));
  (f.regulasi || []).forEach(r => a.push(`${r.peraturan}${r.pasal ? ", " + r.pasal : ""}${r.perlu_verifikasi ? " (perlu verifikasi)" : ""}`));
  return a;
}
function caLines(f) { return (f.tindakan_korektif || []).map(c => `[${c.hierarki}] ${c.tindakan}`); }

/* ======================= EXCEL (format template "Inspection Monitoring") ======================= */
function colLetter(n) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function xDate(iso) { const d = parseISO(iso); return d ? new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())) : null; }
async function buildExcel(items, goods, meta) {
  await loadScript(LIBS.excel);
  const wb = new ExcelJS.Workbook(); wb.creator = "QSHE Department"; wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  const NAVY = "FF002060", WHT = "FFFFFFFF", GRY = "FFF2F2F2", LINE = "FFA6A6A6", RED = "FFC00000", FONT = "Calibri";
  const thin = { style: "thin", color: { argb: LINE } }, med = { style: "medium", color: { argb: "FF000000" } };
  const allThin = { top: thin, left: thin, bottom: thin, right: thin };
  const solid = argb => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const px = { col: w => Math.round(w * 7 + 5), row: h => Math.round(h * 96 / 72) };
  const EMU = 9525;
  const insp = meta.insp;
  const locText = insp ? insp.area : "Semua area";
  const inspText = insp ? ((insp.inspectors || []).join(", ") || "-") : [...new Set(items.concat(goods).map(f => (inspById(f.inspId) || {}).inspectors || []).flat())].join(", ") || "-";

  /* kolom: [judul, lebar] — urutan & judul mengikuti template (Photo Before disatukan dengan Evidence Photo, Related Dept tidak dipakai) */
  const NEED = [["No.", 6], ["Location & Date", 22], ["Deskripsi", 46], ["Evidence\nPhoto", 42], ["Remarks", 36], ["Improvement plan", 40], ["Photo", 36], ["Due date", 15], ["PIC", 18], ["Status - Remark", 28]];
  const GOOD = NEED.slice(0, 5);

  const textLines = (s, w) => { const cpl = Math.max(6, Math.floor(w * 1.05)); return String(s || "").split("\n").reduce((n, l) => n + Math.max(1, Math.ceil(l.length / cpl)), 0); };

  /* tempel 1–4 foto ke satu sel, tersusun rapi (1 foto penuh, 2 foto berdampingan, 3–4 foto grid 2×2) */
  async function putPhotos(ws, list, col0, row0, boxW, boxH) {
    const ph = []; for (const p of (list || []).slice(0, 4)) { const d = await Store.getPhoto(p.id); if (d) ph.push({ d, w: p.w || 4, h: p.h || 3 }); }
    if (!ph.length) return false;
    const n = ph.length, cols = n === 1 ? 1 : 2, rows = n <= 2 ? 1 : 2, gap = 4;
    const sw = (boxW - gap * (cols - 1)) / cols, sh = (boxH - gap * (rows - 1)) / rows;
    ph.forEach((p, k) => {
      const r = Math.min(sw / p.w, sh / p.h), w = Math.floor(p.w * r), h = Math.floor(p.h * r);
      const cx = k % cols, cy = Math.floor(k / cols);
      const x = 4 + cx * (sw + gap) + (sw - w) / 2, y = 4 + cy * (sh + gap) + (sh - h) / 2;
      const id = wb.addImage({ base64: p.d, extension: "jpeg" });
      ws.addImage(id, { tl: { nativeCol: col0, nativeColOff: Math.round(x * EMU), nativeRow: row0, nativeRowOff: Math.round(y * EMU) }, ext: { width: w, height: h }, editAs: "oneCell" });
    });
    return true;
  }

  async function makeSheet(name, cols, list, kind) {
    const need = kind === "need", n = cols.length, last = colLetter(n);
    const ws = wb.addWorksheet(name, {
      views: [{ state: "frozen", xSplit: 0, ySplit: 7, zoomScale: 80, zoomScaleNormal: 80, showGridLines: false }],
      properties: { tabColor: { argb: need ? "FFF36F20" : "FF2B7A4B" }, defaultRowHeight: 15 },
      pageSetup: { orientation: "landscape", paperSize: 8 /* A3 */, fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true, margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.5, header: 0.2, footer: 0.25 } },
      headerFooter: { oddFooter: "&LQSHE Department&CInspection Monitoring&RHalaman &P dari &N" }
    });
    ws.pageSetup.printTitlesRow = "6:7";
    cols.forEach((c, i) => ws.getColumn(i + 1).width = c[1]);

    /* judul & identitas */
    ws.mergeCells(`A1:${last}1`);
    const t = ws.getCell("A1"); t.value = "INSPECTION MONITORING"; t.font = { name: FONT, size: 24, bold: true }; t.alignment = { horizontal: "center", vertical: "middle" }; ws.getRow(1).height = 36;
    ws.mergeCells(`A2:${last}2`);
    const sub = ws.getCell("A2"); sub.value = need ? "Need Improvement — temuan yang perlu ditindaklanjuti" : "Good Point — praktik baik yang teramati"; sub.font = { name: FONT, size: 12, italic: true, color: { argb: "FF595959" } }; sub.alignment = { horizontal: "center" };
    const mw = need ? 5 : n; /* lebar blok identitas (kolom) */
    ws.mergeCells(`A3:${colLetter(mw)}3`); ws.getCell("A3").value = `Location : ${locText}`;
    ws.mergeCells(`A4:${colLetter(mw)}4`); ws.getCell("A4").value = `Inspector : ${inspText}`;
    ws.mergeCells(`A5:${colLetter(mw)}5`); ws.getCell("A5").value = `Tanggal : ${insp ? fmtTgl(insp.tanggal, true) : "Seluruh register"}${meta.isDraft ? "   |   DRAFT — untuk review tim" : ""}`;
    ["A3", "A4", "A5"].forEach(a => { ws.getCell(a).font = { name: FONT, size: 11, bold: a === "A3", color: { argb: a === "A5" && meta.isDraft ? RED : "FF000000" } }; ws.getCell(a).alignment = { horizontal: "left", vertical: "middle" }; });

    /* header 2 baris (baris 6–7), warna navy sesuai template */
    cols.forEach((c, i) => {
      const col = i + 1;
      if (need && c[0] === "Photo") { ws.getCell(6, col).value = "Photo"; ws.getCell(7, col).value = "After"; }
      else { ws.mergeCells(6, col, 7, col); ws.getCell(6, col).value = c[0]; }
    });
    for (let r = 6; r <= 7; r++) for (let c = 1; c <= n; c++) {
      const cell = ws.getCell(r, c); cell.fill = solid(NAVY); cell.font = { name: FONT, size: 11, bold: true, color: { argb: WHT } };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = { top: r === 6 ? med : thin, bottom: r === 7 ? med : thin, left: c === 1 ? med : thin, right: c === n ? med : thin };
    }
    ws.getRow(6).height = 24; ws.getRow(7).height = 20;

    /* blok "Finding Close" (hanya Need Improvement) */
    const R0 = 8, RN = R0 + Math.max(list.length, 1) - 1;
    if (need) {
      ws.mergeCells("H3:I4"); ws.getCell("H3").value = "Finding Close"; ws.getCell("H3").font = { name: FONT, size: 16, bold: true, color: { argb: WHT } };
      ws.getCell("H3").alignment = { horizontal: "center", vertical: "middle" };
      ws.mergeCells("J3:J4"); ws.getCell("J3").value = { formula: `COUNTIF(J${R0}:J${RN},"Closed*")&" / "&COUNT(A${R0}:A${RN})` };
      ws.getCell("J3").font = { name: FONT, size: 18, bold: true, color: { argb: WHT } }; ws.getCell("J3").alignment = { horizontal: "center", vertical: "middle" };
      for (let r = 3; r <= 4; r++) for (let c = 8; c <= 10; c++) { const cell = ws.getCell(r, c); cell.fill = solid("FF595959"); cell.border = allThin; }
    }

    /* isi */
    const bodyFont = { name: FONT, size: 11 };
    if (!list.length) {
      ws.mergeCells(`A${R0}:${last}${R0}`); const c = ws.getCell(`A${R0}`); c.value = need ? "Tidak ada temuan pada inspeksi ini." : "Tidak ada good practice pada inspeksi ini."; c.font = { ...bodyFont, italic: true, color: { argb: "FF595959" } };
      c.alignment = { horizontal: "center", vertical: "middle" }; c.border = allThin; ws.getRow(R0).height = 36;
    }
    for (let i = 0; i < list.length; i++) {
      const f = list[i], rn = R0 + i, row = ws.getRow(rn);
      const loc = `${String(f.area || "").toUpperCase()}, ${fmtTgl(f.tanggal).toUpperCase()}` + (f.lokasi ? `\nLokasi: ${f.lokasi}` : "");
      const desc = (f.judul ? f.judul + "\n" : "") + (f.deskripsi || "") + (f.stop_work ? "\n[STOP WORK — area/pekerjaan dihentikan]" : "");
      const rem = stdLines(f).map(x => "• " + x).join("\n");
      const plan = [f.tindakan_segera ? "Segera: " + f.tindakan_segera : "", ...(f.tindakan_korektif || []).map((c, j) => `${j + 1}. ${c.tindakan}`)].filter(Boolean).join("\n");
      const late = need && f.status !== "Closed" && f.due && f.due < localISO();
      const vals = [i + 1, loc, desc, "", rem, need ? plan : null, need ? "" : null, need ? xDate(f.due) : null, need ? (f.pic || "") : null, need ? (f.status || "Open") + (f.remarks ? "\n" + f.remarks : "") : null];
      cols.forEach((c, j) => {
        const cell = row.getCell(j + 1); cell.value = vals[j] ?? ""; cell.font = bodyFont; cell.border = allThin;
        const center = [0, 1, 7, 8, 9].includes(j);
        cell.alignment = { vertical: "middle", horizontal: center ? "center" : "left", wrapText: true, indent: center ? 0 : 1 };
      });
      row.getCell(3).font = bodyFont;
      if (need) {
        row.getCell(8).numFmt = "dd/mm/yyyy";
        if (late) { row.getCell(8).font = { ...bodyFont, bold: true, color: { argb: RED } }; }
        if (!f.due) row.getCell(8).value = "";
      }
      /* tinggi baris: cukup untuk foto & teks */
      const hasPh = (f.photos || []).length || (need && (f.closePhotos || []).length);
      const textH = Math.max(textLines(desc, cols[2][1]), textLines(rem, cols[4][1]), need ? textLines(plan, cols[5][1]) : 0) * 15 + 12;
      const h = Math.min(400, Math.max(hasPh ? 172 : 60, textH));
      row.height = h;
      const boxH = px.row(h) - 8;
      await putPhotos(ws, f.photos, 3, rn - 1, px.col(cols[3][1]) - 8, boxH);
      if (need) await putPhotos(ws, f.closePhotos, 6, rn - 1, px.col(cols[6][1]) - 8, boxH);
    }

    /* status: warna + daftar-pilihan tidak dipasang agar kolom bebas diisi keterangan */
    if (need && list.length) {
      const rg = `J${R0}:J${RN}`, fill = argb => ({ type: "pattern", pattern: "solid", bgColor: { argb } });
      ws.addConditionalFormatting({ ref: rg, rules: [
        { type: "expression", priority: 1, formulae: [`LEFT($J${R0},6)="Closed"`], style: { fill: fill("FFE3F2E8"), font: { color: { argb: "FF2B7A4B" }, bold: true } } },
        { type: "expression", priority: 2, formulae: [`LEFT($J${R0},11)="On Progress"`], style: { fill: fill("FFFFF0D9"), font: { color: { argb: "FFA35E00" }, bold: true } } },
        { type: "expression", priority: 3, formulae: [`LEFT($J${R0},4)="Open"`], style: { fill: fill("FFFBE7E9"), font: { color: { argb: "FFC4162A" }, bold: true } } }] });
    }

    /* catatan penutup (teks dari template) */
    const nr = RN + 2;
    ws.mergeCells(`A${nr}:${last}${nr}`);
    const nc = ws.getCell(`A${nr}`);
    nc.value = "Note: \n1. Batas waktu untuk menindaklanjuti temuan 1 bulan setelah closing meeting cabang / Project.\n2. Hasil tindak lanjut akan dilaporkan kepada Management setiap 3 bulan.";
    nc.font = { name: FONT, size: 11, color: { argb: RED } }; nc.alignment = { vertical: "top", horizontal: "left", wrapText: true, indent: 1 };
    ws.getRow(nr).height = 62;
    return ws;
  }

  await makeSheet("Good Point", GOOD, goods, "good");
  await makeSheet("Need Improvement", NEED, items, "need");
  wb.views = [{ activeTab: items.length || !goods.length ? 1 : 0, firstSheet: 0, visibility: "visible" }];

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

/* ======================= POWERPOINT (di atas file template perusahaan) ======================= */
const P = { F: "DINPro-Light", FB: "DINPro-Bold", OR: "F26F21", INK: "1E1E1E", GREY: "595959", SOFT: "F2F2F2", LINE: "D0D0D0", HEAD: "262626", ZEBRA: "FDF0E8", SEV: { Critical: "C4162A", Major: "C77700", Minor: "2A6496" }, ST: { Open: "C4162A", "On Progress": "C77700", Closed: "2B7A4B" } };
const TEMPLATE_PPT = "assets/template-ppt.pptx";
function fitRect(iw, ih, bx, by, bw, bh) { const r = Math.min(bw / (iw || 4), bh / (ih || 3)); const w = (iw || 4) * r, h = (ih || 3) * r; return { x: bx + (bw - w) / 2, y: by + (bh - h) / 2, w, h }; }
function fitFont(texts, w, h, max, min) {
  for (let fs = max; fs >= min; fs -= 0.5) {
    const cw = fs * 0.5 / 72, lh = fs * 1.22 / 72, cpl = Math.max(8, Math.floor((w - 0.2) / cw));
    let n = 0; texts.forEach(t => String(t || "").split("\n").forEach(l => { n += Math.max(1, Math.ceil(l.length / cpl)); }));
    if (n * lh <= h - 0.12) return fs;
  }
  return min;
}
function clip(s, n) { s = String(s || ""); return s.length > n ? s.slice(0, n - 1).trimEnd() + "…" : s; }

/* Susun isi slide dengan PptxGenJS (ukuran 16:9 = 10 × 5,625 inci, sama dengan template), lalu tempelkan ke slide template. */
async function buildPPT(insp, sets, opt) {
  await loadScript(LIBS.pptx);
  const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_16x9";
  const { F, FB, OR, INK, GREY } = P;
  const defs = []; /* judul tiap slide isi, urutannya sama dengan slide yang dibuat */
  const X0 = 0.35, W = 9.3;
  const mk = (title, right) => {
    const s = pptx.addSlide(); defs.push({ title });
    if (right) s.addText(right, { x: 5.6, y: 0.8, w: 4.05, h: 0.2, fontFace: F, fontSize: 9, color: GREY, align: "right", margin: 0 });
    if (sets.isDraft) s.addText("DRAFT – untuk review tim", { x: X0, y: 0.8, w: 2.4, h: 0.2, fontFace: FB, bold: true, fontSize: 8.5, color: "C4162A", margin: 0 });
    return s;
  };
  const photo = async (s, p, x, y, w, h, caption) => {
    const d = p ? await Store.getPhoto(p.id) : null;
    if (d) { s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: "F2F2F2" }, line: { color: P.LINE, width: 0.5 } }); const r = fitRect(p.w, p.h, x + 0.03, y + 0.03, w - 0.06, h - 0.06); s.addImage({ data: d, x: r.x, y: r.y, w: r.w, h: r.h }); }
    else { s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: "F7F7F7" }, line: { color: P.LINE, width: 0.5, dashType: "dash" } }); s.addText(caption || "Tidak ada foto", { x, y, w, h, fontFace: F, fontSize: 9, color: GREY, align: "center", valign: "middle" }); }
  };
  /* 1–4 foto tersusun otomatis di dalam satu area (1 penuh, 2 bertumpuk/berdampingan, 3–4 grid 2×2) */
  const photoArea = async (s, list, x, y, w, h, stack) => {
    const ph = (list || []).slice(0, 4), n = ph.length, gap = 0.08;
    if (n <= 1) return photo(s, ph[0], x, y, w, h);
    if (n === 2) { if (stack) { const hh = (h - gap) / 2; await photo(s, ph[0], x, y, w, hh); await photo(s, ph[1], x, y + hh + gap, w, hh); } else { const ww = (w - gap) / 2; await photo(s, ph[0], x, y, ww, h); await photo(s, ph[1], x + ww + gap, y, ww, h); } return; }
    const ww = (w - gap) / 2, hh = (h - gap) / 2;
    for (let k = 0; k < n; k++) await photo(s, ph[k], x + (k % 2) * (ww + gap), y + Math.floor(k / 2) * (hh + gap), ww, hh);
  };
  const sevChip = (s, sev, x, y) => s.addText(sev || "-", { x, y, w: 1.0, h: 0.27, fontFace: FB, bold: true, fontSize: 10.5, color: "FFFFFF", fill: { color: P.SEV[sev] || GREY }, align: "center", valign: "middle", margin: 0 });
  const inspectors = (insp.inspectors || []).join(", ") || "-";
  const allN = sets.items.length;
  let s;

  /* A. tracking temuan sebelumnya */
  if (opt.track && sets.tracked.length) {
    for (let i = 0; i < sets.tracked.length; i += 2) {
      s = mk("Improvement Tracking", `Status tindak lanjut temuan sebelumnya · ${i + 1}–${Math.min(i + 2, sets.tracked.length)} dari ${sets.tracked.length}`);
      const pair = sets.tracked.slice(i, i + 2);
      for (let j = 0; j < pair.length; j++) {
        const f = pair[j], y = 1.05 + j * 2.2, st = f.status || "Open", H = 2.08;
        s.addShape(pptx.ShapeType.rect, { x: X0, y, w: W, h: H, fill: { color: "FFFFFF" }, line: { color: P.LINE, width: 0.75 } });
        s.addShape(pptx.ShapeType.rect, { x: X0, y, w: 0.08, h: H, fill: { color: P.ST[st] || GREY }, line: { color: P.ST[st] || GREY, width: 0 } });
        s.addText(`${i + j + 1}.  ${f.no}`, { x: 0.55, y: y + 0.07, w: 3, h: 0.25, fontFace: FB, bold: true, fontSize: 10, color: GREY, margin: 0 });
        s.addText(st, { x: 4.3, y: y + 0.07, w: 1.1, h: 0.25, fontFace: FB, bold: true, fontSize: 10, color: "FFFFFF", fill: { color: P.ST[st] || GREY }, align: "center", valign: "middle", margin: 0 });
        const body = [
          { text: clip(f.judul, 90), options: { bold: true, fontFace: FB, fontSize: 11, color: INK, breakLine: true } },
          { text: clip(f.deskripsi, 260), options: { fontSize: 9, color: INK, breakLine: true } },
          { text: clip(stdLines(f).join("; "), 150), options: { fontSize: 8, color: GREY, italic: true, breakLine: true } },
          { text: `PIC: ${f.pic || "-"}   |   Batas waktu: ${fmtTgl(f.due)}${st === "Closed" && f.closedDate ? "   |   Ditutup: " + fmtTgl(f.closedDate) : ""}`, options: { fontSize: 9, bold: true, fontFace: FB, color: INK, breakLine: !!f.remarks } }
        ];
        if (f.remarks) body.push({ text: "Progres: " + clip(f.remarks, 130), options: { fontSize: 8.5, color: GREY } });
        s.addText(body, { x: 0.55, y: y + 0.38, w: 4.85, h: H - 0.45, fontFace: F, valign: "top", margin: 0, paraSpaceAfter: 3 });
        await photo(s, (f.photos || [])[0], 5.55, y + 0.1, 1.95, 1.62, "Tidak ada foto");
        await photo(s, (f.closePhotos || [])[0], 7.6, y + 0.1, 1.95, 1.62, st === "Closed" ? "Foto bukti belum diunggah" : "Belum ada bukti penutupan");
        s.addText("Sebelum", { x: 5.55, y: y + 1.75, w: 1.95, h: 0.2, fontFace: F, fontSize: 8, color: GREY, align: "center", margin: 0 });
        s.addText("Sesudah", { x: 7.6, y: y + 1.75, w: 1.95, h: 0.2, fontFace: F, fontSize: 8, color: GREY, align: "center", margin: 0 });
      }
    }
  }

  /* B. overview */
  s = mk("Inspection Overview", "Ringkasan pelaksanaan inspeksi");
  const ins = insp.inspectors || [], ise = insp.inspectees || [];
  const kv = [["Agenda", insp.agenda || "-"], ["Tanggal Pelaksanaan", fmtTgl(insp.tanggal, true)], ["Area", insp.area]];
  for (let k = 0; k < Math.max(2, ins.length); k++) kv.push([`Nama Inspector ${k + 1}`, ins[k] || "-"]);
  for (let k = 0; k < Math.max(2, ise.length); k++) kv.push([`Nama Inspectee ${k + 1}`, ise[k] || "-"]);
  const rowH = Math.min(0.34, 4.2 / kv.length);
  s.addTable(kv.map(([a, b]) => [{ text: a, options: { color: OR, fontFace: FB, bold: true, fontSize: 11 } }, { text: ": " + b, options: { color: INK, fontSize: 11 } }]),
    { x: X0, y: 1.1, w: 5.4, colW: [2.0, 3.4], fontFace: F, border: { type: "none" }, rowH, margin: 0.02 });
  const cnt = sv => sets.items.filter(f => f.severity === sv).length;
  s.addText("Hasil inspeksi", { x: 6.1, y: 1.1, w: 3.5, h: 0.28, fontFace: FB, bold: true, fontSize: 12, color: INK, margin: 0 });
  [["Critical", cnt("Critical")], ["Major", cnt("Major")], ["Minor", cnt("Minor")]].forEach(([sv, n], k) => {
    const y = 1.45 + k * 0.85;
    s.addShape(pptx.ShapeType.rect, { x: 6.1, y, w: 3.55, h: 0.72, fill: { color: "FFFFFF" }, line: { color: P.LINE, width: 0.75 } });
    s.addShape(pptx.ShapeType.rect, { x: 6.1, y, w: 0.1, h: 0.72, fill: { color: P.SEV[sv] }, line: { color: P.SEV[sv], width: 0 } });
    s.addText(String(n), { x: 6.3, y, w: 1.0, h: 0.72, fontFace: FB, bold: true, fontSize: 28, color: P.SEV[sv], valign: "middle", margin: 0 });
    s.addText(sv, { x: 7.3, y, w: 2.3, h: 0.72, fontFace: F, fontSize: 13, color: INK, valign: "middle", margin: 0 });
  });
  s.addText(`${sets.goods.length} good SHE practice   |   ${sets.items.filter(f => f.stop_work).length} stop work`, { x: 6.1, y: 4.05, w: 3.55, h: 0.28, fontFace: F, fontSize: 10, color: GREY, margin: 0 });

  /* C. ringkasan temuan (tabel) */
  if (sets.items.length) {
    const per = 5;
    for (let i = 0; i < sets.items.length; i += per) {
      s = mk("Ringkasan Temuan", sets.items.length > per ? `Halaman ${i / per + 1} dari ${Math.ceil(sets.items.length / per)}` : "Daftar temuan periode berjalan");
      const hdrOpt = { bold: true, fontFace: FB, color: "FFFFFF", fill: { color: P.HEAD }, align: "center", valign: "middle", fontSize: 10 };
      const rows = [[{ text: "No", options: hdrOpt }, { text: "Temuan", options: hdrOpt }, { text: "Standar / Klausul", options: hdrOpt }, { text: "Tingkat", options: hdrOpt }, { text: "PIC", options: hdrOpt }, { text: "Batas Waktu", options: hdrOpt }]];
      sets.items.slice(i, i + per).forEach((f, k) => {
        const zebra = { fill: { color: k % 2 ? "FFFFFF" : P.ZEBRA } };
        rows.push([
          { text: String(i + k + 1), options: { ...zebra, align: "center", bold: true } },
          { text: [{ text: clip(f.judul, 80), options: { bold: true, fontFace: FB, breakLine: true } }, { text: clip(f.lokasi || f.area, 60), options: { color: GREY, fontSize: 8 } }], options: zebra },
          { text: clip((f.klausul_iso || []).map(k2 => `${k2.standar.replace(/:\d{4}/, "")} ${k2.klausul}`).concat((f.kriteria_smk3 || []).map(k2 => "SMK3 " + k2.kode)).join("\n"), 110), options: { ...zebra, fontSize: 8 } },
          { text: f.severity, options: { bold: true, fontFace: FB, color: "FFFFFF", fill: { color: P.SEV[f.severity] || GREY }, align: "center" } },
          { text: f.pic || "-", options: { ...zebra, align: "center" } },
          { text: f.due ? fmtShort(f.due) : "-", options: { ...zebra, align: "center" } }
        ]);
      });
      s.addTable(rows, { x: X0, y: 1.1, w: W, colW: [0.45, 3.6, 2.05, 1.0, 1.1, 1.1], fontFace: F, fontSize: 9, color: INK, valign: "middle", border: { type: "solid", pt: 0.5, color: P.LINE }, rowH: [0.34].concat(Array(Math.min(per, sets.items.length - i)).fill(0.7)), margin: 0.05 });
    }
  }

  /* D. satu slide per temuan: foto otomatis tersusun + teks baku */
  for (let i = 0; i < sets.items.length; i++) {
    const f = sets.items[i];
    s = mk("Current Inspection Findings", `Temuan ${i + 1} dari ${allN}`);
    const ca = caLines(f).map(x => clip(x.replace(/^\[[^\]]+\]\s*/, ""), 150));
    const hasFU = !!(f.tindakan_segera || ca.length || f.pic || f.due || (f.akar_masalah || {}).uraian || f.bukti_penutupan);
    const top = 1.72, bottom = hasFU ? 4.3 : 5.38;
    sevChip(s, f.severity, X0, 1.04);
    s.addText(`${f.no}   |   ${f.area}${f.stop_work ? "   |   STOP WORK" : ""}`, { x: 1.45, y: 1.04, w: 8.2, h: 0.27, fontFace: FB, bold: true, fontSize: 10, color: f.stop_work ? P.SEV.Critical : GREY, valign: "middle", margin: 0 });
    s.addText(clip(f.judul, 110), { x: X0, y: 1.32, w: W, h: 0.36, fontFace: FB, bold: true, fontSize: String(f.judul).length > 70 ? 12.5 : 14, color: INK, valign: "middle", margin: 0 });
    await photoArea(s, f.photos, X0, top, 3.7, bottom - top - 0.27, true);
    s.addText("Lokasi: " + clip(f.lokasi || f.area, 70), { x: X0, y: bottom - 0.24, w: 3.7, h: 0.22, fontFace: F, fontSize: 8, color: GREY, italic: true, margin: 0 });
    const desc = f.deskripsi || "", risk = f.potensi_risiko ? "Potensi risiko: " + f.potensi_risiko : "";
    const sl = stdLines(f, true).map(x => clip(x, 130));
    const T = bottom - top, sh = Math.min(T * 0.55, Math.max(0.8, 0.36 + sl.length * 0.27)), dh = T - sh - 0.08;
    const dfs = fitFont([desc, risk], 5.4, dh - 0.22, 11, 7.5);
    s.addText([{ text: "Deskripsi Temuan", options: { bold: true, fontFace: FB, color: OR, fontSize: 10, breakLine: true } }, { text: desc, options: { fontSize: dfs, color: INK, breakLine: !!risk } }].concat(risk ? [{ text: risk, options: { fontSize: dfs - 0.5, color: GREY, italic: true } }] : []),
      { x: 4.25, y: top, w: 5.4, h: dh, fontFace: F, valign: "top", align: "justify", margin: 0, paraSpaceAfter: 3 });
    const sfs = fitFont(sl, 5.25, sh - 0.3, 9, 6.5);
    s.addShape(pptx.ShapeType.rect, { x: 4.25, y: top + dh + 0.08, w: 5.4, h: sh, fill: { color: P.SOFT }, line: { color: P.SOFT, width: 0 } });
    s.addText([{ text: "Remarks: standar / klausul / regulasi", options: { bold: true, fontFace: FB, color: OR, fontSize: 9, breakLine: true } }].concat(sl.length ? sl.map((x, k) => ({ text: x, options: { bullet: { indent: 10 }, fontSize: sfs, color: INK, breakLine: k < sl.length - 1 } })) : [{ text: "-", options: { fontSize: sfs, color: GREY } }]),
      { x: 4.33, y: top + dh + 0.12, w: 5.25, h: sh - 0.08, fontFace: F, valign: "top", margin: 0 });
    if (hasFU) { /* band tindak lanjut (Improvement plan, PIC, Due date, Status) — hanya tampil bila sudah diisi */
      const by = 4.38, bh = 1.0;
      s.addShape(pptx.ShapeType.rect, { x: X0, y: by, w: W, h: bh, fill: { color: "FFFFFF" }, line: { color: P.LINE, width: 0.75 } });
      s.addShape(pptx.ShapeType.rect, { x: X0, y: by, w: 0.08, h: bh, fill: { color: OR }, line: { color: OR, width: 0 } });
      const leftT = [f.tindakan_segera, ...ca], rightT = [(f.akar_masalah || {}).uraian, f.bukti_penutupan];
      const ffs = Math.min(fitFont(leftT, 5.3, bh - 0.1, 10, 6.5), fitFont(rightT, 3.4, bh - 0.1, 10, 6.5));
      s.addText([{ text: "Improvement plan", options: { bold: true, fontFace: FB, color: OR, fontSize: ffs, breakLine: true } }]
        .concat(f.tindakan_segera ? [{ text: "Segera: " + clip(f.tindakan_segera, 180), options: { fontSize: ffs, color: INK, breakLine: !!ca.length } }] : [])
        .concat(ca.map((x, k) => ({ text: x, options: { bullet: { type: "number" }, fontSize: ffs, color: INK, breakLine: k < ca.length - 1 } })))
        .concat(!f.tindakan_segera && !ca.length ? [{ text: "Belum diisi", options: { fontSize: ffs, color: GREY, italic: true } }] : []),
        { x: 0.55, y: by + 0.05, w: 5.4, h: bh - 0.08, fontFace: F, valign: "top", margin: 0 });
      s.addText([{ text: "PIC: ", options: { bold: true, fontFace: FB, fontSize: ffs, color: INK } }, { text: f.pic || "-", options: { fontSize: ffs, color: INK, breakLine: true } },
        { text: "Due date: ", options: { bold: true, fontFace: FB, fontSize: ffs, color: INK } }, { text: f.due ? fmtTgl(f.due) : "-", options: { fontSize: ffs, color: INK, breakLine: true } },
        { text: "Status: ", options: { bold: true, fontFace: FB, fontSize: ffs, color: INK } }, { text: f.status || "Open", options: { fontSize: ffs, color: P.ST[f.status || "Open"] || INK, bold: true, breakLine: !!f.bukti_penutupan } }]
        .concat(f.bukti_penutupan ? [{ text: "Bukti penutupan: ", options: { bold: true, fontFace: FB, fontSize: ffs, color: INK } }, { text: clip(f.bukti_penutupan, 90), options: { fontSize: ffs, color: INK } }] : []),
        { x: 6.1, y: by + 0.05, w: 3.45, h: bh - 0.08, fontFace: F, valign: "top", margin: 0 });
    }
  }

  /* E. good practice */
  if (opt.good) for (const g of sets.goods) {
    s = mk("Good SHE Practices", g.no);
    await photoArea(s, g.photos, X0, 1.1, 5.2, 4.25, true);
    const txt = [g.deskripsi, g.tindakan_pencegahan ? "Catatan peningkatan: " + g.tindakan_pencegahan : ""].filter(Boolean);
    const std = stdLines(g).length ? "Sesuai: " + clip(stdLines(g).join("; "), 220) : "";
    const gfs = fitFont([g.judul, ...txt, std], 3.9, 4.25, 12, 7.5);
    s.addText([{ text: g.judul, options: { bold: true, fontFace: FB, fontSize: gfs + 1.5, color: INK, breakLine: true } }]
      .concat(txt.map(t => ({ text: t, options: { fontSize: gfs, color: INK, breakLine: true } })))
      .concat(std ? [{ text: std, options: { fontSize: gfs - 1.5, color: GREY, italic: true } }] : []),
      { x: 5.75, y: 1.1, w: 3.9, h: 4.25, fontFace: F, valign: "top", align: "left", margin: 0, paraSpaceAfter: 5 });
  }

  /* F. kesimpulan */
  s = mk("Kesimpulan");
  const conc = opt.kesimpulan || autoConclusion(insp, sets);
  const recs = opt.rekomendasi && opt.rekomendasi.length ? opt.rekomendasi : autoRecs(sets);
  const cfs = fitFont([conc, ...recs.map(r => r.judul + ": " + r.uraian), "Rekomendasi"], 9.1, 4.2, 12, 8);
  s.addText(conc.split(/\n\s*\n|\n/).filter(Boolean).map(p => ({ text: p, options: { fontSize: cfs, color: INK, breakLine: true, paraSpaceAfter: 6 } }))
    .concat(recs.length ? [{ text: "Rekomendasi Tindakan Perbaikan", options: { fontSize: cfs, bold: true, fontFace: FB, color: INK, breakLine: true } }] : [])
    .concat(recs.flatMap((r, k) => [{ text: r.judul + (r.uraian ? ": " : ""), options: { bold: true, fontFace: FB, fontSize: cfs, color: INK, bullet: { indent: 12 } } }, { text: r.uraian || "", options: { fontSize: cfs, color: INK, breakLine: k < recs.length - 1 } }])),
    { x: 0.45, y: 1.1, w: 9.1, h: 4.25, fontFace: F, valign: "top", align: "justify", margin: 0 });

  const gen = await pptx.write({ outputType: "arraybuffer" });
  return await mergeIntoTemplate(gen, defs, insp);
}

/* ---- tempel slide hasil PptxGenJS ke file template: hal. 1–2 dipakai apa adanya, hal. isi memakai latar & judul template, hal. penutup dari template ---- */
async function mergeIntoTemplate(genBuf, defs, insp) {
  const resp = await fetch(TEMPLATE_PPT, { cache: "no-cache" });
  if (!resp.ok) throw new Error("File template PowerPoint (assets/template-ppt.pptx) tidak ditemukan.");
  const tpl = await JSZip.loadAsync(await resp.arrayBuffer());
  const gz = await JSZip.loadAsync(genBuf);
  const rd = async (z, p) => await z.file(p).async("string");
  const xmlEsc = t => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const REL = "http://schemas.openxmlformats.org/package/2006/relationships";

  /* 1. hal. 2 (cover): hanya isi teks placeholder, tata letak & desain tidak diubah */
  let s2 = await rd(tpl, "ppt/slides/slide2.xml");
  s2 = s2.replace("[Division/Department Name]", "QSHE Department");
  const t1 = xmlEsc(`Inspection Monitoring – ${insp.area}`), t2 = xmlEsc(fmtTgl(insp.tanggal));
  s2 = s2.replace('<a:p><a:endParaRPr lang="en-ID" dirty="0"/></a:p>',
    `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2000" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="DINPro-Bold"/></a:rPr><a:t>${t1}</a:t></a:r></a:p>` +
    `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="DINPro-Light"/></a:rPr><a:t>${t2}</a:t></a:r></a:p>`);
  tpl.file("ppt/slides/slide2.xml", s2);

  /* 2. hal. penutup (slide 6 template): judul diisi */
  let s6 = await rd(tpl, "ppt/slides/slide6.xml");
  s6 = s6.replace("<a:t>(Topic)</a:t>", "<a:t>Inspection Monitoring</a:t>");
  tpl.file("ppt/slides/slide6.xml", s6);

  /* 3. slide isi: salin slide contoh template (latar, garis & logo) + judul + bentuk hasil PptxGenJS */
  const base = await rd(tpl, "ppt/slides/slide3.xml");
  const baseRels = await rd(tpl, "ppt/slides/_rels/slide3.xml.rels");
  let ct = await rd(tpl, "[Content_Types].xml");
  let prs = await rd(tpl, "ppt/presentation.xml");
  let prsRels = await rd(tpl, "ppt/_rels/presentation.xml.rels");
  const slideRIds = [];
  for (let i = 0; i < defs.length; i++) {
    const gx = await rd(gz, `ppt/slides/slide${i + 1}.xml`);
    const grel = new DOMParser().parseFromString(await rd(gz, `ppt/slides/_rels/slide${i + 1}.xml.rels`), "application/xml");
    const map = {}; let extra = ""; let rn = 3;
    for (const r of [...grel.getElementsByTagName("Relationship")]) {
      if (!/\/image$/.test(r.getAttribute("Type"))) continue;
      const src = "ppt/" + r.getAttribute("Target").replace(/^\.\.\//, ""), name = `gen${i + 1}_${src.split("/").pop()}`;
      tpl.file("ppt/media/" + name, await gz.file(src).async("uint8array"));
      const id = "rId" + rn++; map[r.getAttribute("Id")] = id;
      extra += `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/>`;
    }
    let inner = (gx.match(/<\/p:grpSpPr>([\s\S]*)<\/p:spTree>/) || [, ""])[1];
    inner = inner.replace(/<p:cNvPr id="(\d+)"/g, (m, n) => `<p:cNvPr id="${+n + 20}"`).replace(/r:embed="(rId\d+)"/g, (m, id) => `r:embed="${map[id] || id}"`);
    let x = base.replace("<a:t>(Topic)</a:t>", `<a:t>${xmlEsc(defs[i].title)}</a:t>`)
      .replace('<a:off x="256511" y="-77031"/><a:ext cx="3922084" cy="970166"/>', '<a:off x="256511" y="-77031"/><a:ext cx="6900000" cy="970166"/>')
      .replace(/<a:solidFill><a:schemeClr val="accent2"\/><\/a:solidFill><a:latin typeface="DINPro-Bold"[^>]*\/>/, '<a:solidFill><a:srgbClr val="F26F21"/></a:solidFill><a:latin typeface="DINPro-Bold"/>')
      .replace(/<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2">[\s\S]*?<\/p:sp>/, "")
      .replace(/<p:extLst><p:ext uri="\{BB962C8B[\s\S]*?<\/p:extLst>/, "")
      .replace("</p:spTree>", inner + "</p:spTree>");
    const n = 7 + i;
    tpl.file(`ppt/slides/slide${n}.xml`, x);
    tpl.file(`ppt/slides/_rels/slide${n}.xml.rels`, baseRels.replace("</Relationships>", extra + "</Relationships>"));
    ct = ct.replace("</Types>", `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
    const rid = "rId" + (100 + i); slideRIds.push([300 + i, rid]);
    prsRels = prsRels.replace("</Relationships>", `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/></Relationships>`);
  }
  /* 4. buang slide contoh 3–5, atur urutan: 1, 2, isi…, penutup */
  [3, 4, 5].forEach(n => {
    tpl.remove(`ppt/slides/slide${n}.xml`); tpl.remove(`ppt/slides/_rels/slide${n}.xml.rels`);
    ct = ct.replace(new RegExp(`<Override PartName="/ppt/slides/slide${n}\\.xml"[^>]*/>`), "");
  });
  prsRels = prsRels.replace(/<Relationship Id="rId(6|7|8)"[^>]*\/>/g, "");
  prs = prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst><p:sldId id="256" r:id="rId4"/><p:sldId id="261" r:id="rId5"/>${slideRIds.map(([id, r]) => `<p:sldId id="${id}" r:id="${r}"/>`).join("")}<p:sldId id="264" r:id="rId9"/></p:sldIdLst>`);
  tpl.file("[Content_Types].xml", ct); tpl.file("ppt/presentation.xml", prs); tpl.file("ppt/_rels/presentation.xml.rels", prsRels);
  tpl.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Company>QSHE Department</Company></Properties>`);
  return await tpl.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", compression: "DEFLATE" });
}
function autoConclusion(insp, sets) {
  const c = s => sets.items.filter(f => f.severity === s).length;
  const top = sets.items.filter(f => f.severity !== "Minor").slice(0, 3).map(f => f.judul.toLowerCase());
  return `Hasil inspeksi di ${insp.area} pada ${fmtTgl(insp.tanggal)} mencatat ${sets.items.length} temuan (${c("Critical")} Critical, ${c("Major")} Major, ${c("Minor")} Minor) dan ${sets.goods.length} good SHE practice.` +
    (top.length ? `\n\nTindakan korektif yang perlu diprioritaskan meliputi: ${top.join("; ")}.` : "");
}
function autoRecs(sets) {
  return sets.items.filter(f => (f.tindakan_korektif || []).length).slice(0, 4).map(f => ({ judul: clip(f.judul, 50), uraian: f.tindakan_korektif[0].tindakan }));
}

/* ======================= BOOT ======================= */
function setTab(t) {
  S.tab = t;
  $$(".tab").forEach(b => b.setAttribute("aria-selected", b.dataset.tab === t));
  $$(".view").forEach(v => v.hidden = v.id !== "view-" + t);
  if (t === "register") renderRegister();
  if (t === "dashboard") renderDashboard();
  if (t === "laporan") renderReportSide();
  if (t === "catat") { renderInspSelect(); renderSession(); }
  window.scrollTo(0, 0);
}
$$(".tab").forEach(b => b.onclick = () => setTab(b.dataset.tab));
let _od = null;
function onData() {
  clearTimeout(_od);
  _od = setTimeout(() => {
    renderDatalists(); renderInspSelect(); renderSession(); updateReviewCount();
    if (S.tab === "register" && $("#mDetail").hidden) renderRegister();
    if (S.tab === "dashboard") renderDashboard();
    if (S.tab === "laporan") renderReportSide();
  }, 80);
}
applyTheme(); renderNotes(); renderDatalists(); renderInspSelect(); renderSession(); updateAINote();
Store.init();
/* identitas: mode perangkat QSHE diatur di Pengaturan (PIN) */

