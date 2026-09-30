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
const Cap = { jenis: "temuan", photos: [], editor: null, ctl: null };
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

/* ---------------- capture view ---------------- */
function baseRecord(insp) {
  return {
    id: uid("f_"), inspId: insp.id, no: "", jenis: Cap.jenis, area: insp.area, tanggal: insp.tanggal,
    lokasi: $("#inLokasi").value.trim(), raw: $("#inRaw").value.trim(), photos: [], closePhotos: [],
    judul: "", deskripsi: "", bahaya: "", potensi_risiko: "", kategori_aspek: ["K3"], kategori_temuan: "Operational",
    klausul_iso: [], kriteria_smk3: [], regulasi: [], severity: Cap.jenis === "good" ? "" : "Minor", alasan_severity: "", stop_work: false,
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

function renderPhotoGrid() {
  const g = $("#photoGrid");
  g.innerHTML = Cap.photos.map((p, i) => `<div class="ph"><img alt="Foto ${i + 1}" src="${p.dataUrl}"><button type="button" data-rm="${i}" aria-label="Hapus foto">×</button></div>`).join("")
    + (Cap.photos.length < 4 ? `<label class="ph-add" for="filePhoto"><span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>Ambil / pilih foto</span></label>` : "");
  $$("[data-rm]", g).forEach(b => b.onclick = () => { Cap.photos.splice(+b.dataset.rm, 1); renderPhotoGrid(); });
}
$("#filePhoto").addEventListener("change", async e => {
  const files = [...e.target.files].slice(0, 4 - Cap.photos.length); e.target.value = "";
  for (const f of files) { try { Cap.photos.push(await compressImage(f)); } catch (err) { toast(err.message || "Foto gagal diproses", "err"); } }
  renderPhotoGrid();
});
$$("#segJenis button").forEach(b => b.onclick = () => { Cap.jenis = b.dataset.v; $$("#segJenis button").forEach(x => x.setAttribute("aria-pressed", x === b)); $("#inRaw").placeholder = Cap.jenis === "good" ? "contoh: toolbox di lemari no. 1 sudah dikelompokkan per jenis & terkunci semua" : "contoh: sepatu safety di rak workshop solnya banyak yg getas & robek, masih dicampur sama yg bagus"; });

function currentInsp() { return inspById(S.currentInsp); }
function requireInsp() {
  const i = currentInsp(); if (i) return i;
  toast("Buat atau pilih inspeksi dulu, supaya temuan tercatat di area & tanggal yang benar.", "err"); openInspModal(); return null;
}

$("#btnAI").onclick = async () => {
  const insp = requireInsp(); if (!insp) return;
  const raw = $("#inRaw").value.trim();
  if (raw.length < 8 && !Cap.photos.length) { toast("Tulis catatan singkat (minimal satu kalimat) atau lampirkan foto.", "err"); $("#inRaw").focus(); return; }
  Cap.ctl = new AbortController();
  const area = $("#aiArea"); const t0 = Date.now();
  area.innerHTML = `<div class="ai-busy"><div style="flex:1"><div class="small" id="aiMsg">Menganalisis catatan${Cap.photos.length ? " dan foto" : ""}… biasanya 20–60 detik</div><div class="bar" style="margin-top:8px"></div></div><button class="btn" id="btnStop">Batalkan</button></div>`;
  $("#btnStop").onclick = () => Cap.ctl && Cap.ctl.abort();
  $("#btnAI").disabled = true;
  try {
    const rec = baseRecord(insp); if (!raw && Cap.photos.length) rec.raw = "(lihat foto)";
    const ai = await runFindingAI({ jenis: Cap.jenis, area: insp.area, lokasi: rec.lokasi, tanggal: fmtTgl(insp.tanggal), raw: rec.raw }, Cap.photos, Cap.ctl.signal,
      n => { const m = $("#aiMsg"); if (m) m.textContent = `Menyusun laporan… ${Math.round((Date.now() - t0) / 1000)} detik`; });
    Object.assign(rec, ai);
    rec.pic = ai.pic_saran || ""; delete rec.pic_saran;
    if (rec.jenis !== "good") rec.due = addDays(insp.tanggal, S.settings.due[rec.severity] || 30);
    if (!rec.lokasi && ai.lokasi) rec.lokasi = ai.lokasi;
    openCaptureEditor(rec);
    area.innerHTML = "";
  } catch (e) {
    area.innerHTML = e && e.code === "cancelled" ? "" : `<div class="note" style="border-left-color:var(--crit)">${esc(aiErrMsg(e))}</div>`;
  } finally { $("#btnAI").disabled = !AI.fn; Cap.ctl = null; }
};
$("#btnManual").onclick = () => {
  const insp = requireInsp(); if (!insp) return;
  const rec = baseRecord(insp); rec.deskripsi = rec.raw;
  openCaptureEditor(rec);
};

$("#btnRapi").onclick = () => {
  const insp = requireInsp(); if (!insp) return;
  const raw = $("#inRaw").value.trim();
  if (raw.length < 8) { toast("Tulis catatan singkat (minimal satu kalimat) dulu.", "err"); $("#inRaw").focus(); return; }
  const rec = baseRecord(insp);
  const r = RAPI.susun({ jenis: rec.jenis, area: insp.area, lokasi: rec.lokasi, raw });
  Object.assign(rec, r.fields);
  const ch = r.changes.slice(0, 14).map(([a, b]) => `<span style="white-space:nowrap"><s>${esc(a)}</s> → <b>${esc(b)}</b></span>`).join(", ");
  $("#aiArea").innerHTML = `<div class="note small">${r.topics.length ? `<b>Topik terdeteksi:</b> ${esc(r.topics.join(", "))}.<br>` : ""}${ch ? `<b>Perbaikan bahasa:</b> ${ch}${r.changes.length > 14 ? ", …" : ""}` : "Tidak ada kata tidak baku yang perlu diganti."}</div>`;
  openCaptureEditor(rec);
};

function openCaptureEditor(rec) {
  const host = $("#editorHost");
  Cap.editor = renderEditor(host, rec, {
    mode: "new",
    footHTML: w => `<button class="btn ghost" data-act="discard">Buang</button><button class="btn" data-act="draft">Simpan draft</button><button class="btn primary" data-act="review">Kirim ke Supervisor</button>`,
    onBind: (h) => {
      $$("[data-act]", h).forEach(b => b.onclick = async () => {
        if (b.dataset.act === "discard") { if (confirm("Buang hasil ini? Foto dan catatan di formulir kiri tetap ada.")) { host.innerHTML = ""; Cap.editor = null; } return; }
        await saveCaptured(b.dataset.act === "review" ? "Review" : "Draft", b);
      });
    }
  });
  host.scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth", block: "start" });
}
async function saveCaptured(review, btn) {
  const insp = currentInsp(); if (!insp || !Cap.editor) return;
  const w = Cap.editor.get();
  if (!w.judul.trim()) { toast("Judul temuan masih kosong.", "err"); return; }
  $$("#editorHost .sheet-foot button").forEach(b => b.disabled = true); btn.textContent = "Menyimpan…";
  try {
    const photos = [];
    for (const p of Cap.photos) { const id = uid("ph_"); await Store.putPhoto(id, p.dataUrl); photos.push({ id, w: p.w, h: p.h }); }
    const rec = { ...w, photos, review, no: nextNo(insp, w.jenis), lokasi: w.lokasi || $("#inLokasi").value.trim() };
    if (review === "Review") rec.sentAt = Date.now();
    await Store.put("findings", rec);
    toast(review === "Review" ? `${rec.no} dikirim ke Supervisor untuk review.` : `${rec.no} tersimpan sebagai draft.`);
    Cap.photos = []; renderPhotoGrid(); $("#inRaw").value = ""; $("#inLokasi").value = "";
    $("#editorHost").innerHTML = ""; Cap.editor = null; $("#aiArea").innerHTML = "";
  } catch (e) {
    toast(e.message || "Gagal menyimpan.", "err");
    $$("#editorHost .sheet-foot button").forEach(b => b.disabled = false); btn.textContent = review === "Review" ? "Kirim ke Supervisor" : "Simpan draft";
  }
}

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
  box.innerHTML = items.length ? items.map(fcardHTML).join("") : `<div class="empty">Belum ada entri. Foto kondisi di lapangan, tulis catatan singkat, lalu tekan <b>Rapikan jadi laporan</b>.</div>`;
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

/* ======================= EXCEL ======================= */
function colLetter(n) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function xDate(iso) { const d = parseISO(iso); return d ? new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())) : null; }
async function buildExcel(items, goods, meta) {
  await loadScript(LIBS.excel);
  const wb = new ExcelJS.Workbook(); wb.creator = "QSHE Department"; wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  const ORG = "FFF36F20", BLK = "FF13294B", WHT = "FFFFFFFF", GRY = "FFF2F3F4", LINE = "FFBFC4C9";
  const SEVF = { Critical: "FFC4162A", Major: "FFC77700", Minor: "FF2A6496" };
  const border = { top: { style: "thin", color: { argb: LINE } }, left: { style: "thin", color: { argb: LINE } }, bottom: { style: "thin", color: { argb: LINE } }, right: { style: "thin", color: { argb: LINE } } };

  const COLS = [
    ["No", "no", 5], ["ID Temuan", "id", 18], ["Tanggal", "tgl", 11], ["Area", "area", 16], ["Lokasi", "lokasi", 18], ["Aspek", "aspek", 12],
    ["Standard", "std", 15], ["Clause", "clause", 16], ["Kriteria SMK3 (PP 50/2012)", "smk3", 30], ["Regulasi Terkait", "reg", 36],
    ["Category Finding", "sev", 11], ["Compliance", "c1", 7], ["Operational", "c2", 7], ["Administration", "c3", 7],
    ["Judul", "judul", 28], ["Finding", "finding", 52], ["Potensi Risiko", "risk", 30], ["Root Cause (dugaan awal)", "rc", 30],
    ["Tindakan Segera", "imm", 30], ["Corrective Action", "ca", 46], ["Preventive Action", "pa", 30], ["Bukti Penutupan", "evreq", 26],
    ["PIC", "pic", 16], ["Inspector", "insp", 16], ["Due Date", "due", 11], ["Status", "status", 12], ["Closed Date", "closed", 11],
    ["Remarks", "remarks", 28], ["Foto Temuan", "ph", 24], ["Foto Penutupan", "ph2", 24], ["Tahap Review", "review", 13]
  ];
  const ci = k => COLS.findIndex(c => c[1] === k) + 1, L = k => colLetter(ci(k));
  const last = colLetter(COLS.length), R0 = 7;
  const ws = wb.addWorksheet("Register Temuan", { views: [{ state: "frozen", xSplit: 2, ySplit: 6, zoomScale: 90 }], pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } }, properties: { tabColor: { argb: ORG } } });
  COLS.forEach((c, i) => ws.getColumn(i + 1).width = c[2]);
  ws.mergeCells(`A1:${last}1`); ws.getCell("A1").value = "REGISTER TEMUAN INSPEKSI QSHE";
  ws.getCell("A1").font = { name: "Arial", size: 16, bold: true, color: { argb: WHT } }; ws.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: ORG } };
  ws.getCell("A1").alignment = { vertical: "middle", indent: 1 }; ws.getRow(1).height = 30;
  ws.mergeCells(`A2:${last}2`); ws.getCell("A2").value = `QSHE Department — ${meta.title}`; ws.getCell("A2").font = { name: "Arial", size: 11, bold: true };
  ws.mergeCells(`A3:${last}3`);
  const insp = meta.insp;
  ws.getCell("A3").value = (insp ? `Agenda: ${insp.agenda || "-"}   |   Inspector: ${(insp.inspectors || []).join(", ") || "-"}   |   Inspectee: ${(insp.inspectees || []).join(", ") || "-"}   |   ` : "") + `Dicetak: ${fmtTgl(localISO())}` + (meta.isDraft ? "   |   DRAFT — untuk review tim" : "   |   Status: Final");
  ws.getCell("A3").font = { name: "Arial", size: 9, color: { argb: meta.isDraft ? "FFC4162A" : "FF5E656D" }, bold: !!meta.isDraft };

  // header (2 rows, "Category" grouped like the audit summary format)
  const r5 = ws.getRow(5), r6 = ws.getRow(6);
  COLS.forEach((c, i) => {
    const col = i + 1;
    if (["c1", "c2", "c3"].includes(c[1])) { r6.getCell(col).value = c[0]; }
    else { ws.mergeCells(5, col, 6, col); r5.getCell(col).value = c[0]; }
  });
  ws.mergeCells(5, ci("c1"), 5, ci("c3")); r5.getCell(ci("c1")).value = "Category";
  [r5, r6].forEach(r => { for (let c = 1; c <= COLS.length; c++) { const cell = r.getCell(c); cell.font = { name: "Arial", size: 9, bold: true, color: { argb: WHT } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLK } }; cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true }; cell.border = border; } });
  r6.getCell(ci("c1")).fill = r6.getCell(ci("c2")).fill = r6.getCell(ci("c3")).fill = r5.getCell(ci("c1")).fill = { type: "pattern", pattern: "solid", fgColor: { argb: ORG } };
  r5.height = 18; r6.height = 30;

  const imgBox = async (list, colKey, rowNo) => {
    const p = (list || [])[0]; if (!p) return;
    const d = await Store.getPhoto(p.id); if (!d) return;
    const id = wb.addImage({ base64: d, extension: "jpeg" });
    const maxW = COLS[ci(colKey) - 1][2] * 7 - 8, maxH = 118; const r = Math.min(maxW / (p.w || 4), maxH / (p.h || 3));
    ws.addImage(id, { tl: { col: ci(colKey) - 1 + 0.08, row: rowNo - 1 + 0.08 }, ext: { width: Math.round((p.w || 4) * r), height: Math.round((p.h || 3) * r) }, editAs: "oneCell" });
  };

  for (let i = 0; i < items.length; i++) {
    const f = items[i]; const rn = R0 + i; const row = ws.getRow(rn);
    const insp2 = inspById(f.inspId);
    const std = [...new Set((f.klausul_iso || []).map(k => k.standar))];
    const v = {
      no: i + 1, id: f.no, tgl: xDate(f.tanggal), area: f.area, lokasi: f.lokasi, aspek: (f.kategori_aspek || []).join(", "),
      std: std.concat((f.kriteria_smk3 || []).length ? ["SMK3 PP 50/2012"] : []).join("\n"), clause: (f.klausul_iso || []).map(k => `${k.standar} Kl. ${k.klausul}${k.judul ? " – " + k.judul : ""}`).join("\n"),
      smk3: (f.kriteria_smk3 || []).map(k => `${k.kode} ${k.ringkas}`).join("\n"), reg: (f.regulasi || []).map(r => `${r.peraturan}${r.pasal ? ", " + r.pasal : ""}${r.perlu_verifikasi ? " (perlu verifikasi)" : ""}`).join("\n"),
      sev: f.severity, c1: f.kategori_temuan === "Compliance" ? "X" : "", c2: f.kategori_temuan === "Operational" ? "X" : "", c3: f.kategori_temuan === "Administration" ? "X" : "",
      judul: f.judul, finding: f.deskripsi + (f.stop_work ? "\n[STOP WORK — area/pekerjaan dihentikan]" : ""), risk: f.potensi_risiko,
      rc: [(f.akar_masalah || {}).kategori, (f.akar_masalah || {}).uraian].filter(Boolean).join(": "), imm: f.tindakan_segera,
      ca: caLines(f).map((x, j) => `${j + 1}. ${x}`).join("\n"), pa: f.tindakan_pencegahan, evreq: f.bukti_penutupan,
      pic: f.pic, insp: (insp2 && insp2.inspectors || []).join(", "), due: xDate(f.due), status: f.status || "Open", closed: xDate(f.closedDate),
      remarks: f.remarks, ph: "", ph2: "", review: REVIEW_LABEL[f.review || "Draft"]
    };
    COLS.forEach((c, j) => { const cell = row.getCell(j + 1); cell.value = v[c[1]] ?? ""; cell.font = { name: "Arial", size: 9 }; cell.alignment = { vertical: "top", wrapText: true, horizontal: ["no", "c1", "c2", "c3", "sev", "status", "tgl", "due", "closed", "review"].includes(c[1]) ? "center" : "left" }; cell.border = border; });
    ["tgl", "due", "closed"].forEach(k => row.getCell(ci(k)).numFmt = "dd/mm/yyyy");
    row.getCell(ci("id")).font = { name: "Arial", size: 9, bold: true };
    row.getCell(ci("judul")).font = { name: "Arial", size: 9, bold: true };
    row.height = (f.photos || []).length || (f.closePhotos || []).length ? 96 : Math.min(160, Math.max(48, Math.ceil(String(v.finding).length / 60) * 12 + 10));
    await imgBox(f.photos, "ph", rn); await imgBox(f.closePhotos, "ph2", rn);
  }
  const RN = Math.max(R0, R0 + items.length - 1), RMAX = R0 + Math.max(items.length, 1) + 300;
  if (!items.length) { ws.mergeCells(`A${R0}:${last}${R0}`); ws.getCell(`A${R0}`).value = "Tidak ada temuan pada filter ini."; ws.getCell(`A${R0}`).font = { name: "Arial", size: 9, italic: true }; }
  ws.autoFilter = { from: { row: 6, column: 1 }, to: { row: 6, column: COLS.length } };
  // validations so the sheet stays usable as a live tracker
  const rng = k => `${L(k)}${R0}:${L(k)}${RMAX}`;
  ws.dataValidations.add(rng("sev"), { type: "list", allowBlank: true, formulae: ['"Critical,Major,Minor"'] });
  ws.dataValidations.add(rng("status"), { type: "list", allowBlank: false, formulae: ['"Open,On Progress,Closed"'], showErrorMessage: true, errorTitle: "Status", error: "Pilih Open, On Progress, atau Closed." });
  ws.dataValidations.add(rng("review"), { type: "list", allowBlank: true, formulae: ['"Draft,Menunggu Supervisor,Menunggu QSHE,Final"'] });
  const fill = argb => ({ type: "pattern", pattern: "solid", bgColor: { argb } });
  ws.addConditionalFormatting({ ref: rng("sev"), rules: SEVS.map((s, k) => ({ type: "expression", priority: k + 1, formulae: [`$${L("sev")}${R0}="${s}"`], style: { fill: fill(SEVF[s]), font: { color: { argb: WHT }, bold: true } } })) });
  ws.addConditionalFormatting({ ref: rng("status"), rules: [
    { type: "expression", priority: 4, formulae: [`$${L("status")}${R0}="Closed"`], style: { fill: fill("FFE3F2E8"), font: { color: { argb: "FF2B7A4B" }, bold: true } } },
    { type: "expression", priority: 5, formulae: [`$${L("status")}${R0}="On Progress"`], style: { fill: fill("FFFFF0D9"), font: { color: { argb: "FFA35E00" }, bold: true } } },
    { type: "expression", priority: 6, formulae: [`$${L("status")}${R0}="Open"`], style: { fill: fill("FFFBE7E9"), font: { color: { argb: "FFC4162A" }, bold: true } } }] });
  ws.addConditionalFormatting({ ref: rng("due"), rules: [{ type: "expression", priority: 7, formulae: [`AND($${L("status")}${R0}<>"Closed",$${L("due")}${R0}<>"",$${L("due")}${R0}<TODAY())`], style: { fill: fill("FFC4162A"), font: { color: { argb: WHT }, bold: true } } }] });

  /* ---- Rekap sheet (live formulas) ---- */
  const rk = wb.addWorksheet("Rekap", { properties: { tabColor: { argb: BLK } }, views: [{ showGridLines: false }], pageSetup: { orientation: "portrait", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  rk.columns = [{ width: 30 }, { width: 11 }, { width: 11 }, { width: 13 }, { width: 11 }, { width: 16 }, { width: 12 }];
  rk.mergeCells("A1:G1"); rk.getCell("A1").value = "REKAP STATUS TINDAK LANJUT"; rk.getCell("A1").font = { name: "Arial", size: 14, bold: true, color: { argb: WHT } }; rk.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: ORG } }; rk.getRow(1).height = 26;
  rk.getCell("A2").value = `Sumber: sheet Register Temuan (baris ${R0}–${RMAX}). Angka otomatis berubah bila status di register diubah.`; rk.getCell("A2").font = { name: "Arial", size: 9, italic: true, color: { argb: "FF5E656D" } };
  const RG = k => `'Register Temuan'!$${L(k)}$${R0}:$${L(k)}$${RMAX}`;
  const hdr = ["", "Jumlah", "Open", "On Progress", "Closed", "Lewat batas waktu", "% Closed"];
  let r = 4;
  const block = (title, keys, critFn, label) => {
    rk.getCell(`A${r}`).value = title; rk.getCell(`A${r}`).font = { name: "Arial", size: 11, bold: true }; r++;
    const hr = rk.getRow(r); hdr.forEach((h, j) => { const c = hr.getCell(j + 1); c.value = j ? h : label; c.font = { name: "Arial", size: 9, bold: true, color: { argb: WHT } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLK } }; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = border; });
    r++; const first = r;
    keys.forEach(k => {
      const row = rk.getRow(r); const crit = critFn(r);
      row.getCell(1).value = k;
      row.getCell(2).value = { formula: `COUNTIFS(${crit})` };
      row.getCell(3).value = { formula: `COUNTIFS(${crit},${RG("status")},"Open")` };
      row.getCell(4).value = { formula: `COUNTIFS(${crit},${RG("status")},"On Progress")` };
      row.getCell(5).value = { formula: `COUNTIFS(${crit},${RG("status")},"Closed")` };
      row.getCell(6).value = { formula: `COUNTIFS(${crit},${RG("status")},"<>Closed",${RG("due")},"<"&TODAY())` };
      row.getCell(7).value = { formula: `IF(B${r}=0,0,E${r}/B${r})` }; row.getCell(7).numFmt = "0%";
      for (let j = 1; j <= 7; j++) { const c = row.getCell(j); c.font = { name: "Arial", size: 9, bold: j === 1 }; c.border = border; if (j > 1) c.alignment = { horizontal: "center" }; }
      if (SEVF[k]) { row.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: SEVF[k] } }; row.getCell(1).font = { name: "Arial", size: 9, bold: true, color: { argb: WHT } }; }
      r++;
    });
    const tr = rk.getRow(r); tr.getCell(1).value = "Total";
    for (let j = 2; j <= 6; j++) tr.getCell(j).value = { formula: `SUM(${colLetter(j)}${first}:${colLetter(j)}${r - 1})` };
    tr.getCell(7).value = { formula: `IF(B${r}=0,0,E${r}/B${r})` }; tr.getCell(7).numFmt = "0%";
    for (let j = 1; j <= 7; j++) { const c = tr.getCell(j); c.font = { name: "Arial", size: 9, bold: true }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRY } }; c.border = border; if (j > 1) c.alignment = { horizontal: "center" }; }
    r += 3;
  };
  block("Per Tingkat Temuan", SEVS, rr => `${RG("sev")},$A${rr}`, "Tingkat");
  const pics = [...new Set(items.map(f => f.pic || "").filter(Boolean))].sort();
  if (pics.length) block("Per PIC", pics, rr => `${RG("pic")},$A${rr}`, "PIC");
  const areas = [...new Set(items.map(f => f.area || "").filter(Boolean))].sort();
  if (areas.length > 1) block("Per Area", areas, rr => `${RG("area")},$A${rr}`, "Area");
  block("Per Aspek", ["K3", "Lingkungan", "Mutu"], rr => `${RG("aspek")},"*"&$A${rr}&"*"`, "Aspek");

  /* ---- Good practice sheet ---- */
  if (goods.length) {
    const gs = wb.addWorksheet("Good Practice", { properties: { tabColor: { argb: "FF2B7A4B" } }, views: [{ state: "frozen", ySplit: 3 }] });
    const GC = [["No", 5], ["ID", 18], ["Tanggal", 11], ["Area", 16], ["Lokasi", 18], ["Judul", 28], ["Uraian", 60], ["Persyaratan yang terpenuhi", 40], ["Saran peningkatan", 34], ["Foto", 24]];
    GC.forEach((c, i) => gs.getColumn(i + 1).width = c[1]);
    gs.mergeCells(`A1:${colLetter(GC.length)}1`); gs.getCell("A1").value = "GOOD SHE PRACTICES"; gs.getCell("A1").font = { name: "Arial", size: 14, bold: true, color: { argb: WHT } }; gs.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2B7A4B" } }; gs.getRow(1).height = 26;
    const h = gs.getRow(3); GC.forEach((c, i) => { const cell = h.getCell(i + 1); cell.value = c[0]; cell.font = { name: "Arial", size: 9, bold: true, color: { argb: WHT } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLK } }; cell.border = border; cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; });
    for (let i = 0; i < goods.length; i++) {
      const f = goods[i], rn = 4 + i, row = gs.getRow(rn);
      [i + 1, f.no, xDate(f.tanggal), f.area, f.lokasi, f.judul, f.deskripsi, stdLines(f).join("\n"), f.tindakan_pencegahan, ""].forEach((v, j) => { const c = row.getCell(j + 1); c.value = v ?? ""; c.font = { name: "Arial", size: 9, bold: j === 5 }; c.alignment = { vertical: "top", wrapText: true }; c.border = border; });
      row.getCell(3).numFmt = "dd/mm/yyyy"; row.height = 96;
      const p = (f.photos || [])[0];
      if (p) { const d = await Store.getPhoto(p.id); if (d) { const id = wb.addImage({ base64: d, extension: "jpeg" }); const rr = Math.min(160 / (p.w || 4), 118 / (p.h || 3)); gs.addImage(id, { tl: { col: 9.08, row: rn - 1 + 0.08 }, ext: { width: Math.round((p.w || 4) * rr), height: Math.round((p.h || 3) * rr) }, editAs: "oneCell" }); } }
    }
  }

  /* ---- Referensi sheet ---- */
  const rf = wb.addWorksheet("Kriteria & Referensi", { views: [{ showGridLines: false }] });
  rf.columns = [{ width: 22 }, { width: 90 }, { width: 18 }];
  const put = (a, b, c, st) => { const row = rf.addRow([a, b, c]); row.eachCell(cell => { cell.font = { name: "Arial", size: 9, bold: !!st }; cell.alignment = { vertical: "top", wrapText: true }; cell.border = border; if (st === "h") { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLK } }; cell.font = { name: "Arial", size: 9, bold: true, color: { argb: WHT } }; } }); return row; };
  rf.addRow(["KRITERIA TINGKAT TEMUAN & BATAS WAKTU"]).font = { name: "Arial", size: 12, bold: true };
  put("Tingkat", "Kriteria", "Batas waktu tindakan korektif", "h");
  const d = S.settings.due;
  const kr = [["Critical", "Potensi cedera berat/fatal atau LTI; ATAU pelanggaran regulasi pemerintah/perizinan yang berisiko sanksi; ATAU potensi kerusakan/pencemaran lingkungan signifikan. Pekerjaan/area dihentikan atau diisolasi segera (stop work).", `${d.Critical} hari (tindakan segera ≤ 24 jam)`],
    ["Major", "Potensi atau aktual cedera ringan; ATAU temuan/peringatan regulator yang tidak fatal; ATAU ketidaksesuaian proses yang berdampak pada operasional.", `${d.Major} hari`],
    ["Minor", "Selain Critical dan Major: dampak minimal, observasi, atau celah kecil dokumentasi.", `${d.Minor} hari`]];
  kr.forEach(k => { const row = put(...k); row.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: SEVF[k[0]] } }; row.getCell(1).font = { name: "Arial", size: 9, bold: true, color: { argb: WHT } }; });
  rf.addRow([]); rf.addRow(["HIERARKI PENGENDALIAN (ISO 45001:2018 Klausul 8.1.2)"]).font = { name: "Arial", size: 12, bold: true };
  put("Level", "Contoh penerapan", "", "h");
  [["Eliminasi", "Menghilangkan bahaya, mis. menyingkirkan APD rusak dari stok, membongkar instalasi yang tidak dipakai."], ["Substitusi", "Mengganti dengan yang kurang berbahaya, mis. bahan pembersih berbasis air."], ["Rekayasa", "Guarding, jack stand, barikade fisik, ventilasi, penataan ulang area kerja."], ["Administratif", "Prosedur/IK, rambu, inspeksi berkala, pelatihan, izin kerja."], ["APD", "Penyediaan APD yang sesuai dan layak pakai beserta instruksi pemakaiannya."]].forEach(x => put(x[0], x[1], ""));
  rf.addRow([]); rf.addRow(["STATUS"]).font = { name: "Arial", size: 12, bold: true };
  put("Status", "Arti", "", "h");
  [["Open", "Belum ada tindakan korektif yang berjalan."], ["On Progress", "Tindakan korektif sedang berjalan, belum ada bukti penutupan."], ["Closed", "Tindakan selesai dan bukti penutupan telah diverifikasi QSHE."]].forEach(x => put(x[0], x[1], ""));

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

/* ======================= POWERPOINT ======================= */
const P = { F: "Arial", OR: "F36F20", NAVY: "13294B", INK: "16191E", GREY: "4A515B", SOFT: "E9ECEF", LINE: "D3D8DF", SEV: { Critical: "C4162A", Major: "C77700", Minor: "2A6496" }, ST: { Open: "C4162A", "On Progress": "C77700", Closed: "2B7A4B" } };
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
async function buildPPT(insp, sets, opt) {
  await loadScript(LIBS.pptx);
  const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_4x3";
  pptx.author = "QSHE Department"; pptx.company = "QSHE Department"; pptx.title = `QSHE Inspection — ${insp.area} — ${insp.tanggal}`;
  pptx.defineSlideMaster({ title: "TITLE", background: { data: ASSET.bgTitle }, objects: [
    { rect: { x: 7.05, y: 0.52, w: 0.07, h: 0.34, fill: { color: P.OR } } },
    { text: { text: "QSHE Department", options: { x: 7.2, y: 0.5, w: 2.5, h: 0.38, fontFace: P.F, fontSize: 16, bold: true, color: "FFFFFF", margin: 0, valign: "middle" } } }
  ] });
  pptx.defineSlideMaster({ title: "CONTENT", background: { data: ASSET.bgContent }, objects: [
    { text: { text: "QSHE Department", options: { x: 7.85, y: 0.3, w: 1.8, h: 0.3, fontFace: P.F, fontSize: 11, bold: true, color: P.NAVY, align: "right", margin: 0 } } },
    { text: { text: "Laporan Inspeksi QSHE", options: { x: 0.45, y: 7.24, w: 4, h: 0.24, fontFace: P.F, fontSize: 8, color: P.GREY, margin: 0, valign: "middle" } } }
  ], slideNumber: { x: 9.2, y: 7.1, w: 0.5, h: 0.25, fontFace: P.F, fontSize: 8, color: P.GREY, align: "right" } });
  const { F, OR, INK, GREY } = P;
  const head = (s, t, sub, right) => {
    s.addText(t, { x: 0.35, y: 0.3, w: 7.4, h: 0.42, fontFace: F, fontSize: 20, bold: true, color: P.NAVY, margin: 0 });
    if (sub) s.addText(sub, { x: 0.35, y: 0.7, w: 7.4, h: 0.3, fontFace: F, fontSize: 12, color: OR, margin: 0 });
    if (right) s.addText(right, { x: 5.6, y: 1.17, w: 4.05, h: 0.22, fontFace: F, fontSize: 9, color: GREY, align: "right", margin: 0 });
    if (sets.isDraft) s.addText("DRAFT – untuk review tim", { x: 0.35, y: 1.17, w: 2.2, h: 0.22, fontFace: F, fontSize: 8, bold: true, color: "C4162A", margin: 0 });
  };
  const photo = async (s, p, x, y, w, h, caption) => {
    const d = p ? await Store.getPhoto(p.id) : null;
    if (d) { s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: "F2F3F4" }, line: { color: P.LINE, width: 0.5 } }); const r = fitRect(p.w, p.h, x + 0.04, y + 0.04, w - 0.08, h - 0.08); s.addImage({ data: d, x: r.x, y: r.y, w: r.w, h: r.h }); }
    else { s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: "F2F3F4" }, line: { color: P.LINE, width: 0.5, dashType: "dash" } }); s.addText(caption || "Tidak ada foto", { x, y, w, h, fontFace: F, fontSize: 9, color: GREY, align: "center", valign: "middle" }); }
  };
  const sevChip = (s, sev, x, y) => s.addText(sev || "-", { x, y, w: 1.05, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: "FFFFFF", fill: { color: P.SEV[sev] || GREY }, align: "center", valign: "middle", margin: 0 });
  const inspectors = (insp.inspectors || []).join(", ") || "-";
  const allN = sets.items.length;

  /* 1. cover */
  let s = pptx.addSlide({ masterName: "TITLE" });
  s.addText([{ text: "Executive Summary", options: { fontSize: 30, breakLine: true } }, { text: `${insp.area} Inspection Result`, options: { fontSize: insp.area.length > 26 ? 20 : insp.area.length > 14 ? 24 : 30, breakLine: true } }, { text: fmtTgl(insp.tanggal), options: { fontSize: 20 } }],
    { x: 1.6, y: 2.15, w: 6.3, h: 2.1, fontFace: F, color: "FFFFFF", valign: "top", margin: 0 });
  s.addText("Inspector by: " + inspectors, { x: 3.1, y: 4.55, w: 5.2, h: 0.45, fontFace: F, fontSize: 16, color: "FFFFFF", margin: 0 });
  if (sets.isDraft) s.addText("DRAFT – untuk review tim", { x: 3.1, y: 5.15, w: 2.6, h: 0.36, fontFace: F, fontSize: 11, bold: true, color: INK, fill: { color: "FFFFFF" }, align: "center", valign: "middle", margin: 0 });

  /* 2. tracking */
  if (opt.track && sets.tracked.length) {
    for (let i = 0; i < sets.tracked.length; i += 2) {
      s = pptx.addSlide({ masterName: "CONTENT" });
      head(s, "IMPROVEMENT TRACKING", "Status Tindak Lanjut Temuan Sebelumnya", `Tracking ${i + 1}–${Math.min(i + 2, sets.tracked.length)} dari ${sets.tracked.length}`);
      const pair = sets.tracked.slice(i, i + 2);
      for (let j = 0; j < pair.length; j++) {
        const f = pair[j], y = 1.5 + j * 2.9, st = f.status || "Open";
        s.addShape(pptx.ShapeType.rect, { x: 0.35, y, w: 9.3, h: 2.7, fill: { color: "FFFFFF" }, line: { color: P.LINE, width: 0.75 } });
        s.addShape(pptx.ShapeType.rect, { x: 0.35, y, w: 0.08, h: 2.7, fill: { color: P.ST[st] || GREY }, line: { color: P.ST[st] || GREY, width: 0 } });
        s.addText(`${i + j + 1}.  ${f.no}`, { x: 0.55, y: y + 0.08, w: 3, h: 0.28, fontFace: F, fontSize: 10, bold: true, color: GREY, margin: 0 });
        s.addText(st, { x: 4.3, y: y + 0.08, w: 1.1, h: 0.28, fontFace: F, fontSize: 10, bold: true, color: "FFFFFF", fill: { color: P.ST[st] || GREY }, align: "center", valign: "middle", margin: 0 });
        const body = [
          { text: clip(f.judul, 90), options: { bold: true, fontSize: 11, color: INK, breakLine: true } },
          { text: clip(f.deskripsi, 300), options: { fontSize: 9, color: INK, breakLine: true } },
          { text: clip(stdLines(f).join("; "), 170), options: { fontSize: 8, color: GREY, italic: true, breakLine: true } },
          { text: `PIC: ${f.pic || "-"}   |   Batas waktu: ${fmtTgl(f.due)}${st === "Closed" && f.closedDate ? "   |   Ditutup: " + fmtTgl(f.closedDate) : ""}`, options: { fontSize: 9, bold: true, color: INK, breakLine: !!f.remarks } }
        ];
        if (f.remarks) body.push({ text: "Progres: " + clip(f.remarks, 140), options: { fontSize: 8.5, color: GREY } });
        s.addText(body, { x: 0.55, y: y + 0.42, w: 4.85, h: 2.2, fontFace: F, valign: "top", margin: 0, paraSpaceAfter: 3 });
        await photo(s, (f.photos || [])[0], 5.55, y + 0.12, 1.95, 2.2, "Tidak ada foto");
        await photo(s, (f.closePhotos || [])[0], 7.6, y + 0.12, 1.95, 2.2, st === "Closed" ? "Foto bukti belum diunggah" : "Belum ada bukti penutupan");
        s.addText("Sebelum", { x: 5.55, y: y + 2.36, w: 1.95, h: 0.22, fontFace: F, fontSize: 8, color: GREY, align: "center", margin: 0 });
        s.addText("Sesudah", { x: 7.6, y: y + 2.36, w: 1.95, h: 0.22, fontFace: F, fontSize: 8, color: GREY, align: "center", margin: 0 });
      }
    }
  }

  /* 3. overview */
  s = pptx.addSlide({ masterName: "CONTENT" });
  head(s, "INSPECTION OVERVIEW", "(Ringkasan Pelaksanaan Inspeksi)");
  const ins = insp.inspectors || [], ise = insp.inspectees || [];
  const kv = [["Agenda", insp.agenda || "-"], ["Tanggal Pelaksanaan", fmtTgl(insp.tanggal, true)], ["Area", insp.area]];
  for (let k = 0; k < Math.max(3, ins.length); k++) kv.push([`Nama Inspector ${k + 1}`, ins[k] || "-"]);
  for (let k = 0; k < Math.max(3, ise.length); k++) kv.push([`Nama Inspectee ${k + 1}`, ise[k] || "-"]);
  s.addTable(kv.map(([a, b]) => [{ text: a, options: { color: OR, fontSize: 12 } }, { text: ": " + b, options: { color: INK, fontSize: 12 } }]),
    { x: 0.35, y: 1.55, w: 5.3, colW: [2.1, 3.2], fontFace: F, border: { type: "none" }, rowH: 0.36, margin: 0.02 });
  const cnt = sv => sets.items.filter(f => f.severity === sv).length;
  s.addText("Hasil inspeksi", { x: 6.0, y: 1.55, w: 3.6, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: INK, margin: 0 });
  [["Critical", cnt("Critical")], ["Major", cnt("Major")], ["Minor", cnt("Minor")]].forEach(([sv, n], k) => {
    const y = 1.95 + k * 0.95;
    s.addShape(pptx.ShapeType.rect, { x: 6.0, y, w: 3.6, h: 0.8, fill: { color: "FFFFFF" }, line: { color: P.LINE, width: 0.75 } });
    s.addShape(pptx.ShapeType.rect, { x: 6.0, y, w: 0.1, h: 0.8, fill: { color: P.SEV[sv] }, line: { color: P.SEV[sv], width: 0 } });
    s.addText(String(n), { x: 6.2, y, w: 1.0, h: 0.8, fontFace: F, fontSize: 30, bold: true, color: P.SEV[sv], valign: "middle", margin: 0 });
    s.addText(sv, { x: 7.2, y, w: 2.3, h: 0.8, fontFace: F, fontSize: 13, color: INK, valign: "middle", margin: 0 });
  });
  s.addText(`${sets.goods.length} good SHE practice   |   ${sets.items.filter(f => f.stop_work).length} stop work`, { x: 6.0, y: 4.85, w: 3.6, h: 0.3, fontFace: F, fontSize: 10, color: GREY, margin: 0 });

  /* 4. summary table */
  if (sets.items.length) {
    const per = 7;
    for (let i = 0; i < sets.items.length; i += per) {
      s = pptx.addSlide({ masterName: "CONTENT" });
      head(s, "RINGKASAN TEMUAN", "(Daftar Temuan Periode Berjalan)", sets.items.length > per ? `Halaman ${i / per + 1} dari ${Math.ceil(sets.items.length / per)}` : "");
      const hdrOpt = { bold: true, color: "FFFFFF", fill: { color: P.NAVY }, align: "center", valign: "middle", fontSize: 10 };
      const rows = [[{ text: "No", options: hdrOpt }, { text: "Temuan", options: hdrOpt }, { text: "Standar / Klausul", options: hdrOpt }, { text: "Tingkat", options: hdrOpt }, { text: "PIC", options: hdrOpt }, { text: "Batas Waktu", options: hdrOpt }]];
      sets.items.slice(i, i + per).forEach((f, k) => {
        const zebra = { fill: { color: k % 2 ? "FFFFFF" : "FDF0E8" } };
        rows.push([
          { text: String(i + k + 1), options: { ...zebra, align: "center", bold: true } },
          { text: [{ text: clip(f.judul, 80), options: { bold: true, breakLine: true } }, { text: clip(f.lokasi || f.area, 60), options: { color: GREY, fontSize: 8 } }], options: zebra },
          { text: clip((f.klausul_iso || []).map(k2 => `${k2.standar.replace(/:\d{4}/, "")} ${k2.klausul}`).concat((f.kriteria_smk3 || []).map(k2 => "SMK3 " + k2.kode)).join("\n"), 110), options: { ...zebra, fontSize: 8 } },
          { text: f.severity, options: { bold: true, color: "FFFFFF", fill: { color: P.SEV[f.severity] || GREY }, align: "center" } },
          { text: f.pic || "-", options: { ...zebra, align: "center" } },
          { text: fmtShort(f.due), options: { ...zebra, align: "center" } }
        ]);
      });
      s.addTable(rows, { x: 0.35, y: 1.5, w: 9.3, colW: [0.45, 3.55, 2.05, 1.0, 1.15, 1.1], fontFace: F, fontSize: 9, color: INK, valign: "middle", border: { type: "solid", pt: 0.5, color: P.LINE }, rowH: [0.36].concat(Array(Math.min(per, sets.items.length - i)).fill(0.72)), margin: 0.05 });
    }
  }

  /* 5. one slide per finding */
  for (let i = 0; i < sets.items.length; i++) {
    const f = sets.items[i];
    s = pptx.addSlide({ masterName: "CONTENT" });
    head(s, "CURRENT INSPECTION FINDINGS", "(Temuan Inspeksi Periode Berjalan)", `Temuan ${i + 1} dari ${allN}`);
    sevChip(s, f.severity, 0.35, 1.5);
    s.addText(`${f.no}   |   PIC: ${f.pic || "-"}   |   Batas waktu: ${fmtTgl(f.due)}${f.stop_work ? "   |   STOP WORK" : ""}`, { x: 1.5, y: 1.5, w: 8.15, h: 0.3, fontFace: F, fontSize: 10, bold: true, color: f.stop_work ? P.SEV.Critical : GREY, valign: "middle", margin: 0 });
    s.addText(clip(f.judul, 110), { x: 0.35, y: 1.88, w: 9.3, h: 0.45, fontFace: F, fontSize: f.judul.length > 70 ? 13 : 15, bold: true, color: INK, valign: "middle", margin: 0 });
    // photos (left)
    const ph = (f.photos || []).slice(0, 2);
    if (ph.length <= 1) await photo(s, ph[0], 0.35, 2.45, 3.75, 2.85);
    else { await photo(s, ph[0], 0.35, 2.45, 3.75, 1.4); await photo(s, ph[1], 0.35, 3.9, 3.75, 1.4); }
    s.addText("Lokasi: " + clip(f.lokasi || f.area, 70), { x: 0.35, y: 5.33, w: 3.75, h: 0.24, fontFace: F, fontSize: 8, color: GREY, italic: true, margin: 0 });
    // description (right)
    const desc = f.deskripsi || "", risk = f.potensi_risiko ? "Potensi risiko: " + f.potensi_risiko : "";
    const dfs = fitFont([desc, risk], 5.35, 1.7, 11, 8);
    s.addText([{ text: "Deskripsi Temuan", options: { bold: true, color: OR, fontSize: 10, breakLine: true } }, { text: desc, options: { fontSize: dfs, color: INK, breakLine: !!risk } }].concat(risk ? [{ text: risk, options: { fontSize: dfs - 0.5, color: GREY, italic: true } }] : []),
      { x: 4.3, y: 2.42, w: 5.35, h: 1.95, fontFace: F, valign: "top", align: "justify", margin: 0, paraSpaceAfter: 3 });
    const sl = stdLines(f, true).map(x => clip(x, 115));
    const sfs = fitFont(sl, 5.2, 0.85, 9, 7);
    s.addShape(pptx.ShapeType.rect, { x: 4.3, y: 4.4, w: 5.35, h: 1.15, fill: { color: P.SOFT }, line: { color: P.SOFT, width: 0 } });
    s.addText([{ text: "Standar / Klausul / Regulasi", options: { bold: true, color: OR, fontSize: 9, breakLine: true } }].concat(sl.map((x, k) => ({ text: x, options: { bullet: { indent: 10 }, fontSize: sfs, color: INK, breakLine: k < sl.length - 1 } }))),
      { x: 4.38, y: 4.44, w: 5.2, h: 1.08, fontFace: F, valign: "top", margin: 0 });
    // follow-up band
    s.addShape(pptx.ShapeType.rect, { x: 0.35, y: 5.68, w: 9.3, h: 1.45, fill: { color: "FFFFFF" }, line: { color: P.LINE, width: 0.75 } });
    s.addShape(pptx.ShapeType.rect, { x: 0.35, y: 5.68, w: 0.08, h: 1.45, fill: { color: P.NAVY }, line: { color: OR, width: 0 } });
    const ca = caLines(f).map(x => clip(x, 150));
    const leftT = [f.tindakan_segera, ...ca], rightT = [(f.akar_masalah || {}).uraian, f.bukti_penutupan];
    const ffs = Math.min(fitFont(leftT, 5.3, 1.1, 10, 6.5), fitFont(rightT, 3.4, 1.1, 10, 6.5));
    s.addText([{ text: "Tindakan segera: ", options: { bold: true, fontSize: ffs, color: INK } }, { text: clip(f.tindakan_segera || "-", 200), options: { fontSize: ffs, color: INK, breakLine: true } }, { text: "Tindakan korektif:", options: { bold: true, fontSize: ffs, color: INK, breakLine: true } }]
      .concat(ca.map((x, k) => ({ text: x, options: { bullet: { type: "number" }, fontSize: ffs, color: INK, breakLine: k < ca.length - 1 } }))),
      { x: 0.55, y: 5.74, w: 5.4, h: 1.34, fontFace: F, valign: "top", margin: 0 });
    s.addText([{ text: "Akar masalah (dugaan awal): ", options: { bold: true, fontSize: ffs, color: INK } }, { text: clip([(f.akar_masalah || {}).kategori, String((f.akar_masalah || {}).uraian || "").replace(/^Dugaan awal:\s*/i, "")].filter(Boolean).join(" – "), 160), options: { fontSize: ffs, color: INK, breakLine: true } },
      { text: "Bukti penutupan: ", options: { bold: true, fontSize: ffs, color: INK } }, { text: clip(f.bukti_penutupan || "-", 140), options: { fontSize: ffs, color: INK } }],
      { x: 6.1, y: 5.74, w: 3.45, h: 1.34, fontFace: F, valign: "top", margin: 0 });
  }

  /* 6. good practice */
  if (opt.good) for (const g of sets.goods) {
    s = pptx.addSlide({ masterName: "CONTENT" });
    head(s, "GOOD SHE PRACTICES", "(Implementasi Standar K3L Yang Baik)", g.no);
    const gp = (g.photos || []).slice(0, 2);
    if (gp.length <= 1) await photo(s, gp[0], 0.35, 1.5, 9.3, 3.55);
    else { await photo(s, gp[0], 0.35, 1.5, 4.6, 3.55); await photo(s, gp[1], 5.05, 1.5, 4.6, 3.55); }
    const txt = [g.deskripsi, g.tindakan_pencegahan ? "Catatan peningkatan: " + g.tindakan_pencegahan : ""].filter(Boolean);
    const gfs = fitFont([g.judul, ...txt, stdLines(g).join("; ")], 9.3, 1.95, 12, 8);
    s.addText([{ text: g.judul, options: { bold: true, fontSize: gfs + 1, color: INK, breakLine: true } }].concat(txt.map((t, k) => ({ text: t, options: { fontSize: gfs, color: INK, breakLine: true } })))
      .concat(stdLines(g).length ? [{ text: "Sesuai: " + clip(stdLines(g).join("; "), 200), options: { fontSize: gfs - 1.5, color: GREY, italic: true } }] : []),
      { x: 0.35, y: 5.15, w: 9.3, h: 2.0, fontFace: F, valign: "top", align: "justify", margin: 0, paraSpaceAfter: 4 });
  }

  /* 7. conclusion */
  s = pptx.addSlide({ masterName: "CONTENT" });
  head(s, "KESIMPULAN", "");
  const conc = opt.kesimpulan || autoConclusion(insp, sets);
  const recs = opt.rekomendasi && opt.rekomendasi.length ? opt.rekomendasi : autoRecs(sets);
  const cfs = fitFont([conc, ...recs.map(r => r.judul + ": " + r.uraian), "Rekomendasi"], 9.1, 5.6, 12, 8);
  const paras = conc.split(/\n\s*\n|\n/).filter(Boolean);
  s.addText(paras.map((p, k) => ({ text: p, options: { fontSize: cfs, color: INK, breakLine: true, paraSpaceAfter: 6 } }))
    .concat([{ text: "Rekomendasi Tindakan Perbaikan", options: { fontSize: cfs, bold: true, underline: { style: "sng" }, color: INK, breakLine: true } }])
    .concat(recs.flatMap((r, k) => [{ text: r.judul + (r.uraian ? ": " : ""), options: { bold: true, fontSize: cfs, color: INK, bullet: { indent: 12 } } }, { text: r.uraian || "", options: { fontSize: cfs, color: INK, breakLine: k < recs.length - 1 } }])),
    { x: 0.45, y: 1.45, w: 9.1, h: 5.7, fontFace: F, valign: "top", align: "justify", margin: 0 });

  /* 8. closing */
  s = pptx.addSlide({ masterName: "TITLE" });
  s.addText("Terima Kasih", { x: 1.6, y: 2.6, w: 6, h: 0.9, fontFace: F, fontSize: 36, color: "FFFFFF", margin: 0 });

  return await pptx.write({ outputType: "blob" });
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
applyTheme(); renderPhotoGrid(); renderDatalists(); renderInspSelect(); renderSession(); updateAINote();
Store.init();
/* identitas: mode perangkat QSHE diatur di Pengaturan (PIN) */

