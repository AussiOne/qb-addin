/* Quickbase Live Dashboard – PowerPoint content add-in
 *
 * Each instance of the add-in on a slide keeps its own settings (URL, zoom,
 * crop, pan, refresh) inside the .pptx via Office.context.document.settings.
 *
 * Geometry model (all "page px" = CSS px of the dashboard page itself):
 *   - The Quickbase page is rendered in an iframe of size  vw × vh  (the "page").
 *   - crop {l,t,r,b} cuts that page down to the region of interest.
 *   - zoom z scales page px -> screen px (computed from the frame size in fit modes).
 *   - pan {x,y} scrolls inside the cropped region when it is larger than the frame.
 *   - The PowerPoint frame itself is moved/resized with PowerPoint's own handles.
 *
 * Two zoom behaviours (state.zoomMode):
 *   "browser" (default) – behaves like Ctrl +/- in a browser. The page is laid out in a
 *       "window" of  (vw / zoom) px  wide with the box's aspect ratio, then scaled to fill the
 *       box. Zooming in narrows the layout, so Quickbase charts/tables reflow and resize.
 *       The layout does not depend on PowerPoint's editor zoom, so edit view and slideshow match.
 *   "magnify" – the page is laid out at a fixed vw × vh and simply scaled (like a picture).
 */
(function () {
  "use strict";

  const KEY = "qbLiveDashboard";
  const MIN_Z = 0.05, MAX_Z = 5;
  // Same zoom steps as Chrome / Edge
  const BROWSER_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5];
  const SWAP_DELAY_MS = 2500;   // let Quickbase render its widgets before swapping frames
  const SLOW_LOAD_MS = 25000;

  const DEFAULTS = {
    v: 1,
    url: "",
    vw: 1600, vh: 1000,
    zoomMode: "browser",     // browser | magnify
    fit: "width",            // width | contain | cover | none  (magnify mode)
    zoom: 1,
    crop: { l: 0, t: 0, r: 0, b: 0 },
    pan: { x: 0, y: 0 },
    refreshMin: 0,
    reloadOnShow: true,
    toolbarInShow: false,
    locked: false
  };

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const stage = $("stage"), clip = $("clip"), canvas = $("canvas");
  const shield = $("shield"), cropRect = $("cropRect"), hint = $("hint");
  const loading = $("loading"), loadingText = $("loadingText"), empty = $("empty");
  const zoomLabel = $("zoomLabel"), statusEl = $("status");
  const settingsPanel = $("settings"), form = $("settingsForm");
  let frame = $("qb");

  // ---------- State ----------
  let state = clone(DEFAULTS);
  let inOffice = false;
  let mode = "none";          // none | pan | crop
  let view = "edit";          // edit | read (slideshow)
  let lastLoaded = 0;
  let refreshTimer = null;
  let pendingFrame = null;
  let slowTimer = null;
  let geo = null;             // last computed geometry

  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function clamp(v, a, b) { return Math.min(b, Math.max(a, v)); }
  function num(v, d) { const n = Number(v); return Number.isFinite(n) ? n : d; }

  function normalize(s) {
    const o = Object.assign(clone(DEFAULTS), s || {});
    o.crop = Object.assign({ l: 0, t: 0, r: 0, b: 0 }, (s && s.crop) || {});
    o.pan = Object.assign({ x: 0, y: 0 }, (s && s.pan) || {});
    o.vw = clamp(num(o.vw, 1600), 320, 6000);
    o.vh = clamp(num(o.vh, 1000), 200, 12000);
    o.zoom = clamp(num(o.zoom, 1), MIN_Z, MAX_Z);
    if (!["width", "contain", "cover", "none"].includes(o.fit)) o.fit = "width";
    if (!["browser", "magnify"].includes(o.zoomMode)) o.zoomMode = "browser";
    ["l", "t", "r", "b"].forEach((k) => { o.crop[k] = Math.max(0, num(o.crop[k], 0)); });
    // never crop away the whole page
    if (o.crop.l + o.crop.r > o.vw - 20) { o.crop.l = 0; o.crop.r = 0; }
    if (o.crop.t + o.crop.b > o.vh - 20) { o.crop.t = 0; o.crop.b = 0; }
    o.refreshMin = clamp(num(o.refreshMin, 0), 0, 1440);
    return o;
  }

  // ---------- Persistence (document settings in Office, localStorage outside) ----------
  function loadState() {
    let raw = null;
    try {
      if (inOffice) raw = Office.context.document.settings.get(KEY);
      else raw = JSON.parse(localStorage.getItem(KEY) || "null");
    } catch (e) { raw = null; }
    state = normalize(raw);
  }

  let saveTimer = null;
  function saveState() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      try {
        if (inOffice) {
          Office.context.document.settings.set(KEY, clone(state));
          Office.context.document.settings.saveAsync((r) => {
            if (r.status !== Office.AsyncResultStatus.Succeeded) console.warn("Save failed", r.error);
          });
        } else {
          localStorage.setItem(KEY, JSON.stringify(state));
        }
      } catch (e) { console.warn(e); }
    }, 300);
  }

  // ---------- Geometry / rendering ----------
  function computeGeometry(full) {
    const sw = Math.max(1, stage.clientWidth), sh = Math.max(1, stage.clientHeight);
    const c = full ? { l: 0, t: 0, r: 0, b: 0 } : state.crop;

    if (state.zoomMode === "browser") {
      // Virtual browser window: vw/zoom wide, same aspect ratio as the box.
      const pageW = state.vw / state.zoom;
      const pageH = pageW * sh / sw;
      const cw = Math.max(20, pageW - c.l - c.r);
      const ch = Math.max(20, pageH - c.t - c.b);
      const z = Math.min(sw / cw, sh / ch);   // crop area is enlarged to fill the box
      const g = { sw, sh, c, fit: "contain", pageW, pageH, cw, ch, z,
                  visW: cw, visH: ch, maxPanX: 0, maxPanY: 0 };
      g.clipW = cw * z; g.clipH = ch * z;
      g.left = (sw - g.clipW) / 2; g.top = (sh - g.clipH) / 2;
      return g;
    }

    const fit = full ? "contain" : state.fit;
    const cw = Math.max(20, state.vw - c.l - c.r);
    const ch = Math.max(20, state.vh - c.t - c.b);
    let z;
    switch (fit) {
      case "width": z = sw / cw; break;
      case "contain": z = Math.min(sw / cw, sh / ch); break;
      case "cover": z = Math.max(sw / cw, sh / ch); break;
      default: z = state.zoom;
    }
    z = clamp(z, MIN_Z, MAX_Z);
    const visW = Math.min(cw, sw / z), visH = Math.min(ch, sh / z);
    const g = { sw, sh, c, fit, pageW: state.vw, pageH: state.vh, cw, ch, z,
                visW, visH, maxPanX: cw - visW, maxPanY: ch - visH };
    const clipW = visW * z, clipH = visH * z;
    g.clipW = clipW; g.clipH = clipH;
    g.left = clipW < sw ? (sw - clipW) / 2 : 0;
    g.top = clipH < sh && fit === "contain" ? (sh - clipH) / 2 : 0;
    return g;
  }

  function render() {
    const full = mode === "crop";
    const g = computeGeometry(full);
    const pan = full ? { x: 0, y: 0 } : state.pan;
    if (!full) {
      pan.x = clamp(pan.x, 0, g.maxPanX);
      pan.y = clamp(pan.y, 0, g.maxPanY);
    }
    geo = g;

    clip.style.left = g.left + "px";
    clip.style.top = g.top + "px";
    clip.style.width = g.clipW + "px";
    clip.style.height = g.clipH + "px";

    canvas.style.width = g.pageW + "px";
    canvas.style.height = g.pageH + "px";
    const tx = -(g.c.l + pan.x) * g.z, ty = -(g.c.t + pan.y) * g.z;
    canvas.style.transform = `translate(${tx}px, ${ty}px) scale(${g.z})`;

    const browser = state.zoomMode === "browser";
    zoomLabel.textContent = Math.round((browser ? state.zoom : g.z) * 100) + "%";
    zoomLabel.title = browser ? "Browser zoom – click to reset to 100%"
      : state.fit === "none" ? "Reset to 100%" : `Scaling: ${fitName(state.fit)} – click for 100%`;
    document.body.classList.toggle("zoom-browser", browser);

    document.body.classList.toggle("mode-pan", mode === "pan");
    document.body.classList.toggle("mode-crop", mode === "crop");
    document.body.classList.toggle("toolbar-pinned", mode !== "none");
    shield.hidden = !(mode !== "none" || state.locked);
    $("btnPan").classList.toggle("active", mode === "pan");
    $("btnCrop").classList.toggle("active", mode === "crop");
    empty.hidden = !!state.url;
  }

  function fitName(f) {
    return { width: "Fit width", contain: "Fit whole area", cover: "Fill frame", none: "Fixed" }[f];
  }

  // ---------- Loading / refresh ----------
  function isHttpsUrl(u) {
    try { return new URL(u).protocol === "https:"; } catch (e) { return false; }
  }

  function makeFrame() {
    const f = document.createElement("iframe");
    f.title = "Quickbase dashboard";
    f.setAttribute("referrerpolicy", "strict-origin-when-cross-origin");
    f.setAttribute("allow", "clipboard-read; clipboard-write; fullscreen");
    return f;
  }

  /** Load (or reload) the dashboard. Reloads are double-buffered so the slide never flashes blank. */
  function loadDashboard() {
    clearTimeout(slowTimer);
    if (!state.url || !isHttpsUrl(state.url)) {
      frame.removeAttribute("src");
      loading.hidden = true;
      render();
      return;
    }
    const firstLoad = !frame.getAttribute("src");

    if (firstLoad) {
      if (pendingFrame) { pendingFrame.remove(); pendingFrame = null; }
      loading.hidden = false;
      loadingText.textContent = "Loading dashboard…";
      frame.onload = () => { loaded(); };
      frame.src = state.url;
    } else {
      if (pendingFrame) pendingFrame.remove();
      const nf = makeFrame();
      nf.style.position = "absolute"; nf.style.inset = "0"; nf.style.visibility = "hidden";
      canvas.appendChild(nf);
      pendingFrame = nf;
      setStatus("Refreshing…");
      let swapped = false;
      const swap = () => {
        if (swapped || pendingFrame !== nf) return;
        swapped = true;
        frame.remove();
        nf.style.position = ""; nf.style.inset = ""; nf.style.visibility = "";
        nf.id = "qb";
        frame = nf; pendingFrame = null;
        loaded();
      };
      nf.onload = () => setTimeout(swap, SWAP_DELAY_MS);
      setTimeout(swap, SLOW_LOAD_MS);     // swap anyway if load never fires
      nf.src = state.url;
    }
    slowTimer = setTimeout(() => {
      if (!loading.hidden) {
        loadingText.innerHTML = "Still loading… If this stays blank, open <b>⚙ Settings › Troubleshooting</b>. " +
          "Your realm may block embedding, or you may need to sign in.";
      }
    }, SLOW_LOAD_MS);
    render();
  }

  function loaded() {
    clearTimeout(slowTimer);
    loading.hidden = true;
    lastLoaded = Date.now();
    const t = new Date(lastLoaded);
    setStatus("Updated " + t.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }));
  }

  function setStatus(s) { statusEl.textContent = s || ""; }

  function scheduleRefresh() {
    clearInterval(refreshTimer);
    if (state.refreshMin > 0) refreshTimer = setInterval(loadDashboard, state.refreshMin * 60000);
  }

  // ---------- Zoom ----------
  function setZoom(newZ, anchorX, anchorY) {
    if (state.zoomMode === "browser") {
      state.zoom = clamp(newZ, BROWSER_STEPS[0], BROWSER_STEPS[BROWSER_STEPS.length - 1]);
      render(); saveState();
      return;
    }
    const g = geo || computeGeometry(false);
    newZ = clamp(Math.round(newZ * 100) / 100, MIN_Z, MAX_Z);
    // keep the content point under the anchor (default: frame centre) fixed
    const ax = anchorX == null ? g.sw / 2 : anchorX;
    const ay = anchorY == null ? g.sh / 2 : anchorY;
    const px = state.pan.x + (ax - g.left) / g.z;
    const py = state.pan.y + (ay - g.top) / g.z;
    state.fit = "none";
    state.zoom = newZ;
    const ng = computeGeometry(false);
    state.pan.x = px - (ax - ng.left) / newZ;
    state.pan.y = py - (ay - ng.top) / newZ;
    render(); saveState();
  }
  const zoomStep = (dir) => {
    if (state.zoomMode === "browser") {
      const cur = state.zoom;
      const next = dir > 0 ? BROWSER_STEPS.find((s) => s > cur + 0.001)
                           : BROWSER_STEPS.slice().reverse().find((s) => s < cur - 0.001);
      if (next) setZoom(next);
      return;
    }
    const z = (geo || computeGeometry(false)).z;
    setZoom(dir > 0 ? z * 1.1 : z / 1.1);
  };

  function setFit(f) {
    state.fit = f; state.pan = { x: 0, y: 0 };
    render(); saveState();
  }

  // ---------- Pan / crop interaction ----------
  let drag = null;

  function setMode(m) {
    mode = mode === m ? "none" : m;
    cropRect.hidden = true;
    if (mode === "pan") showHint("Drag to move · scroll to pan · Ctrl+scroll to zoom · Esc to finish");
    else if (mode === "crop") showHint("Drag a rectangle around the area to keep · Esc to cancel");
    else showHint("");
    render();
  }

  function showHint(text) {
    hint.textContent = text;
    hint.hidden = !text;
  }

  function stagePoint(e) {
    const r = stage.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  shield.addEventListener("pointerdown", (e) => {
    if (mode === "none") return;
    e.preventDefault();
    shield.setPointerCapture(e.pointerId);
    const p = stagePoint(e);
    drag = { start: p, pan: Object.assign({}, state.pan) };
    document.body.classList.add("dragging");
  });

  shield.addEventListener("pointermove", (e) => {
    if (!drag) return;
    const p = stagePoint(e);
    if (mode === "pan") {
      const z = geo.z;
      state.pan.x = drag.pan.x - (p.x - drag.start.x) / z;
      state.pan.y = drag.pan.y - (p.y - drag.start.y) / z;
      render();
    } else if (mode === "crop") {
      const x = Math.min(p.x, drag.start.x), y = Math.min(p.y, drag.start.y);
      const w = Math.abs(p.x - drag.start.x), h = Math.abs(p.y - drag.start.y);
      Object.assign(cropRect.style, { left: x + "px", top: y + "px", width: w + "px", height: h + "px" });
      cropRect.hidden = false;
    }
  });

  function endDrag(e) {
    if (!drag) return;
    const p = stagePoint(e);
    document.body.classList.remove("dragging");
    if (mode === "pan") {
      saveState();
    } else if (mode === "crop") {
      const g = geo; // full-page geometry while cropping
      const toPage = (sx, sy) => ({
        x: clamp((sx - g.left) / g.z, 0, g.pageW),
        y: clamp((sy - g.top) / g.z, 0, g.pageH)
      });
      const a = toPage(Math.min(p.x, drag.start.x), Math.min(p.y, drag.start.y));
      const b = toPage(Math.max(p.x, drag.start.x), Math.max(p.y, drag.start.y));
      if (b.x - a.x >= 20 && b.y - a.y >= 20) {
        state.crop = {
          l: Math.round(a.x), t: Math.round(a.y),
          r: Math.round(g.pageW - b.x), b: Math.round(g.pageH - b.y)
        };
        state.pan = { x: 0, y: 0 };
        if (state.zoomMode === "magnify") state.fit = "contain";
        saveState();
      }
      cropRect.hidden = true;
      mode = "none"; showHint("");
      render();
    }
    drag = null;
  }
  shield.addEventListener("pointerup", endDrag);
  shield.addEventListener("pointercancel", endDrag);

  shield.addEventListener("wheel", (e) => {
    if (mode === "crop") return;
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) {
      const p = stagePoint(e);
      if (state.zoomMode === "browser") zoomStep(e.deltaY < 0 ? 1 : -1);
      else setZoom(geo.z * (e.deltaY < 0 ? 1.1 : 1 / 1.1), p.x, p.y);
    } else if (mode === "pan") {
      state.pan.x += e.deltaX / geo.z;
      state.pan.y += e.deltaY / geo.z;
      render(); saveState();
    }
  }, { passive: false });

  // ---------- Toolbar visibility ----------
  // A cross-origin iframe swallows mouse events, so a thin hot zone along the top edge
  // (plus any movement over our own UI) reveals the toolbar.
  const hot = document.createElement("div");
  hot.id = "hotzone";
  Object.assign(hot.style, { position: "absolute", left: 0, right: 0, top: 0, height: "14px", zIndex: 9 });
  document.body.appendChild(hot);

  let tbTimer = null;
  function pokeToolbar() {
    document.body.classList.add("show-toolbar");
    clearTimeout(tbTimer);
    tbTimer = setTimeout(() => {
      if (!$("toolbar").matches(":hover")) document.body.classList.remove("show-toolbar");
      else pokeToolbar();
    }, 2500);
  }
  hot.addEventListener("mouseenter", pokeToolbar);
  document.addEventListener("mousemove", pokeToolbar);
  $("toolbar").addEventListener("mouseenter", pokeToolbar);

  // ---------- Settings form ----------
  function openSettings() {
    mode = "none"; showHint(""); render();
    const f = form.elements;
    f.url.value = state.url;
    f.vw.value = state.vw; f.vh.value = state.vh;
    f.fit.value = state.fit;
    f.zoomMode.value = state.zoomMode;
    f.zoom.value = Math.round(state.zoom * 100);
    f.cl.value = state.crop.l; f.ct.value = state.crop.t;
    f.cr.value = state.crop.r; f.cb.value = state.crop.b;
    f.refreshMin.value = state.refreshMin;
    f.reloadOnShow.checked = state.reloadOnShow;
    f.toolbarInShow.checked = state.toolbarInShow;
    f.locked.checked = state.locked;
    settingsPanel.hidden = false;
    setTimeout(() => f.url.focus(), 0);
  }
  function closeSettings() { settingsPanel.hidden = true; }

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const f = form.elements;
    const url = f.url.value.trim();
    if (!isHttpsUrl(url)) {
      f.url.setCustomValidity("Enter a full https:// URL");
      f.url.reportValidity();
      return;
    }
    f.url.setCustomValidity("");
    const urlChanged = url !== state.url;
    const zoomInput = num(f.zoom.value, 100) / 100;
    const fitWas = state.fit + state.zoomMode;
    state = normalize({
      url,
      vw: f.vw.value, vh: f.vh.value,
      fit: f.fit.value,
      zoomMode: f.zoomMode.value,
      zoom: zoomInput,
      crop: { l: f.cl.value, t: f.ct.value, r: f.cr.value, b: f.cb.value },
      pan: state.pan,
      refreshMin: f.refreshMin.value,
      reloadOnShow: f.reloadOnShow.checked,
      toolbarInShow: f.toolbarInShow.checked,
      locked: f.locked.checked
    });
    if (urlChanged || fitWas !== state.fit + state.zoomMode) state.pan = { x: 0, y: 0 };
    applyBodyFlags();
    saveState();
    closeSettings();
    scheduleRefresh();
    if (urlChanged) { frame.removeAttribute("src"); }
    loadDashboard();
  });
  form.elements.url.addEventListener("input", () => form.elements.url.setCustomValidity(""));

  function applyBodyFlags() {
    document.body.classList.toggle("toolbar-in-show", !!state.toolbarInShow);
    document.body.classList.toggle("slideshow", view === "read");
  }

  // ---------- External windows ----------
  function currentOrigin() {
    const raw = (form.elements.url.value || state.url || "").trim();
    try { return new URL(raw).origin; } catch (e) { return null; }
  }

  function signIn() {
    const origin = currentOrigin();
    if (!origin) { alertInline("Enter your Quickbase URL first."); return; }
    const dlg = new URL("dialog.html", location.href);
    dlg.searchParams.set("to", origin + "/");
    if (inOffice && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(dlg.href, { height: 75, width: 45, promptBeforeOpen: false }, (r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          window.open(origin + "/", "_blank");
          return;
        }
        const d = r.value;
        // Fires when the user closes the window (12006) or it navigates somewhere unreachable.
        d.addEventHandler(Office.EventType.DialogEventReceived, () => {
          try { d.close(); } catch (e) { /* already closed */ }
          frame.removeAttribute("src");
          loadDashboard();
        });
      });
    } else {
      window.open(origin + "/", "_blank");
    }
  }

  function openExternal() {
    const url = (form.elements.url.value || state.url || "").trim();
    if (!isHttpsUrl(url)) return;
    try {
      if (inOffice && Office.context.ui && Office.context.ui.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(url);
        return;
      }
    } catch (e) { /* fall through */ }
    window.open(url, "_blank");
  }

  function alertInline(msg) {
    const f = form.elements.url;
    f.setCustomValidity(msg); f.reportValidity();
  }

  // ---------- Actions ----------
  const actions = {
    zoomIn: () => zoomStep(1),
    zoomOut: () => zoomStep(-1),
    zoomReset: () => setZoom(1),
    fitWidth: () => { if (state.zoomMode === "magnify") setFit("width"); },
    fitAll: () => { if (state.zoomMode === "magnify") setFit("contain"); },
    pan: () => { if (state.zoomMode === "magnify") setMode("pan"); },
    crop: () => setMode("crop"),
    uncrop: () => { state.crop = { l: 0, t: 0, r: 0, b: 0 }; state.pan = { x: 0, y: 0 }; render(); saveState(); },
    reload: () => loadDashboard(),
    settings: openSettings,
    closeSettings,
    signin: signIn,
    openExternal
  };

  document.addEventListener("click", (e) => {
    const el = e.target.closest("[data-action]");
    if (!el) return;
    const fn = actions[el.dataset.action];
    if (fn) { e.preventDefault(); fn(); }
  });

  document.addEventListener("keydown", (e) => {
    if (e.target.closest("input, select, textarea")) {
      if (e.key === "Escape") closeSettings();
      return;
    }
    if (!settingsPanel.hidden) { if (e.key === "Escape") closeSettings(); return; }
    const k = e.key.toLowerCase();
    const map = { "+": "zoomIn", "=": "zoomIn", "-": "zoomOut", "_": "zoomOut", "0": "zoomReset",
                  w: "fitWidth", f: "fitAll", p: "pan", c: "crop", r: "reload" };
    if (k === "escape") { if (mode !== "none") { mode = "none"; cropRect.hidden = true; showHint(""); render(); } return; }
    if (map[k] && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); actions[map[k]](); }
  });

  // Re-fit whenever the frame is resized on the slide.
  if (window.ResizeObserver) new ResizeObserver(() => render()).observe(stage);
  window.addEventListener("resize", render);

  // Refresh stale data when the add-in becomes visible again.
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && state.url && state.refreshMin > 0 &&
        Date.now() - lastLoaded > state.refreshMin * 60000) loadDashboard();
  });

  // ---------- Slideshow detection ----------
  function setView(v) {
    const prev = view;
    view = v === "read" ? "read" : "edit";
    applyBodyFlags();
    if (view === "read" && prev !== "read" && state.reloadOnShow && lastLoaded && Date.now() - lastLoaded > 5000) {
      loadDashboard();
    }
  }

  // ---------- Boot ----------
  function start() {
    loadState();
    applyBodyFlags();
    render();
    scheduleRefresh();
    if (state.url) loadDashboard();
    // Expose for debugging / tests
    window.__qb = { get state() { return state; }, get geo() { return geo; }, render, setZoom, setFit, setMode };
  }

  function bootOffice() {
    inOffice = true;
    try {
      const doc = Office.context.document;
      if (doc.getActiveViewAsync) {
        doc.getActiveViewAsync((r) => { if (r.status === Office.AsyncResultStatus.Succeeded) setView(r.value); });
        doc.addHandlerAsync(Office.EventType.ActiveViewChanged, (ev) => setView(ev.activeView));
      }
    } catch (e) { console.warn("View detection unavailable", e); }
    start();
  }

  if (window.Office && Office.onReady) {
    Office.onReady((info) => {
      if (info && info.host) bootOffice();
      else start(); // opened directly in a browser (testing)
    });
  } else {
    start();
  }
})();
