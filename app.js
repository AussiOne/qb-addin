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
    alignX: "center",        // left | center | right   (snap when content is smaller than the box)
    alignY: "middle",        // top | middle | bottom
    bg: "",                  // "" = white, else #rrggbb (picked or sampled)
    shapeId: "",             // this add-in's shape id on the slide (learned on first enlarge)
    restore: null,           // original shape rect while temporarily enlarged
    refreshMin: 0,
    reloadOnShow: true,
    toolbarInShow: false,
    locked: false,
    sharp: false,            // browser mode: window = box size (1:1 pixels, sharpest)
    renderer: "auto",        // auto | zoom | transform  (how the page is scaled)
    snap: null,              // {pictureId, home, takenAt} while the snapshot picture is on the slide
    snapMode: false,         // true = picture on the slide, live box parked beside it
    snapDelay: 8,            // seconds to wait after the dashboard loads before capturing
    snapHiRes: true,         // enlarge the box while capturing for a sharper picture
    slideId: ""              // slide this box was last found on
  };

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const stage = $("stage"), clip = $("clip"), canvas = $("canvas");
  const shield = $("shield"), cropRect = $("cropRect"), hint = $("hint");
  const loading = $("loading"), loadingText = $("loadingText"), empty = $("empty");
  const zoomLabel = $("zoomLabel"), statusEl = $("status");
  const settingsPanel = $("settings"), form = $("settingsForm");
  const posPanel = $("posPanel"), bgPick = $("bgPick"), toolbar = $("toolbar");
  const shotPanel = $("shotPanel");
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
    if (!["left", "center", "right"].includes(o.alignX)) o.alignX = "center";
    if (!["top", "middle", "bottom"].includes(o.alignY)) o.alignY = "middle";
    if (!/^#[0-9a-f]{6}$/i.test(o.bg || "")) o.bg = "";
    const r = o.restore;
    if (!r || typeof r.shapeId !== "string" || typeof r.slideId !== "string" ||
        !["left", "top", "width", "height"].every((k) => Number.isFinite(r[k]))) o.restore = null;
    if (typeof o.shapeId !== "string") o.shapeId = "";
    o.sharp = !!o.sharp;
    if (!["auto", "zoom", "transform"].includes(o.renderer)) o.renderer = "auto";
    const sn = o.snap;
    if (!sn || typeof sn.pictureId !== "string" || !sn.home) o.snap = null;
    o.snapMode = !!o.snapMode;
    o.snapDelay = clamp(num(o.snapDelay, 8), 0, 300);
    o.snapHiRes = o.snapHiRes !== false;
    if (typeof o.slideId !== "string") o.slideId = "";
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
  const AX = { left: 0, center: 0.5, right: 1 }, AY = { top: 0, middle: 0.5, bottom: 1 };
  function place(g) {
    g.left = Math.max(0, g.sw - g.clipW) * AX[state.alignX];
    g.top = Math.max(0, g.sh - g.clipH) * AY[state.alignY];
    return g;
  }
  function computeGeometry(full) {
    const sw = Math.max(1, stage.clientWidth), sh = Math.max(1, stage.clientHeight);
    const c = full ? { l: 0, t: 0, r: 0, b: 0 } : state.crop;

    if (state.zoomMode === "browser") {
      // Virtual browser window: vw/zoom wide, same aspect ratio as the box.
      // Sharp: lay out at the box's real pixel size, so 100% zoom = 1:1 pixels (no blur,
      // thin lines like text carets stay visible). Otherwise use the fixed window width.
      const pageW = Math.round((state.sharp ? sw : state.vw) / state.zoom);
      const pageH = Math.round(pageW * sh / sw);
      const cw = Math.max(20, pageW - c.l - c.r);
      const ch = Math.max(20, pageH - c.t - c.b);
      const z = Math.min(sw / cw, sh / ch);   // crop area is enlarged to fill the box
      const g = { sw, sh, c, fit: "contain", pageW, pageH, cw, ch, z,
                  visW: cw, visH: ch, maxPanX: 0, maxPanY: 0 };
      g.clipW = cw * z; g.clipH = ch * z;
      return place(g);
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
    return place(g);
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
    if (useNativeZoom()) {
      // CSS zoom is passed into the embedded page as a real browser zoom level
      // (its devicePixelRatio changes), exactly like Ctrl +/- in Edge. Unlike transform:
      // scale() this keeps the page's own custom mouse pointers working (e.g. Excel's
      // white cross over the cell grid) and lets it render text/canvas at full sharpness.
      canvas.style.transform = "none";
      canvas.style.zoom = String(g.z);
      // left/top of a zoomed element are themselves multiplied by its zoom
      canvas.style.left = -(g.c.l + pan.x) + "px";
      canvas.style.top = -(g.c.t + pan.y) + "px";
    } else {
      canvas.style.zoom = "";
      canvas.style.left = "0px"; canvas.style.top = "0px";
      // Whole-pixel offsets, and no scale() at all when it is 1:1, keep rendering crisp.
      canvas.style.transform = Math.abs(g.z - 1) < 0.001
        ? `translate(${Math.round(tx)}px, ${Math.round(ty)}px)`
        : `translate(${Math.round(tx)}px, ${Math.round(ty)}px) scale(${g.z})`;
    }

    const browser = state.zoomMode === "browser";
    zoomLabel.textContent = Math.round((browser ? state.zoom : g.z) * 100) + "%";
    zoomLabel.title = browser ? "Browser zoom – click to reset to 100%"
      : state.fit === "none" ? "Reset to 100%" : `Scaling: ${fitName(state.fit)} – click for 100%`;
    document.body.classList.toggle("zoom-browser", browser);

    document.body.classList.toggle("mode-pan", mode === "pan");
    document.body.classList.toggle("mode-crop", mode === "crop");
    document.body.classList.toggle("toolbar-pinned", mode !== "none" || !posPanel.hidden || !shotPanel.hidden || document.body.classList.contains("menu-open"));
    document.documentElement.style.setProperty("--bg", state.bg || "#ffffff");
    layoutToolbar();
    shield.hidden = !(mode !== "none" || state.locked || hoverGuard);
    $("btnPan").classList.toggle("active", mode === "pan");
    $("btnCrop").classList.toggle("active", mode === "crop");
    empty.hidden = !!state.url;
  }

  // Native CSS zoom on Chromium-based hosts (PowerPoint for Windows = WebView2, Edge, Chrome).
  // Mac PowerPoint (Safari engine) and other browsers fall back to transform: scale().
  const IS_CHROMIUM = /(Chrome|Chromium|Edg|EdgA|EdgiOS)\//.test(navigator.userAgent) &&
    !!(window.CSS && CSS.supports && CSS.supports("zoom", "0.5"));
  function useNativeZoom() {
    if (state.renderer === "transform") return false;
    if (state.renderer === "zoom") return true;
    return IS_CHROMIUM;
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

  let loadWaiters = [];
  function waitForLoad(timeoutMs) {
    return new Promise((res) => {
      const t = setTimeout(() => res(false), timeoutMs);
      loadWaiters.push(() => { clearTimeout(t); res(true); });
    });
  }

  function loaded() {
    const w = loadWaiters; loadWaiters = []; w.forEach((f) => f());
    clearTimeout(slowTimer);
    loading.hidden = true;
    lastLoaded = Date.now();
    const t = new Date(lastLoaded);
    setStatus("Updated " + t.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }));
  }

  function setStatus(s) { statusEl.textContent = s || ""; }

  function scheduleRefresh() {
    clearInterval(refreshTimer);
    if (state.refreshMin > 0) refreshTimer = setInterval(() => (state.snapMode ? snapCycle(false) : loadDashboard()), state.refreshMin * 60000);
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
    if (mode === "crop") enlarge("crop");
    else maybeRestore("crop");
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
      maybeRestore("crop");
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

  // ---------- Responsive toolbar ----------
  // The add-in can only draw inside its own box on the slide, so when the box is small the
  // toolbar shrinks: full labels -> icons only -> a single ☰ button with a drop-down menu.
  let tbKey = "";
  function layoutToolbar() {
    const sw = stage.clientWidth, sh = stage.clientHeight;
    const key = sw + "x" + sh + state.zoomMode;
    if (key === tbKey) return;
    tbKey = key;
    const b = document.body;
    b.classList.remove("tb-full", "tb-compact", "tb-mini");
    if (toolbar.offsetParent === null && getComputedStyle(toolbar).display === "none") {
      b.classList.add(sw >= 600 ? "tb-full" : sw >= 340 ? "tb-compact" : "tb-mini");
      tbKey = ""; // measure properly next time it is visible
      return;
    }
    const fits = () => toolbar.offsetHeight <= 40 && toolbar.offsetHeight < sh * 0.45 && toolbar.scrollWidth <= sw;
    b.classList.add("tb-full");
    if (fits()) return;
    b.classList.replace("tb-full", "tb-compact");
    if (fits()) return;
    b.classList.replace("tb-compact", "tb-mini");
  }

  function closeMenus() {
    if (!posPanel.hidden) maybeRestore("position");
    posPanel.hidden = true;
    if (!shotPanel.hidden) maybeRestore("shot");
    shotPanel.hidden = true;
    $("btnShot").classList.remove("active");
    document.body.classList.remove("menu-open");
    $("btnPos").classList.remove("active");
  }

  // ---------- Snap position & background ----------
  function openPosPanel() {
    const wasOpen = !posPanel.hidden;
    closeMenus();
    if (wasOpen) { render(); return; }
    posPanel.querySelectorAll(".snapgrid button").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.x === state.alignX && btn.dataset.y === state.alignY);
    });
    bgPick.value = state.bg || "#ffffff";
    const canSample = typeof window.EyeDropper === "function";
    $("btnSample").hidden = !canSample;
    $("sampleNote").textContent = canSample
      ? "Eyedropper: click Sample, then click the dashboard's edge to match its colour."
      : "Colour sampling isn't available in this version of PowerPoint – use the colour picker.";
    posPanel.hidden = false;
    $("btnPos").classList.add("active");
    enlarge("position");
    render();
  }

  function snap(el) {
    state.alignX = el.dataset.x; state.alignY = el.dataset.y;
    posPanel.querySelectorAll(".snapgrid button").forEach((b) => b.classList.toggle("active", b === el));
    render(); saveState();
  }

  function toHex(c) {
    if (/^#[0-9a-f]{6}$/i.test(c)) return c.toLowerCase();
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(c || "");
    return m ? "#" + [m[1], m[2], m[3]].map((v) => (+v).toString(16).padStart(2, "0")).join("") : "";
  }

  function setBg(hex) {
    state.bg = toHex(hex) || "";
    bgPick.value = state.bg || "#ffffff";
    render(); saveState();
  }

  function sampleBg() {
    if (typeof window.EyeDropper !== "function") return;
    closeMenus(); render();
    // EyeDropper samples the screen, so it can read the (cross-origin) dashboard's pixels
    // when the user clicks – scripts can't read them directly.
    new window.EyeDropper().open()
      .then((r) => setBg(r.sRGBHex))
      .catch(() => { /* cancelled */ });
  }

  bgPick.addEventListener("input", () => setBg(bgPick.value));

  // ---------- Temporary enlarge while editing (PowerPoint shape API) ----------
  // When the box is small, opening Settings / Position / Crop (or ☰) makes the add-in's own
  // shape on the slide bigger, and puts it back afterwards. Needs PowerPointApi 1.5
  // (current Microsoft 365 desktop & web); older versions just use the compact toolbar.
  const SMALL_W = 680, SMALL_H = 380;      // CSS px: below this, auto-enlarge
  const TARGET_W = 760, MIN_H = 330;       // CSS px to aim for when enlarged
  let enlargeTrigger = null;
  let resizeBusy = false;

  function canResize() {
    try {
      return inOffice && !!window.PowerPoint && Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
    } catch (e) { return false; }
  }
  const isSmall = () => stage.clientWidth < SMALL_W || stage.clientHeight < SMALL_H;

  const SHAPE_PROPS = "items/id,items/type,items/left,items/top,items/width,items/height";

  /** Find this add-in's own shape on the current slide.
   *  Known ids (learned earlier) are used first. strict = only accept the known shape, never guess
   *  (used by automatic, timer-driven actions so we never move another box by mistake). */
  async function locateSelf(ctx, strict) {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    const sel = ctx.presentation.getSelectedShapes();
    sel.load(SHAPE_PROPS);
    await ctx.sync();
    const slide = slides.items[0];
    if (!slide) return null;
    const isApp = (sh) => String(sh.type).toLowerCase() === "contentapp";
    const all = slide.shapes;
    all.load(SHAPE_PROPS);
    await ctx.sync();
    const apps = all.items.filter(isApp);
    if (state.shapeId && state.slideId === slide.id) {
      const me = apps.find((x) => x.id === state.shapeId);
      if (me) return { slide, shape: me };
    }
    if (strict) return null;
    let shape = null;
    const selApps = sel.items.filter(isApp);
    if (selApps.length === 1) shape = apps.find((x) => x.id === selApps[0].id) || selApps[0];
    if (!shape) shape = apps.length === 1 ? apps[0] : null;
    if (!shape && apps.length > 1) {
      const aspect = stage.clientWidth / Math.max(1, stage.clientHeight);
      const ranked = apps
        .map((x) => ({ x, d: Math.abs(x.width / x.height - aspect) / aspect }))
        .sort((a, b) => a.d - b.d);
      if (ranked[0].d < 0.02 && ranked[1].d > 0.05) shape = ranked[0].x;
    }
    if (shape) { state.shapeId = shape.id; state.slideId = slide.id; }
    return shape ? { slide, shape } : null;
  }

  async function slideSize(ctx) {
    try {
      if (Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
        const ps = ctx.presentation.pageSetup;
        ps.load("slideWidth,slideHeight");
        await ctx.sync();
        if (ps.slideWidth > 0 && ps.slideHeight > 0) return { w: ps.slideWidth, h: ps.slideHeight };
      }
    } catch (e) { /* older API */ }
    return { w: 960, h: 540 }; // standard 16:9 slide, in points
  }

  /** Returns true if the box was enlarged. */
  async function enlarge(trigger, force, strict) {
    if (state.restore || resizeBusy || view === "read" || !canResize()) return !!state.restore;
    if (!force && !isSmall()) return false;
    resizeBusy = true;
    let result = false;
    try {
      result = await PowerPoint.run(async (ctx) => {
        const found = await locateSelf(ctx, strict);
        if (!found) return "notfound";
        const { slide, shape } = found;
        const size = await slideSize(ctx);
        const pxPerPt = stage.clientWidth / shape.width;      // depends on PowerPoint's view zoom
        const aspect = shape.width / shape.height;
        let w = Math.max(shape.width, TARGET_W / pxPerPt);
        let h = w / aspect;
        if (h * pxPerPt < MIN_H) h = MIN_H / pxPerPt;          // very flat boxes get some height
        const maxW = size.w * 0.96, maxH = size.h * 0.96;
        if (w > maxW) { h = h * maxW / w; w = maxW; }
        if (h > maxH) { w = w * maxH / h; h = maxH; }
        w = Math.max(w, shape.width); h = Math.max(h, shape.height);
        if (w < shape.width * 1.05 && h < shape.height * 1.05) return "nochange";
        const cx = shape.left + shape.width / 2, cy = shape.top + shape.height / 2;
        state.shapeId = shape.id;
        state.restore = { slideId: slide.id, shapeId: shape.id,
                          left: shape.left, top: shape.top, width: shape.width, height: shape.height };
        shape.left = clamp(cx - w / 2, 0, Math.max(0, size.w - w));
        shape.top = clamp(cy - h / 2, 0, Math.max(0, size.h - h));
        shape.width = w;
        shape.height = h;
        await ctx.sync();
        return true;
      });
    } catch (e) {
      console.warn("Enlarge failed", e);
      state.restore = null;
      result = "error";
    } finally {
      resizeBusy = false;
    }
    if (result === true) {
      enlargeTrigger = trigger;
      saveState();
      flashHint("Enlarged for editing – click ✓ Done to shrink back");
    } else if (force) {
      flashHint(result === "nochange"
        ? "Already as big as the slide allows – zoom PowerPoint's view in for more room"
        : "Couldn't resize the box automatically – drag its handles to make it bigger");
    }
    applyBodyFlags(); render();
    return result === true;
  }

  /** Restore only if the box was enlarged by this trigger (Settings, Position, Crop). */
  function maybeRestore(trigger) {
    if (state.restore && enlargeTrigger === trigger) restoreSize();
  }

  async function restoreSize(onStartup) {
    const r = state.restore;
    if (!r || resizeBusy || !canResize()) return;
    resizeBusy = true;
    try {
      await PowerPoint.run(async (ctx) => {
        if (onStartup) {
          // A copied slide carries these settings too – only resize if we are that shape.
          const found = await locateSelf(ctx, true);
          if (!found || found.shape.id !== r.shapeId) return;
        }
        const shape = ctx.presentation.slides.getItem(r.slideId).shapes.getItem(r.shapeId);
        shape.left = r.left; shape.top = r.top; shape.width = r.width; shape.height = r.height;
        await ctx.sync();
      });
    } catch (e) {
      console.warn("Restore failed", e);
    } finally {
      resizeBusy = false;
      state.restore = null;
      enlargeTrigger = null;
      saveState(); applyBodyFlags(); render();
    }
  }

  let hintTimer = null;
  function flashHint(text) {
    if (mode !== "none") return; // keep pan/crop instructions visible
    showHint(text);
    clearTimeout(hintTimer);
    hintTimer = setTimeout(() => { if (mode === "none") showHint(""); }, 4000);
  }

  // ---------- Snapshot mode – a sharp picture on the slide for PDFs ----------
  // PowerPoint exports add-ins to PDF as a low-resolution snapshot. In snapshot mode a real
  // picture sits on the slide instead, and this box waits just off the slide's right edge as a
  // small control bar ("chip"); off-slide objects never appear in PDFs or the slideshow.
  // Each update (⟳, auto-refresh, turning the mode on):
  //   bring the box back over the picture's spot → enlarge (sharper) → reload → wait for the
  //   page, then the user's delay → block hover (no tooltips) → screen-capture → crop →
  //   replace the picture → park the box again.
  // Screen capture needs one click per box per PowerPoint session (browser rule); after that
  // updates run on their own while this slide is the one being shown.
  const SNAP_NAME = "Quickbase snapshot";
  const CHIP_W = 230, CHIP_H = 30;         // points
  const HOVER_GUARD_MS = 1500;
  let snapBusy = false;
  let capStream = null, capVideo = null;
  let hoverGuard = false;

  function canSnapshot() {
    try {
      return canResize() && Office.context.requirements.isSetSupported("ImageCoercion", "1.1");
    } catch (e) { return false; }
  }
  const canCapture = () => !!(navigator.mediaDevices && navigator.mediaDevices.getDisplayMedia);
  const streamAlive = () => !!capStream && capStream.getVideoTracks().some((t) => t.readyState === "live");
  const parked = () => state.snapMode && !!state.snap && !snapBusy;
  const rectOf = (x) => ({ left: x.left, top: x.top, width: x.width, height: x.height });
  const setRect = (x, r) => { x.left = r.left; x.top = r.top; x.width = r.width; x.height = r.height; };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const frames = (n) => new Promise((r) => { const f = () => (--n <= 0 ? r() : requestAnimationFrame(f)); requestAnimationFrame(f); });

  /** Ask for screen capture once; the stream is kept open and reused for every update. */
  async function openStream() {
    capStream = await navigator.mediaDevices.getDisplayMedia({
      video: { cursor: "never", frameRate: 5, width: { ideal: 7680 }, height: { ideal: 4320 } },
      audio: false, selfBrowserSurface: "include", surfaceSwitching: "exclude"
    });
    capStream.getVideoTracks().forEach((t) => t.addEventListener("ended", () => {
      capStream = null; capVideo = null;
      chipNote("Screen sharing stopped – click ⟳ to allow it again");
      applyBodyFlags(); render();
    }));
    capVideo = document.createElement("video");
    capVideo.muted = true; capVideo.srcObject = capStream;
    await capVideo.play();
    applyBodyFlags();
  }

  function grabFrame() {
    const c = document.createElement("canvas");
    c.width = capVideo.videoWidth; c.height = capVideo.videoHeight;
    c.getContext("2d", { willReadFrequently: true }).drawImage(capVideo, 0, 0);
    return c;
  }

  /** Bounding box of the magenta marker in a captured frame (physical pixels). */
  function findMarker(canvas) {
    const w = canvas.width, h = canvas.height;
    const d = canvas.getContext("2d").getImageData(0, 0, w, h).data;
    let x0 = w, y0 = h, x1 = -1, y1 = -1, n = 0;
    const step = 2;
    for (let y = 0; y < h; y += step) {
      for (let x = 0; x < w; x += step) {
        const i = (y * w + x) * 4;
        if (d[i] > 200 && d[i + 1] < 70 && d[i + 2] > 200) {
          n++;
          if (x < x0) x0 = x; if (x > x1) x1 = x;
          if (y < y0) y0 = y; if (y > y1) y1 = y;
        }
      }
    }
    if (n < 100) return null;
    const box = { x: x0, y: y0, w: x1 - x0 + step, h: y1 - y0 + step };
    return (n * step * step) / (box.w * box.h) > 0.7 ? box : null;
  }

  /** Capture just this box (UI hidden) and return it as base64 PNG. */
  async function captureBox() {
    document.body.classList.add("capturing", "marking");
    await frames(3); await sleep(350);
    const a = grabFrame();
    document.body.classList.remove("marking");
    await frames(3); await sleep(450);
    const b = grabFrame();
    document.body.classList.remove("capturing");
    const box = findMarker(a);
    if (!box) throw new Error("nobox");
    const want = stage.clientWidth / Math.max(1, stage.clientHeight);
    if (Math.abs(box.w / box.h - want) / want > 0.04) throw new Error("partial");
    const inset = 2;
    const out = document.createElement("canvas");
    out.width = box.w - inset * 2; out.height = box.h - inset * 2;
    out.getContext("2d").drawImage(b, box.x + inset, box.y + inset, out.width, out.height, 0, 0, out.width, out.height);
    return { base64: out.toDataURL("image/png").split(",")[1], w: out.width, h: out.height };
  }

  function insertImage(base64, rect) {
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(base64, {
        coercionType: Office.CoercionType.Image,
        imageLeft: rect.left, imageTop: rect.top, imageWidth: rect.width, imageHeight: rect.height
      }, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)));
    });
  }

  async function slideShapes(ctx, slide) {
    const all = slide.shapes;
    all.load(SHAPE_PROPS + ",items/name");
    await ctx.sync();
    return all.items;
  }

  function chipRect(size, home) {
    return { left: size.w + 12, top: clamp(home.top, 0, Math.max(0, size.h - CHIP_H)), width: CHIP_W, height: CHIP_H };
  }
  const picName = () => SNAP_NAME + " (" + (state.url ? new URL(state.url).hostname : "dashboard") + ")";

  /**
   * Take / update the snapshot. gesture = started by a click (may ask for screen capture and
   * may learn which box is ours); otherwise only runs if everything is already set up and this
   * box is on the slide being shown.
   */
  async function snapCycle(gesture, ownPicture) {
    if (snapBusy || view === "read") return;
    if (!canSnapshot()) { flashHint("Snapshots need a current Microsoft 365 PowerPoint"); return; }
    if (!ownPicture && !streamAlive()) {
      if (!gesture) { chipNote("Click ⟳ to update the snapshot"); return; }
      if (!canCapture()) { openShotPanel("Screen capture isn't available here – use “Use my own picture”."); return; }
      try { await openStream(); }            // first await: still counts as the click
      catch (e) {
        openShotPanel(e && e.name === "NotAllowedError"
          ? "Screen capture was cancelled or isn't allowed here. Try again, or use “Use my own picture”."
          : "Screen capture failed (" + (e && e.name || e) + "). Use “Use my own picture”.");
        return;
      }
    }
    if (state.restore && enlargeTrigger === "shot") enlargeTrigger = "snapshot"; // don't shrink mid-click
    closeMenus();
    snapBusy = true;
    document.body.classList.add("busy");
    applyBodyFlags(); render();
    let home = null, oldPicId = state.snap ? state.snap.pictureId : null;
    try {
      if (state.restore) await restoreSize();   // start from the box's real size and spot
      // 1. Bring the box back to the picture's spot and move the old picture aside.
      await PowerPoint.run(async (ctx) => {
        const found = await locateSelf(ctx, !gesture);
        if (!found) throw new Error(gesture ? "notfound" : "notshown");
        const { slide, shape } = found;
        const size = await slideSize(ctx);
        home = state.snap ? state.snap.home : rectOf(shape);
        if (oldPicId) {
          const pic = (await slideShapes(ctx, slide)).find((x) => x.id === oldPicId);
          if (pic) { home = rectOf(pic); pic.left = size.w + CHIP_W + 40; } else oldPicId = null;
        }
        setRect(shape, home);
        await ctx.sync();
      });
      if (ownPicture) {
        await PowerPoint.run(async (ctx) => {
          const { slide, shape } = await locateSelf(ctx, true);
          const shapes = await slideShapes(ctx, slide);
          const isPic = (x) => String(x.type).toLowerCase() === "image";
          const pic = shapes.filter((x) => isPic(x) && x.id !== oldPicId && !String(x.name || "").startsWith(SNAP_NAME)).pop();
          if (!pic) throw new Error("nopicture");
          setRect(pic, home); pic.name = picName();
          if (oldPicId) { const old = shapes.find((x) => x.id === oldPicId); if (old) old.delete(); }
          setRect(shape, chipRect(await slideSize(ctx), home));
          await ctx.sync();
          state.snap = { pictureId: pic.id, home, takenAt: Date.now() };
        });
        state.snapMode = true;
        flashHint("Your picture is now the snapshot");
        return;
      }
      // 2. Sharper capture: enlarge (not in Sharp mode, where the layout follows the box size).
      if (state.snapHiRes && !state.sharp) await enlarge("snapshot", true, true);
      // 3. Reload, wait for the page and then the user's delay, with hover blocked.
      hoverGuard = true; render();
      const waiter = waitForLoad(90000);
      loadDashboard();
      await waiter;
      const total = Math.max(state.snapDelay, HOVER_GUARD_MS / 1000);
      for (let left = Math.ceil(total); left > 0; left--) {
        showHint(`Snapshot in ${left}s…`);
        await sleep(Math.min(1000, total * 1000));
      }
      showHint("");
      // 4. Capture.
      const shot = await captureBox();
      if (state.restore) await restoreSize();
      // 5. Replace the picture and park the box.
      await PowerPoint.run(async (ctx) => {
        const { slide, shape } = await locateSelf(ctx, true);
        const before = await slideShapes(ctx, slide);
        const ids = new Set(before.map((x) => x.id));
        await insertImage(shot.base64, home);
        const after = await slideShapes(ctx, slide);
        const pic = after.filter((x) => !ids.has(x.id)).pop();
        if (!pic) throw new Error("insertfailed");
        pic.name = picName();
        if (oldPicId) { const old = after.find((x) => x.id === oldPicId); if (old) old.delete(); }
        const me = after.find((x) => x.id === shape.id) || shape;
        setRect(me, chipRect(await slideSize(ctx), home));
        await ctx.sync();
        state.snap = { pictureId: pic.id, home, takenAt: Date.now() };
      });
      state.snapMode = true;
      chipNote("");
      flashHint(`Snapshot updated (${shot.w}×${shot.h} px)`);
    } catch (e) {
      console.warn("Snapshot failed", e);
      document.body.classList.remove("capturing", "marking");
      if (state.restore) await restoreSize();
      // Put things back the way they were.
      if (home && oldPicId) {
        try {
          await PowerPoint.run(async (ctx) => {
            const found = await locateSelf(ctx, true);
            if (!found) return;
            const shapes = await slideShapes(ctx, found.slide);
            const old = shapes.find((x) => x.id === oldPicId);
            if (old) setRect(old, home);
            setRect(found.shape, chipRect(await slideSize(ctx), home));
            await ctx.sync();
          });
        } catch (e2) { /* ignore */ }
      }
      const msg = {
        notshown: "",   // automatic update while another slide is shown – just skip
        nobox: "Couldn't find the box in the capture – share the screen or PowerPoint window that shows it.",
        partial: "Part of the box was hidden or off screen – make the whole slide visible, then click ⟳.",
        nopicture: "No picture found – insert one first (Insert › Screenshot › Screen Clipping).",
        notfound: "Couldn't identify this box – click its border to select it, then try again."
      }[e && e.message];
      if (msg !== "") { if (state.snapMode) chipNote(msg || "Snapshot failed – " + (e && e.message || e)); else openShotPanel(msg || "Snapshot failed – " + (e && e.message || e)); }
    } finally {
      hoverGuard = false;
      snapBusy = false;
      document.body.classList.remove("busy");
      showHint("");
      saveState(); applyBodyFlags(); render();
    }
  }

  /** Live view: delete the picture and put the live box where the picture was. */
  async function goLive() {
    if (snapBusy) return;
    try {
      await PowerPoint.run(async (ctx) => {
        const found = await locateSelf(ctx, false);
        if (!found) throw new Error("notfound");
        const shapes = await slideShapes(ctx, found.slide);
        const pic = state.snap && shapes.find((x) => x.id === state.snap.pictureId);
        const home = pic ? rectOf(pic) : (state.snap ? state.snap.home : rectOf(found.shape));
        if (pic) pic.delete();
        setRect(found.shape, home);
        await ctx.sync();
      });
    } catch (e) {
      console.warn("Live view failed", e);
      flashHint("Couldn't move the box back automatically – drag it onto the slide");
    }
    state.snap = null; state.snapMode = false;
    chipNote("");
    saveState(); applyBodyFlags(); render();
    loadDashboard();
  }

  function chipNote(text) {
    $("chipNote").textContent = text || "";
    $("chipNote").title = text || "";
    $("chip").classList.toggle("has-note", !!text);
  }
  function chipStatus() {
    if (!state.snap) return "Snapshot";
    const t = new Date(state.snap.takenAt);
    const sameDay = t.toDateString() === new Date().toDateString();
    return "Snapshot " + (sameDay ? t.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
                                  : t.toLocaleDateString([], { month: "short", day: "numeric" }));
  }

  function openShotPanel(note) {
    const wasOpen = !shotPanel.hidden && !note;
    closeMenus();
    if (wasOpen) { render(); return; }
    $("shotDelay").value = state.snapDelay;
    $("shotHiRes").checked = state.snapHiRes;
    $("shotNote").textContent = note || "";
    $("shotNote").classList.toggle("warn", !!note);
    shotPanel.hidden = false;
    $("btnShot").classList.add("active");
    render();
    enlarge("shot");
  }
  $("shotDelay").addEventListener("change", () => { state.snapDelay = clamp(num($("shotDelay").value, 8), 0, 300); saveState(); });
  $("shotHiRes").addEventListener("change", () => { state.snapHiRes = $("shotHiRes").checked; saveState(); });

  // ---------- Settings form ----------
  function openSettings() {
    closeMenus();
    mode = "none"; showHint(""); render();
    const f = form.elements;
    f.url.value = state.url;
    qbTabSeg = "";
    renderQbOptions();
    f.vw.value = state.vw; f.vh.value = state.vh;
    f.fit.value = state.fit;
    f.zoomMode.value = state.zoomMode;
    f.zoom.value = Math.round(state.zoom * 100);
    f.cl.value = state.crop.l; f.ct.value = state.crop.t;
    f.cr.value = state.crop.r; f.cb.value = state.crop.b;
    f.refreshMin.value = state.refreshMin;
    f.snapDelay.value = state.snapDelay;
    f.snapHiRes.checked = state.snapHiRes;
    f.reloadOnShow.checked = state.reloadOnShow;
    f.toolbarInShow.checked = state.toolbarInShow;
    f.locked.checked = state.locked;
    f.sharp.checked = state.sharp;
    f.renderer.value = state.renderer;
    settingsPanel.hidden = false;
    setTimeout(() => f.url.focus(), 0);
    enlarge("settings");
  }
  function closeSettings() { settingsPanel.hidden = true; maybeRestore("settings"); }

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const f = form.elements;
    const url = normalizeQbUrl(unwrapEmbed(f.url.value));
    f.url.value = url;
    if (!isHttpsUrl(url)) {
      f.url.setCustomValidity("Enter a full https:// URL");
      f.url.reportValidity();
      return;
    }
    f.url.setCustomValidity("");
    const urlChanged = url !== state.url;
    const zoomInput = num(f.zoom.value, 100) / 100;
    const fitWas = state.fit + state.zoomMode;
    state = normalize(Object.assign({}, state, {
      url,
      vw: f.vw.value, vh: f.vh.value,
      fit: f.fit.value,
      zoomMode: f.zoomMode.value,
      zoom: zoomInput,
      crop: { l: f.cl.value, t: f.ct.value, r: f.cr.value, b: f.cb.value },
      pan: state.pan,
      refreshMin: f.refreshMin.value,
      snapDelay: f.snapDelay.value,
      snapHiRes: f.snapHiRes.checked,
      reloadOnShow: f.reloadOnShow.checked,
      toolbarInShow: f.toolbarInShow.checked,
      locked: f.locked.checked,
      sharp: f.sharp.checked,
      renderer: f.renderer.value
    }));
    if (urlChanged || fitWas !== state.fit + state.zoomMode) state.pan = { x: 0, y: 0 };
    applyBodyFlags();
    saveState();
    closeSettings();
    scheduleRefresh();
    if (urlChanged) { frame.removeAttribute("src"); }
    if (state.snapMode) setTimeout(() => snapCycle(true), 900);   // update the picture with the new settings
    else loadDashboard();
  });
  // ---------- Quickbase address options ----------
  // Mirrors the toggles in Quickbase's "Share dashboard" dialog. Each toggle reads its state
  // from the address and writes it back, so a pasted address and the toggles never disagree
  // and a parameter can never end up in the address twice.
  // Only parameter names confirmed from real Quickbase embed codes are listed here.
  // sense "present": toggle is ON when key=on is in the address.
  // sense "absent":  toggle is ON when the key is NOT in the address (Quickbase adds key=1 to turn it off).
  const QB_PARAMS = [
    { key: "embedMode", on: "1", sense: "present", label: "Embed mode",
      hint: "Quickbase's embedded layout. Added automatically by the Share dialog." },
    { key: "hidefilters", on: "1", sense: "absent", label: "Show filters",
      hint: "Dashboard filters bar." },
    { key: "hidealldashboardtabs", on: "1", sense: "absent", label: "Show all dashboard tabs",
      hint: "Tabs for the dashboard's other pages." },
    { key: "denydashboardediting", on: "1", sense: "absent", label: "Allow dashboard editing",
      hint: "Lets people with permission edit the dashboard from inside the box." },
    { key: "ifv", on: "1", sense: "present", label: "Hide Quickbase header (ifv=1)", extra: true,
      hint: "Hides Quickbase's own header bars. Part of the Share dialog's code; also works on reports and forms." }
  ];
  // Keys that look like embed switches but aren't listed above are shown as "Other" so they can
  // still be turned off (e.g. the remaining Share-dialog toggles until their names are confirmed).
  const QB_SWITCHY = /^(hide|deny|show|allow|embed|no)/i;
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  let qbTabSeg = "";   // remembered dashboard-tab segment, so "first tab" can be undone

  function splitUrl(u) {
    const h = u.indexOf("#");
    const hash = h >= 0 ? u.slice(h) : "";
    const base = h >= 0 ? u.slice(0, h) : u;
    const q = base.indexOf("?");
    return { path: q >= 0 ? base.slice(0, q) : base, parts: q >= 0 ? base.slice(q + 1).split("&").filter(Boolean) : [], hash };
  }
  const joinUrl = (x) => x.path + (x.parts.length ? "?" + x.parts.join("&") : "") + x.hash;
  const keyOf = (part) => { try { return decodeURIComponent(part.split("=")[0]); } catch (e) { return part.split("=")[0]; } };
  const valOf = (part) => { const i = part.indexOf("="); return i < 0 ? "" : part.slice(i + 1); };

  function getParam(u, key) {
    const p = splitUrl(u).parts.find((x) => keyOf(x).toLowerCase() === key.toLowerCase());
    return p === undefined ? null : valOf(p);
  }
  /** Remove every copy of key (any capitalisation), then add key=val once (val null = just remove). */
  function setParam(u, key, val) {
    const x = splitUrl(u);
    const k = key.toLowerCase();
    const at = x.parts.findIndex((p) => keyOf(p).toLowerCase() === k);   // keep its original position
    x.parts = x.parts.filter((p, i) => i === at || keyOf(p).toLowerCase() !== k);
    if (val === null || val === undefined) { if (at >= 0) x.parts.splice(at, 1); }
    else if (at >= 0) x.parts[at] = key + "=" + val;
    else x.parts.push(key + "=" + val);
    return joinUrl(x);
  }
  const truthy = (v) => v !== null && !/^(0|false|no|off)$/i.test(v);

  /** Is this a Quickbase address? Returns info or null. */
  function parseQb(u) {
    let url;
    try { url = new URL(u); } catch (e) { return null; }
    if (url.protocol !== "https:" || !/(^|\.)quickbase\.com$/i.test(url.hostname)) return null;
    const a = (getParam(u, "a") || "").toLowerCase();
    const navAction = (url.pathname.match(/\/action\/([^/?]+)/i) || [])[1] || "";
    const action = (a || navAction).toLowerCase();
    const kind = action === "showpage" || getParam(u, "pageIdV2") ? "dashboard"
      : action === "q" ? "report"
      : /^(dr|er|nwr)$/.test(action) ? "form" : "page";
    const segs = url.pathname.split("/").filter(Boolean);
    const tabSeg = segs.length >= 3 && segs[0].toLowerCase() === "db" && GUID.test(segs[2]) ? segs[2] : "";
    return { kind, tabSeg };
  }

  /** If an <iframe …> embed code was pasted, keep just its address. */
  function unwrapEmbed(v) {
    const m = /<iframe[^>]*\ssrc\s*=\s*["']([^"']+)["']/i.exec(v);
    if (!m) return v.trim();
    const t = document.createElement("textarea");
    t.innerHTML = m[1];               // decode &amp; etc.
    return t.value.trim();
  }

  /** Clean up the address: one copy of each known option, canonical spelling and value. */
  function normalizeQbUrl(u) {
    if (!parseQb(u)) return u;
    QB_PARAMS.forEach((p) => {
      const v = getParam(u, p.key);
      if (v === null) return;
      u = truthy(v) ? setParam(u, p.key, p.on) : setParam(u, p.key, null);
    });
    return u;
  }

  function renderQbOptions() {
    const input = form.elements.url;
    const u = input.value.trim();
    const info = parseQb(u);
    $("qbOpts").hidden = !info;
    if (!info) return;
    if (info.tabSeg) qbTabSeg = info.tabSeg;
    $("qbKind").textContent = "– " + { dashboard: "dashboard / page", report: "report", form: "form", page: "Quickbase page" }[info.kind];
    const box = $("qbToggles");
    box.innerHTML = "";
    const add = (container, id, label, checked, hint, onChange, extra) => {
      const l = document.createElement("label");
      l.className = "check" + (extra ? " qb-extra" : "");
      l.title = hint || "";
      const c = document.createElement("input");
      c.type = "checkbox"; c.checked = checked; c.id = id;
      c.addEventListener("change", () => onChange(c.checked));
      l.appendChild(c);
      l.appendChild(document.createTextNode(" " + label));
      if (extra) { const b = document.createElement("span"); b.className = "qb-badge"; b.textContent = "extra"; l.appendChild(b); }
      container.appendChild(l);
    };
    const apply = (nu) => { input.value = nu; renderQbOptions(); };
    QB_PARAMS.forEach((p) => {
      if (p.key !== "ifv" && info.kind !== "dashboard") return;   // dashboard-only switches
      const present = truthy(getParam(u, p.key));
      const checked = p.sense === "present" ? present : !present;
      add(box, "qb_" + p.key, p.label, checked, p.hint, (on) => {
        const wantPresent = p.sense === "present" ? on : !on;
        apply(setParam(input.value.trim(), p.key, wantPresent ? p.on : null));
      }, p.extra);
    });
    if (info.kind === "dashboard" && (info.tabSeg || qbTabSeg)) {
      add(box, "qb_firsttab", "Always open the first tab", !info.tabSeg,
        "Share links point at one dashboard tab. Removing that part of the address always opens the first tab.",
        (on) => {
          const url = new URL(input.value.trim());
          const segs = url.pathname.split("/");
          if (on) { const i = segs.findIndex((s) => GUID.test(s)); if (i > 0) segs.splice(i, 1); }
          else if (qbTabSeg && segs.length >= 3) segs.splice(3, 0, qbTabSeg);
          const x = splitUrl(input.value.trim());
          x.path = url.origin + segs.join("/").replace(/\/{2,}/g, "/");
          apply(joinUrl(x));
        }, true);
    }
    // Other switch-like parameters already in the address
    const other = $("qbOther");
    other.innerHTML = "";
    const known = new Set(QB_PARAMS.map((p) => p.key.toLowerCase()));
    splitUrl(u).parts.forEach((part) => {
      const k = keyOf(part);
      if (known.has(k.toLowerCase()) || !QB_SWITCHY.test(k)) return;
      add(other, "qbo_" + k, k + "=" + valOf(part) + " (other)", true,
        "Found in the address. Untick to remove it.", (on) => { if (!on) apply(setParam(input.value.trim(), k, null)); });
    });
    $("qbHelp").textContent = info.kind === "dashboard"
      ? "These match Quickbase's Share dashboard dialog and always reflect the address above. “Show dashboard name”, “Allow full screen” and “Show link to dashboard” will be added once their address codes are confirmed. Anything else found in the address is listed as “other”."
      : "These always reflect the address above.";
  }

  form.elements.url.addEventListener("input", () => {
    const input = form.elements.url;
    input.setCustomValidity("");
    const raw = input.value;
    if (/<iframe/i.test(raw)) input.value = unwrapEmbed(raw);
    const norm = normalizeQbUrl(input.value.trim());
    if (norm !== input.value.trim()) input.value = norm;
    renderQbOptions();
  });



  function applyBodyFlags() {
    document.body.classList.toggle("toolbar-in-show", !!state.toolbarInShow);
    document.body.classList.toggle("enlarged", !!state.restore);
    document.body.classList.toggle("can-resize", canResize());
    document.body.classList.toggle("can-snapshot", canSnapshot());
    document.body.classList.toggle("parked", parked());
    document.body.classList.toggle("snap-mode", !!state.snapMode);
    document.body.classList.toggle("stream-on", streamAlive());
    $("chipStatus").textContent = chipStatus();
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
    reload: () => (state.snapMode ? snapCycle(true) : loadDashboard()),
    settings: openSettings,
    closeSettings,
    signin: signIn,
    openExternal,
    menu: () => {
      const open = !document.body.classList.contains("menu-open");
      closeMenus();
      if (open && canResize()) {
        // Prefer making the box big enough for the full toolbar; fall back to the drop-down.
        enlarge("menu", true).then((ok) => {
          if (!ok) { document.body.classList.add("menu-open"); render(); }
        });
        return;
      }
      document.body.classList.toggle("menu-open", open);
      render();
    },
    enlarge: () => enlarge("manual", true),
    done: () => restoreSize(),
    position: openPosPanel,
    snap,
    sampleBg,
    resetBg: () => setBg(""),
    shot: () => openShotPanel(),
    shotTake: () => snapCycle(true),
    shotOwn: () => snapCycle(true, true),
    liveView: () => goLive()
  };
  // Actions that keep the ☰ menu open (so you can click them repeatedly)
  const KEEP_MENU = new Set(["enlarge", "zoomIn", "zoomOut", "zoomReset", "menu", "position", "snap", "resetBg"]);

  document.addEventListener("click", (e) => {
    const el = e.target.closest("[data-action]");
    if (!el) {
      // click outside any control closes open pop-ups
      if (!e.target.closest("#posPanel, #shotPanel, #toolbar") && (!posPanel.hidden || !shotPanel.hidden || document.body.classList.contains("menu-open"))) {
        closeMenus(); render();
      }
      return;
    }
    const name = el.dataset.action;
    const fn = actions[name];
    if (fn) {
      e.preventDefault();
      if (document.body.classList.contains("menu-open") && !KEEP_MENU.has(name)) {
        document.body.classList.remove("menu-open");
      }
      fn(el);
      render();
    }
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
    if (k === "escape" && (!posPanel.hidden || !shotPanel.hidden || document.body.classList.contains("menu-open"))) { closeMenus(); render(); return; }
    if (k === "escape") { if (mode !== "none") { mode = "none"; cropRect.hidden = true; showHint(""); render(); maybeRestore("crop"); } return; }
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
    if (view === "read" && state.restore) restoreSize();
    if (view === "read" && prev !== "read" && state.reloadOnShow && lastLoaded && Date.now() - lastLoaded > 5000) {
      loadDashboard();
    }
  }

  // ---------- Boot ----------
  function start() {
    $("verLabel").textContent = window.QB_BUILD || "dev";
    loadState();
    applyBodyFlags();
    render();
    scheduleRefresh();
    if (state.url && !parked()) loadDashboard();   // parked: loads when the snapshot updates
    // Expose for debugging / tests
    // Left enlarged last time (e.g. PowerPoint closed mid-edit)? Put the box back.
    if (state.restore) setTimeout(() => restoreSize(true), 1500);
    window.__qb = { get state() { return state; }, get geo() { return geo; }, render, setZoom, setFit, setMode, snapCycle };
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
