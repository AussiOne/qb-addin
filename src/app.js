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
    locked: false
  };

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const stage = $("stage"), clip = $("clip"), canvas = $("canvas");
  const shield = $("shield"), cropRect = $("cropRect"), hint = $("hint");
  const loading = $("loading"), loadingText = $("loadingText"), empty = $("empty");
  const zoomLabel = $("zoomLabel"), statusEl = $("status");
  const settingsPanel = $("settings"), form = $("settingsForm");
  const posPanel = $("posPanel"), bgPick = $("bgPick"), toolbar = $("toolbar");
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
      const pageW = state.vw / state.zoom;
      const pageH = pageW * sh / sw;
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
    canvas.style.transform = `translate(${tx}px, ${ty}px) scale(${g.z})`;

    const browser = state.zoomMode === "browser";
    zoomLabel.textContent = Math.round((browser ? state.zoom : g.z) * 100) + "%";
    zoomLabel.title = browser ? "Browser zoom – click to reset to 100%"
      : state.fit === "none" ? "Reset to 100%" : `Scaling: ${fitName(state.fit)} – click for 100%`;
    document.body.classList.toggle("zoom-browser", browser);

    document.body.classList.toggle("mode-pan", mode === "pan");
    document.body.classList.toggle("mode-crop", mode === "crop");
    document.body.classList.toggle("toolbar-pinned", mode !== "none" || !posPanel.hidden || document.body.classList.contains("menu-open"));
    document.documentElement.style.setProperty("--bg", state.bg || "#ffffff");
    layoutToolbar();
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

  /** Find this add-in's own shape: the selected content add-in, else by remembered id / only one / matching aspect. */
  async function locateSelf(ctx) {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    const sel = ctx.presentation.getSelectedShapes();
    sel.load(SHAPE_PROPS);
    await ctx.sync();
    const slide = slides.items[0];
    if (!slide) return null;
    const isApp = (sh) => String(sh.type).toLowerCase() === "contentapp";
    let shape = null;
    const selApps = sel.items.filter(isApp);
    if (selApps.length === 1) shape = selApps[0];
    if (!shape) {
      const all = slide.shapes;
      all.load(SHAPE_PROPS);
      await ctx.sync();
      const apps = all.items.filter(isApp);
      shape = apps.find((x) => x.id === state.shapeId) || (apps.length === 1 ? apps[0] : null);
      if (!shape && apps.length > 1) {
        const aspect = stage.clientWidth / Math.max(1, stage.clientHeight);
        const ranked = apps
          .map((x) => ({ x, d: Math.abs(x.width / x.height - aspect) / aspect }))
          .sort((a, b) => a.d - b.d);
        if (ranked[0].d < 0.02 && ranked[1].d > 0.05) shape = ranked[0].x;
      }
    }
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
  async function enlarge(trigger, force) {
    if (state.restore || resizeBusy || view === "read" || !canResize()) return !!state.restore;
    if (!force && !isSmall()) return false;
    resizeBusy = true;
    let result = false;
    try {
      result = await PowerPoint.run(async (ctx) => {
        const found = await locateSelf(ctx);
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
          const found = await locateSelf(ctx);
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

  // ---------- Settings form ----------
  function openSettings() {
    closeMenus();
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
    enlarge("settings");
  }
  function closeSettings() { settingsPanel.hidden = true; maybeRestore("settings"); }

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
    state = normalize(Object.assign({}, state, {
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
    }));
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
    document.body.classList.toggle("enlarged", !!state.restore);
    document.body.classList.toggle("can-resize", canResize());
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
    resetBg: () => setBg("")
  };
  // Actions that keep the ☰ menu open (so you can click them repeatedly)
  const KEEP_MENU = new Set(["enlarge", "zoomIn", "zoomOut", "zoomReset", "menu", "position", "snap", "resetBg"]);

  document.addEventListener("click", (e) => {
    const el = e.target.closest("[data-action]");
    if (!el) {
      // click outside any control closes open pop-ups
      if (!e.target.closest("#posPanel, #toolbar") && (!posPanel.hidden || document.body.classList.contains("menu-open"))) {
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
    if (k === "escape" && (!posPanel.hidden || document.body.classList.contains("menu-open"))) { closeMenus(); render(); return; }
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
    loadState();
    applyBodyFlags();
    render();
    scheduleRefresh();
    if (state.url) loadDashboard();
    // Expose for debugging / tests
    // Left enlarged last time (e.g. PowerPoint closed mid-edit)? Put the box back.
    if (state.restore) setTimeout(() => restoreSize(true), 1500);
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
