/* ============================================================
   app.js — upload handling, toolbar state, navigation, export.
   ============================================================ */
(function () {
  const P = window.DiffParse;
  const E = window.DiffEngine;
  const R = window.DiffRender;
  const GD = window.GridDiff;
  const GR = window.GridRender;

  // ---- App state ----
  const state = {
    a: null,            // {name,text,kind,sheets?} or pending {pasted}
    b: null,
    view: "split",      // split | inline
    focus: false,
    ignoreWs: false,
    granularity: "word",// word | char
    tabular: false,     // both files are spreadsheets/csv
    gridMode: true,     // show tabular data as a cell grid
    rows: [],
    gridData: null,
    navTargets: [],
    currentHunk: -1,
    forced: new Set(),
    searchMarks: [],
    searchIdx: -1,
    search: ""
  };

  // pending slot inputs (before compare)
  const slot = { a: null, b: null };

  // ---- DOM refs ----
  const $ = sel => document.querySelector(sel);
  const uploadEl = $("#upload");
  const toolbarEl = $("#toolbar");
  const diffEl = $("#diffview");
  const loadingEl = $("#loading");
  const errEl = $("#uploadErr");

  // ============ Upload screen ============
  function setupSlot(side) {
    const slotEl = $(`.slot[data-side="${side}"]`);
    const dz = slotEl.querySelector(".dropzone");
    const input = slotEl.querySelector("input[type=file]");
    const pasteLink = slotEl.querySelector(".dz-paste");
    const pasteBox = slotEl.querySelector(".pastebox");
    const textarea = pasteBox.querySelector("textarea");
    const pasteCancel = pasteBox.querySelector(".pb-cancel");
    const pasteDone = pasteBox.querySelector(".pb-done");
    const chip = slotEl.querySelector(".filechip");
    const removeBtn = chip.querySelector(".fc-remove");

    dz.addEventListener("click", () => input.click());
    input.addEventListener("change", e => {
      if (e.target.files[0]) loadFileToSlot(side, e.target.files[0]);
      input.value = "";
    });

    dz.addEventListener("dragover", e => { e.preventDefault(); dz.classList.add("dragover"); });
    dz.addEventListener("dragleave", () => dz.classList.remove("dragover"));
    dz.addEventListener("drop", e => {
      e.preventDefault(); dz.classList.remove("dragover");
      if (e.dataTransfer.files[0]) loadFileToSlot(side, e.dataTransfer.files[0]);
    });

    pasteLink.addEventListener("click", e => {
      e.stopPropagation();
      slotEl.classList.add("pasting");
      textarea.focus();
    });
    pasteCancel.addEventListener("click", () => {
      slotEl.classList.remove("pasting");
      textarea.value = "";
    });
    pasteDone.addEventListener("click", () => {
      const text = textarea.value;
      if (!text.trim()) { slotEl.classList.remove("pasting"); return; }
      slot[side] = { name: side === "a" ? "Pasted (original)" : "Pasted (changed)", text: text.replace(/\r\n/g, "\n"), kind: "paste" };
      fillChip(side, slot[side], "Pasted text · " + text.split("\n").length + " lines");
      slotEl.classList.remove("pasting");
      slotEl.classList.add("filled");
      maybeAutoCompare();
    });

    removeBtn.addEventListener("click", () => {
      slot[side] = null;
      slotEl.classList.remove("filled", "pasting");
      clearErr();
    });
  }

  async function loadFileToSlot(side, file) {
    clearErr();
    const slotEl = $(`.slot[data-side="${side}"]`);
    fillChip(side, { name: file.name }, "Parsing…", true);
    slotEl.classList.add("filled");
    try {
      const parsed = await P.parseFile(file);
      slot[side] = parsed;
      const lines = parsed.text.split("\n").length;
      fillChip(side, parsed, `${P.kindLabel(parsed.kind)} · ${formatSize(file.size)} · ${lines} lines`);
      maybeAutoCompare();
    } catch (err) {
      slot[side] = null;
      slotEl.classList.remove("filled");
      showErr(err.message || "Could not read that file.");
    }
  }

  function fillChip(side, info, sub, loading) {
    const chip = $(`.slot[data-side="${side}"] .filechip`);
    chip.querySelector(".fc-name").textContent = info.name;
    chip.querySelector(".fc-sub").textContent = sub;
    const ic = chip.querySelector(".fc-icon");
    ic.innerHTML = loading
      ? `<div class="spinner" style="width:18px;height:18px;border-width:2px"></div>`
      : fileGlyph(info.kind);
  }

  function fileGlyph() {
    return `<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>`;
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1024 / 1024).toFixed(1) + " MB";
  }

  function showErr(msg) { errEl.textContent = msg; }
  function clearErr() { errEl.textContent = ""; }

  // Text and pasted content belong to the same "text" group.
  // CSV and spreadsheets are both tabular. Docs and PDFs are their own groups.
  function kindGroup(kind) {
    if (kind === "text" || kind === "paste") return "text";
    if (kind === "csv"  || kind === "sheet") return "spreadsheet";
    return kind; // "doc" | "pdf"
  }

  function kindsCompatible(kA, kB) {
    return kindGroup(kA) === kindGroup(kB);
  }

  function maybeAutoCompare() {
    if (!slot.a || !slot.b) { $("#compareBtn").disabled = true; return; }
    if (!kindsCompatible(slot.a.kind, slot.b.kind)) {
      showErr(
        `Can't compare a ${P.kindLabel(slot.a.kind)} with a ${P.kindLabel(slot.b.kind)}. ` +
        `Both files must be the same type.`
      );
      $("#compareBtn").disabled = true;
      return;
    }
    clearErr();
    $("#compareBtn").disabled = false;
    runCompare();
  }

  // ============ Run comparison ============
  function runCompare() {
    if (!slot.a || !slot.b) return;
    if (!kindsCompatible(slot.a.kind, slot.b.kind)) {
      showErr(
        `Can't compare a ${P.kindLabel(slot.a.kind)} with a ${P.kindLabel(slot.b.kind)}. ` +
        `Both files must be the same type.`
      );
      return;
    }
    state.a = slot.a; state.b = slot.b;
    state.forced = new Set();
    state.currentHunk = -1;
    state.tabular = !!(state.a.sheets && state.b.sheets);
    if (state.tabular) state.gridMode = true;
    rebuildDiff();
    uploadEl.classList.add("hidden");
    toolbarEl.classList.remove("hidden");
    diffEl.classList.remove("hidden");
    $("#resetBtn").classList.remove("hidden");
  }

  function isGrid() { return state.tabular && state.gridMode; }

  function rebuildDiff() {
    state.forced = new Set();
    if (isGrid()) {
      state.gridData = GD.buildGridDiff(state.a.sheets, state.b.sheets, { ignoreWs: state.ignoreWs });
    } else {
      state.rows = E.buildRows(state.a.text, state.b.text, { ignoreWs: state.ignoreWs });
    }
    updateToolbarForMode();
    renderAll();
    updateStats();
  }

  function renderAll() {
    if (isGrid()) {
      GR.render(diffEl, state);
    } else {
      R.render(diffEl, {
        rows: state.rows,
        view: state.view,
        focus: state.focus,
        granularity: state.granularity,
        forced: state.forced,
        nameA: state.a.name, kindA: P.kindLabel(state.a.kind),
        nameB: state.b.name, kindB: P.kindLabel(state.b.kind)
      });
    }
    wireFolds();
    scanNav();
    updateNav();
    if (state.search) applySearchNow(state.search, true);
  }

  // Show/hide controls that only make sense in a given mode.
  function updateToolbarForMode() {
    const grid = isGrid();
    $("#gridSeg").classList.toggle("hidden", !state.tabular);
    $("#highlightGroup").classList.toggle("hidden", grid);
    $("#statMod").classList.toggle("hidden", !grid);
    if (state.tabular) {
      document.querySelectorAll("[data-tab]").forEach(b =>
        b.classList.toggle("active", b.dataset.tab === (state.gridMode ? "grid" : "text")));
    }
  }

  function wireFolds() {
    diffEl.querySelectorAll(".fold").forEach(f => {
      f.addEventListener("click", () => {
        const from = +f.dataset.from, to = +f.dataset.to;
        for (let i = from; i <= to; i++) state.forced.add(i);
        renderAll();
      });
    });
    diffEl.querySelectorAll(".gfold").forEach(f => {
      f.addEventListener("click", () => {
        const si = f.dataset.si, from = +f.dataset.from, to = +f.dataset.to;
        for (let i = from; i <= to; i++) state.forced.add(`${si}:${i}`);
        renderAll();
      });
    });
  }

  function updateStats() {
    if (isGrid()) {
      const s = state.gridData.stats;
      $("#statAdd").textContent = "+" + s.addCells;
      $("#statDel").textContent = "−" + s.delCells;
      $("#statMod").textContent = "~" + s.modCells;
    } else {
      const st = E.stats(state.rows);
      $("#statAdd").textContent = "+" + st.add;
      $("#statDel").textContent = "−" + st.del;
    }
  }

  // ============ Change navigation ============
  // Build a flat list of DOM elements (first row of each change run).
  function scanNav() {
    state.navTargets = [];
    let scope, sel;
    if (isGrid()) {
      scope = diffEl.querySelector(".grid-primary");
      sel = ".gtr[data-type], .gfold";
    } else {
      scope = diffEl.querySelector(".diff-grid");
      sel = ".row[data-type], .fold";
    }
    if (!scope) return;
    let prevChanged = false;
    scope.querySelectorAll(sel).forEach(el => {
      if (el.classList.contains("fold") || el.classList.contains("gfold")) { prevChanged = false; return; }
      const ch = el.dataset.type && el.dataset.type !== "eq";
      if (ch && !prevChanged) state.navTargets.push(el);
      prevChanged = !!ch;
    });
    if (state.currentHunk >= state.navTargets.length) state.currentHunk = -1;
  }

  function updateNav() {
    const n = state.navTargets.length;
    const cur = state.currentHunk >= 0 ? state.currentHunk + 1 : "–";
    $("#navCount").textContent = n ? `${cur} / ${n}` : "0 changes";
    $("#navPrev").disabled = n === 0;
    $("#navNext").disabled = n === 0;
  }

  function gotoHunk(dir) {
    if (!state.navTargets.length) return;
    let i = state.currentHunk + dir;
    if (i < 0) i = state.navTargets.length - 1;
    if (i >= state.navTargets.length) i = 0;
    state.currentHunk = i;
    diffEl.querySelectorAll(".change-target").forEach(r => r.classList.remove("change-target"));
    const target = state.navTargets[i];
    if (target) {
      target.classList.add("change-target");
      scrollToEl(target);
    }
    updateNav();
  }

  function scrollToEl(el) {
    const cRect = diffEl.getBoundingClientRect();
    const eRect = el.getBoundingClientRect();
    const top = diffEl.scrollTop + (eRect.top - cRect.top) - 90;
    diffEl.scrollTo({ top: Math.max(0, top), behavior: "smooth" });
  }

  // ============ Search ============
  function applySearchNow(query, keepIdx) {
    state.search = query;
    state.searchMarks = R.applySearch(diffEl, query);
    const count = state.searchMarks.length;
    $("#searchCount").textContent = query ? (count ? `${Math.min(state.searchIdx + 1, count) || 1}/${count}` : "0/0") : "";
    if (!keepIdx) state.searchIdx = count ? 0 : -1;
    if (count && state.searchIdx >= 0) focusMatch(state.searchIdx);
  }

  function focusMatch(i) {
    state.searchMarks.forEach(m => m.classList.remove("current"));
    if (i < 0 || i >= state.searchMarks.length) return;
    const m = state.searchMarks[i];
    m.classList.add("current");
    scrollToEl(m);
    $("#searchCount").textContent = `${i + 1}/${state.searchMarks.length}`;
  }

  function cycleSearch(dir) {
    if (!state.searchMarks.length) return;
    state.searchIdx = (state.searchIdx + dir + state.searchMarks.length) % state.searchMarks.length;
    focusMatch(state.searchIdx);
  }

  // ============ Export ============
  function copyDiff() {
    const text = E.unified(state.a.name, state.b.name, state.a.text, state.b.text);
    navigator.clipboard.writeText(text).then(
      () => toast("Unified diff copied to clipboard"),
      () => toast("Copy failed — try Export instead")
    );
  }

  function exportDiff() {
    const text = E.unified(state.a.name, state.b.name, state.a.text, state.b.text);
    const blob = new Blob([text], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const base = (state.a.name || "diff").replace(/\.[^.]+$/, "");
    link.href = url;
    link.download = `${base}.diff`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast("Diff exported");
  }

  let toastTimer;
  function toast(msg) {
    const t = $("#toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), 2200);
  }

  // ============ Reset ============
  function reset() {
    slot.a = null; slot.b = null;
    state.a = null; state.b = null;
    state.search = ""; state.searchIdx = -1; state.currentHunk = -1;
    state.view = "split"; state.focus = false; state.gridMode = true;
    state.forced = new Set();
    $("#searchInput").value = "";
    $("#searchCount").textContent = "";
    $("#focusToggle").classList.remove("on");
    document.querySelectorAll("[data-view]").forEach(b => b.classList.toggle("active", b.dataset.view === "split"));
    document.querySelectorAll("[data-tab]").forEach(b => b.classList.toggle("active", b.dataset.tab === "grid"));
    document.querySelectorAll(".slot").forEach(s => s.classList.remove("filled", "pasting"));
    document.querySelectorAll(".pastebox textarea").forEach(t => t.value = "");
    uploadEl.classList.remove("hidden");
    toolbarEl.classList.add("hidden");
    diffEl.classList.add("hidden");
    diffEl.innerHTML = "";
    $("#resetBtn").classList.add("hidden");
    $("#compareBtn").disabled = true;
    clearErr();
  }

  // ============ Toolbar wiring ============
  function setupToolbar() {
    // view segmented
    document.querySelectorAll("[data-view]").forEach(b => {
      b.addEventListener("click", () => {
        if (state.view === b.dataset.view) return;
        state.view = b.dataset.view;
        document.querySelectorAll("[data-view]").forEach(x => x.classList.toggle("active", x === b));
        renderAll();
      });
    });
    // granularity segmented
    document.querySelectorAll("[data-gran]").forEach(b => {
      b.addEventListener("click", () => {
        if (state.granularity === b.dataset.gran) return;
        state.granularity = b.dataset.gran;
        document.querySelectorAll("[data-gran]").forEach(x => x.classList.toggle("active", x === b));
        renderAll();
      });
    });
    // grid / text segmented (spreadsheets only)
    document.querySelectorAll("[data-tab]").forEach(b => {
      b.addEventListener("click", () => {
        const wantGrid = b.dataset.tab === "grid";
        if (state.gridMode === wantGrid) return;
        state.gridMode = wantGrid;
        state.currentHunk = -1;
        rebuildDiff();
      });
    });
    // focus toggle
    $("#focusToggle").addEventListener("click", () => {
      state.focus = !state.focus;
      $("#focusToggle").classList.toggle("on", state.focus);
      if (!state.focus) state.forced = new Set();
      renderAll();
    });
    // ignore whitespace toggle
    $("#wsToggle").addEventListener("click", () => {
      state.ignoreWs = !state.ignoreWs;
      $("#wsToggle").classList.toggle("on", state.ignoreWs);
      state.currentHunk = -1;
      rebuildDiff();
    });
    // search
    const si = $("#searchInput");
    let searchTimer;
    si.addEventListener("input", () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => applySearchNow(si.value, false), 160);
    });
    si.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); cycleSearch(e.shiftKey ? -1 : 1); }
      if (e.key === "Escape") { si.value = ""; applySearchNow("", false); }
    });
    // nav
    $("#navPrev").addEventListener("click", () => gotoHunk(-1));
    $("#navNext").addEventListener("click", () => gotoHunk(1));
    // export
    $("#copyBtn").addEventListener("click", copyDiff);
    $("#exportBtn").addEventListener("click", exportDiff);
    // reset / swap
    $("#resetBtn").addEventListener("click", reset);
    $("#swapBtn").addEventListener("click", () => {
      const tmp = state.a; state.a = state.b; state.b = tmp;
      slot.a = state.a; slot.b = state.b;
      state.currentHunk = -1;
      rebuildDiff();
      toast("Swapped sides");
    });

    // keyboard shortcuts
    document.addEventListener("keydown", e => {
      if (diffEl.classList.contains("hidden")) return;
      if (document.activeElement === si) return;
      if (e.altKey && e.key === "n") { e.preventDefault(); gotoHunk(1); }
      if (e.altKey && e.key === "p") { e.preventDefault(); gotoHunk(-1); }
      if ((e.metaKey || e.ctrlKey) && e.key === "f") { e.preventDefault(); si.focus(); }
    });
  }

  // ============ Init ============
  function init() {
    setupSlot("a");
    setupSlot("b");
    setupToolbar();
    $("#compareBtn").addEventListener("click", runCompare);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else { init(); }
})();
