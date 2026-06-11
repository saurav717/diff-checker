/* ============================================================
   app.js — upload handling, toolbar state, navigation, export.
   ============================================================ */
(function () {
  const P = window.DiffParse;
  const E = window.DiffEngine;
  const R = window.DiffRender;
  const GD = window.GridDiff;
  const GR = window.GridRender;
  const PV = window.PdfVisual;
  const DV = window.DocxVisual;
  const NV = window.NotebookVisual;

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
    sheetA: null,       // selected sheet name in file A (multi-sheet compare)
    sheetB: null,       // selected sheet name in file B
    pdfVisual: false,   // both files are PDFs (visual diff available)
    pdfMode: "visual",  // visual | text  (when pdfVisual)
    pdfData: null,      // rendered pages + highlight boxes
    ocrActive: false,   // text was recovered with OCR (forces word diff)
    docVisual: false,   // both files are Word docs (visual diff available)
    docMode: "visual",  // visual | text  (when docVisual)
    docGran: "word",    // word | sentence  (Word visual highlight granularity)
    linkMode: "changes",// changes | all  (hover-link mapping scope)
    docData: null,      // marked HTML + stats
    nbVisual: false,    // both files are Jupyter notebooks
    nbMode: "visual",   // visual | text  (when nbVisual)
    nbData: null,       // rendered cells + stats
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
      populateSlotSheet(side, null);
      slotEl.classList.remove("pasting");
      slotEl.classList.add("filled");
      maybeAutoCompare();
    });

    removeBtn.addEventListener("click", () => {
      slot[side] = null;
      slotEl.classList.remove("filled", "pasting");
      populateSlotSheet(side, null);
      maybeAutoCompare();
      clearErr();
    });
  }

  async function loadFileToSlot(side, file) {
    clearErr();
    const slotEl = $(`.slot[data-side="${side}"]`);
    populateSlotSheet(side, null);            // clear any prior picker while parsing
    fillChip(side, { name: file.name }, "Parsing…", true);
    slotEl.classList.add("filled");
    try {
      const parsed = await P.parseFile(file);
      slot[side] = parsed;
      const lines = parsed.text.split("\n").length;
      fillChip(side, parsed, `${P.kindLabel(parsed.kind)} · ${formatSize(file.size)} · ${lines} lines`);
      populateSlotSheet(side, parsed);        // offer sheet choice for multi-sheet workbooks
      maybeAutoCompare();
    } catch (err) {
      slot[side] = null;
      slotEl.classList.remove("filled");
      populateSlotSheet(side, null);
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

  // Validate the two slots and toggle the Compare button. The comparison itself
  // only runs when the user clicks Compare (see #compareBtn click handler) —
  // it is intentionally NOT started automatically on upload.
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
    if (state.tabular) {
      state.gridMode = true;
      const aNames = state.a.sheets.map(s => s.name);
      const bNames = state.b.sheets.map(s => s.name);
      // Honor the sheet picked on the upload screen; otherwise default to A's
      // first sheet, matched to B by name when possible.
      const pickA = slotSheetValue("a"), pickB = slotSheetValue("b");
      state.sheetA = (pickA && aNames.indexOf(pickA) !== -1) ? pickA : (aNames[0] || null);
      state.sheetB = (pickB && bNames.indexOf(pickB) !== -1) ? pickB
        : (bNames.indexOf(state.sheetA) !== -1 ? state.sheetA : (bNames[0] || null));
    }
    state.pdfVisual = !!(state.a.kind === "pdf" && state.b.kind === "pdf" && state.a.pdfPages && state.b.pdfPages);
    if (state.pdfVisual) state.pdfMode = "visual";
    state.pdfData = null;
    state.docVisual = !!(state.a.kind === "doc" && state.b.kind === "doc" && state.a.docHtml && state.b.docHtml);
    if (state.docVisual) state.docMode = "visual";
    state.docData = null;
    state.nbVisual = !!(state.a.kind === "notebook" && state.b.kind === "notebook" && state.a.nbCells && state.b.nbCells);
    if (state.nbVisual) state.nbMode = "visual";
    state.nbData = null;
    rebuildDiff();
    uploadEl.classList.add("hidden");
    toolbarEl.classList.remove("hidden");
    diffEl.classList.remove("hidden");
    $("#resetBtn").classList.remove("hidden");
    $("#swapBtn").classList.remove("hidden");
  }

  function isGrid() { return state.tabular && state.gridMode; }
  function multiSheet() {
    return state.tabular && (
      (state.a.sheets && state.a.sheets.length > 1) ||
      (state.b.sheets && state.b.sheets.length > 1));
  }
  // Fill the sheet pickers from each workbook, reflecting the current choice.
  function populateSheetSelectors() {
    const selA = $("#sheetA"), selB = $("#sheetB");
    if (!selA || !selB || !state.a.sheets || !state.b.sheets) return;
    const fill = (sel, sheets, current) => {
      sel.innerHTML = "";
      sheets.forEach(s => {
        const o = document.createElement("option");
        o.value = s.name; o.textContent = s.name;
        if (s.name === current) o.selected = true;
        sel.appendChild(o);
      });
    };
    fill(selA, state.a.sheets, state.sheetA);
    fill(selB, state.b.sheets, state.sheetB);
    // Selectors now reflect the active comparison, so nothing is staged.
    const applyBtn = $("#sheetApplyBtn");
    if (applyBtn) applyBtn.disabled = true;
  }

  // ---- Upload-screen per-file sheet pickers (choose sheets before comparing) ----
  // Populate (or hide) the sheet dropdown inside a slot's filechip. Only shown
  // when the parsed file is a workbook with more than one sheet.
  function populateSlotSheet(side, parsed) {
    const wrap = $(`.slot[data-side="${side}"] .fc-sheet`);
    const sel = $(`.slot[data-side="${side}"] .fc-sheet-sel`);
    if (!wrap || !sel) return;
    const sheets = parsed && parsed.sheets;
    if (!sheets || sheets.length < 2) { wrap.hidden = true; sel.innerHTML = ""; return; }
    sel.innerHTML = "";
    sheets.forEach((s, i) => {
      const o = document.createElement("option");
      const nrows = Array.isArray(s.grid) ? s.grid.length : 0;
      o.value = s.name;
      o.textContent = nrows ? `${s.name} · ${nrows} row${nrows === 1 ? "" : "s"}` : s.name;
      if (i === 0) o.selected = true;
      sel.appendChild(o);
    });
    wrap.hidden = false;
  }
  // The sheet chosen on the upload screen for a side, or null if no picker shown.
  function slotSheetValue(side) {
    const wrap = $(`.slot[data-side="${side}"] .fc-sheet`);
    const sel = $(`.slot[data-side="${side}"] .fc-sheet-sel`);
    if (!wrap || wrap.hidden || !sel || !sel.value) return null;
    return sel.value;
  }
  function isVisualPdf() { return state.pdfVisual && state.pdfMode === "visual"; }
  function isVisualDoc() { return state.docVisual && state.docMode === "visual"; }
  // Text source for the line-diff: recovered OCR text when active, else extracted text.
  function srcText(f) { return (state.ocrActive && f && f.ocrText) ? f.ocrText : (f ? f.text : ""); }
  function isVisualNb() { return state.nbVisual && state.nbMode === "visual"; }

  function showLoading(on, label) {
    if (label) loadingEl.querySelector("span").textContent = label;
    loadingEl.classList.toggle("show", !!on);
  }

  function rebuildDiff() {
    state.forced = new Set();

    if (isVisualPdf()) {
      if (!state.pdfData) {
        showLoading(true, "Comparing pages…");
        PV.build(state.a, state.b, { ocr: state.ocrActive }).then(d => {
          state.pdfData = d;
          showLoading(false);
          updateToolbarForMode();
          renderAll();
          updateStats();
        });
        return;
      }
      updateToolbarForMode();
      renderAll();
      updateStats();
      return;
    }

    if (isVisualDoc()) {
      if (!state.docData) {
        showLoading(true, "Rendering documents…");
        DV.build(state.a, state.b, state.docGran).then(d => {
          state.docData = d;
          showLoading(false);
          updateToolbarForMode();
          renderAll();
          updateStats();
        });
        return;
      }
      updateToolbarForMode();
      renderAll();
      updateStats();
      return;
    }

    if (isVisualNb()) {
      if (!state.nbData) state.nbData = NV.build(state.a, state.b, state.docGran);
      updateToolbarForMode();
      renderAll();
      updateStats();
      return;
    }

    if (isGrid()) {
      state.gridData = GD.buildGridDiff(state.a.sheets, state.b.sheets,
        { ignoreWs: state.ignoreWs, selA: state.sheetA, selB: state.sheetB });
    } else {
      state.rows = E.buildRows(srcText(state.a), srcText(state.b), { ignoreWs: state.ignoreWs });
    }
    updateToolbarForMode();
    renderAll();
    updateStats();
  }

  /* Recover text from PDFs whose text layer is unreliable, using OCR on
     the already-rendered page bitmaps, then re-run the precise word diff.
     Only the file(s) with a weak text layer are OCR'd — a file that already
     has good extracted text keeps it (faster, and more accurate). */
  async function runOcr() {
    if (!window.OCR || !OCR.available()) { toast("OCR engine isn’t available."); return; }
    const btn = document.getElementById("ocrBtn");
    if (btn) { btn.disabled = true; btn.textContent = "Reading…"; }

    // OCR BOTH files so recognition is consistent — comparing OCR text on one
    // side against natively-extracted text on the other produces noisy diffs
    // on common text (punctuation/bullets tokenise differently). Identical
    // content recognised by the same engine cancels out cleanly.
    const todo = [];
    [["A", state.a], ["B", state.b]].forEach(([label, f]) => {
      if (!f.ocrDone && f.pdfPages) todo.push([label, f]);
    });

    showLoading(true, "Starting OCR…");
    try {
      // Warm the engine (downloads the language model on first use) with feedback.
      await OCR.ensure((frac, status) => showLoading(true, `Preparing OCR — ${status}… ${Math.round(frac * 100)}%`));

      for (let t = 0; t < todo.length; t++) {
        const [label, f] = todo[t];
        const pages = f.pdfPages.length;

        // Re-render at high resolution so small text is legible to OCR.
        let images = f.pdfPages.map(p => ({ url: p.url, vw: p.vw, vh: p.vh }));
        if (f.ocrBuffer && DiffParse.renderPdfImages) {
          try {
            showLoading(true, `Rendering ${todo.length > 1 ? "file " + label + " " : ""}at high resolution for OCR…`);
            const hi = await DiffParse.renderPdfImages(f.ocrBuffer, 3.2);
            if (hi.length === pages) images = hi;
          } catch (_) { /* fall back to the on-screen bitmaps */ }
        }

        const res = await OCR.run(images, (frac, status) => {
          const pageNo = Math.min(pages, Math.floor(frac * pages) + 1);
          const where = pages > 1 ? `page ${pageNo}/${pages} · ` : "";
          const scope = todo.length > 1 ? `file ${label} · ` : "";
          showLoading(true, `Reading text with OCR — ${scope}${where}${status}… ${Math.round(frac * 100)}%`);
        });

        // Map recognised word boxes from the (hi-res) OCR image space back to
        // the on-screen page coordinate space the highlights are drawn in.
        f.pdfPages.forEach((pg, i) => {
          const src = images[i];
          const fx = src && src.vw ? pg.vw / src.vw : 1;
          const fy = src && src.vh ? pg.vh / src.vh : 1;
          const ws = (res.pages[i] && res.pages[i].words) || [];
          pg.words = ws.map(w => ({ str: w.str, x: w.x * fx, y: w.y * fy, w: w.w * fx, h: w.h * fy }));
        });
        f.ocrText = res.text;
        f.ocrDone = true;
      }

      state.ocrActive = true;
      state.currentHunk = -1;
      showLoading(true, "Comparing recovered text…");
      state.pdfData = await PV.build(state.a, state.b, { ocr: true });
      if (state.pdfMode === "text") {
        state.rows = E.buildRows(srcText(state.a), srcText(state.b), { ignoreWs: state.ignoreWs });
      }
      showLoading(false);
      updateToolbarForMode();
      renderAll();
      updateStats();
      toast("Text recovered with OCR");
    } catch (e) {
      console.error(e);
      showLoading(false);
      toast("OCR failed — please try again.");
      if (btn) { btn.disabled = false; btn.textContent = "Read text with OCR & compare"; }
    }
  }

  function scheduleMinimap() {
    if (!window.DiffMinimap) return;
    // Centre ruler for side-by-side; move to the right edge for inline view.
    const inlineView = state.view === "inline" && !isVisualPdf() && !isVisualDoc() && !isVisualNb();
    DiffMinimap.setPlacement(inlineView ? "right" : "center");
    requestAnimationFrame(() => DiffMinimap.refresh());
    setTimeout(() => DiffMinimap.refresh(), 450);   // re-measure after images/fonts settle
  }

  function renderAll() {
    if (isVisualPdf()) {
      PV.render(diffEl, state);
      scanNav();
      updateNav();
      scheduleMinimap();
      return;
    }
    if (isVisualDoc()) {
      DV.render(diffEl, state);
      scanNav();
      updateNav();
      scheduleMinimap();
      return;
    }
    if (isVisualNb()) {
      NV.render(diffEl, state);
      scanNav();
      updateNav();
      scheduleMinimap();
      return;
    }
    if (isGrid()) {
      GR.render(diffEl, state);
      syncGridScroll();
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
    scheduleMinimap();
    if (state.search) applySearchNow(state.search, true);
  }

  // Show/hide controls that only make sense in a given mode.
  function updateToolbarForMode() {
    const grid = isGrid();
    const vpdf = isVisualPdf();
    const vdoc = isVisualDoc();
    const vnb = isVisualNb();
    const visual = vpdf || vdoc || vnb;   // a rendered view (PDF / Word / Notebook)
    $("#gridSeg").classList.toggle("hidden", !state.tabular);
    $("#sheetSel").classList.toggle("hidden", !(grid && multiSheet()));
    if (state.tabular) populateSheetSelectors();
    $("#pdfSeg").classList.toggle("hidden", !state.pdfVisual);
    $("#docSeg").classList.toggle("hidden", !state.docVisual);
    $("#docGranGroup").classList.toggle("hidden", !(isVisualDoc() || isVisualNb()));
    $("#linkGroup").classList.toggle("hidden", isVisualPdf());
    $("#nbSeg").classList.toggle("hidden", !state.nbVisual);
    // In a visual view, line-oriented text controls don't apply — but
    // "Focus changes" does (it collapses to just the changed content).
    $("#viewGroup").classList.toggle("hidden", visual);
    $("#optGroup").classList.remove("hidden");
    $("#wsToggle").classList.toggle("hidden", visual);
    $("#highlightGroup").classList.toggle("hidden", grid || visual);
    $("#searchGroup").classList.toggle("hidden", visual);
    $("#statMod").classList.toggle("hidden", !grid);
    $("#changesBtnGroup").classList.toggle("hidden", !grid);
    if (!grid) closeChangesPanel();
    if (state.tabular) {
      document.querySelectorAll("[data-tab]").forEach(b =>
        b.classList.toggle("active", b.dataset.tab === (state.gridMode ? "grid" : "text")));
    }
    if (state.pdfVisual) {
      document.querySelectorAll("[data-pdftab]").forEach(b =>
        b.classList.toggle("active", b.dataset.pdftab === state.pdfMode));
    }
    if (state.docVisual) {
      document.querySelectorAll("[data-doctab]").forEach(b =>
        b.classList.toggle("active", b.dataset.doctab === state.docMode));
    }
    document.querySelectorAll("[data-docgran]").forEach(b =>
      b.classList.toggle("active", b.dataset.docgran === state.docGran));
    document.querySelectorAll("[data-link]").forEach(b =>
      b.classList.toggle("active", b.dataset.link === state.linkMode));
    if (state.nbVisual) {
      document.querySelectorAll("[data-nbtab]").forEach(b =>
        b.classList.toggle("active", b.dataset.nbtab === state.nbMode));
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
    // Default: hide the third (modified / moved) pill; branches that use it
    // re-show it below. Prevents a stale "⇅N" lingering across file switches.
    { const _m = $("#statMod"); if (_m) _m.classList.add("hidden"); }
    if (isVisualPdf()) {
      const s = state.pdfData ? state.pdfData.stats : { add: 0, del: 0, mod: 0 };
      $("#statAdd").textContent = "+" + s.add;
      $("#statDel").textContent = "−" + s.del;
      const modEl = $("#statMod");
      if (modEl) {
        modEl.textContent = "~" + (s.mod || 0);
        modEl.classList.toggle("hidden", !s.mod);
      }
      return;
    }
    if (isVisualDoc()) {
      const s = state.docData ? state.docData.stats : { add: 0, del: 0 };
      $("#statAdd").textContent = "+" + s.add;
      $("#statDel").textContent = "−" + s.del;
      return;
    }
    if (isVisualNb()) {
      const s = state.nbData ? state.nbData.stats : { add: 0, del: 0, moved: 0 };
      $("#statAdd").textContent = "+" + s.add;
      $("#statDel").textContent = "−" + s.del;
      const modEl = $("#statMod");
      if (modEl) {
        modEl.textContent = "⇅" + (s.moved || 0);
        modEl.title = "moved cells";
        modEl.classList.toggle("hidden", !s.moved);
      }
      return;
    }
    if (isGrid()) {
      const s = state.gridData.stats;
      $("#statAdd").textContent = "+" + s.addCells;
      $("#statDel").textContent = "−" + s.delCells;
      $("#statMod").textContent = "~" + s.modCells;
      updateChangesButton();
    } else {
      const st = E.stats(state.rows);
      $("#statAdd").textContent = "+" + st.add;
      $("#statDel").textContent = "−" + st.del;
    }
  }

  // ============ Changes panel (spreadsheets) ============
  function gridChangeRowCount() {
    if (!state.gridData) return 0;
    let n = 0;
    state.gridData.sheets.forEach(sh => sh.rows.forEach(r => { if (r.type !== "eq") n++; }));
    return n;
  }
  function updateChangesButton() {
    const el = $("#changesCount");
    if (!el) return;
    const s = state.gridData ? state.gridData.stats : null;
    const colChanges = s ? (s.addedCols + s.removedCols + s.movedCols) : 0;
    el.textContent = gridChangeRowCount() + colChanges;
  }
  function closeChangesPanel() {
    const p = $("#changesPanel");
    if (p) { p.classList.remove("open"); p.setAttribute("aria-hidden", "true"); }
  }
  function openChangesPanel() {
    const p = $("#changesPanel");
    if (!p) return;
    populateChangesPanel();
    p.classList.add("open"); p.setAttribute("aria-hidden", "false");
  }
  function toggleChangesPanel() {
    const p = $("#changesPanel");
    if (p && p.classList.contains("open")) closeChangesPanel(); else openChangesPanel();
  }
  function colLabelFor(sh, c) {
    const letter = c.bIdx != null ? GD.colName(c.bIdx) : GD.colName(c.aIdx);
    return sh.useHeaders && c.name ? c.name : letter;
  }
  function populateChangesPanel() {
    const body = $("#cpBody"), sub = $("#cpSub");
    if (!body || !state.gridData) return;
    const gd = state.gridData, multi = gd.sheets.length > 1;
    let html = "", totalRows = 0;
    gd.sheets.forEach((sh, si) => {
      const colItems = [], rowItems = [];
      sh.columns.forEach(c => {
        if (c.status === "same") return;
        const label = esc(colLabelFor(sh, c));
        if (c.status === "added") colItems.push(`<div class="cp-item cp-coli"><span class="cp-pill add">col +</span><span class="cp-rowno">${label}</span></div>`);
        else if (c.status === "removed") colItems.push(`<div class="cp-item cp-coli"><span class="cp-pill del">col −</span><span class="cp-rowno">${label}</span></div>`);
        else colItems.push(`<div class="cp-item cp-coli"><span class="cp-pill move">col ⇄</span><span class="cp-rowno">${label}</span><span class="cp-cols">${GD.colName(c.aIdx)}→${GD.colName(c.bIdx)}</span></div>`);
      });
      sh.rows.forEach((r, ri) => {
        if (r.type === "eq") return;
        totalRows++;
        const pill = r.type === "add" ? "add" : r.type === "del" ? "del" : r.type === "move" ? "move" : "mod";
        const tag = r.type === "add" ? "added" : r.type === "del" ? "removed" : r.type === "move" ? "moved" : "edited";
        const no = r.type === "del" ? r.left.no : r.right.no;
        let colsTxt = "";
        if (r.changed && r.changed.length) colsTxt = r.changed.map(ci => colLabelFor(sh, sh.columns[ci])).join(", ");
        if (r.type === "move") colsTxt = `row ${r.left ? r.left.no : "?"} → ${no}` + (colsTxt ? " · " + colsTxt : "");
        rowItems.push(`<div class="cp-item" data-rk="${si}-${ri}"><span class="cp-pill ${pill}">${tag}</span><span class="cp-rowno">Row ${no}</span><span class="cp-cols">${esc(colsTxt)}</span></div>`);
      });
      if (multi && (colItems.length || rowItems.length)) html += `<div class="cp-section">${esc(sh.name)}</div>`;
      if (colItems.length) html += `<div class="cp-section">Columns</div>` + colItems.join("");
      if (rowItems.length) html += `<div class="cp-section">Rows</div>` + rowItems.join("");
    });
    body.innerHTML = html || `<div class="cp-empty">No changes.</div>`;
    if (sub) {
      const s = gd.stats;
      const bits = [];
      if (totalRows) bits.push(totalRows + " row" + (totalRows === 1 ? "" : "s"));
      const cc = s.addedCols + s.removedCols + s.movedCols;
      if (cc) bits.push(cc + " column" + (cc === 1 ? "" : "s"));
      sub.textContent = bits.join(" · ") || "no changes";
    }
  }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  /* Keep the two side-by-side grid panes scrolling together (both axes), so
     aligned rows/columns stay lined up while panning a wide sheet. */
  function syncGridScroll() {
    diffEl.querySelectorAll(".gridpair").forEach(pair => {
      const sides = pair.querySelectorAll(".gridside");
      if (sides.length !== 2) return;
      const [l, r] = sides;
      let lock = false;
      const mirror = (from, to) => {
        if (lock) return;
        lock = true;
        to.scrollTop = from.scrollTop;
        to.scrollLeft = from.scrollLeft;
        lock = false;
      };
      l.addEventListener("scroll", () => mirror(l, r));
      r.addEventListener("scroll", () => mirror(r, l));
    });
  }

  function jumpToGridRow(rk) {
    const v = (window.CSS && CSS.escape) ? CSS.escape(rk) : rk;
    let trs;
    try { trs = diffEl.querySelectorAll(`.gtr[data-rk="${v}"]`); } catch (_) { return; }
    if (!trs.length) return;
    const first = trs[0];
    // scroll within the row's own scroll box (.gridside / .grid-scroll) when it
    // is capped-height, else fall back to the page scroller.
    const box = first.closest(".gridside, .grid-scroll");
    if (box && box.scrollHeight > box.clientHeight + 4) {
      const bRect = box.getBoundingClientRect(), rRect = first.getBoundingClientRect();
      box.scrollTop += (rRect.top - bRect.top) - box.clientHeight / 3;
    }
    const cRect = diffEl.getBoundingClientRect(), rRect2 = first.getBoundingClientRect();
    diffEl.scrollTop += (rRect2.top - cRect.top) - diffEl.clientHeight / 3;
    trs.forEach(tr => { tr.classList.remove("flash"); void tr.offsetWidth; tr.classList.add("flash"); });
  }

  // ============ Change navigation ============
  // Build a flat list of DOM elements (first row of each change run).
  function scanNav() {
    state.navTargets = [];
    let scope, sel;
    if (isVisualPdf()) {
      scope = diffEl.querySelector(".pdf-pages");
      if (!scope) return;
      const seen = new Set();
      scope.querySelectorAll(".pdf-hl").forEach(el => {
        const row = el.closest(".pdf-row");
        if (!row) return;
        // dedupe the mirrored A/B highlight of the same change line
        const key = row.dataset.page + ":" + Math.round(parseFloat(el.style.top));
        if (seen.has(key)) return;
        seen.add(key);
        state.navTargets.push(el);
      });
      if (state.currentHunk >= state.navTargets.length) state.currentHunk = -1;
      return;
    }
    if (isVisualDoc()) {
      scope = diffEl.querySelector(".doc-pages");
      if (!scope) return;
      const seen = new Set();
      scope.querySelectorAll(".dv-mark").forEach(el => {
        const id = el.dataset.chg;
        if (seen.has(id)) return;
        seen.add(id);
        state.navTargets.push(el);
      });
      if (state.currentHunk >= state.navTargets.length) state.currentHunk = -1;
      return;
    }
    if (isVisualNb()) {
      scope = diffEl.querySelector(".nb-cells");
      if (!scope) return;
      const seen = new Set();
      scope.querySelectorAll(".dv-mark").forEach(el => {
        const id = el.dataset.chg;
        if (!id || seen.has(id)) return;
        seen.add(id);
        state.navTargets.push(el);
      });
      if (state.currentHunk >= state.navTargets.length) state.currentHunk = -1;
      return;
    }
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
    // If the element lives inside a capped-height grid scroll box, scroll that
    // box first so the row is actually revealed.
    const box = el.closest(".gridside, .grid-scroll");
    if (box && box.scrollHeight > box.clientHeight + 4) {
      const bRect = box.getBoundingClientRect(), eR = el.getBoundingClientRect();
      box.scrollTop += (eR.top - bRect.top) - box.clientHeight / 3;
    }
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
  // Swap the two sides (A ↔ B), including the chosen sheets, then re-diff.
  function swapSides() {
    const tmp = state.a; state.a = state.b; state.b = tmp;
    slot.a = state.a; slot.b = state.b;
    const ts = state.sheetA; state.sheetA = state.sheetB; state.sheetB = ts;
    state.currentHunk = -1;
    state.pdfData = null;
    state.docData = null;
    state.nbData = null;
    state.gridData = null;
    rebuildDiff();
    toast("Swapped sides");
  }

  function reset() {
    slot.a = null; slot.b = null;
    state.a = null; state.b = null;
    state.search = ""; state.searchIdx = -1; state.currentHunk = -1;
    state.view = "split"; state.focus = false; state.gridMode = true;
    state.sheetA = null; state.sheetB = null;
    state.pdfVisual = false; state.pdfMode = "visual"; state.pdfData = null; state.ocrActive = false;
    state.docVisual = false; state.docMode = "visual"; state.docData = null; state.docGran = "word"; state.linkMode = "changes";
    state.nbVisual = false; state.nbMode = "visual"; state.nbData = null;
    state.forced = new Set();
    $("#searchInput").value = "";
    $("#searchCount").textContent = "";
    $("#focusToggle").classList.remove("on");
    document.querySelectorAll("[data-view]").forEach(b => b.classList.toggle("active", b.dataset.view === "split"));
    document.querySelectorAll("[data-tab]").forEach(b => b.classList.toggle("active", b.dataset.tab === "grid"));
    document.querySelectorAll(".slot").forEach(s => s.classList.remove("filled", "pasting"));
    document.querySelectorAll(".fc-sheet").forEach(w => { w.hidden = true; });
    document.querySelectorAll(".fc-sheet-sel").forEach(s => { s.innerHTML = ""; });
    document.querySelectorAll(".pastebox textarea").forEach(t => t.value = "");
    uploadEl.classList.remove("hidden");
    toolbarEl.classList.add("hidden");
    diffEl.classList.add("hidden");
    diffEl.innerHTML = "";
    if (window.DiffMinimap) DiffMinimap.refresh();
    $("#resetBtn").classList.add("hidden");
    $("#swapBtn").classList.add("hidden");
    $("#compareBtn").disabled = true;
    closeChangesPanel();
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
    // sheet pickers (multi-sheet spreadsheets) — choose which sheet of each file
    // to compare. Changing a dropdown only STAGES the choice; the diff re-runs
    // when the user clicks "Compare sheets".
    const sheetSelA = $("#sheetA"), sheetSelB = $("#sheetB"), sheetApplyBtn = $("#sheetApplyBtn");
    const syncSheetApply = () => {
      if (!sheetApplyBtn) return;
      const changed = (sheetSelA && sheetSelA.value !== state.sheetA) ||
                      (sheetSelB && sheetSelB.value !== state.sheetB);
      sheetApplyBtn.disabled = !changed;
    };
    if (sheetSelA) sheetSelA.addEventListener("change", syncSheetApply);
    if (sheetSelB) sheetSelB.addEventListener("change", syncSheetApply);
    if (sheetApplyBtn) sheetApplyBtn.addEventListener("click", () => {
      if (sheetSelA) state.sheetA = sheetSelA.value;
      if (sheetSelB) state.sheetB = sheetSelB.value;
      state.currentHunk = -1;
      sheetApplyBtn.disabled = true;
      rebuildDiff();
    });
    // Changes panel (spreadsheets)
    const changesBtn = $("#changesBtn"), cpClose = $("#cpClose"), cpBody = $("#cpBody");
    if (changesBtn) changesBtn.addEventListener("click", toggleChangesPanel);
    if (cpClose) cpClose.addEventListener("click", closeChangesPanel);
    if (cpBody) cpBody.addEventListener("click", (e) => {
      const item = e.target.closest(".cp-item[data-rk]");
      if (item) jumpToGridRow(item.getAttribute("data-rk"));
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
    // visual / text segmented (PDFs only)
    document.querySelectorAll("[data-pdftab]").forEach(b => {
      b.addEventListener("click", () => {
        if (state.pdfMode === b.dataset.pdftab) return;
        state.pdfMode = b.dataset.pdftab;
        state.currentHunk = -1;
        if (state.pdfMode === "text") {
          state.rows = E.buildRows(srcText(state.a), srcText(state.b), { ignoreWs: state.ignoreWs });
        }
        document.querySelectorAll("[data-pdftab]").forEach(x => x.classList.toggle("active", x === b));
        updateToolbarForMode();
        renderAll();
        updateStats();
      });
    });
    // OCR button inside the PDF "pixel-compared" note (delegated — note is re-rendered)
    diffEl.addEventListener("click", (e) => {
      if (e.target.closest("#ocrBtn")) runOcr();
    });
    // visual / text segmented (Word docs only)
    document.querySelectorAll("[data-doctab]").forEach(b => {
      b.addEventListener("click", () => {
        if (state.docMode === b.dataset.doctab) return;
        state.docMode = b.dataset.doctab;
        state.currentHunk = -1;
        if (state.docMode === "text") {
          state.rows = E.buildRows(state.a.text, state.b.text, { ignoreWs: state.ignoreWs });
        }
        document.querySelectorAll("[data-doctab]").forEach(x => x.classList.toggle("active", x === b));
        updateToolbarForMode();
        renderAll();
        updateStats();
      });
    });
    // Hover-link scope toggle: changes-only vs all text
    document.querySelectorAll("[data-link]").forEach(b => {
      b.addEventListener("click", () => {
        if (state.linkMode === b.dataset.link) return;
        state.linkMode = b.dataset.link;
        document.querySelectorAll("[data-link]").forEach(x => x.classList.toggle("active", x === b));
      });
    });
    // Highlight granularity: per-word vs whole-sentence (Word docs & notebooks)
    document.querySelectorAll("[data-docgran]").forEach(b => {
      b.addEventListener("click", () => {
        if (state.docGran === b.dataset.docgran) return;
        state.docGran = b.dataset.docgran;
        state.currentHunk = -1;
        document.querySelectorAll("[data-docgran]").forEach(x => x.classList.toggle("active", x === b));
        if (isVisualNb()) {
          state.nbData = NV.build(state.a, state.b, state.docGran);
          renderAll();
          updateStats();
          return;
        }
        showLoading(true, "Updating highlights…");
        DV.build(state.a, state.b, state.docGran).then(d => {
          state.docData = d;
          showLoading(false);
          renderAll();
          updateStats();
        });
      });
    });
    document.querySelectorAll("[data-nbtab]").forEach(b => {
      b.addEventListener("click", () => {
        if (state.nbMode === b.dataset.nbtab) return;
        state.nbMode = b.dataset.nbtab;
        state.currentHunk = -1;
        if (state.nbMode === "text") {
          state.rows = E.buildRows(state.a.text, state.b.text, { ignoreWs: state.ignoreWs });
        }
        document.querySelectorAll("[data-nbtab]").forEach(x => x.classList.toggle("active", x === b));
        updateToolbarForMode();
        renderAll();
        updateStats();
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
    $("#swapBtn").addEventListener("click", swapSides);
    const sheetSwapBtn = $("#sheetSwapBtn");
    if (sheetSwapBtn) sheetSwapBtn.addEventListener("click", swapSides);

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
    setupHoverLink();
    if (window.DiffMinimap) DiffMinimap.mount(diffEl);
    $("#compareBtn").addEventListener("click", runCompare);
  }

  /* Hover a piece of one document → highlight the piece it maps to in the
     other. "changes" mode links only changed pieces (data-chg / changed grid
     cells); "all" mode also links unchanged text by block (data-map) and any
     grid cell (data-gk). */
  function setupHoverLink() {
    let activeKey = null;
    const clear = () => {
      if (!activeKey) return;
      diffEl.querySelectorAll(".peer-hl").forEach(el => el.classList.remove("peer-hl"));
      activeKey = null;
    };
    const lite = (attr, v) => {
      const key = attr + "=" + v;
      if (key === activeKey) return true;
      clear();
      let matches;
      try { matches = diffEl.querySelectorAll("[" + attr + '="' + (window.CSS && CSS.escape ? CSS.escape(v) : v) + '"]'); }
      catch (_) { return true; }
      matches.forEach(m => m.classList.add("peer-hl"));
      activeKey = key;
      return true;
    };
    diffEl.addEventListener("mouseover", (e) => {
      const all = state.linkMode === "all";
      // 1. fine change marks (Word / notebook), incl. reworded text
      const chgEl = e.target.closest("[data-chg]");
      if (chgEl) return void lite("data-chg", chgEl.getAttribute("data-chg"));
      // 2. spreadsheet cell — in "changes" mode only changed cells map
      const cell = e.target.closest("[data-gk]");
      if (cell && (all || /gc-(add|del|mod)/.test(cell.className))) return void lite("data-gk", cell.getAttribute("data-gk"));
      // 3. generic block / line mapping (Word paragraphs, code/HTML lines).
      //    In "changes" mode only changed units (data-changed) map.
      const blk = e.target.closest("[data-map]");
      if (blk && (all || blk.hasAttribute("data-changed"))) return void lite("data-map", blk.getAttribute("data-map"));
      clear();
    });
    diffEl.addEventListener("mouseleave", clear);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else { init(); }
})();
