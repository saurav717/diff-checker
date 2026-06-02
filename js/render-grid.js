/* ============================================================
   render-grid.js — render structured spreadsheet diffs as
   real cell grids with per-cell change highlighting.
   Depends on DiffEngine (esc), GridDiff, DiffParse. Exposes window.GridRender
   ============================================================ */
(function () {
  const E = window.DiffEngine;
  const G = window.GridDiff;
  const esc = s => E.esc(String(s == null ? "" : s));

  function render(container, state) {
    const gd = state.gridData;
    if (!gd.stats.changed) { container.innerHTML = identical(); return; }

    const parts = [fileHeader(state)];
    gd.sheets.forEach((sheet, si) => {
      parts.push(`<div class="sheet">`);
      if (gd.sheets.length > 1) {
        parts.push(`<div class="sheet-name"><svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><rect x="1.5" y="1.5" width="13" height="13" rx="1.5"/><path d="M1.5 6h13M1.5 10.5h13M6 1.5v13"/></svg> ${esc(sheet.name)}</div>`);
      }
      parts.push(state.view === "inline" ? mergedTable(sheet, state, si) : pairTables(sheet, state, si));
      parts.push(`</div>`);
    });
    container.innerHTML = parts.join("");
  }

  function fileHeader(state) {
    if (state.view === "inline") {
      return `<div class="fileheader inline"><div class="fh-cell"><span class="fh-dot" style="background:var(--del-sign)"></span><span class="fh-name">${esc(state.a.name)}</span><span class="fh-kind">→ ${esc(state.b.name)} · grid</span></div></div>`;
    }
    const lbl = window.DiffParse.kindLabel;
    return `<div class="fileheader split">` +
      `<div class="fh-cell a"><span class="fh-dot"></span><span class="fh-name">${esc(state.a.name)}</span><span class="fh-kind">${lbl(state.a.kind)}</span></div>` +
      `<div class="fh-cell b"><span class="fh-dot"></span><span class="fh-name">${esc(state.b.name)}</span><span class="fh-kind">${lbl(state.b.kind)}</span></div>` +
      `</div>`;
  }

  /* fold visibility list, forced keyed "si:idx".
     Row 0 (the header row) is always kept so column titles stay visible. */
  function visibleItems(rows, focus, forced, si) {
    if (!focus) return rows.map((r, i) => ({ type: "row", row: r, i }));
    const ctx = 3;
    const keep = new Array(rows.length).fill(false);
    if (rows.length) keep[0] = true; // pin header row
    rows.forEach((r, i) => { if (r.type !== "eq") for (let j = i - ctx; j <= i + ctx; j++) if (j >= 0 && j < rows.length) keep[j] = true; });
    forced.forEach(k => { const [s, idx] = k.split(":"); if (+s === si && +idx < rows.length) keep[+idx] = true; });
    const out = []; let i = 0;
    while (i < rows.length) {
      if (keep[i]) { out.push({ type: "row", row: rows[i], i }); i++; }
      else { let j = i; while (j < rows.length && !keep[j]) j++; out.push({ type: "fold", from: i, to: j - 1, count: j - i }); i = j; }
    }
    return out;
  }

  function colHeader(cols) {
    let h = `<th class="ghead corner"></th>`;
    for (let c = 0; c < cols; c++) h += `<th class="ghead colhead">${G.colName(c)}</th>`;
    return `<tr class="gtr ghead-row">${h}</tr>`;
  }

  /* ---------- Side-by-side: two tables ---------- */
  function pairTables(sheet, state, si) {
    const items = visibleItems(sheet.rows, state.focus, state.forced, si);
    const left = oneSide(items, sheet.cols, si, "left");
    const right = oneSide(items, sheet.cols, si, "right");
    return `<div class="gridpair">` +
      `<div class="gridside left grid-primary">${left}</div>` +
      `<div class="gridside right">${right}</div>` +
      `</div>`;
  }

  function oneSide(items, cols, si, side) {
    const isLeft = side === "left";
    let body = "";
    for (const it of items) {
      if (it.type === "fold") {
        body += `<tr class="gfold" data-si="${si}" data-from="${it.from}" data-to="${it.to}"><td class="gfold-cell" colspan="${cols + 1}">⋯ Show ${it.count} unchanged row${it.count === 1 ? "" : "s"}</td></tr>`;
        continue;
      }
      const r = it.row;
      // Which row this side displays (old on left, new on right). null = no row here.
      let payload;
      if (r.type === "eq") payload = isLeft ? r.left : r.right;
      else if (isLeft) payload = r.type === "add" ? null : r.left;
      else payload = r.type === "del" ? null : r.right;

      const no = payload ? payload.no : "";
      let tds = `<td class="ghead rowhead">${no}</td>`;
      for (let c = 0; c < cols; c++) {
        if (!payload) { tds += `<td class="gcell gc-blank">&nbsp;</td>`; continue; }
        const val = esc(payload.cells[c] != null ? payload.cells[c] : "");
        const status = r.cls ? (r.cls[c] || "same") : "same";
        // Box-level colour: red where this side LOST content, green where it GAINED it.
        let cls = "";
        if (isLeft) { if (status === "del" || status === "mod") cls = "gc-del"; }
        else { if (status === "add" || status === "mod") cls = "gc-add"; }
        // Shared key so hovering a cell highlights the SAME cell in the other table.
        tds += `<td class="gcell ${cls}" data-gk="${si}-${it.i}-${c}">${val || "&nbsp;"}</td>`;
      }
      body += `<tr class="gtr${payload ? "" : " gr-blank"}" data-type="${r.type}">${tds}</tr>`;
    }
    return `<table class="gtable"><thead>${colHeader(cols)}</thead><tbody>${body}</tbody></table>`;
  }

  /* ---------- Inline: one merged table ---------- */
  function mergedTable(sheet, state, si) {
    const items = visibleItems(sheet.rows, state.focus, state.forced, si);
    const cols = sheet.cols;
    let body = "";
    for (const it of items) {
      if (it.type === "fold") {
        body += `<tr class="gfold" data-si="${si}" data-from="${it.from}" data-to="${it.to}"><td class="gfold-cell" colspan="${cols + 2}">⋯ Show ${it.count} unchanged row${it.count === 1 ? "" : "s"}</td></tr>`;
        continue;
      }
      const r = it.row;
      if (r.type === "eq") {
        let tds = `<td class="ghead rowhead">${r.right.no}</td><td class="gsign"></td>`;
        for (let c = 0; c < cols; c++) tds += `<td class="gcell">${esc(r.right.cells[c] != null ? r.right.cells[c] : "") || "&nbsp;"}</td>`;
        body += `<tr class="gtr" data-type="eq">${tds}</tr>`;
        continue;
      }
      const oldC = r.type === "add" ? [] : r.left.cells;
      const newC = r.type === "del" ? [] : r.right.cells;
      const sign = r.type === "add" ? "+" : r.type === "del" ? "−" : "~";
      const signCls = r.type === "add" ? "plus" : r.type === "del" ? "minus" : "tilde";
      const no = r.type === "del" ? r.left.no : r.right.no;
      let tds = `<td class="ghead rowhead">${no}</td><td class="gsign ${signCls}">${sign}</td>`;
      for (let c = 0; c < cols; c++) {
        const status = r.cls ? (r.cls[c] || "same") : "same";
        let cls = "", inner;
        if (status === "add") {
          cls = "gc-add"; inner = esc(newC[c] != null ? newC[c] : "") || "&nbsp;";
        } else if (status === "del") {
          cls = "gc-del"; const o = esc(oldC[c] != null ? oldC[c] : ""); inner = o ? `<span class="cell-gone">${o}</span>` : "&nbsp;";
        } else if (status === "mod") {
          cls = "gc-mod"; const nv = esc(newC[c] != null ? newC[c] : ""), ov = esc(oldC[c] != null ? oldC[c] : "");
          inner = `<span class="cell-new">${nv || "&nbsp;"}</span>` + (ov ? `<span class="cell-old">${ov}</span>` : "");
        } else {
          const v = newC[c] != null ? newC[c] : (oldC[c] != null ? oldC[c] : "");
          inner = esc(v) || "&nbsp;";
        }
        tds += `<td class="gcell ${cls}">${inner}</td>`;
      }
      body += `<tr class="gtr" data-type="${r.type}">${tds}</tr>`;
    }
    return `<table class="gtable merged grid-primary"><thead><tr class="gtr ghead-row"><th class="ghead corner"></th><th class="ghead corner sign"></th>${colHeaderCells(cols)}</tr></thead><tbody>${body}</tbody></table>`;
  }

  function colHeaderCells(cols) {
    let h = "";
    for (let c = 0; c < cols; c++) h += `<th class="ghead colhead">${G.colName(c)}</th>`;
    return h;
  }

  function identical() {
    return `<div class="identical">` +
      `<svg class="ic" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="9"/><path d="M8.5 12.5l2.2 2.2L16 9.5"/></svg>` +
      `<h2>The spreadsheets are identical</h2>` +
      `<p>No cell-level differences were found.</p>` +
      `</div>`;
  }

  window.GridRender = { render };
})();
