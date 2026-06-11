/* ============================================================
   render-grid.js — render structured spreadsheet diffs as real cell
   grids with header-aware column alignment, per-cell change colour,
   moved-row / moved-column markers, and per-row change indicators.
   Depends on DiffEngine (esc), GridDiff, DiffParse. Exposes window.GridRender
   ============================================================ */
(function () {
  const E = window.DiffEngine;
  const G = window.GridDiff;
  const esc = s => E.esc(String(s == null ? "" : s));
  const colName = G.colName;

  function colLetter(c) { return c.bIdx != null ? colName(c.bIdx) : colName(c.aIdx); }
  function colLabel(c, useHeaders) {
    if (useHeaders && c.name) return esc(c.name);
    return colLetter(c);
  }

  function render(container, state) {
    const gd = state.gridData;
    if (!gd.stats.changed) { container.innerHTML = identical(); return; }

    const parts = [fileHeader(state)];
    gd.sheets.forEach((sheet, si) => {
      parts.push(`<div class="sheet">`);
      if (gd.sheets.length > 1) {
        parts.push(`<div class="sheet-name"><svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><rect x="1.5" y="1.5" width="13" height="13" rx="1.5"/><path d="M1.5 6h13M1.5 10.5h13M6 1.5v13"/></svg> ${esc(sheet.name)}${sheet.key ? `<span class="sheet-key">matched by “${esc(sheet.key)}”</span>` : ""}</div>`);
      }
      parts.push(colBanner(sheet));
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

  /* Column-level change summary banner (added / removed / moved columns). */
  function colBanner(sheet) {
    const moved = sheet.columns.filter(c => c.status === "moved");
    const added = sheet.columns.filter(c => c.status === "added");
    const removed = sheet.columns.filter(c => c.status === "removed");
    if (!moved.length && !added.length && !removed.length) return "";
    const chip = (cls, label, c, useH) =>
      `<span class="colchip ${cls}">${label} <strong>${colLabel(c, useH)}</strong>${cls === "move" ? ` <span class="cc-arrow">${colName(c.aIdx)}→${colName(c.bIdx)}</span>` : ""}</span>`;
    let h = `<div class="colbanner">`;
    removed.forEach(c => h += chip("del", "− column", c, sheet.useHeaders));
    added.forEach(c => h += chip("add", "+ column", c, sheet.useHeaders));
    moved.forEach(c => h += chip("move", "⇄ column", c, sheet.useHeaders));
    h += `</div>`;
    return h;
  }

  /* fold visibility list, forced keyed "si:idx". */
  function visibleItems(rows, focus, forced, si) {
    if (!focus) return rows.map((r, i) => ({ type: "row", row: r, i }));
    const ctx = 3;
    const keep = new Array(rows.length).fill(false);
    rows.forEach((r, i) => { if (r.type !== "eq") for (let j = i - ctx; j <= i + ctx; j++) if (j >= 0 && j < rows.length) keep[j] = true; });
    forced.forEach(k => { const [s, idx] = k.split(":"); if (+s === si && +idx < rows.length) keep[+idx] = true; });
    const out = []; let i = 0;
    while (i < rows.length) {
      if (keep[i]) { out.push({ type: "row", row: rows[i], i }); i++; }
      else { let j = i; while (j < rows.length && !keep[j]) j++; out.push({ type: "fold", from: i, to: j - 1, count: j - i }); i = j; }
    }
    return out;
  }

  function colHeaderRow(sheet, extra) {
    let h = `<th class="ghead corner"></th>` + (extra || "");
    sheet.columns.forEach((c, ci) => {
      const st = c.status === "added" ? " gch-add" : c.status === "removed" ? " gch-del" : c.status === "moved" ? " gch-move" : "";
      const tag = c.status === "added" ? `<span class="ch-tag add">new</span>`
        : c.status === "removed" ? `<span class="ch-tag del">removed</span>`
        : c.status === "moved" ? `<span class="ch-tag move">${colName(c.aIdx)}→${colName(c.bIdx)}</span>` : "";
      h += `<th class="ghead colhead${st}" data-col="${ci}" title="${sheet.useHeaders && c.name ? esc(c.name) + " · " : ""}column ${colLetter(c)}">` +
        `<span class="ch-name">${colLabel(c, sheet.useHeaders)}</span>${tag}</th>`;
    });
    return `<tr class="gtr ghead-row">${h}</tr>`;
  }

  /* status → short tag for the row indicator */
  const ROW_TAG = { add: "added", del: "removed", mod: "edited", move: "moved" };

  /* ---------- Side-by-side: two aligned tables ---------- */
  function pairTables(sheet, state, si) {
    const items = visibleItems(sheet.rows, state.focus, state.forced, si);
    return `<div class="gridpair">` +
      `<div class="gridside left grid-primary">${oneSide(items, sheet, si, "left")}</div>` +
      `<div class="gridside right">${oneSide(items, sheet, si, "right")}</div>` +
      `</div>`;
  }

  function oneSide(items, sheet, si, side) {
    const isLeft = side === "left";
    const cols = sheet.columns;
    let body = "";
    for (const it of items) {
      if (it.type === "fold") {
        body += `<tr class="gfold" data-si="${si}" data-from="${it.from}" data-to="${it.to}"><td class="gfold-cell" colspan="${cols.length + 1}">⋯ Show ${it.count} unchanged row${it.count === 1 ? "" : "s"}</td></tr>`;
        continue;
      }
      const r = it.row;
      // which underlying row this side displays
      let payload;
      if (isLeft) payload = (r.type === "add") ? null : r.left;
      else payload = (r.type === "del") ? null : r.right;

      const no = payload ? payload.no : "";
      const moveBadge = (r.type === "move")
        ? `<span class="rh-move" title="row moved">${isLeft ? "↧" : "↥"}</span>` : "";
      let tds = `<td class="ghead rowhead" data-type="${r.type}">${no}${moveBadge}</td>`;
      for (let ci = 0; ci < cols.length; ci++) {
        const c = cols[ci];
        // column present on this side?
        const present = isLeft ? c.aIdx != null : c.bIdx != null;
        if (!payload || !present) { tds += `<td class="gcell gc-blank">&nbsp;</td>`; continue; }
        const idx = isLeft ? c.aIdx : c.bIdx;
        const val = esc(payload.cells[idx] != null ? payload.cells[idx] : "");
        const status = r.cls ? (r.cls[ci] || "same") : "same";
        let cls = "";
        if (c.status === "added") cls = isLeft ? "" : "gc-coladd";
        else if (c.status === "removed") cls = isLeft ? "gc-coldel" : "";
        else if (isLeft) { if (status === "del" || status === "mod") cls = "gc-del"; }
        else { if (status === "add" || status === "mod") cls = "gc-add"; }
        tds += `<td class="gcell ${cls}" data-gk="${si}-${it.i}-${ci}">${val || "&nbsp;"}</td>`;
      }
      const rowAttr = r.type !== "eq" ? ` data-rk="${si}-${it.i}" data-rtype="${r.type}"` : "";
      body += `<tr class="gtr${payload ? "" : " gr-blank"}${r.type === "move" ? " gtr-move" : ""}" data-type="${r.type}"${rowAttr}>${tds}</tr>`;
    }
    return `<table class="gtable"><thead>${colHeaderRow(sheet)}</thead><tbody>${body}</tbody></table>`;
  }

  /* ---------- Inline: one merged table ---------- */
  function mergedTable(sheet, state, si) {
    const items = visibleItems(sheet.rows, state.focus, state.forced, si);
    const cols = sheet.columns;
    let body = "";
    for (const it of items) {
      if (it.type === "fold") {
        body += `<tr class="gfold" data-si="${si}" data-from="${it.from}" data-to="${it.to}"><td class="gfold-cell" colspan="${cols.length + 2}">⋯ Show ${it.count} unchanged row${it.count === 1 ? "" : "s"}</td></tr>`;
        continue;
      }
      const r = it.row;
      if (r.type === "eq") {
        let tds = `<td class="ghead rowhead">${r.right.no}</td><td class="gsign"></td>`;
        for (let ci = 0; ci < cols.length; ci++) {
          const c = cols[ci];
          if (c.status === "removed") { tds += `<td class="gcell gc-blank">&nbsp;</td>`; continue; }
          const cls = c.status === "added" ? "gc-coladd" : "";
          const v = r.right.cells[c.bIdx] != null ? r.right.cells[c.bIdx] : "";
          tds += `<td class="gcell ${cls}">${esc(v) || "&nbsp;"}</td>`;
        }
        body += `<tr class="gtr" data-type="eq">${tds}</tr>`;
        continue;
      }
      const oldC = r.type === "add" ? null : r.left.cells;
      const newC = r.type === "del" ? null : r.right.cells;
      const sign = r.type === "add" ? "+" : r.type === "del" ? "−" : r.type === "move" ? "⇅" : "~";
      const signCls = r.type === "add" ? "plus" : r.type === "del" ? "minus" : r.type === "move" ? "move" : "tilde";
      const no = r.type === "del" ? r.left.no : r.right.no;
      let tds = `<td class="ghead rowhead" data-type="${r.type}">${no}</td><td class="gsign ${signCls}">${sign}</td>`;
      for (let ci = 0; ci < cols.length; ci++) {
        const c = cols[ci];
        const status = r.cls ? (r.cls[ci] || "same") : "same";
        let cls = "", inner;
        if (c.status === "added") {
          cls = "gc-coladd"; inner = esc(newC && newC[c.bIdx] != null ? newC[c.bIdx] : "") || "&nbsp;";
        } else if (c.status === "removed") {
          cls = "gc-coldel"; const o = esc(oldC && oldC[c.aIdx] != null ? oldC[c.aIdx] : ""); inner = o ? `<span class="cell-gone">${o}</span>` : "&nbsp;";
        } else if (status === "add") {
          cls = "gc-add"; inner = esc(newC && newC[c.bIdx] != null ? newC[c.bIdx] : "") || "&nbsp;";
        } else if (status === "del") {
          cls = "gc-del"; const o = esc(oldC && oldC[c.aIdx] != null ? oldC[c.aIdx] : ""); inner = o ? `<span class="cell-gone">${o}</span>` : "&nbsp;";
        } else if (status === "mod") {
          cls = "gc-mod"; const nv = esc(newC[c.bIdx] != null ? newC[c.bIdx] : ""), ov = esc(oldC[c.aIdx] != null ? oldC[c.aIdx] : "");
          inner = `<span class="cell-new">${nv || "&nbsp;"}</span>` + (ov ? `<span class="cell-old">${ov}</span>` : "");
        } else {
          const src = newC || oldC, idx = newC ? c.bIdx : c.aIdx;
          inner = esc(src && src[idx] != null ? src[idx] : "") || "&nbsp;";
        }
        tds += `<td class="gcell ${cls}">${inner}</td>`;
      }
      body += `<tr class="gtr${r.type === "move" ? " gtr-move" : ""}" data-type="${r.type}" data-rk="${si}-${it.i}" data-rtype="${r.type}">${tds}</tr>`;
    }
    return `<div class="grid-scroll"><table class="gtable merged grid-primary"><thead>${colHeaderRow(sheet, `<th class="ghead corner sign"></th>`)}</thead><tbody>${body}</tbody></table></div>`;
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
