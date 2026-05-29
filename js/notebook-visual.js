/* ============================================================
   notebook-visual.js — show two Jupyter notebooks side by side as
   rendered cells (markdown formatted, code in monospace, outputs
   incl. images/tables), with changed words highlighted inline.

   Works off the normalized cell model produced by parsers.js for
   both .ipynb and nbconvert .html:
     nbCells: [{ type:'code'|'markdown', source, html?, outputs:[...] }]

   Cells are aligned (diffArrays on a per-cell signature), paired
   cells are word-diffed, and each change region gets a shared
   data-chg id (reusing the .dv-mark classes so navigation +
   highlight styling already apply).

   Exposes: window.NotebookVisual.build(a, b) -> { rows, stats, changes }
            window.NotebookVisual.render(container, state)
   ============================================================ */
(function () {
  function esc(s) {
    return String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  // ---- word tokenisation + diff ----
  function tokenize(text) {
    const re = /\S+/g, t = []; let m;
    while ((m = re.exec(text))) t.push({ s: m[0], start: m.index, end: m.index + m[0].length });
    return t;
  }

  function diffWords(aText, bText, counter) {
    const at = tokenize(aText), bt = tokenize(bText);
    const aMark = new Map(), bMark = new Map();
    let del = 0, add = 0;
    if (typeof Diff !== "undefined") {
      const parts = Diff.diffArrays(at.map(x => x.s), bt.map(x => x.s));
      let ai = 0, bi = 0, prev = false;
      for (const p of parts) {
        const n = p.value.length, chg = p.added || p.removed;
        if (chg && !prev) counter.n++;
        if (p.removed) { for (let k = 0; k < n; k++) aMark.set(ai + k, counter.n); ai += n; del += n; }
        else if (p.added) { for (let k = 0; k < n; k++) bMark.set(bi + k, counter.n); bi += n; add += n; }
        else { ai += n; bi += n; }
        prev = chg;
      }
    }
    return { at, bt, aMark, bMark, del, add };
  }

  function renderMarked(text, toks, markMap, cls) {
    let html = "", pos = 0;
    toks.forEach((t, k) => {
      html += esc(text.slice(pos, t.start));
      const seg = esc(text.slice(t.start, t.end));
      html += markMap.has(k)
        ? `<mark class="dv-mark dv-${cls}" data-chg="${markMap.get(k)}">${seg}</mark>`
        : seg;
      pos = t.end;
    });
    html += esc(text.slice(pos));
    return html;
  }

  function markTextPair(aText, bText, counter) {
    const d = diffWords(aText || "", bText || "", counter);
    return {
      aHtml: renderMarked(aText || "", d.at, d.aMark, "del"),
      bHtml: renderMarked(bText || "", d.bt, d.bMark, "add"),
      del: d.del, add: d.add
    };
  }

  // ---- HTML (rendered markdown) tokenisation + marking (docx-style) ----
  function tokenizeHtml(html) {
    const root = document.createElement("div");
    root.innerHTML = html || "";
    const tokens = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    let node;
    while ((node = walker.nextNode())) {
      const text = node.nodeValue;
      if (!text || !text.trim()) continue;
      const re = /\S+/g; let m;
      while ((m = re.exec(text))) tokens.push({ node, start: m.index, end: m.index + m[0].length, word: m[0] });
    }
    return { root, tokens };
  }

  function applyHtmlMarks(tokens, markMap, cls) {
    const byNode = new Map();
    tokens.forEach((t, i) => {
      if (!markMap.has(i)) return;
      if (!byNode.has(t.node)) byNode.set(t.node, []);
      byNode.get(t.node).push({ start: t.start, end: t.end, chg: markMap.get(i) });
    });
    for (const [node, list] of byNode) {
      const text = node.nodeValue;
      list.sort((a, b) => a.start - b.start);
      let html = "", pos = 0;
      for (const t of list) {
        html += esc(text.slice(pos, t.start));
        html += `<mark class="dv-mark dv-${cls}" data-chg="${t.chg}">` + esc(text.slice(t.start, t.end)) + `</mark>`;
        pos = t.end;
      }
      html += esc(text.slice(pos));
      const span = document.createElement("span");
      span.className = "dv-wrap";
      span.innerHTML = html;
      node.parentNode.replaceChild(span, node);
    }
  }

  function markHtmlPair(htmlA, htmlB, counter) {
    const A = tokenizeHtml(htmlA), B = tokenizeHtml(htmlB);
    const aMark = new Map(), bMark = new Map();
    let del = 0, add = 0;
    if (typeof Diff !== "undefined") {
      const parts = Diff.diffArrays(A.tokens.map(t => t.word), B.tokens.map(t => t.word));
      let ai = 0, bi = 0, prev = false;
      for (const p of parts) {
        const n = p.value.length, chg = p.added || p.removed;
        if (chg && !prev) counter.n++;
        if (p.removed) { for (let k = 0; k < n; k++) aMark.set(ai + k, counter.n); ai += n; del += n; }
        else if (p.added) { for (let k = 0; k < n; k++) bMark.set(bi + k, counter.n); bi += n; add += n; }
        else { ai += n; bi += n; }
        prev = chg;
      }
    }
    applyHtmlMarks(A.tokens, aMark, "del");
    applyHtmlMarks(B.tokens, bMark, "add");
    return { aHtml: A.root.innerHTML, bHtml: B.root.innerHTML, del, add };
  }

  // ---- output rendering + diff ----
  function outText(outs) {
    return outs.filter(o => o.kind === "text").map(o => o.text).join("\n");
  }
  function richOuts(outs) {
    return outs.filter(o => o.kind === "image" || o.kind === "html");
  }
  function richKey(outs) {
    return JSON.stringify(richOuts(outs).map(o => o.kind === "image" ? "img:" + o.src : "html:" + o.html));
  }

  function renderOutputs(side, cell, otherCell, counter, refs) {
    const outs = cell.outputs || [];
    if (!outs.length && (!otherCell || !(otherCell.outputs || []).length)) return "";

    let html = "";
    // text / stream / error outputs (diffed)
    const aT = side === "a" ? outText(outs) : outText(otherCell ? otherCell.outputs : []);
    const bT = side === "a" ? outText(otherCell ? otherCell.outputs : []) : outText(outs);
    const myText = side === "a" ? aT : bT;
    if (myText.trim() || (aT.trim() && bT.trim())) {
      // reuse precomputed pair if present (so both sides share counter ids)
      const pair = refs.textPair || (refs.textPair = markTextPair(aT, bT, counter));
      const isErr = outs.some(o => o.kind === "text" && o.err);
      html += `<pre class="nb-out-text${isErr ? " nb-out-err" : ""}">${side === "a" ? pair.aHtml : pair.bHtml}</pre>`;
      if (!refs.counted) { refs.del += pair.del; refs.add += pair.add; }
    }

    // rich outputs (images / html tables) — flagged changed if the
    // ordered set differs between sides
    const mine = richOuts(outs);
    const otherChanged = otherCell ? (richKey(cell.outputs || []) !== richKey(otherCell.outputs || [])) : true;
    mine.forEach(o => {
      const changedCls = otherChanged ? ` nb-out-changed dv-mark dv-${side === "a" ? "del" : "add"}` : "";
      const chg = otherChanged ? ` data-chg="rich-${++counter.n}"` : "";
      if (o.kind === "image") {
        html += `<div class="nb-out-rich${changedCls}"${chg}><img src="${o.src}" alt="output"></div>`;
      } else {
        html += `<div class="nb-out-rich${changedCls}"${chg}>${o.html}</div>`;
      }
    });
    refs.counted = true;
    return html ? `<div class="nb-out">${html}</div>` : "";
  }

  // ---- one cell's rendered HTML for a given side ----
  function cellHtml(side, cell, body, outputsHtml, lone) {
    if (!cell) return `<div class="nb-cell nb-empty"><span>no matching cell</span></div>`;
    const loneCls = lone ? (side === "a" ? " nb-lone-del" : " nb-lone-add") : "";
    const title = cell.title ? `<div class="nb-celltitle">${esc(cell.title)}</div>` : "";
    if (cell.type === "markdown") {
      return `<div class="nb-cell nb-md${loneCls}">${title}${body}</div>`;
    }
    return `<div class="nb-cell nb-code${loneCls}">` +
      title +
      `<div class="nb-codewrap-outer">` +
      `<div class="nb-prompt">${side === "a" ? "[ ]" : "[ ]"}</div>` +
      `<div class="nb-code-wrap"><pre class="nb-src">${body}</pre>${outputsHtml || ""}</div>` +
      `</div>` +
    `</div>`;
  }

  // Build full A/B HTML for a paired (or lone) row.
  function buildRow(aCell, bCell, counter) {
    let aBody = "", bBody = "", del = 0, add = 0;
    const refs = { del: 0, add: 0, counted: false, textPair: null };

    if (aCell && bCell) {
      if (aCell.type === "markdown" && bCell.type === "markdown") {
        const mp = markHtmlPair(aCell.html, bCell.html, counter);
        aBody = mp.aHtml; bBody = mp.bHtml; del += mp.del; add += mp.add;
      } else {
        const cp = markTextPair(aCell.source, bCell.source, counter);
        aBody = cp.aHtml; bBody = cp.bHtml; del += cp.del; add += cp.add;
      }
      const aOut = renderOutputs("a", aCell, bCell, counter, refs);
      const bOut = renderOutputs("b", bCell, aCell, counter, refs);
      del += refs.del; add += refs.add;
      const aHtml = cellHtml("a", aCell, aBody, aOut, false);
      const bHtml = cellHtml("b", bCell, bBody, bOut, false);
      const changed = del + add > 0;
      return { aHtml, bHtml, changed, del, add };
    }

    if (aCell) { // removed cell
      const body = aCell.type === "markdown"
        ? markHtmlPair(aCell.html, "", counter).aHtml
        : markTextPair(aCell.source, "", counter).aHtml;
      const out = renderOutputs("a", aCell, null, counter, { del: 0, add: 0, counted: true, textPair: null });
      const toks = tokenize(aCell.source).length + tokenize((aCell.html || "").replace(/<[^>]+>/g, " ")).length;
      return { aHtml: cellHtml("a", aCell, body, out, true), bHtml: cellHtml("b", null), changed: true, del: Math.max(1, toks), add: 0 };
    }

    // added cell
    const body = bCell.type === "markdown"
      ? markHtmlPair("", bCell.html, counter).bHtml
      : markTextPair("", bCell.source, counter).bHtml;
    const out = renderOutputs("b", bCell, null, counter, { del: 0, add: 0, counted: true, textPair: null });
    const toks = tokenize(bCell.source).length + tokenize((bCell.html || "").replace(/<[^>]+>/g, " ")).length;
    return { aHtml: cellHtml("a", null), bHtml: cellHtml("b", bCell, body, out, true), changed: true, del: 0, add: Math.max(1, toks) };
  }

  function sig(cell) {
    if (cell._sig != null) return cell._sig;
    let base;
    if (cell.type === "markdown") {
      const d = document.createElement("div"); d.innerHTML = cell.html || "";
      base = (d.textContent || "").replace(/\s+/g, " ").trim();
    } else {
      base = cell.source.replace(/\s+/g, " ").trim();
    }
    return (cell._sig = cell.type + "\u0000" + base);
  }

  function build(a, b) {
    const ac = a.nbCells || [], bc = b.nbCells || [];
    const counter = { n: 0 };
    const rows = [];
    let del = 0, add = 0;

    if (typeof Diff === "undefined") {
      // no diff lib — just show both, no highlights
      const n = Math.max(ac.length, bc.length);
      for (let i = 0; i < n; i++) rows.push(buildRow(ac[i] || null, bc[i] || null, counter));
    } else {
      const parts = Diff.diffArrays(ac.map(sig), bc.map(sig));
      let ai = 0, bi = 0, i = 0;
      while (i < parts.length) {
        const part = parts[i];
        const n = part.value.length;
        if (!part.added && !part.removed) {
          for (let k = 0; k < n; k++) rows.push(buildRow(ac[ai + k], bc[bi + k], counter));
          ai += n; bi += n; i++;
        } else if (part.removed && parts[i + 1] && parts[i + 1].added) {
          const rem = n, addn = parts[i + 1].value.length, m = Math.min(rem, addn);
          for (let k = 0; k < m; k++) rows.push(buildRow(ac[ai + k], bc[bi + k], counter));
          for (let k = m; k < rem; k++) rows.push(buildRow(ac[ai + k], null, counter));
          for (let k = m; k < addn; k++) rows.push(buildRow(null, bc[bi + k], counter));
          ai += rem; bi += addn; i += 2;
        } else if (part.removed) {
          for (let k = 0; k < n; k++) rows.push(buildRow(ac[ai + k], null, counter));
          ai += n; i++;
        } else {
          for (let k = 0; k < n; k++) rows.push(buildRow(null, bc[bi + k], counter));
          bi += n; i++;
        }
      }
    }

    rows.forEach(r => { del += r.del; add += r.add; });
    return { rows, stats: { add, del }, changes: counter.n };
  }

  function render(container, state) {
    const data = state.nbData;
    if (!data) return;

    let html = `<div class="nbdiff">`;
    html += `<div class="nb-colhead">` +
      `<div class="nb-h"><span class="nb-badge a">A</span><span class="nb-hname">${esc(state.a.name)}</span><span class="nb-htag del">removed</span></div>` +
      `<div class="nb-h"><span class="nb-badge b">B</span><span class="nb-hname">${esc(state.b.name)}</span><span class="nb-htag add">added</span></div>` +
      `</div>`;

    if (data.stats.add + data.stats.del === 0) {
      html += `<div class="nb-identical"><div class="nb-id-ic">` +
        `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M20 6L9 17l-5-5"/></svg>` +
        `</div><h2>The notebooks look identical</h2><p>No differences detected across ${data.rows.length} cell${data.rows.length > 1 ? "s" : ""}.</p></div>`;
    }

    const list = data.rows.map((r, i) => ({ r, n: i + 1 })).filter(x => state.focus ? x.r.changed : true);
    if (state.focus && data.stats.add + data.stats.del > 0) {
      const hidden = data.rows.length - list.length;
      if (hidden > 0) {
        html += `<div class="nb-focusnote">Showing ${list.length} changed cell${list.length > 1 ? "s" : ""} · ${hidden} unchanged cell${hidden > 1 ? "s" : ""} hidden</div>`;
      }
    }

    html += `<div class="nb-cells">`;
    list.forEach(({ r, n }) => {
      html += `<div class="nb-row${r.changed ? " has-change" : ""}" data-cell="${n}">` +
        `<div class="nb-cellcol">${r.aHtml}</div>` +
        `<div class="nb-cellnum">${n}</div>` +
        `<div class="nb-cellcol">${r.bHtml}</div>` +
        `</div>`;
    });
    html += `</div></div>`;
    container.innerHTML = html;
  }

  window.NotebookVisual = { build, render };
})();
