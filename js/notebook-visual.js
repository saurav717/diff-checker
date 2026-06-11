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

  // Highlight granularity: "word" (default) or "sentence" (extend a change to
  // the whole sentence / line that contains it). Set per build().
  let gran = "word";

  /* Segment id per token: a new segment begins after a newline in the gap
     between tokens, or after a token ending in sentence punctuation. So in
     prose a unit is a sentence; in code it is effectively a line. */
  function segIds(text, toks) {
    const ids = []; let s = 0;
    for (let i = 0; i < toks.length; i++) {
      if (i > 0) {
        const gap = text.slice(toks[i - 1].end, toks[i].start);
        const prevWord = toks[i - 1].s != null ? toks[i - 1].s : toks[i - 1].word;
        if (gap.indexOf("\n") >= 0 || /[.!?:][)"'\]]?$/.test(prevWord)) s++;
      }
      ids.push(s);
    }
    return ids;
  }
  /* Sentence mode: extend a per-token change map to cover every token of any
     segment that contains a change. */
  function expandToSeg(markMap, ids) {
    if (gran !== "sentence" || !markMap.size) return;
    const seg = new Map();
    for (const [i, v] of markMap) if (!seg.has(ids[i])) seg.set(ids[i], v);
    for (let i = 0; i < ids.length; i++) if (seg.has(ids[i])) markMap.set(i, seg.get(ids[i]));
  }

  // ---- syntax highlighting (highlight.js) ----
  // Map our language ids to highlight.js ids.
  const HLJS_LANG = { python: "python", py: "python", sql: "sql", scala: "scala",
                      r: "r", bash: "bash", sh: "bash", shell: "bash", json: "json",
                      javascript: "javascript", js: "javascript", java: "java", plaintext: null };

  /* Return syntax-highlighted HTML for `source`, falling back to escaped text
     when highlight.js (or the language) isn't available. Whitespace/newlines
     are preserved verbatim, so the word-diff over the visible text is identical
     to diffing the raw source. */
  function highlightCode(source, lang) {
    const src = source || "";
    if (typeof hljs === "undefined") return esc(src);
    const id = HLJS_LANG[(lang || "").toLowerCase()];
    try {
      if (id && hljs.getLanguage(id)) return hljs.highlight(src, { language: id, ignoreIllegals: true }).value;
      return hljs.highlightAuto(src).value;   // unknown language → let hljs guess
    } catch (_) { return esc(src); }
  }

  /* Highlight one or both code sources, then overlay the word-level diff marks
     on the highlighted HTML (so colour + change highlights coexist). */
  function markCodePair(aSource, bSource, lang, counter) {
    const aHi = aSource != null ? highlightCode(aSource, lang) : "";
    const bHi = bSource != null ? highlightCode(bSource, lang) : "";
    return markHtmlPair(aHi, bHi, counter);
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
    expandToSeg(aMark, segIds(aText || "", at));
    expandToSeg(bMark, segIds(bText || "", bt));
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
    // Sentence mode: expand using segment ids derived from the plain text
    // (token order matches the node tokens 1:1).
    if (gran === "sentence") {
      const aText = A.root.textContent || "", bText = B.root.textContent || "";
      expandToSeg(aMark, segIds(aText, tokenize(aText)));
      expandToSeg(bMark, segIds(bText, tokenize(bText)));
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
  // `move` (optional): { role:'from'|'to', partnerRow, id } marks this cell as
  // a relocated block — it gets the move accent, a badge linking to where it
  // came from / went to, and a shared data-chg so hovering links both halves.
  function cellHtml(side, cell, body, outputsHtml, lone, mapId, move) {
    const mapAttr = mapId != null ? ` data-map="nb${mapId}"` : "";
    if (!cell) return `<div class="nb-cell nb-empty"${mapAttr}><span>no matching cell</span></div>`;
    const loneCls = lone ? (side === "a" ? " nb-lone-del" : " nb-lone-add") : "";
    let moveCls = "", moveAttr = "", badge = "";
    if (move) {
      moveCls = " nb-moved nb-moved-" + move.role;
      moveAttr = ` data-chg="mv${move.id}"`;
      const arrow = move.role === "from" ? "\u2193" : "\u2191";
      const label = move.role === "from"
        ? `Moved to cell ${move.partnerRow}`
        : `Moved from cell ${move.partnerRow}`;
      badge = `<div class="nb-movebadge"><span class="nb-move-ic">\u21c5</span>${esc(label)} <span class="nb-move-arrow">${arrow}</span></div>`;
    }
    const title = cell.title ? `<div class="nb-celltitle">${esc(cell.title)}</div>` : "";
    if (cell.type === "markdown") {
      return `<div class="nb-cell nb-md${loneCls}${moveCls}"${mapAttr}${moveAttr}>${badge}${title}${body}</div>`;
    }
    return `<div class="nb-cell nb-code${loneCls}${moveCls}"${mapAttr}${moveAttr}>` +
      badge +
      title +
      `<div class="nb-codewrap-outer">` +
      `<div class="nb-prompt">${side === "a" ? "[ ]" : "[ ]"}</div>` +
      `<div class="nb-code-wrap"><pre class="nb-src">${body}</pre>${outputsHtml || ""}</div>` +
      `</div>` +
    `</div>`;
  }

  // Placeholder shown on the opposite column from a moved cell, pointing at the
  // cell's other location so the relocation reads clearly on both sides.
  function moveStub(role, partnerRow, moveId) {
    const arrow = role === "to" ? "\u2193" : "\u2191";
    const label = role === "to"
      ? `moved down to cell ${partnerRow}`
      : `moved up from cell ${partnerRow}`;
    return `<div class="nb-cell nb-movestub" data-chg="mv${moveId}">` +
      `<span class="nb-move-ic">\u21c5</span><span>${esc(label)}</span>` +
      `<span class="nb-move-arrow">${arrow}</span></div>`;
  }

  // Compute the diffed A/B body + outputs for a matched pair, sharing one
  // change counter so the marks on both sides carry identical data-chg ids.
  function pairBodies(aCell, bCell, counter) {
    let aBody = "", bBody = "", del = 0, add = 0;
    const refs = { del: 0, add: 0, counted: false, textPair: null };
    if (aCell.type === "markdown" && bCell.type === "markdown") {
      const mp = markHtmlPair(aCell.html, bCell.html, counter);
      aBody = mp.aHtml; bBody = mp.bHtml; del += mp.del; add += mp.add;
    } else {
      const cp = markCodePair(aCell.source, bCell.source, bCell.lang || aCell.lang, counter);
      aBody = cp.aHtml; bBody = cp.bHtml; del += cp.del; add += cp.add;
    }
    const aOut = renderOutputs("a", aCell, bCell, counter, refs);
    const bOut = renderOutputs("b", bCell, aCell, counter, refs);
    del += refs.del; add += refs.add;
    return { aBody, bBody, aOut, bOut, del, add };
  }

  // Build full A/B HTML for a paired (or lone) row.
  function buildRow(aCell, bCell, counter, mapId) {
    if (aCell && bCell) {
      const pb = pairBodies(aCell, bCell, counter);
      const aHtml = cellHtml("a", aCell, pb.aBody, pb.aOut, false, mapId);
      const bHtml = cellHtml("b", bCell, pb.bBody, pb.bOut, false, mapId);
      const changed = pb.del + pb.add > 0;
      return { aHtml, bHtml, changed, del: pb.del, add: pb.add };
    }

    if (aCell) { // removed cell
      const body = aCell.type === "markdown"
        ? markHtmlPair(aCell.html, "", counter).aHtml
        : markCodePair(aCell.source, null, aCell.lang, counter).aHtml;
      const out = renderOutputs("a", aCell, null, counter, { del: 0, add: 0, counted: true, textPair: null });
      const toks = tokenize(aCell.source).length + tokenize((aCell.html || "").replace(/<[^>]+>/g, " ")).length;
      return { aHtml: cellHtml("a", aCell, body, out, true, mapId), bHtml: cellHtml("b", null, "", "", false, mapId), changed: true, del: Math.max(1, toks), add: 0 };
    }

    // added cell
    const body = bCell.type === "markdown"
      ? markHtmlPair("", bCell.html, counter).bHtml
      : markCodePair(null, bCell.source, bCell.lang, counter).bHtml;
    const out = renderOutputs("b", bCell, null, counter, { del: 0, add: 0, counted: true, textPair: null });
    const toks = tokenize(bCell.source).length + tokenize((bCell.html || "").replace(/<[^>]+>/g, " ")).length;
    return { aHtml: cellHtml("a", null, "", "", false, mapId), bHtml: cellHtml("b", bCell, body, out, true, mapId), changed: true, del: 0, add: Math.max(1, toks) };
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

  // ---- move (relocated cell/block) detection ----
  // Comparable plain text for similarity scoring (markdown -> rendered text).
  function cmpText(cell) {
    if (cell.type === "markdown") {
      const d = document.createElement("div"); d.innerHTML = cell.html || "";
      return (d.textContent || "").replace(/\s+/g, " ").trim();
    }
    return (cell.source || "").replace(/\s+/g, " ").trim();
  }
  // Token Dice coefficient: 2*common / (lenA+lenB). 1 = identical content.
  function similarity(aCell, bCell) {
    if (!aCell || !bCell || aCell.type !== bCell.type) return 0;
    const a = cmpText(aCell), b = cmpText(bCell);
    if (a === b) return 1;
    if (typeof Diff === "undefined") return 0;
    const at = a ? a.split(/\s+/) : [], bt = b ? b.split(/\s+/) : [];
    const denom = at.length + bt.length;
    if (!denom) return 1;
    let common = 0;
    for (const p of Diff.diffArrays(at, bt)) if (!p.added && !p.removed) common += p.value.length;
    return (2 * common) / denom;
  }
  /* Reconnect unpaired removed (A-only) and added (B-only) ops that are really
     the same block relocated up or down. A pairing counts as a MOVE only when
     the two ops land more than one row apart — adjacent near-matches stay plain
     add/remove so ordinary in-place edits are untouched. Most-similar pairs win,
     ties broken by proximity; each cell is matched at most once. Mutates ops
     (attaches .move) and returns the number of moves found. */
  function detectMoves(ops) {
    if (typeof Diff === "undefined") return 0;
    const dels = [], adds = [];
    ops.forEach((op, i) => {
      if (op.a && !op.b) dels.push({ op, i });
      else if (!op.a && op.b) adds.push({ op, i });
    });
    if (!dels.length || !adds.length) return 0;
    const THRESH = 0.5;            // min similarity to treat as the same block
    const cands = [];
    for (const d of dels) for (const a of adds) {
      const dist = Math.abs(d.i - a.i);
      if (dist < 2) continue;      // adjacent -> ordinary edit, not a relocation
      const s = similarity(d.op.a, a.op.b);
      if (s >= THRESH) cands.push({ d, a, s, dist });
    }
    cands.sort((x, y) => (y.s - x.s) || (x.dist - y.dist));
    const usedD = new Set(), usedA = new Set();
    let moves = 0, id = 0;
    for (const c of cands) {
      if (usedD.has(c.d.op) || usedA.has(c.a.op)) continue;
      usedD.add(c.d.op); usedA.add(c.a.op);
      id++;
      c.d.op.move = { role: "from", partner: c.a.op, id };
      c.a.op.move = { role: "to", partner: c.d.op, id };
      moves++;
    }
    return moves;
  }
  /* Build the two rows for a relocated cell. The A side keeps the content with
     a "moved to" badge (opposite a stub pointing forward); the B side shows the
     content with a "moved from" badge (opposite a stub pointing back). Any
     within-cell text/output edits are still diffed and counted; the relocation
     itself is reported via the `moved` stat, not as add/remove churn. */
  function buildMoveRows(delOp, addOp, counter) {
    const aCell = delOp.a, bCell = addOp.b;
    const pb = pairBodies(aCell, bCell, counter);
    const id = delOp.move.id;
    const aHtml = cellHtml("a", aCell, pb.aBody, pb.aOut, false, null, { role: "from", partnerRow: addOp.row, id });
    const bHtml = cellHtml("b", bCell, pb.bBody, pb.bOut, false, null, { role: "to", partnerRow: delOp.row, id });
    const fromRow = { aHtml, bHtml: moveStub("to", addOp.row, id), changed: true, del: pb.del, add: pb.add, moved: true, moveId: id };
    const toRow = { aHtml: moveStub("from", delOp.row, id), bHtml, changed: true, del: 0, add: 0, moved: true, moveId: id };
    return { fromRow, toRow };
  }

  function build(a, b, granularity) {
    gran = granularity === "sentence" ? "sentence" : "word";
    const ac = a.nbCells || [], bc = b.nbCells || [];
    const counter = { n: 0 };

    // 1. Alignment pass -> ordered list of ops { a, b } (either may be null).
    const ops = [];
    if (typeof Diff === "undefined") {
      const n = Math.max(ac.length, bc.length);
      for (let i = 0; i < n; i++) ops.push({ a: ac[i] || null, b: bc[i] || null });
    } else {
      const parts = Diff.diffArrays(ac.map(sig), bc.map(sig));
      let ai = 0, bi = 0, i = 0;
      while (i < parts.length) {
        const part = parts[i];
        const n = part.value.length;
        if (!part.added && !part.removed) {
          for (let k = 0; k < n; k++) ops.push({ a: ac[ai + k], b: bc[bi + k] });
          ai += n; bi += n; i++;
        } else if (part.removed && parts[i + 1] && parts[i + 1].added) {
          const rem = n, addn = parts[i + 1].value.length, m = Math.min(rem, addn);
          for (let k = 0; k < m; k++) ops.push({ a: ac[ai + k], b: bc[bi + k] });
          for (let k = m; k < rem; k++) ops.push({ a: ac[ai + k], b: null });
          for (let k = m; k < addn; k++) ops.push({ a: null, b: bc[bi + k] });
          ai += rem; bi += addn; i += 2;
        } else if (part.removed) {
          for (let k = 0; k < n; k++) ops.push({ a: ac[ai + k], b: null });
          ai += n; i++;
        } else {
          for (let k = 0; k < n; k++) ops.push({ a: null, b: bc[bi + k] });
          bi += n; i++;
        }
      }
    }

    // 2. Reconnect relocated blocks among the unpaired add/remove ops.
    const moved = detectMoves(ops);

    // 3. Number rows (one per op, in order) so move badges can reference the
    //    partner's row, then materialise each row.
    ops.forEach((op, i) => { op.row = i + 1; });
    for (const op of ops) {
      if (op.move && op.move.role === "from") {
        const built = buildMoveRows(op, op.move.partner, counter);
        op.rowObj = built.fromRow;
        op.move.partner.rowObj = built.toRow;
      }
    }
    const rid = { n: 0 };
    const rows = ops.map(op => op.rowObj || buildRow(op.a, op.b, counter, ++rid.n));

    let del = 0, add = 0;
    rows.forEach(r => { del += r.del; add += r.add; });
    return { rows, stats: { add, del, moved }, changes: counter.n };
  }

  function render(container, state) {
    const data = state.nbData;
    if (!data) return;

    let html = `<div class="nbdiff">`;
    html += `<div class="nb-colhead">` +
      `<div class="nb-h"><span class="nb-badge a">A</span><span class="nb-hname">${esc(state.a.name)}</span><span class="nb-htag del">removed</span></div>` +
      `<div class="nb-h"><span class="nb-badge b">B</span><span class="nb-hname">${esc(state.b.name)}</span><span class="nb-htag add">added</span></div>` +
      `</div>`;

    if (data.stats.add + data.stats.del + (data.stats.moved || 0) === 0) {
      html += `<div class="nb-identical"><div class="nb-id-ic">` +
        `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M20 6L9 17l-5-5"/></svg>` +
        `</div><h2>The notebooks look identical</h2><p>No differences detected across ${data.rows.length} cell${data.rows.length > 1 ? "s" : ""}.</p></div>`;
    }

    const list = data.rows.map((r, i) => ({ r, n: i + 1 })).filter(x => state.focus ? x.r.changed : true);
    if (state.focus && data.stats.add + data.stats.del + (data.stats.moved || 0) > 0) {
      const hidden = data.rows.length - list.length;
      if (hidden > 0) {
        html += `<div class="nb-focusnote">Showing ${list.length} changed cell${list.length > 1 ? "s" : ""} · ${hidden} unchanged cell${hidden > 1 ? "s" : ""} hidden</div>`;
      }
    }

    html += `<div class="nb-cells">`;
    list.forEach(({ r, n }) => {
      const moveCls = r.moved ? " nb-row-moved" : "";
      const moveAttr = r.moveId ? ` data-moverow="${r.moveId}"` : "";
      html += `<div class="nb-row${r.changed ? " has-change" : ""}${moveCls}" data-cell="${n}"${moveAttr}>` +
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
