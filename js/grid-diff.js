/* ============================================================
   grid-diff.js — structured spreadsheet diff (cell-level), robust to
   columns and rows being moved / inserted / deleted.

   Pipeline (per sheet pair):
     1. Detect a header row in each grid (independently).
     2. Align COLUMNS by header name → ordered "unified" column list, each
        tagged same / added / removed / moved. Falls back to positional
        alignment when headers aren't present on both sides.
     3. Detect a KEY column (values unique + filled on both sides). If found,
        match rows by key (order-independent); otherwise match by the content
        signature of the comparable columns.
     4. Present rows in the NEW file's order, classify each comparable cell,
        and flag genuinely relocated rows as "moved" (suppressed on a full
        re-sort to avoid noise).

   Depends on global `Diff`. Exposes window.GridDiff.
   ============================================================ */
(function () {
  function colName(i) {
    let s = ""; i++;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  function norm(v, iw) { const s = v == null ? "" : String(v); return iw ? s.trim().replace(/\s+/g, " ") : s; }
  function maxCols(g) { let m = 0; for (const r of g) if (r.length > m) m = r.length; return m; }

  /* ---- header detection ---- */
  const NUMERIC_RE = /^[-+]?[$£€]?\s*[\d.,]+\s*%?$/;
  function looksLikeHeader(grid) {
    if (!grid || grid.length < 2) return false;
    const r0 = (grid[0] || []).map(c => String(c == null ? "" : c).trim());
    if (!r0.length) return false;
    const nonEmpty = r0.filter(c => c !== "");
    if (nonEmpty.length < Math.max(1, Math.ceil(r0.length * 0.6))) return false;     // mostly filled
    const lower = nonEmpty.map(c => c.toLowerCase());
    if (new Set(lower).size !== nonEmpty.length) return false;                        // labels unique
    const numeric = nonEmpty.filter(c => NUMERIC_RE.test(c)).length;
    if (numeric > nonEmpty.length * 0.3) return false;                                // headers aren't numbers
    return true;
  }

  /* ---- column alignment ---- */
  function alignColumns(headerA, headerB, nA, nB) {
    if (headerA && headerB && typeof Diff !== "undefined") {
      const na = headerA.map(h => String(h == null ? "" : h).trim());
      const nb = headerB.map(h => String(h == null ? "" : h).trim());
      const ka = na.map(s => s.toLowerCase()), kb = nb.map(s => s.toLowerCase());
      const parts = Diff.diffArrays(ka, kb);
      const cols = []; let ai = 0, bi = 0;
      for (const p of parts) {
        const n = p.value.length;
        if (!p.added && !p.removed) { for (let k = 0; k < n; k++) { cols.push({ aIdx: ai, bIdx: bi, status: "same", name: nb[bi] }); ai++; bi++; } }
        else if (p.removed) { for (let k = 0; k < n; k++) { cols.push({ aIdx: ai, bIdx: null, status: "removed", name: na[ai] }); ai++; } }
        else { for (let k = 0; k < n; k++) { cols.push({ aIdx: null, bIdx: bi, status: "added", name: nb[bi] }); bi++; } }
      }
      // merge a removed + added pair that share a header name → a single "moved" column
      const added = cols.filter(c => c.status === "added");
      for (const rm of cols) {
        if (rm.status !== "removed") continue;
        const key = (rm.name || "").toLowerCase();
        if (!key) continue;
        const m = added.find(a => !a._used && (a.name || "").toLowerCase() === key);
        if (m) { m.status = "moved"; m.aIdx = rm.aIdx; m.fromName = colName(rm.aIdx); m._used = true; rm._drop = true; }
      }
      for (let i = cols.length - 1; i >= 0; i--) if (cols[i]._drop) cols.splice(i, 1);
      return cols;
    }
    // positional fallback (no usable headers)
    return positionalCols(nA, nB);
  }

  function positionalCols(nA, nB) {
    const n = Math.max(nA, nB), cols = [];
    for (let i = 0; i < n; i++) cols.push({ aIdx: i < nA ? i : null, bIdx: i < nB ? i : null, status: "same", name: null });
    return cols;
  }

  /* ---- content-based column alignment (no usable headers) ----
     Identify a column by the (order-independent) multiset of values down it,
     so a column that was moved to a new position is recognised as the SAME
     column rather than showing up as a wall of cell edits. Only overrides the
     positional layout when it confidently explains a real reordering; otherwise
     it returns the positional columns so unrelated data still degrades to a
     plain cell-level diff. */
  const COL_MATCH = 0.7;            // min Dice overlap to call two columns "the same"
  function columnValues(data, idx, iw) {
    const out = [];
    for (const r of data) { const v = norm(r[idx], iw); if (v !== "") out.push(v); }
    return out;
  }
  function multisetDice(a, b) {
    if (!a.length || !b.length) return 0;            // need real content on both sides
    const counts = new Map();
    for (const v of a) counts.set(v, (counts.get(v) || 0) + 1);
    let inter = 0;
    for (const v of b) { const c = counts.get(v) || 0; if (c > 0) { inter++; counts.set(v, c - 1); } }
    return (2 * inter) / (a.length + b.length);
  }
  function alignColumnsByContent(dataA, dataB, nA, nB, iw) {
    const positional = positionalCols(nA, nB);
    if (typeof Diff === "undefined") return positional;
    if (nA * nB > 4096 || nA === 0 || nB === 0) return positional;   // keep it cheap; nothing to match

    const sigA = [], sigB = [];
    for (let i = 0; i < nA; i++) sigA.push(columnValues(dataA, i, iw));
    for (let j = 0; j < nB; j++) sigB.push(columnValues(dataB, j, iw));

    // greedy best-similarity one-to-one matching of A↔B columns
    const cand = [];
    for (let i = 0; i < nA; i++) for (let j = 0; j < nB; j++) {
      const s = multisetDice(sigA[i], sigB[j]);
      if (s >= COL_MATCH) cand.push({ i, j, s });
    }
    cand.sort((x, y) => y.s - x.s);
    const aUsed = new Array(nA).fill(false), bUsed = new Array(nB).fill(false);
    const matchA = new Array(nB).fill(-1);           // B col j → A col i (or -1)
    for (const p of cand) { if (aUsed[p.i] || bUsed[p.j]) continue; aUsed[p.i] = true; bUsed[p.j] = true; matchA[p.j] = p.i; }
    // positional mop-up: pair any leftover columns that still sit at the same index
    for (let k = 0; k < Math.min(nA, nB); k++) if (!aUsed[k] && !bUsed[k]) { aUsed[k] = true; bUsed[k] = true; matchA[k] = k; }

    // same vs moved: longest increasing run of A-positions taken in B order stays put
    const seqJ = [], seqA = [];
    for (let j = 0; j < nB; j++) if (matchA[j] >= 0) { seqJ.push(j); seqA.push(matchA[j]); }
    const keep = lisMask(seqA);
    const inPlace = new Map();
    seqJ.forEach((j, k) => inPlace.set(j, keep[k]));
    let movedAny = false; inPlace.forEach(v => { if (!v) movedAny = true; });

    // only trust content alignment when it actually found a reordering AND it
    // explains most of the columns — otherwise the positional diff is safer.
    const matched = seqJ.length;
    if (!movedAny || matched < Math.max(2, Math.min(nA, nB) * 0.7)) return positional;

    const cols = [];
    for (let j = 0; j < nB; j++) {
      const i = matchA[j];
      if (i >= 0) cols.push({ aIdx: i, bIdx: j, status: inPlace.get(j) ? "same" : "moved", name: null, fromName: inPlace.get(j) ? undefined : colName(i) });
      else cols.push({ aIdx: null, bIdx: j, status: "added", name: null });
    }
    // slot removed A columns back in next to where they used to sit
    for (let i = 0; i < nA; i++) {
      if (aUsed[i]) continue;
      let pos = 0;
      for (let k = 0; k < cols.length; k++) if (cols[k].aIdx != null && cols[k].aIdx < i) pos = k + 1;
      cols.splice(pos, 0, { aIdx: i, bIdx: null, status: "removed", name: null });
    }
    return cols;
  }

  /* ---- key column detection (among columns present on both sides) ---- */
  function detectKey(dataA, dataB, cols, iw) {
    if (!dataA.length || !dataB.length) return null;
    const cands = cols.filter(c => c.aIdx != null && c.bIdx != null && (c.status === "same" || c.status === "moved"));
    let best = null, bestOverlap = 0;
    for (const c of cands) {
      const va = dataA.map(r => norm(r[c.aIdx], iw)).filter(x => x !== "");
      const vb = dataB.map(r => norm(r[c.bIdx], iw)).filter(x => x !== "");
      if (va.length < dataA.length * 0.95 || vb.length < dataB.length * 0.95) continue;   // well filled
      const ua = new Set(va), ub = new Set(vb);
      if (ua.size < va.length * 0.98 || ub.size < vb.length * 0.98) continue;             // ~unique
      let overlap = 0; for (const k of ua) if (ub.has(k)) overlap++;
      if (overlap > bestOverlap) { bestOverlap = overlap; best = c; }
    }
    // require meaningful overlap so we don't lock onto an unrelated unique column
    return bestOverlap >= Math.min(dataA.length, dataB.length) * 0.3 ? best : null;
  }

  function comparableCols(cols) { return cols.filter(c => c.status === "same" || c.status === "moved"); }
  function rowSig(row, cols, side, iw) {
    const out = [];
    for (const c of cols) { const idx = side === "a" ? c.aIdx : c.bIdx; out.push(norm(row[idx], iw)); }
    return out.join("\u0001");
  }

  function classifyCell(o, n, iw) {
    const a = norm(o, iw), b = norm(n, iw);
    if (a === b) return "same";
    if (a === "") return "add";
    if (b === "") return "del";
    return "mod";
  }
  /* Per-row cell classes over the unified columns. Added/removed COLUMNS get
     their own column-level class and do NOT count as a row change. */
  function classifyRow(aCells, bCells, cols, iw) {
    const cls = [], changed = [];
    cols.forEach((c, ci) => {
      if (c.status === "added") { cls[ci] = aCells ? "coladd" : "add"; return; }
      if (c.status === "removed") { cls[ci] = bCells ? "coldel" : "del"; return; }
      const o = aCells && c.aIdx != null ? aCells[c.aIdx] : "";
      const n = bCells && c.bIdx != null ? bCells[c.bIdx] : "";
      const s = classifyCell(o, n, iw);
      cls[ci] = s;
      if (s !== "same") changed.push(ci);
    });
    return { cls, changed };
  }

  /* longest strictly-increasing subsequence indices of arr (values = A positions) */
  function lisMask(arr) {
    const n = arr.length, keep = new Array(n).fill(false);
    if (!n) return keep;
    const tails = [], tailIdx = [], prev = new Array(n).fill(-1);
    for (let i = 0; i < n; i++) {
      let lo = 0, hi = tails.length;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (tails[mid] < arr[i]) lo = mid + 1; else hi = mid; }
      tails[lo] = arr[i]; tailIdx[lo] = i; prev[i] = lo > 0 ? tailIdx[lo - 1] : -1;
    }
    let k = tailIdx[tails.length - 1];
    while (k >= 0) { keep[k] = true; k = prev[k]; }
    return keep;
  }

  /* ---- row matching ---- */
  function makeRow(type, aRow, aNo, bRow, bNo) {
    const r = { type };
    if (aRow !== undefined) r.left = { cells: aRow, no: aNo };
    if (bRow !== undefined) r.right = { cells: bRow, no: bNo };
    return r;
  }

  function alignRows(dataA, dataB, cols, key, iw, baseA, baseB) {
    const rows = [];
    if (key && typeof Diff !== "undefined") {
      // ---- key-based: order-independent match, presented in B order ----
      const ka = dataA.map(r => norm(r[key.aIdx], iw));
      const aByKey = new Map();
      ka.forEach((k, i) => { if (!aByKey.has(k)) aByKey.set(k, []); aByKey.get(k).push(i); });
      const matchedA = new Array(dataA.length).fill(false);
      const bMatch = dataB.map(r => {
        const k = norm(r[key.bIdx], iw), q = aByKey.get(k);
        if (q && q.length) { const ai = q.shift(); matchedA[ai] = true; return ai; }
        return -1;
      });
      // sequence of A-positions of matched rows (B order) → detect relocations
      const seqA = bMatch.filter(ai => ai >= 0);
      const inPlace = lisMask(seqA);
      let movedCount = 0; inPlace.forEach(v => { if (!v) movedCount++; });
      const allowMoves = seqA.length > 0 && movedCount <= Math.max(3, seqA.length * 0.3);
      let aPtr = 0, seqPtr = 0, moveId = 0;
      const emitRemovedUpTo = (lim) => { while (aPtr < lim) { if (!matchedA[aPtr]) rows.push(makeRow("del", dataA[aPtr], aPtr + baseA, undefined)); aPtr++; } };
      for (let j = 0; j < dataB.length; j++) {
        const ai = bMatch[j];
        if (ai >= 0) {
          emitRemovedUpTo(ai); aPtr = Math.max(aPtr, ai + 1);
          const moved = allowMoves && !inPlace[seqPtr];
          seqPtr++;
          const row = makeRow(moved ? "move" : "eq", dataA[ai], ai + baseA, dataB[j], j + baseB);
          if (moved) row.moveId = ++moveId;
          rows.push(row);
        } else rows.push(makeRow("add", undefined, undefined, dataB[j], j + baseB));
      }
      emitRemovedUpTo(dataA.length);
    } else if (typeof Diff !== "undefined") {
      // ---- content-based: signature diff + adjacent mod-pairing + exact moves ----
      const cc = comparableCols(cols);
      const sigA = dataA.map(r => rowSig(r, cc, "a", iw));
      const sigB = dataB.map(r => rowSig(r, cc, "b", iw));
      const parts = Diff.diffArrays(sigA, sigB);
      let ai = 0, bi = 0, pd = [], pa = [];
      const flush = () => {
        const m = Math.min(pd.length, pa.length);
        for (let k = 0; k < m; k++) rows.push(makeRow("mod", pd[k].cells, pd[k].no, pa[k].cells, pa[k].no));
        for (let k = m; k < pd.length; k++) rows.push(makeRow("del", pd[k].cells, pd[k].no, undefined));
        for (let k = m; k < pa.length; k++) rows.push(makeRow("add", undefined, undefined, pa[k].cells, pa[k].no));
        pd = []; pa = [];
      };
      for (const p of parts) {
        const n = p.value.length;
        if (p.added) { for (let k = 0; k < n; k++) { pa.push({ cells: dataB[bi], no: bi + baseB }); bi++; } }
        else if (p.removed) { for (let k = 0; k < n; k++) { pd.push({ cells: dataA[ai], no: ai + baseA }); ai++; } }
        else { flush(); for (let k = 0; k < n; k++) { rows.push(makeRow("eq", dataA[ai], ai + baseA, dataB[bi], bi + baseB)); ai++; bi++; } }
      }
      flush();
      detectExactMoves(rows, cc, iw);
    } else {
      const n = Math.max(dataA.length, dataB.length);
      for (let i = 0; i < n; i++) rows.push(makeRow(i < dataA.length && i < dataB.length ? "eq" : (i < dataA.length ? "del" : "add"),
        dataA[i], i + baseA, dataB[i], i + baseB));
    }
    return rows;
  }

  /* Reconnect a deleted row and an added row that are identical (relocated).
     Only for content matching; capped so a re-sort doesn't turn into noise. */
  function detectExactMoves(rows, cc, iw) {
    const dels = [], adds = [];
    rows.forEach((r, i) => { if (r.type === "del") dels.push(i); else if (r.type === "add") adds.push(i); });
    if (!dels.length || !adds.length) return;
    const sigOf = (cells, side) => rowSig(cells || [], cc, side, iw);
    const addBySig = new Map();
    adds.forEach(i => { const s = sigOf(rows[i].right.cells, "b"); if (!addBySig.has(s)) addBySig.set(s, []); addBySig.get(s).push(i); });
    const pairs = [];
    for (const di of dels) {
      const s = sigOf(rows[di].left.cells, "a");
      const q = addBySig.get(s);
      if (q && q.length) pairs.push([di, q.shift()]);
    }
    if (pairs.length > Math.max(5, (dels.length + adds.length) * 0.3)) return;   // looks like a re-sort
    let moveId = 0;
    for (const [di, aj] of pairs) {
      const dr = rows[di], ar = rows[aj];
      dr.type = "move"; dr.right = ar.right; dr.moveId = ++moveId; dr._mergedAdd = aj;
    }
    // drop the now-merged add rows
    const drop = new Set(pairs.map(p => p[1]));
    for (let i = rows.length - 1; i >= 0; i--) if (drop.has(i)) rows.splice(i, 1);
  }

  /* ---- sheet pairing ---- */
  function pairSheets(A, B) {
    if (A.length === 1 && B.length === 1) {
      const name = A[0].name === B[0].name ? A[0].name : `${A[0].name} → ${B[0].name}`;
      return [{ name, gA: A[0].grid, gB: B[0].grid }];
    }
    const mapA = new Map(A.map(s => [s.name, s.grid]));
    const mapB = new Map(B.map(s => [s.name, s.grid]));
    const names = [];
    A.forEach(s => names.push(s.name));
    B.forEach(s => { if (!mapA.has(s.name)) names.push(s.name); });
    return names.map(n => ({ name: n, gA: mapA.get(n) || [], gB: mapB.get(n) || [] }));
  }

  function diffSheet(gA, gB, name, opts) {
    const iw = opts.ignoreWs;
    const hA = looksLikeHeader(gA), hB = looksLikeHeader(gB);
    const useHeaders = hA && hB;
    const headerA = useHeaders ? gA[0] : null, headerB = useHeaders ? gB[0] : null;
    const dataA = useHeaders ? gA.slice(1) : gA;
    const dataB = useHeaders ? gB.slice(1) : gB;
    const baseA = useHeaders ? 2 : 1, baseB = useHeaders ? 2 : 1;     // spreadsheet row numbers

    const cols = useHeaders
      ? alignColumns(headerA, headerB, maxCols(gA), maxCols(gB))
      : alignColumnsByContent(dataA, dataB, maxCols(dataA), maxCols(dataB), iw);
    const key = detectKey(dataA, dataB, cols, iw);
    const rows = alignRows(dataA, dataB, cols, key, iw, baseA, baseB);

    // classify changed rows + collect per-row changed columns
    const stats = { addCells: 0, delCells: 0, modCells: 0, addedRows: 0, removedRows: 0, movedRows: 0, modifiedRows: 0 };
    for (const r of rows) {
      if (r.type === "eq") {
        // key-matched "eq" rows may still differ in non-key cells
        if (key) {
          const { cls, changed } = classifyRow(r.left.cells, r.right.cells, cols, iw);
          if (changed.length) {
            r.type = "mod"; r.cls = cls; r.changed = changed;
            stats.modifiedRows++;
            for (const ci of changed) {
              const s = cls[ci];
              if (s === "add") stats.addCells++; else if (s === "del") stats.delCells++; else if (s === "mod") stats.modCells++;
            }
          }
        }
        continue;
      }
      const aCells = (r.type === "add") ? null : r.left.cells;
      const bCells = (r.type === "del") ? null : r.right.cells;
      const { cls, changed } = classifyRow(aCells, bCells, cols, iw);
      r.cls = cls; r.changed = changed;
      for (const ci of changed) {
        const s = cls[ci];
        if (s === "add") stats.addCells++; else if (s === "del") stats.delCells++; else if (s === "mod") stats.modCells++;
      }
      if (r.type === "add") stats.addedRows++;
      else if (r.type === "del") stats.removedRows++;
      else if (r.type === "move") { stats.movedRows++; if (changed.length) stats.modifiedRows++; }
      else if (r.type === "mod") stats.modifiedRows++;
    }

    const colStats = { added: 0, removed: 0, moved: 0 };
    cols.forEach(c => { if (c.status === "added") colStats.added++; else if (c.status === "removed") colStats.removed++; else if (c.status === "moved") colStats.moved++; });

    return { name, rows, cols, columns: cols, useHeaders, key: key ? (key.name || colName(key.bIdx)) : null, stats, colStats };
  }

  function buildGridDiff(sheetsA, sheetsB, opts) {
    opts = opts || {};
    let pairs;
    if (opts.selA != null && opts.selB != null) {
      const a = sheetsA.find(s => s.name === opts.selA) || sheetsA[0];
      const b = sheetsB.find(s => s.name === opts.selB) || sheetsB[0];
      const an = a ? a.name : "?", bn = b ? b.name : "?";
      pairs = [{ name: an === bn ? an : `${an} → ${bn}`, gA: (a && a.grid) || [], gB: (b && b.grid) || [] }];
    } else {
      pairs = pairSheets(sheetsA, sheetsB);
    }
    let addCells = 0, delCells = 0, modCells = 0;
    let addedRows = 0, removedRows = 0, movedRows = 0, modifiedRows = 0, addedCols = 0, removedCols = 0, movedCols = 0;
    const sheets = pairs.map(p => {
      const s = diffSheet(p.gA, p.gB, p.name, opts);
      addCells += s.stats.addCells; delCells += s.stats.delCells; modCells += s.stats.modCells;
      addedRows += s.stats.addedRows; removedRows += s.stats.removedRows; movedRows += s.stats.movedRows; modifiedRows += s.stats.modifiedRows;
      addedCols += s.colStats.added; removedCols += s.colStats.removed; movedCols += s.colStats.moved;
      // expose a flat width for renderers that still want a column count
      s.cols = s.columns.length;
      return s;
    });
    const changed = sheets.some(s => s.rows.some(r => r.type !== "eq")) || addedCols + removedCols + movedCols > 0;
    return {
      sheets,
      stats: { addCells, delCells, modCells, addedRows, removedRows, movedRows, modifiedRows, addedCols, removedCols, movedCols, changed },
      colName
    };
  }

  window.GridDiff = { buildGridDiff, colName };
})();
