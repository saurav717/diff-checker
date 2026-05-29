/* ============================================================
   grid-diff.js — structured spreadsheet diff (cell-level).
   Depends on global `Diff`. Exposes window.GridDiff
   ============================================================ */
(function () {
  function colName(i) {
    let s = "";
    i++;
    while (i > 0) {
      const m = (i - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  function norm(v, iw) {
    const s = v == null ? "" : String(v);
    return iw ? s.trim().replace(/\s+/g, " ") : s;
  }
  function sig(row, iw) { return row.map(c => norm(c, iw)).join("\u0001"); }
  function maxCols(g) { let m = 0; for (const r of g) if (r.length > m) m = r.length; return m; }

  /* Match sheets between the two workbooks. */
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

  /* Align rows of two grids -> [{type, left:{cells,no}, right:{cells,no}, changed?:Set}] */
  function alignGrid(A, B, opts) {
    const iw = opts.ignoreWs;
    const sigA = A.map(r => sig(r, iw));
    const sigB = B.map(r => sig(r, iw));
    const parts = Diff.diffArrays(sigA, sigB);

    const rows = [];
    let ai = 0, bi = 0, pd = [], pa = [];
    function flush() {
      const m = Math.min(pd.length, pa.length);
      for (let k = 0; k < m; k++) rows.push({ type: "mod", left: pd[k], right: pa[k] });
      for (let k = m; k < pd.length; k++) rows.push({ type: "del", left: pd[k] });
      for (let k = m; k < pa.length; k++) rows.push({ type: "add", right: pa[k] });
      pd = []; pa = [];
    }
    for (const part of parts) {
      const n = part.count != null ? part.count : part.value.length;
      if (part.added) { for (let k = 0; k < n; k++) { pa.push({ cells: B[bi], no: bi + 1 }); bi++; } }
      else if (part.removed) { for (let k = 0; k < n; k++) { pd.push({ cells: A[ai], no: ai + 1 }); ai++; } }
      else { flush(); for (let k = 0; k < n; k++) { rows.push({ type: "eq", left: { cells: A[ai], no: ai + 1 }, right: { cells: B[bi], no: bi + 1 } }); ai++; bi++; } }
    }
    flush();
    return rows;
  }

  function classify(oldV, newV, iw) {
    const o = norm(oldV, iw), n = norm(newV, iw);
    if (o === n) return "same";
    if (o === "") return "add";   // empty -> value
    if (n === "") return "del";   // value -> empty
    return "mod";                 // value -> different value
  }

  function buildGridDiff(sheetsA, sheetsB, opts) {
    const pairs = pairSheets(sheetsA, sheetsB);
    let addCells = 0, delCells = 0, modCells = 0;
    const sheets = pairs.map(p => {
      const rows = alignGrid(p.gA, p.gB, opts);
      const cols = Math.max(maxCols(p.gA), maxCols(p.gB), 1);
      for (const r of rows) {
        if (r.type === "eq") continue;
        // Classify EVERY cell independently so one row can mix add/del/mod.
        const oldC = r.type === "add" ? [] : r.left.cells;
        const newC = r.type === "del" ? [] : r.right.cells;
        const n = Math.max(oldC.length, newC.length, cols);
        const cls = [];
        for (let c = 0; c < n; c++) {
          const s = classify(oldC[c], newC[c], opts.ignoreWs);
          cls[c] = s;
          if (s === "add") addCells++;
          else if (s === "del") delCells++;
          else if (s === "mod") modCells++;
        }
        r.cls = cls;
      }
      return { name: p.name, rows, cols };
    });
    return {
      sheets,
      stats: { addCells, delCells, modCells, changed: sheets.some(s => s.rows.some(r => r.type !== "eq")) },
      colName
    };
  }

  window.GridDiff = { buildGridDiff, colName };
})();
