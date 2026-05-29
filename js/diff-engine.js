/* ============================================================
   diff-engine.js — line alignment + intra-line diff + folding
   Depends on global `Diff` (jsdiff).
   Exposes window.DiffEngine
   ============================================================ */
(function () {
  function norm(line, ignoreWs) {
    if (!ignoreWs) return line;
    return line.replace(/\s+/g, " ").trim();
  }

  /* Build aligned rows from two texts.
     Row shapes:
       {type:'eq',  left:{no,raw}, right:{no,raw}}
       {type:'mod', left:{no,raw}, right:{no,raw}}
       {type:'del', left:{no,raw}}
       {type:'add', right:{no,raw}}
  */
  function buildRows(textA, textB, opts) {
    const ignoreWs = !!opts.ignoreWs;
    const A = textA.split("\n");
    const B = textB.split("\n");
    const nA = A.map(l => norm(l, ignoreWs));
    const nB = B.map(l => norm(l, ignoreWs));

    const parts = Diff.diffArrays(nA, nB);

    const rows = [];
    let ai = 0, bi = 0;
    let pendDel = [], pendAdd = [];

    function flush() {
      const m = Math.min(pendDel.length, pendAdd.length);
      for (let k = 0; k < m; k++) {
        rows.push({ type: "mod", left: pendDel[k], right: pendAdd[k] });
      }
      for (let k = m; k < pendDel.length; k++) rows.push({ type: "del", left: pendDel[k] });
      for (let k = m; k < pendAdd.length; k++) rows.push({ type: "add", right: pendAdd[k] });
      pendDel = []; pendAdd = [];
    }

    for (const part of parts) {
      const count = part.count != null ? part.count : part.value.length;
      if (part.added) {
        for (let k = 0; k < count; k++) { pendAdd.push({ raw: B[bi] }); bi++; }
      } else if (part.removed) {
        for (let k = 0; k < count; k++) { pendDel.push({ raw: A[ai] }); ai++; }
      } else {
        flush();
        for (let k = 0; k < count; k++) {
          rows.push({ type: "eq", left: { raw: A[ai] }, right: { raw: B[bi] } });
          ai++; bi++;
        }
      }
    }
    flush();

    // Assign line numbers
    let lno = 0, rno = 0;
    for (const r of rows) {
      if (r.type === "eq" || r.type === "mod") { r.left.no = ++lno; r.right.no = ++rno; }
      else if (r.type === "del") { r.left.no = ++lno; }
      else if (r.type === "add") { r.right.no = ++rno; }
    }
    return rows;
  }

  /* Intra-line diff -> {leftHtml, rightHtml} with <mark> spans. */
  function intraLine(a, b, granularity) {
    const parts = granularity === "char"
      ? Diff.diffChars(a, b)
      : Diff.diffWordsWithSpace(a, b);
    let L = "", R = "";
    for (const p of parts) {
      const t = esc(p.value);
      if (p.added) R += `<mark class="w-add">${t}</mark>`;
      else if (p.removed) L += `<mark class="w-del">${t}</mark>`;
      else { L += t; R += t; }
    }
    return { leftHtml: L, rightHtml: R };
  }

  function esc(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  /* Group consecutive change rows into hunks (for navigation). */
  function computeHunks(rows) {
    const hunks = [];
    let start = -1;
    for (let i = 0; i < rows.length; i++) {
      const changed = rows[i].type !== "eq";
      if (changed && start === -1) start = i;
      if (!changed && start !== -1) { hunks.push({ start, end: i - 1 }); start = -1; }
    }
    if (start !== -1) hunks.push({ start, end: rows.length - 1 });
    return hunks;
  }

  /* Visible list for focus mode. forced = Set of row indices to force-show.
     Returns array of {type:'row', row, i} | {type:'fold', from, to, count} */
  function computeVisible(rows, focus, forced, ctx) {
    if (!focus) return rows.map((row, i) => ({ type: "row", row, i }));
    ctx = ctx == null ? 3 : ctx;
    const keep = new Array(rows.length).fill(false);
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].type !== "eq") {
        for (let j = i - ctx; j <= i + ctx; j++) if (j >= 0 && j < rows.length) keep[j] = true;
      }
    }
    if (forced) forced.forEach(i => { if (i >= 0 && i < rows.length) keep[i] = true; });

    const out = [];
    let i = 0;
    while (i < rows.length) {
      if (keep[i]) { out.push({ type: "row", row: rows[i], i }); i++; }
      else {
        let j = i;
        while (j < rows.length && !keep[j]) j++;
        out.push({ type: "fold", from: i, to: j - 1, count: j - i });
        i = j;
      }
    }
    return out;
  }

  function stats(rows) {
    let add = 0, del = 0;
    for (const r of rows) {
      if (r.type === "add") add++;
      else if (r.type === "del") del++;
      else if (r.type === "mod") { add++; del++; }
    }
    return { add, del, changed: rows.some(r => r.type !== "eq") };
  }

  /* Unified diff text for export (uses jsdiff). */
  function unified(nameA, nameB, textA, textB) {
    try {
      return Diff.createTwoFilesPatch(nameA, nameB, textA, textB, "", "", { context: 3 });
    } catch (e) {
      return `--- ${nameA}\n+++ ${nameB}\n(diff unavailable)`;
    }
  }

  window.DiffEngine = { buildRows, intraLine, computeHunks, computeVisible, stats, unified, esc };
})();
