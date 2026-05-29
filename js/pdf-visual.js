/* ============================================================
   pdf-visual.js — show two PDFs side by side and highlight the
   regions that changed, page by page.

   All pdf.js work (rendering pages + extracting word boxes) is done
   once at parse time (see parsers.js) and cached on each file as
   `pdfPages: [{url, vw, vh, words}]`. This module is pure + sync: it
   diffs the cached word sequences per page and assembles highlight
   rectangles over removed words (left / A) and added words (right / B).

   Exposes: window.PdfVisual.build(a, b)  -> { pages, stats }
            window.PdfVisual.render(container, state)
   ============================================================ */
(function () {
  function build(a, b) {
    const pa = a.pdfPages || [];
    const pb = b.pdfPages || [];
    const nPages = Math.max(pa.length, pb.length);

    const pages = [];
    let addCount = 0, delCount = 0;

    for (let p = 0; p < nPages; p++) {
      const A = pa[p] || null;
      const B = pb[p] || null;
      const d = diffPage(A, B);
      addCount += d.add;
      delCount += d.del;
      pages.push({
        a: A ? { url: A.url, vw: A.vw, vh: A.vh, boxes: d.hlA } : null,
        b: B ? { url: B.url, vw: B.vw, vh: B.vh, boxes: d.hlB } : null,
        changed: d.add + d.del > 0
      });
    }

    return { pages, stats: { add: addCount, del: delCount } };
  }

  /* Diff word sequences of two pages; return merged highlight boxes. */
  function diffPage(A, B) {
    const aw = A ? A.words : [];
    const bw = B ? B.words : [];
    if (typeof Diff === "undefined") {
      // no diff lib — highlight nothing, just show pages
      return { hlA: [], hlB: [], del: 0, add: 0 };
    }
    const parts = Diff.diffArrays(aw.map(w => w.str), bw.map(w => w.str));

    let ai = 0, bi = 0;
    const rem = [], add = [];
    for (const part of parts) {
      const n = part.value.length;
      if (part.removed) { for (let k = 0; k < n; k++) rem.push(aw[ai + k]); ai += n; }
      else if (part.added) { for (let k = 0; k < n; k++) add.push(bw[bi + k]); bi += n; }
      else { ai += n; bi += n; }
    }
    return {
      hlA: mergeBoxes(rem),
      hlB: mergeBoxes(add),
      del: rem.length,
      add: add.length
    };
  }

  /* Merge contiguous same-line word boxes into clean line segments. */
  function mergeBoxes(words) {
    if (!words.length) return [];
    const sorted = [...words].sort((p, q) => (p.y - q.y) || (p.x - q.x));
    const boxes = [];
    let cur = null;
    for (const w of sorted) {
      if (cur &&
          Math.abs(w.y - cur.y) <= cur.h * 0.6 &&
          w.x <= cur.x + cur.w + cur.h * 0.9) {
        const right = Math.max(cur.x + cur.w, w.x + w.w);
        cur.x = Math.min(cur.x, w.x);
        cur.w = right - cur.x;
        cur.y = Math.min(cur.y, w.y);
        cur.h = Math.max(cur.h, w.h);
      } else {
        cur = { x: w.x, y: w.y, w: w.w, h: w.h };
        boxes.push(cur);
      }
    }
    return boxes;
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  function pageSideHtml(side, cls) {
    if (!side) {
      return `<div class="pdf-page pdf-page-empty"><span>no matching page</span></div>`;
    }
    let boxes = "";
    for (const bx of side.boxes) {
      const padY = bx.h * 0.18, padX = bx.h * 0.12;
      const l = (bx.x - padX) / side.vw * 100;
      const t = (bx.y - padY) / side.vh * 100;
      const w = (bx.w + padX * 2) / side.vw * 100;
      const h = (bx.h + padY * 2) / side.vh * 100;
      boxes += `<i class="pdf-hl ${cls}" style="left:${l.toFixed(3)}%;top:${t.toFixed(3)}%;width:${w.toFixed(3)}%;height:${h.toFixed(3)}%"></i>`;
    }
    return `<div class="pdf-page" style="aspect-ratio:${side.vw} / ${side.vh}">` +
             `<img src="${side.url}" alt="" loading="lazy">` +
             `<div class="pdf-ov">${boxes}</div>` +
           `</div>`;
  }

  function render(container, state) {
    const data = state.pdfData;
    if (!data) return;

    let html = `<div class="pdfdiff">`;
    html += `<div class="pdf-colhead">` +
      `<div class="pdf-h"><span class="pdf-badge a">A</span><span class="pdf-hname">${esc(state.a.name)}</span><span class="pdf-htag del">removed</span></div>` +
      `<div class="pdf-h"><span class="pdf-badge b">B</span><span class="pdf-hname">${esc(state.b.name)}</span><span class="pdf-htag add">added</span></div>` +
      `</div>`;

    if (data.stats.add + data.stats.del === 0) {
      html += `<div class="pdf-identical"><div class="pdf-id-ic">` +
        `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M20 6L9 17l-5-5"/></svg>` +
        `</div><h2>The PDFs look identical</h2><p>No text differences detected across ${data.pages.length} page${data.pages.length > 1 ? "s" : ""}.</p></div>`;
    }

    html += `<div class="pdf-pages">`;
    data.pages.forEach((pg, i) => {
      html += `<div class="pdf-row${pg.changed ? " has-change" : ""}" data-page="${i + 1}">` +
        `<div class="pdf-cell">${pageSideHtml(pg.a, "del")}</div>` +
        `<div class="pdf-pnum">${i + 1}</div>` +
        `<div class="pdf-cell">${pageSideHtml(pg.b, "add")}</div>` +
        `</div>`;
    });
    html += `</div></div>`;
    container.innerHTML = html;
  }

  window.PdfVisual = { build, render };
})();
