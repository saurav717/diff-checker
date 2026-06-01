/* ============================================================
   pdf-visual.js — show two PDFs side by side and highlight the
   regions that changed, page by page.

   Two comparison strategies, chosen automatically per page:

   1. WORD diff — when both pages expose a healthy, comparable text
      layer. Diffs the cached word sequences (in reading order) and
      highlights removed (A) / added (B) words. Handles reflow well.

   2. PIXEL diff — when a page's text layer is missing or unreliable
      (e.g. a filled / flattened / scanned PDF where most text can't be
      extracted). Comparing word sequences there produces nonsense
      ("everything deleted"), so instead we compare the rendered page
      bitmaps directly and highlight the regions that actually look
      different. Robust to broken text layers.

   pdf.js work (rendering pages + extracting word boxes) is done once at
   parse time (parsers.js) and cached on each file as
   `pdfPages: [{url, vw, vh, words}]`.

   Exposes: window.PdfVisual.build(a, b)  -> Promise<{ pages, stats, usedPixel }>
            window.PdfVisual.render(container, state)
   ============================================================ */
(function () {
  /* ---------- shared helpers ---------- */
  function esc(s) {
    return String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  /* pdf.js returns text items in content-stream order, which is NOT
     necessarily reading order. Canonicalise to reading order so
     identical content always yields an identical word sequence. */
  function readingOrder(words) {
    if (!words || !words.length) return [];
    const ws = [...words].sort((a, b) => (a.y - b.y) || (a.x - b.x));
    const lines = [];
    let cur = null;
    for (const w of ws) {
      if (cur && Math.abs((w.y + w.h / 2) - cur.cy) <= Math.max(w.h, cur.h) * 0.6) {
        cur.items.push(w);
        cur.cy = (cur.cy * (cur.items.length - 1) + (w.y + w.h / 2)) / cur.items.length;
        cur.h = Math.max(cur.h, w.h);
      } else {
        cur = { cy: w.y + w.h / 2, h: w.h, items: [w] };
        lines.push(cur);
      }
    }
    const out = [];
    for (const ln of lines) {
      ln.items.sort((a, b) => a.x - b.x);
      for (const w of ln.items) out.push(w);
    }
    return out;
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

  /* ---------- strategy 1: word diff ---------- */
  // Loose token key for OCR comparisons: lowercase and strip surrounding
  // punctuation — so identical text doesn't flag just because recognition
  // attached punctuation differently ("CMS?" vs "CMS", "etc.)," vs "etc.").
  function normTok(s) {
    return String(s).toLowerCase().replace(/^[^a-z0-9]+|[^a-z0-9]+$/g, "");
  }
  function wordDiffPage(A, B, loose) {
    const aw = A ? readingOrder(A.words) : [];
    const bw = B ? readingOrder(B.words) : [];
    if (typeof Diff === "undefined") return { hlA: [], hlB: [], del: 0, add: 0, mod: 0 };

    const opts = loose ? { comparator: (l, r) => l === r || normTok(l) === normTok(r) } : undefined;
    const parts = Diff.diffArrays(aw.map(w => w.str), bw.map(w => w.str), opts);
    let ai = 0, bi = 0;
    const rem = [], add = [];
    for (const part of parts) {
      const n = part.value.length;
      if (part.removed) { for (let k = 0; k < n; k++) rem.push(aw[ai + k]); ai += n; }
      else if (part.added) { for (let k = 0; k < n; k++) add.push(bw[bi + k]); bi += n; }
      else { ai += n; bi += n; }
    }
    return { hlA: mergeBoxes(rem), hlB: mergeBoxes(add), del: rem.length, add: add.length, mod: 0 };
  }

  /* ---------- strategy 2: pixel diff ---------- */
  function imgToData(url, W, H) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        try {
          const c = document.createElement("canvas");
          c.width = W; c.height = H;
          const ctx = c.getContext("2d", { willReadFrequently: true });
          ctx.fillStyle = "#fff";
          ctx.fillRect(0, 0, W, H);
          ctx.drawImage(img, 0, 0, W, H);
          const d = ctx.getImageData(0, 0, W, H);
          c.width = c.height = 0;
          resolve(d);
        } catch (e) { reject(e); }
      };
      img.onerror = reject;
      img.src = url;
    });
  }

  /* grow a boolean cell mask so nearby strokes/words merge into one region */
  function dilate(src, cols, rows, rx, ry) {
    const out = new Uint8Array(cols * rows);
    for (let y = 0; y < rows; y++) {
      for (let x = 0; x < cols; x++) {
        let on = 0;
        for (let dy = -ry; dy <= ry && !on; dy++) {
          for (let dx = -rx; dx <= rx; dx++) {
            const nx = x + dx, ny = y + dy;
            if (nx < 0 || ny < 0 || nx >= cols || ny >= rows) continue;
            if (src[ny * cols + nx]) { on = 1; break; }
          }
        }
        out[y * cols + x] = on;
      }
    }
    return out;
  }

  /* flood-fill connected components (8-connectivity) over a cell mask */
  function labelRegions(mask, cols, rows) {
    const seen = new Uint8Array(cols * rows);
    const regions = [];
    const stack = [];
    for (let s = 0; s < cols * rows; s++) {
      if (!mask[s] || seen[s]) continue;
      stack.length = 0; stack.push(s); seen[s] = 1;
      let x0 = cols, y0 = rows, x1 = 0, y1 = 0;
      const cells = [];
      while (stack.length) {
        const c = stack.pop();
        const cx = c % cols, cy = (c / cols) | 0;
        cells.push(c);
        if (cx < x0) x0 = cx; if (cx > x1) x1 = cx;
        if (cy < y0) y0 = cy; if (cy > y1) y1 = cy;
        for (let dy = -1; dy <= 1; dy++) {
          for (let dx = -1; dx <= 1; dx++) {
            if (!dx && !dy) continue;
            const nx = cx + dx, ny = cy + dy;
            if (nx < 0 || ny < 0 || nx >= cols || ny >= rows) continue;
            const ni = ny * cols + nx;
            if (mask[ni] && !seen[ni]) { seen[ni] = 1; stack.push(ni); }
          }
        }
      }
      regions.push({ x0, y0, x1, y1, cells });
    }
    return regions;
  }

  /* Estimate a per-band (dx,dy) offset between the two renderings. Two PDFs
     of the same document are often shifted relative to each other, sometimes
     unevenly down the page; aligning each horizontal band independently lets
     us keep the per-pixel match window small (so real additions aren't masked)
     while still recognising identical-but-shifted text as the same.

     The offset is chosen to MAXIMISE the count of well-matched content pixels
     (not to minimise total difference) — this makes the alignment robust to
     added/removed content, which would otherwise drag a sum-of-differences
     toward a wrong offset. Each entry is the offset to add to a B coordinate
     to land on A. */
  function estimateBandOffsets(gA, gB, W, H) {
    const BH = 32;                       // band height (full-res px)
    const ALIGN_T = 45;                  // luminance gap under which two pixels "match"
    const nb = Math.ceil(H / BH);
    const offs = new Array(nb);
    for (let b = 0; b < nb; b++) {
      const y0 = b * BH, y1 = Math.min(H, y0 + BH);
      let bestScore = -1, by = 0, bx = 0, content = 0;
      for (let oy = -8; oy <= 8; oy++) {
        for (let ox = -4; ox <= 4; ox++) {
          let match = 0, seen = 0;
          for (let y = y0; y < y1; y += 3) {
            const ay = y + oy; if (ay < 0 || ay >= H) continue;
            const baseB = y * W, baseA = ay * W;
            for (let x = 0; x < W; x += 3) {
              const ax = x + ox; if (ax < 0 || ax >= W) continue;
              const vb = gB[baseB + x], va = gA[baseA + ax];
              if (vb < 235 || va < 235) {            // a content pixel on either side
                seen++;
                if (Math.abs(vb - va) < ALIGN_T) match++;
              }
            }
          }
          if (match > bestScore) { bestScore = match; by = oy; bx = ox; content = seen; }
        }
      }
      offs[b] = { gx: bx, gy: by, has: bestScore > 12 };
    }
    // Bands that were mostly blank inherit the nearest aligned band's offset.
    let last = null;
    for (let b = 0; b < nb; b++) { if (offs[b].has) last = offs[b]; else if (last) offs[b] = { gx: last.gx, gy: last.gy, has: false }; }
    last = null;
    for (let b = nb - 1; b >= 0; b--) { if (offs[b].has) last = offs[b]; else if (last && !offs[b].has) offs[b] = { gx: last.gx, gy: last.gy, has: false }; }
    return { offs, BH, nb };
  }

  const CELL = 8;          // px per analysis cell
  const INK = 205;         // luminance below this counts as "ink" (darker than light grey)
  const MATCH_T = 55;      // luminance gap above which ink is "not explained" by the other side
  const MIN_PX = 5;        // changed ink pixels in a cell before it counts as changed

  async function pixelDiffPage(A, B) {
    // A page present on only one side: flag the whole page.
    if (A && !B) return { hlA: [{ x: 0, y: 0, w: A.vw, h: A.vh, whole: true }], hlB: [], del: 1, add: 0, mod: 0 };
    if (B && !A) return { hlA: [], hlB: [{ x: 0, y: 0, w: B.vw, h: B.vh, whole: true }], del: 0, add: 1, mod: 0 };
    if (!A || !B) return { hlA: [], hlB: [], del: 0, add: 0, mod: 0 };

    const W = A.vw, H = A.vh;
    const ia = await imgToData(A.url, W, H);
    const ib = await imgToData(B.url, W, H);     // scaled onto A's geometry
    const da = ia.data, db = ib.data;

    // Grayscale (luminance) buffers.
    const gA = new Float32Array(W * H), gB = new Float32Array(W * H);
    for (let p = 0, i = 0; p < W * H; p++, i += 4) {
      gA[p] = 0.299 * da[i] + 0.587 * da[i + 1] + 0.114 * da[i + 2];
      gB[p] = 0.299 * db[i] + 0.587 * db[i + 1] + 0.114 * db[i + 2];
    }

    const cols = Math.ceil(W / CELL), rows = Math.ceil(H / CELL);
    const N = cols * rows;
    const chgA = new Uint16Array(N);   // ink pixels in A not explained by B (removed)
    const chgB = new Uint16Array(N);   // ink pixels in B not explained by A (added)

    // Correct for an overall shift between the two renderings, then allow a
    // generous local search around that — so identical text shifted by a few
    // px (even unevenly down the page) is still recognised as the same.
    const { offs, BH, nb } = estimateBandOffsets(gA, gB, W, H);
    const RX = 2, RY = 4;   // small local search — bands already handle the bulk shift

    /* A pixel of ink only registers as a change if NO pixel in the aligned
       search window of the other image has a similar luminance. This makes
       anti-aliasing and shifts of the SAME content "explained" (so common
       text never flags), while genuinely added/removed ink is kept. */
    for (let y = 0; y < H; y++) {
      const cellRow = ((y / CELL) | 0) * cols;
      const band = offs[Math.min(nb - 1, (y / BH) | 0)];
      const gx = band.gx, gy = band.gy;
      for (let x = 0; x < W; x++) {
        const p = y * W + x;
        const lib = gB[p], lia = gA[p];
        if (lib >= INK && lia >= INK) continue;   // both blank here → skip fast
        const ci = cellRow + ((x / CELL) | 0);

        if (lib < INK) {                          // candidate added ink (B): match against A
          const cyA = y + gy, cxA = x + gx;
          let m = Infinity;
          for (let dy = -RY; dy <= RY && m > MATCH_T; dy++) {
            const ay = cyA + dy; if (ay < 0 || ay >= H) continue;
            const base = ay * W;
            for (let dx = -RX; dx <= RX; dx++) {
              const ax = cxA + dx; if (ax < 0 || ax >= W) continue;
              const d = Math.abs(lib - gA[base + ax]);
              if (d < m) { m = d; if (m <= MATCH_T) break; }
            }
          }
          if (m > MATCH_T) chgB[ci]++;
        }
        if (lia < INK) {                          // candidate removed ink (A): match against B
          const cyB = y - gy, cxB = x - gx;
          let m = Infinity;
          for (let dy = -RY; dy <= RY && m > MATCH_T; dy++) {
            const by = cyB + dy; if (by < 0 || by >= H) continue;
            const base = by * W;
            for (let dx = -RX; dx <= RX; dx++) {
              const bx = cxB + dx; if (bx < 0 || bx >= W) continue;
              const d = Math.abs(lia - gB[base + bx]);
              if (d < m) { m = d; if (m <= MATCH_T) break; }
            }
          }
          if (m > MATCH_T) chgA[ci]++;
        }
      }
    }

    const changed = new Uint8Array(N);
    for (let c = 0; c < N; c++) {
      if (chgA[c] >= MIN_PX || chgB[c] >= MIN_PX) changed[c] = 1;
    }

    // Bridge gaps between letters/words on a line before grouping.
    const grown = dilate(changed, cols, rows, 3, 1);
    const regions = labelRegions(grown, cols, rows);

    const hlA = [], hlB = [];
    let add = 0, del = 0;
    for (const r of regions) {
      if (r.cells.length < 4) continue;   // drop tiny specks
      let ca = 0, cb = 0;
      for (const ci of r.cells) { ca += chgA[ci]; cb += chgB[ci]; }
      const box = {
        x: r.x0 * CELL, y: r.y0 * CELL,
        w: (r.x1 - r.x0 + 1) * CELL, h: (r.y1 - r.y0 + 1) * CELL
      };
      const lo = Math.min(ca, cb), hi = Math.max(ca, cb) || 1;
      if (lo / hi >= 0.45) {
        // Comparable ink changed on both sides → content replaced in place:
        // show as a deletion on the left AND an addition on the right.
        hlA.push(Object.assign({}, box));
        hlB.push(Object.assign({}, box));
        del++; add++;
      } else if (cb > ca) {
        hlB.push(box); add++;             // ink only/mostly added in B
      } else {
        hlA.push(box); del++;             // ink only/mostly removed from A
      }
    }
    return { hlA, hlB, add, del, mod: 0 };
  }

  /* Decide whether a page's text layer is trustworthy enough for a
     word diff, or whether we should fall back to a pixel comparison.
     We only trust the word diff when BOTH sides expose a substantial,
     comparable amount of text — otherwise (scanned / flattened / filled
     PDFs whose text can't be extracted) we compare pixels instead. */
  function textLayerReliable(A, B) {
    const na = A && A.words ? A.words.length : 0;
    const nb = B && B.words ? B.words.length : 0;
    const hi = Math.max(na, nb), lo = Math.min(na, nb);
    if (lo < 20) return false;           // a side lacks a usable text layer
    return lo / hi >= 0.6;               // and the two counts are comparable
  }

  /* ---------- build ---------- */
  async function build(a, b, opts) {
    const forceWord = !!(opts && opts.ocr);
    const pa = a.pdfPages || [];
    const pb = b.pdfPages || [];
    const nPages = Math.max(pa.length, pb.length);

    const pages = [];
    let addCount = 0, delCount = 0, modCount = 0, usedPixel = false;

    for (let p = 0; p < nPages; p++) {
      const A = pa[p] || null;
      const B = pb[p] || null;

      let d;
      if (A && B && (forceWord || textLayerReliable(A, B))) {
        d = wordDiffPage(A, B, forceWord);
      } else {
        usedPixel = true;
        try { d = await pixelDiffPage(A, B); }
        catch (_) { d = wordDiffPage(A, B); }   // last-resort fallback
      }

      addCount += d.add; delCount += d.del; modCount += d.mod || 0;
      pages.push({
        a: A ? { url: A.url, vw: A.vw, vh: A.vh, boxes: d.hlA } : null,
        b: B ? { url: B.url, vw: B.vw, vh: B.vh, boxes: d.hlB } : null,
        changed: (d.add + d.del + (d.mod || 0)) > 0
      });
    }

    return { pages, stats: { add: addCount, del: delCount, mod: modCount }, usedPixel, usedOcr: forceWord };
  }

  /* ---------- render ---------- */
  function pageSideHtml(side, cls) {
    if (!side) {
      return `<div class="pdf-page pdf-page-empty"><span>no matching page</span></div>`;
    }
    let boxes = "";
    for (const bx of side.boxes) {
      const c = bx.mod ? "mod" : cls;
      if (bx.whole) {
        boxes += `<i class="pdf-hl ${c} pdf-hl-whole" style="left:0;top:0;width:100%;height:100%"></i>`;
        continue;
      }
      const padY = Math.max(bx.h * 0.04, 1), padX = Math.max(bx.h * 0.04, 1);
      const l = (bx.x - padX) / side.vw * 100;
      const t = (bx.y - padY) / side.vh * 100;
      const w = (bx.w + padX * 2) / side.vw * 100;
      const h = (bx.h + padY * 2) / side.vh * 100;
      boxes += `<i class="pdf-hl ${c}" style="left:${l.toFixed(3)}%;top:${t.toFixed(3)}%;width:${w.toFixed(3)}%;height:${h.toFixed(3)}%"></i>`;
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

    const total = data.stats.add + data.stats.del + (data.stats.mod || 0);

    if (total === 0) {
      html += `<div class="pdf-identical"><div class="pdf-id-ic">` +
        `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M20 6L9 17l-5-5"/></svg>` +
        `</div><h2>The PDFs look identical</h2><p>No differences detected across ${data.pages.length} page${data.pages.length > 1 ? "s" : ""}.</p></div>`;
    }

    if (data.usedOcr && total > 0) {
      html += `<div class="pdf-pixelnote pdf-ocrnote">` +
        `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 7V5a1 1 0 0 1 1-1h2M4 17v2a1 1 0 0 0 1 1h2M20 7V5a1 1 0 0 0-1-1h-2M20 17v2a1 1 0 0 1-1 1h-2M8 12h8"/></svg>` +
        `Text was recovered with <strong>OCR</strong> and compared word-by-word. Recognition isn’t perfect — a few highlights may reflect OCR misreads rather than real edits.` +
        `</div>`;
    } else if (data.usedPixel && total > 0) {
      html += `<div class="pdf-pixelnote">` +
        `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 9v4m0 4h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>` +
        `<span class="pn-text">A file’s text layer was incomplete (often the case with filled, flattened or scanned PDFs), so those pages were compared <strong>visually, pixel-by-pixel</strong> — highlighting regions that look different rather than extracted words.</span>` +
        `<button class="pdf-ocr-btn" id="ocrBtn" type="button">` +
        `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9"><path d="M4 7V5a1 1 0 0 1 1-1h2M4 17v2a1 1 0 0 0 1 1h2M20 7V5a1 1 0 0 0-1-1h-2M20 17v2a1 1 0 0 1-1 1h-2M8 12h8"/></svg>` +
        `Read text with OCR &amp; compare</button>` +
        `</div>`;
    }

    // "Focus changes" collapses to just the pages that changed.
    const list = data.pages
      .map((pg, i) => ({ pg, n: i + 1 }))
      .filter(x => (state.focus ? x.pg.changed : true));

    if (state.focus && list.length && total > 0) {
      const hidden = data.pages.length - list.length;
      if (hidden > 0) {
        html += `<div class="pdf-focusnote">Showing ${list.length} changed page${list.length > 1 ? "s" : ""} · ${hidden} unchanged page${hidden > 1 ? "s" : ""} hidden</div>`;
      }
    }

    html += `<div class="pdf-pages">`;
    list.forEach(({ pg, n }) => {
      html += `<div class="pdf-row${pg.changed ? " has-change" : ""}" data-page="${n}">` +
        `<div class="pdf-cell">${pageSideHtml(pg.a, "del")}</div>` +
        `<div class="pdf-pnum">${n}</div>` +
        `<div class="pdf-cell">${pageSideHtml(pg.b, "add")}</div>` +
        `</div>`;
    });
    html += `</div></div>`;
    container.innerHTML = html;
  }

  window.PdfVisual = { build, render };
})();
