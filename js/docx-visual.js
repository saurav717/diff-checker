/* ============================================================
   docx-visual.js — show two Word documents as formatted "pages"
   side by side, with changed words highlighted inline.

   Word files have no page bitmap, so we render the formatted HTML
   (produced by mammoth at parse time, cached as `docHtml`) inside a
   page-styled container. We diff the word sequences of the two
   documents and wrap removed words (A) / added words (B) in <mark>s,
   tagging each change region with a shared id for navigation.

   Exposes: window.DocxVisual.build(a, b)  -> { htmlA, htmlB, stats, changes }
            window.DocxVisual.render(container, state)
   ============================================================ */
(function () {
  function esc(s) {
    return String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  /* Map a Word font name to a stack that leads with a metric-compatible
     web font (loaded in index.html) so the page matches Word's metrics
     even when the original font isn't installed. */
  function mapFont(name) {
    if (!name) return null;
    const n = name.toLowerCase();
    if (n.includes("calibri")) return "'Carlito','Calibri',sans-serif";
    if (n.includes("cambria")) return "'Caladea','Cambria',Georgia,serif";
    if (n.includes("times")) return "'Tinos','Times New Roman',Times,serif";
    if (n.includes("arial") || n.includes("helvetica")) return "'Arimo','Arial',Helvetica,sans-serif";
    if (n.includes("garamond")) return "'EB Garamond',Garamond,'Times New Roman',serif";
    if (n.includes("georgia")) return "Georgia,'Times New Roman',serif";
    if (n.includes("verdana")) return "Verdana,Geneva,sans-serif";
    if (n.includes("tahoma")) return "Tahoma,Geneva,sans-serif";
    if (n.includes("courier") || n.includes("consolas") || n.includes("mono"))
      return "'Cousine','Courier New',monospace";
    // Unknown named font: try it, then fall back to Calibri's clone.
    return `'${name.replace(/'/g, "")}','Carlito',sans-serif`;
  }

  /* Build the inline style that makes a .doc-page reproduce the source
     document's font, size and spacing as Microsoft Word would show it. */
  function pageStyle(d) {
    const s = (d && d.docStyle) || {};
    const out = [];
    const fam = mapFont(s.font);
    if (fam) out.push(`font-family:${fam}`);
    out.push(`font-size:${s.sizePt ? s.sizePt : 11}pt`);            // Word default body = 11pt
    out.push(`line-height:${s.lineHeight ? s.lineHeight : 1.08}`);   // Word default = 1.08
    out.push(`--doc-after:${s.afterPt != null ? s.afterPt : 8}pt`);  // space after para
    out.push(`--doc-before:${s.beforePt != null ? s.beforePt : 0}pt`);
    return out.join(";");
  }

  /* Walk an element's text nodes and collect word tokens with references
     to their text node + offsets. */
  /* nearest block-level ancestor — sentences never span block boundaries */
  function blockOf(node, root) {
    let el = node.parentElement;
    while (el && el !== root && !/^(P|LI|TD|TH|H1|H2|H3|H4|H5|H6|DIV|SECTION|ARTICLE|HEADER|FOOTER|BLOCKQUOTE|FIGCAPTION|TR)$/.test(el.tagName)) el = el.parentElement;
    return el || root;
  }
  function collectTokens(root) {
    const tokens = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    let node;
    while ((node = walker.nextNode())) {
      const text = node.nodeValue;
      if (!text || !text.trim()) continue;
      const block = blockOf(node, root);
      const re = /\S+/g;
      let m;
      while ((m = re.exec(text))) {
        tokens.push({ node, start: m.index, end: m.index + m[0].length, word: m[0], block });
      }
    }
    return tokens;
  }
  function tokenizeHtml(html) {
    const root = document.createElement("div");
    root.innerHTML = html || "";
    return { root, tokens: collectTokens(root) };
  }

  /* Group tokens into sentences: a new sentence begins at a block change or
     right after a token that ends with sentence punctuation. */
  function assignSentences(tokens) {
    let sid = 0, prevBlock = null, prevEnded = true;
    for (const t of tokens) {
      if (prevBlock !== t.block || prevEnded) sid++;
      t.sentence = sid;
      prevEnded = /[.!?:][)"'\]]?$/.test(t.word);
      prevBlock = t.block;
    }
  }
  /* Sentence mode: extend the per-word change map to cover every word of any
     sentence that contains a change (so a reworded sentence highlights whole). */
  function expandToSentences(tokens, markMap) {
    const changedSents = new Map();   // sentence id -> first change id
    tokens.forEach((t, i) => {
      if (markMap.has(i) && !changedSents.has(t.sentence)) changedSents.set(t.sentence, markMap.get(i));
    });
    tokens.forEach((t, i) => {
      if (changedSents.has(t.sentence)) markMap.set(i, changedSents.get(t.sentence));
    });
  }

  /* Render a .docx buffer with docx-preview into a detached element,
     returning the rendered DOM + the generated CSS (namespaced per side so
     the two documents' stylesheets don't collide). */
  async function renderDocx(buffer, className) {
    const host = document.createElement("div");
    host.style.cssText = "position:absolute;left:-99999px;top:0;width:820px;visibility:hidden";
    document.body.appendChild(host);
    const styleHost = document.createElement("div");
    try {
      await window.docx.renderAsync(buffer.slice(0), host, styleHost, {
        className,
        inWrapper: true,
        ignoreWidth: false,                 // real page width → headers/tabs/right-alignment correct
        ignoreHeight: false,                // real page height → faithful pagination
        breakPages: true,                   // honour authored page & section breaks
        ignoreLastRenderedPageBreak: true,  // ...but NOT Word's stale soft markers (they cause blank pages)
        renderHeaders: true,                // brings in header content (e.g. a logo + "Version 2.0")
        renderFooters: true,
        renderFootnotes: true,
        useBase64URL: true,                 // embed images so they survive innerHTML round-trip
        experimental: true
      });
    } finally {
      document.body.removeChild(host);
    }
    return { root: host, css: styleHost.innerHTML };
  }

  /* Replace each affected text node with marked HTML. When merge=true
     (sentence mode), consecutive marked words separated only by whitespace
     are joined into one continuous highlight. */
  function applyMarks(tokens, markMap, cls, merge) {
    const byNode = new Map();
    tokens.forEach((t, i) => {
      if (!markMap.has(i)) return;
      if (!byNode.has(t.node)) byNode.set(t.node, []);
      byNode.get(t.node).push({ start: t.start, end: t.end, chg: markMap.get(i) });
    });
    for (const [node, list] of byNode) {
      const text = node.nodeValue;
      list.sort((a, b) => a.start - b.start);
      // Build the ranges to wrap (merge whitespace-separated runs in sentence mode).
      const ranges = [];
      for (const t of list) {
        const last = ranges[ranges.length - 1];
        if (merge && last && /^\s*$/.test(text.slice(last.end, t.start))) {
          last.end = t.end;                       // extend continuous highlight
        } else {
          ranges.push({ start: t.start, end: t.end, chg: t.chg });
        }
      }
      let html = "", pos = 0;
      for (const r of ranges) {
        html += esc(text.slice(pos, r.start));
        html += `<mark class="dv-mark dv-${cls}" data-chg="${r.chg}">` + esc(text.slice(r.start, r.end)) + `</mark>`;
        pos = r.end;
      }
      html += esc(text.slice(pos));
      const span = document.createElement("span");
      span.className = "dv-wrap";
      span.innerHTML = html;
      node.parentNode.replaceChild(span, node);
    }
  }

  /* Render (or reuse cached) docx-preview HTML for one file. Caching the raw
     rendered HTML on the file object makes granularity toggles cheap — no need
     to re-run docx-preview, just re-tokenize + re-mark. */
  async function prep(file, cls) {
    if (file.docPreviewHtml != null) return { html: file.docPreviewHtml, css: file.docPreviewCss, preview: true };
    if (window.docx && file.docBuffer) {
      try {
        const r = await renderDocx(file.docBuffer, cls);
        file.docPreviewHtml = r.root.innerHTML;
        file.docPreviewCss = r.css;
        return { html: file.docPreviewHtml, css: file.docPreviewCss, preview: true };
      } catch (_) { /* fall through to mammoth */ }
    }
    return { html: file.docHtml || "", css: "", preview: false };
  }

  async function build(a, b, granularity) {
    const sentence = granularity === "sentence";
    const pa = await prep(a, "dpa"), pb = await prep(b, "dpb");
    const usedPreview = pa.preview && pb.preview;
    const A = tokenizeHtml(pa.html), B = tokenizeHtml(pb.html);
    assignSentences(A.tokens);
    assignSentences(B.tokens);

    const aMarks = new Map(), bMarks = new Map();
    const pairCount = new Map();   // aBlock -> Map(bBlock -> shared-token count)
    let del = 0, add = 0, chg = 0;

    if (typeof Diff !== "undefined") {
      const parts = Diff.diffArrays(A.tokens.map(t => t.word), B.tokens.map(t => t.word));
      let ai = 0, bi = 0, prevChange = false;
      for (const part of parts) {
        const n = part.value.length;
        const isChange = part.added || part.removed;
        if (isChange && !prevChange) chg++;
        if (part.removed) { for (let k = 0; k < n; k++) aMarks.set(ai + k, chg); ai += n; del += n; }
        else if (part.added) { for (let k = 0; k < n; k++) bMarks.set(bi + k, chg); bi += n; add += n; }
        else {
          // equal run: corresponding tokens map their blocks A↔B
          for (let k = 0; k < n; k++) {
            const ab = A.tokens[ai + k].block, bb = B.tokens[bi + k].block;
            if (ab && bb) {
              if (!pairCount.has(ab)) pairCount.set(ab, new Map());
              const m = pairCount.get(ab);
              m.set(bb, (m.get(bb) || 0) + 1);
            }
          }
          ai += n; bi += n;
        }
        prevChange = isChange;
      }
    }

    if (sentence) {
      expandToSentences(A.tokens, aMarks);
      expandToSentences(B.tokens, bMarks);
    }

    applyMarks(A.tokens, aMarks, "del", sentence);
    applyMarks(B.tokens, bMarks, "add", sentence);

    // Assign a shared map id to each A block and its best-matching B block, so
    // hovering ANY text (changed or not) can highlight its mapped counterpart.
    let mid = 0;
    const usedB = new Set(), seenA = new Set();
    for (const t of A.tokens) {
      const ab = t.block;
      if (!ab || seenA.has(ab)) continue;
      seenA.add(ab);
      const cands = pairCount.get(ab);
      if (!cands) continue;
      let best = null, bestN = 0;
      for (const [bb, c] of cands) { if (!usedB.has(bb) && c > bestN) { best = bb; bestN = c; } }
      if (best) {
        const id = "m" + (mid++);
        ab.setAttribute("data-map", id);
        best.setAttribute("data-map", id);
        usedB.add(best);
      }
    }

    return {
      htmlA: A.root.innerHTML,
      htmlB: B.root.innerHTML,
      cssA: pa.css, cssB: pb.css, usedPreview,
      stats: { add, del },
      changes: chg
    };
  }

  /* docx-preview renders pages at their true width (e.g. 816px = 8.5in).
     Scale each rendered document down so it fits its diff column, preserving
     the exact Word layout (margins, header tabs, pagination). */
  /* Word's PAGE field can't be computed by docx-preview, so every page shows
     the document's cached value. Replace the page-number text in each page's
     footer with its real sequential position. */
  function fixPageNumbers(host) {
    const footers = host.querySelectorAll("footer");
    footers.forEach((f, i) => {
      const walker = document.createTreeWalker(f, NodeFilter.SHOW_TEXT, null);
      let n;
      while ((n = walker.nextNode())) {
        if (/^\s*\d+\s*$/.test(n.nodeValue)) { n.nodeValue = String(i + 1); break; }
      }
    });
  }

  /* docx-preview doesn't always honour a header's right tab stop (so text
     meant for the right margin lands mid-line). For the common logo-left /
     text-right header & footer pattern, lay it out so the image block stays
     left and the remaining text is pushed to the right edge — matching Word. */
  function fixHeaderLayout(host) {
    host.querySelectorAll("header, footer").forEach(hf => {
      const img = hf.querySelector("img");
      if (!img || !(hf.textContent || "").trim()) return;   // only logo + text headers
      let imgBlock = img;
      while (imgBlock.parentElement && imgBlock.parentElement !== hf) imgBlock = imgBlock.parentElement;
      hf.style.display = "flex";
      hf.style.alignItems = "center";
      hf.style.flexWrap = "nowrap";
      imgBlock.style.marginRight = "auto";   // logo left, everything after → right
      imgBlock.style.flex = "0 0 auto";
    });
  }

  function scaleHosts(container) {
    container.querySelectorAll(".docx-host").forEach(host => {
      const inner = host.firstElementChild;
      if (!inner) return;
      inner.style.transform = "none";       // reset before measuring
      host.style.height = "auto";
      const page = host.querySelector("section");
      const naturalW = page ? page.offsetWidth : inner.scrollWidth;
      const availW = host.clientWidth;
      if (!naturalW || !availW) return;
      const scale = Math.min(1, availW / naturalW);
      inner.style.transformOrigin = "top left";
      inner.style.transform = `scale(${scale})`;
      host.style.height = (inner.scrollHeight * scale) + "px";
    });
  }

  let resizeBound = false;
  function bindResize() {
    if (resizeBound) return;
    resizeBound = true;
    let t;
    window.addEventListener("resize", () => {
      clearTimeout(t);
      t = setTimeout(() => {
        const dv = document.getElementById("diffview");
        if (dv && dv.querySelector(".docx-host")) scaleHosts(dv);
      }, 120);
    });
  }

  function render(container, state) {
    const data = state.docData;
    if (!data) return;

    let html = `<div class="docdiff">`;
    html += `<div class="doc-colhead">` +
      `<div class="doc-h"><span class="doc-badge a">A</span><span class="doc-hname">${esc(state.a.name)}</span><span class="doc-htag del">removed</span></div>` +
      `<div class="doc-h"><span class="doc-badge b">B</span><span class="doc-hname">${esc(state.b.name)}</span><span class="doc-htag add">added</span></div>` +
      `</div>`;

    if (data.stats.add + data.stats.del === 0) {
      html += `<div class="doc-identical"><div class="doc-id-ic">` +
        `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M20 6L9 17l-5-5"/></svg>` +
        `</div><h2>The documents look identical</h2><p>No text differences detected.</p></div>`;
    }

    const cell = (side, file) => {
      if (data.usedPreview) {
        const css = side === "a" ? data.cssA : data.cssB;
        const html = side === "a" ? data.htmlA : data.htmlB;
        return `<div class="doc-cell">${css}<div class="docx-host" data-side="${side}">${html}</div></div>`;
      }
      const html = side === "a" ? data.htmlA : data.htmlB;
      return `<div class="doc-cell"><div class="doc-page" data-side="${side}" style="${pageStyle(file)}">${html}</div></div>`;
    };

    html += `<div class="doc-pages${state.focus ? " focused" : ""}${data.usedPreview ? " preview" : ""}">` +
      cell("a", state.a) + cell("b", state.b) +
      `</div></div>`;
    container.innerHTML = html;

    // Focus changes: hide top-level blocks that contain no change.
    if (state.focus && data.stats.add + data.stats.del > 0) {
      let hiddenBlocks = 0;
      const pageSel = data.usedPreview ? ".docx-host article" : ".doc-page";
      container.querySelectorAll(pageSel).forEach(page => {
        [...page.children].forEach(kid => {
          if (!kid.querySelector(".dv-mark") && !kid.classList.contains("dv-mark")) {
            kid.classList.add("dv-hidden");
            hiddenBlocks++;
          }
        });
      });
      if (hiddenBlocks > 0) {
        const note = document.createElement("div");
        note.className = "doc-focusnote";
        note.textContent = `Showing changed sections only · ${hiddenBlocks} unchanged block${hiddenBlocks > 1 ? "s" : ""} hidden`;
        const pages = container.querySelector(".doc-pages");
        pages.parentNode.insertBefore(note, pages);
      }
    }

    if (data.usedPreview) {
      // Fix logo-left/text-right headers and sequential page numbers, then scale.
      container.querySelectorAll(".docx-host").forEach(host => {
        fixHeaderLayout(host);
        fixPageNumbers(host);
      });
      scaleHosts(container);
      requestAnimationFrame(() => scaleHosts(container));
      bindResize();
    }
  }

  window.DocxVisual = { build, render };
})();
