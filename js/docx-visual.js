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

  /* Parse HTML into a detached root and collect word tokens with
     references to their text node + offsets. */
  function tokenize(html) {
    const root = document.createElement("div");
    root.innerHTML = html || "";
    const tokens = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    let node;
    while ((node = walker.nextNode())) {
      const text = node.nodeValue;
      if (!text || !text.trim()) continue;
      const re = /\S+/g;
      let m;
      while ((m = re.exec(text))) {
        tokens.push({ node, start: m.index, end: m.index + m[0].length, word: m[0] });
      }
    }
    return { root, tokens };
  }

  /* Replace each affected text node with marked HTML. */
  function applyMarks(tokens, markMap, cls) {
    const byNode = new Map();
    tokens.forEach((t, i) => {
      if (!markMap.has(i)) return;
      if (!byNode.has(t.node)) byNode.set(t.node, []);
      byNode.get(t.node).push({ start: t.start, end: t.end, chg: markMap.get(i) });
    });
    for (const [node, list] of byNode) {
      const text = node.nodeValue;
      list.sort((a, b) => a.start - b.start);
      let html = "";
      let pos = 0;
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

  function build(a, b) {
    const A = tokenize(a.docHtml);
    const B = tokenize(b.docHtml);

    const aMarks = new Map(), bMarks = new Map();
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
        else { ai += n; bi += n; }
        prevChange = isChange;
      }
    }

    applyMarks(A.tokens, aMarks, "del");
    applyMarks(B.tokens, bMarks, "add");

    return {
      htmlA: A.root.innerHTML,
      htmlB: B.root.innerHTML,
      stats: { add, del },
      changes: chg
    };
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

    html += `<div class="doc-pages${state.focus ? " focused" : ""}">` +
      `<div class="doc-cell"><div class="doc-page" data-side="a">${data.htmlA}</div></div>` +
      `<div class="doc-cell"><div class="doc-page" data-side="b">${data.htmlB}</div></div>` +
      `</div></div>`;
    container.innerHTML = html;

    // Focus changes: hide top-level blocks that contain no change.
    if (state.focus && data.stats.add + data.stats.del > 0) {
      let hiddenBlocks = 0, totalBlocks = 0;
      container.querySelectorAll(".doc-page").forEach(page => {
        [...page.children].forEach(kid => {
          totalBlocks++;
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
  }

  window.DocxVisual = { build, render };
})();
