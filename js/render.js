/* ============================================================
   render.js — render aligned rows into the diff view, plus
   search highlighting and change navigation.
   Depends on DiffEngine. Exposes window.DiffRender
   ============================================================ */
(function () {
  const E = window.DiffEngine;

  /* Render full diff. ctx state passed in.
     state: {rows, view, focus, granularity, forced, nameA, nameB, kindA, kindB} */
  function render(container, state) {
    const { rows, view, focus, granularity, forced } = state;
    const st = E.stats(rows);

    if (!st.changed) {
      container.innerHTML = identicalHtml();
      return;
    }

    const visible = E.computeVisible(rows, focus, forced, 3);
    const parts = [];

    // sticky file header
    parts.push(fileHeader(state));

    parts.push(`<div class="diff-grid">`);
    for (const item of visible) {
      if (item.type === "fold") {
        parts.push(foldHtml(item, view));
      } else {
        parts.push(view === "inline"
          ? rowInline(item.row, item.i, granularity)
          : rowSplit(item.row, item.i, granularity));
      }
    }
    parts.push(`</div>`);

    container.innerHTML = parts.join("");
  }

  function fileHeader(state) {
    const cls = state.view === "inline" ? "inline" : "split";
    const a = `<div class="fh-cell a"><span class="fh-dot"></span><span class="fh-name">${esc(state.nameA)}</span><span class="fh-kind">${esc(state.kindA)}</span></div>`;
    const b = `<div class="fh-cell b"><span class="fh-dot"></span><span class="fh-name">${esc(state.nameB)}</span><span class="fh-kind">${esc(state.kindB)}</span></div>`;
    if (state.view === "inline") {
      return `<div class="fileheader inline"><div class="fh-cell"><span class="fh-dot" style="background:var(--del-sign)"></span><span class="fh-name">${esc(state.nameA)}</span><span class="fh-kind">→ ${esc(state.nameB)}</span></div></div>`;
    }
    return `<div class="fileheader split">${a}${b}</div>`;
  }

  function gut(no) { return no == null ? "" : no; }

  function rowSplit(row, idx, gran) {
    let lNo = "", rNo = "", lContent = "", rContent = "", lCls = "", rCls = "";
    if (row.type === "eq") {
      lNo = row.left.no; rNo = row.right.no;
      lContent = esc(row.left.raw); rContent = esc(row.right.raw);
    } else if (row.type === "mod") {
      lNo = row.left.no; rNo = row.right.no; lCls = "del-bg"; rCls = "add-bg";
      const im = E.intraLine(row.left.raw, row.right.raw, gran);
      lContent = im.leftHtml; rContent = im.rightHtml;
    } else if (row.type === "del") {
      lNo = row.left.no; lCls = "del-bg";
      lContent = esc(row.left.raw);
    } else if (row.type === "add") {
      rNo = row.right.no; rCls = "add-bg";
      rContent = esc(row.right.raw);
    }
    return `<div class="row split" data-i="${idx}" data-type="${row.type}">` +
      `<div class="ln left ${lCls}">${gut(lNo)}</div>` +
      `<div class="cell left ${lCls}">${lContent || "&nbsp;"}</div>` +
      `<div class="ln right ${rCls}">${gut(rNo)}</div>` +
      `<div class="cell right ${rCls}">${rContent || "&nbsp;"}</div>` +
      `</div>`;
  }

  /* Inline (unified): a 'mod' becomes a minus row then a plus row. */
  function rowInline(row, idx, gran) {
    if (row.type === "eq") {
      return inlineLine(idx, row.type, row.left.no, row.right.no, " ", esc(row.left.raw), "");
    }
    if (row.type === "del") {
      return inlineLine(idx, row.type, row.left.no, "", "-", esc(row.left.raw), "minus");
    }
    if (row.type === "add") {
      return inlineLine(idx, row.type, "", row.right.no, "+", esc(row.right.raw), "plus");
    }
    // mod -> two lines, with intra-line marks
    const im = E.intraLine(row.left.raw, row.right.raw, gran);
    return inlineLine(idx, "del", row.left.no, "", "-", im.leftHtml, "minus") +
           inlineLine(idx, "add", "", row.right.no, "+", im.rightHtml, "plus");
  }

  function inlineLine(idx, type, lNo, rNo, sign, content, signCls) {
    const bg = signCls === "plus" ? "add-bg" : signCls === "minus" ? "del-bg" : "";
    return `<div class="row inline" data-i="${idx}" data-type="${type}">` +
      `<div class="ln ${bg}">${gut(lNo)}</div>` +
      `<div class="sign ${signCls}">${sign}</div>` +
      `<div class="cell ${bg}">${content || "&nbsp;"}</div>` +
      `</div>`;
  }

  function foldHtml(item, view) {
    const cols = view === "inline" ? 3 : 4;
    return `<div class="fold" data-from="${item.from}" data-to="${item.to}" style="grid-column:1/-1">` +
      `<svg class="fold-icon" width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M5 6l3 3 3-3M5 10l3-3 3 3" opacity="0"/><path d="M2 8h3M11 8h3"/><path d="M8 5v6"/></svg>` +
      `<span>Show ${item.count} unchanged line${item.count === 1 ? "" : "s"}</span>` +
      `</div>`;
  }

  function identicalHtml() {
    return `<div class="identical">` +
      `<svg class="ic" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="9"/><path d="M8.5 12.5l2.2 2.2L16 9.5"/></svg>` +
      `<h2>The files are identical</h2>` +
      `<p>No differences were found between the two files.</p>` +
      `</div>`;
  }

  function esc(s) { return E.esc(String(s == null ? "" : s)); }

  /* ---------- Search highlight ---------- */
  function clearSearch(container) {
    container.querySelectorAll("mark.s-hit").forEach(m => {
      const parent = m.parentNode;
      parent.replaceChild(document.createTextNode(m.textContent), m);
      parent.normalize();
    });
  }

  /* Highlight all matches of `query` (case-insensitive) inside .cell elements.
     Returns array of <mark> elements. */
  function applySearch(container, query) {
    clearSearch(container);
    if (!query) return [];
    const q = query.toLowerCase();
    const cells = container.querySelectorAll(".cell, .gcell");
    const marks = [];
    cells.forEach(cell => {
      const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT, null);
      const textNodes = [];
      let n;
      while ((n = walker.nextNode())) textNodes.push(n);
      textNodes.forEach(node => {
        const text = node.nodeValue;
        const lower = text.toLowerCase();
        let idx = lower.indexOf(q);
        if (idx === -1) return;
        const frag = document.createDocumentFragment();
        let last = 0;
        while (idx !== -1) {
          if (idx > last) frag.appendChild(document.createTextNode(text.slice(last, idx)));
          const mark = document.createElement("mark");
          mark.className = "s-hit";
          mark.textContent = text.slice(idx, idx + q.length);
          frag.appendChild(mark);
          marks.push(mark);
          last = idx + q.length;
          idx = lower.indexOf(q, last);
        }
        if (last < text.length) frag.appendChild(document.createTextNode(text.slice(last)));
        node.parentNode.replaceChild(frag, node);
      });
    });
    return marks;
  }

  window.DiffRender = { render, applySearch, clearSearch };
})();
