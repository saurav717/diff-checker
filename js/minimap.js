/* ============================================================
   minimap.js — a transparent overview "minimap" ruler down the
   middle of the diff view. It marks where additions (green),
   deletions (red) and modifications (amber) occur across the WHOLE
   document, with a draggable viewport thumb, so you can see at a
   glance where the changes are and jump straight to them.

   Works for every diff type (code/text, spreadsheet grid, PDF, Word,
   notebook) by scanning the rendered output for known change markers.

   Lives in the diff view's parent (not inside the scroll element) so
   re-renders that replace the scroll element's innerHTML don't wipe it.

   Exposes: window.DiffMinimap.mount(scrollEl)
            window.DiffMinimap.refresh()
   ============================================================ */
(function () {
  const SOURCES = [
    ['[data-type="add"]', "add"], ['[data-type="del"]', "del"], ['[data-type="mod"]', "mod"],
    ["mark.dv-add", "add"], ["mark.dv-del", "del"],
    [".pdf-hl.add", "add"], [".pdf-hl.del", "del"], [".pdf-hl.mod", "mod"],
    [".nb-out-rich.dv-add", "add"], [".nb-out-rich.dv-del", "del"],
    [".nb-row.has-change", "mod"]
  ];
  const COLORS = { add: "#1f883d", del: "#cf222e", mod: "#b5740a" };
  const W = 22;

  let scrollEl = null, host = null, wrap = null, canvas = null, thumb = null;
  let markers = [], rafId = 0, mounted = false, placement = "center";

  function position() {
    if (!wrap || !scrollEl) return;
    const hr = host.getBoundingClientRect(), sr = scrollEl.getBoundingClientRect();
    wrap.style.top = (sr.top - hr.top) + "px";
    const cx = placement === "right"
      ? (scrollEl.clientWidth - W - 6)        // right edge (before native scrollbar)
      : (scrollEl.clientWidth / 2 - W / 2);   // centre (side-by-side)
    wrap.style.left = (sr.left - hr.left + cx) + "px";
    wrap.style.height = scrollEl.clientHeight + "px";
  }

  function collect() {
    markers = [];
    if (!scrollEl) return;
    const top0 = scrollEl.getBoundingClientRect().top;
    const sTop = scrollEl.scrollTop;
    const sh = scrollEl.scrollHeight || 1;
    for (const [sel, type] of SOURCES) {
      let nodes;
      try { nodes = scrollEl.querySelectorAll(sel); } catch (_) { continue; }
      nodes.forEach(el => {
        const r = el.getBoundingClientRect();
        if (!r.height && !r.width) return;            // skip hidden
        const y = (r.top - top0) + sTop + r.height / 2;
        markers.push({ frac: Math.max(0, Math.min(1, y / sh)), type });
      });
    }
  }

  function draw() {
    const h = scrollEl.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = W * dpr; canvas.height = h * dpr;
    canvas.style.width = W + "px"; canvas.style.height = h + "px";
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, h);
    // a faint center guide line
    ctx.fillStyle = "rgba(120,130,150,0.14)";
    ctx.fillRect(W / 2 - 0.5, 0, 1, h);
    for (const m of markers) {
      ctx.fillStyle = COLORS[m.type] || COLORS.mod;
      const y = m.frac * h;
      ctx.fillRect(3, Math.max(0, y - 1.5), W - 6, 3);
    }
  }

  function updateThumb() {
    const sh = scrollEl.scrollHeight || 1, ch = scrollEl.clientHeight;
    const ratio = Math.min(1, ch / sh);
    thumb.style.height = Math.max(26, ratio * ch) + "px";
    thumb.style.top = (scrollEl.scrollTop / sh) * ch + "px";
  }

  function refresh() {
    if (!mounted) return;
    // Hide when there's nothing to scroll/compare (e.g. upload screen).
    const visible = scrollEl.clientHeight > 0 && scrollEl.scrollHeight > scrollEl.clientHeight + 4;
    position();
    collect();
    wrap.style.display = (visible && markers.length) ? "block" : "none";
    if (wrap.style.display === "none") return;
    draw();
    updateThumb();
  }

  function jumpTo(clientY) {
    const r = canvas.getBoundingClientRect();
    const frac = Math.max(0, Math.min(1, (clientY - r.top) / r.height));
    scrollEl.scrollTop = frac * scrollEl.scrollHeight - scrollEl.clientHeight / 2;
  }

  function mount(el) {
    scrollEl = el;
    host = el.parentElement;
    if (getComputedStyle(host).position === "static") host.style.position = "relative";

    wrap = document.createElement("div");
    wrap.className = "diff-minimap";
    canvas = document.createElement("canvas");
    canvas.className = "mm-canvas";
    thumb = document.createElement("div");
    thumb.className = "mm-thumb";
    wrap.appendChild(canvas);
    wrap.appendChild(thumb);
    host.appendChild(wrap);
    mounted = true;

    // click / drag anywhere on the ruler scrolls there
    const startDrag = (e) => {
      e.preventDefault();
      jumpTo(e.clientY);
      const mv = ev => jumpTo(ev.clientY);
      const up = () => { document.removeEventListener("mousemove", mv); document.removeEventListener("mouseup", up); };
      document.addEventListener("mousemove", mv);
      document.addEventListener("mouseup", up);
    };
    canvas.addEventListener("mousedown", startDrag);
    thumb.addEventListener("mousedown", startDrag);

    let tRaf = 0;
    scrollEl.addEventListener("scroll", () => {
      if (tRaf) return;
      tRaf = requestAnimationFrame(() => { tRaf = 0; if (wrap.style.display !== "none") { position(); updateThumb(); } });
    });
    window.addEventListener("resize", () => {
      if (rafId) return;
      rafId = requestAnimationFrame(() => { rafId = 0; refresh(); });
    });
  }

  window.DiffMinimap = { mount, refresh, setPlacement(p) { placement = p === "right" ? "right" : "center"; if (mounted) position(); } };
})();
