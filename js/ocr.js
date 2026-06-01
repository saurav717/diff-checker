/* ============================================================
   ocr.js — optical character recognition fallback for PDFs whose
   text layer is missing or unreliable (scanned / flattened / filled
   forms). Runs entirely in the browser via Tesseract.js on the page
   bitmaps we already rendered at parse time.

   Returns, per page, a list of recognised words WITH bounding boxes
   in the SAME pixel coordinate system as the rendered page image
   (vw × vh) — so they drop straight into the existing word-diff /
   highlight pipeline — plus a reconstructed plain-text rendering.

   Exposes: window.OCR.available()
            window.OCR.run(pages, onProgress) -> { pages:[{words}], text }
            (onProgress(fraction, label) is called continuously)
   ============================================================ */
(function () {
  let workerPromise = null;
  let logProgress = null;   // live hook updated per call so the worker logger can report

  function available() { return typeof Tesseract !== "undefined"; }

  function humanStatus(s) {
    if (!s) return "working";
    if (s.indexOf("loading") === 0 || s.indexOf("initiali") === 0) return "loading OCR engine";
    if (s.indexOf("recogniz") === 0) return "recognising text";
    return s;
  }

  async function ensureWorker() {
    if (!available()) throw new Error("OCR engine failed to load.");
    if (!workerPromise) {
      // The logger fires continuously (engine download, init, and recognition
      // progress 0→1). Forward it to whatever callback the current run set.
      workerPromise = Tesseract.createWorker("eng", 1, {
        logger: (m) => { if (logProgress && m) logProgress(m.progress || 0, humanStatus(m.status)); }
      });
    }
    return workerPromise;
  }

  /* Warm the engine up (downloads model on first call) so the first page
     isn't penalised by init time with no feedback. */
  async function ensure(onStatus) {
    logProgress = onStatus || null;
    return ensureWorker();
  }

  function collectWords(data) {
    const out = [];
    const push = (w) => {
      if (!w || !w.text || !w.text.trim()) return;
      const b = w.bbox || w;
      if (b.x0 == null) return;
      out.push({ str: w.text.trim(), x: b.x0, y: b.y0, w: b.x1 - b.x0, h: b.y1 - b.y0 });
    };
    if (Array.isArray(data.words) && data.words.length) { data.words.forEach(push); return out; }
    if (Array.isArray(data.blocks)) {
      for (const bl of data.blocks)
        for (const p of (bl.paragraphs || []))
          for (const ln of (p.lines || []))
            for (const w of (ln.words || [])) push(w);
    }
    return out;
  }

  /* Reconstruct readable text from recognised words by grouping them
     into lines (by vertical position) and ordering left→right. */
  function reconstructText(words) {
    if (!words.length) return "";
    const ws = [...words].sort((a, b) => (a.y - b.y) || (a.x - b.x));
    const lines = [];
    let cur = null;
    for (const w of ws) {
      const cy = w.y + w.h / 2;
      if (cur && Math.abs(cy - cur.cy) <= Math.max(w.h, cur.h) * 0.6) {
        cur.items.push(w);
        cur.cy = (cur.cy * (cur.items.length - 1) + cy) / cur.items.length;
        cur.h = Math.max(cur.h, w.h);
      } else {
        cur = { cy, h: w.h, items: [w] };
        lines.push(cur);
      }
    }
    return lines.map(ln => ln.items.sort((a, b) => a.x - b.x).map(w => w.str).join(" ")).join("\n");
  }

  /* OCR every page of one file.
     onProgress(fraction 0..1, label) is called continuously. */
  async function run(pages, onProgress) {
    const total = pages.length || 1;
    // Route live engine/recognition progress into an overall fraction.
    logProgress = (frac, status) => {
      if (onProgress) onProgress((run._page + frac) / total, status);
    };
    const worker = await ensureWorker();

    const outPages = [];
    const textParts = [];
    for (let i = 0; i < pages.length; i++) {
      run._page = i;
      let words = [];
      try {
        const ret = await worker.recognize(pages[i].url, {}, { blocks: true });
        words = collectWords(ret.data || {});
      } catch (_) { words = []; }
      outPages.push({ words });
      textParts.push(`──── Page ${i + 1} ────\n${reconstructText(words)}`);
      if (onProgress) onProgress((i + 1) / total, "recognising text");
    }
    logProgress = null;
    return { pages: outPages, text: textParts.join("\n\n").trim() };
  }
  run._page = 0;

  window.OCR = { available, run, ensure, ensureWorker };
})();
