/* ============================================================
   parsers.js — turn any supported file into plain text
   Exposes: window.DiffParse.parseFile(file) -> {name, text, kind}
            window.DiffParse.kindLabel(kind)
   ============================================================ */
(function () {
  const TEXT_EXT = new Set([
    "py","js","jsx","ts","tsx","java","c","cc","cpp","h","hpp","cs","go","rs",
    "rb","php","swift","kt","scala","sh","bash","zsh","sql","r","m","pl","lua",
    "html","htm","css","scss","sass","less","xml","json","yaml","yml","toml","ini",
    "cfg","conf","md","markdown","txt","text","log","csv","tsv","env","gitignore",
    "vue","svelte","dart","ex","exs","erl","clj","hs","jl","f90","vb","asm","bat",
    "ps1","tex","rst","properties","gradle","make","mk","dockerfile","tf"
  ]);

  function ext(name) {
    const i = name.lastIndexOf(".");
    return i === -1 ? "" : name.slice(i + 1).toLowerCase();
  }

  function kindLabel(kind) {
    return ({
      text: "Text / Code",
      csv: "CSV",
      sheet: "Spreadsheet",
      doc: "Word document",
      pdf: "PDF",
      notebook: "Jupyter notebook",
      paste: "Pasted text"
    })[kind] || "Text";
  }

  async function parseFile(file) {
    const e = ext(file.name);

    if (e === "xlsx" || e === "xls" || e === "xlsm") return parseSheet(file);
    if (e === "docx") return parseDocx(file);
    if (e === "pdf") return parsePdf(file);
    if (e === "ipynb") return parseNotebookJson(file);

    // HTML may be a Jupyter notebook export (nbconvert). Detect and, if so,
    // parse it as a notebook; otherwise fall through to plain-text handling.
    if (e === "html" || e === "htm") {
      const raw = await file.text();
      const nb = tryParseNotebookHtml(raw, file.name);
      if (nb) return nb;
      return { name: file.name, text: normalizeNewlines(raw), kind: "text" };
    }

    // Everything else: read as text. CSV gets a structured grid too.
    if (TEXT_EXT.has(e) || e === "") {
      const text = await file.text();
      if (e === "csv" || e === "tsv") {
        let sheets = null;
        try {
          if (typeof XLSX !== "undefined") {
            const wb = XLSX.read(text, { type: "string", raw: true, FS: e === "tsv" ? "\t" : "," });
            sheets = workbookToSheets(wb);
          }
        } catch (_) { sheets = null; }
        return { name: file.name, text: normalizeNewlines(text), kind: "csv", sheets };
      }
      return { name: file.name, text: normalizeNewlines(text), kind: "text" };
    }

    // Unknown extension — attempt text anyway, but flag if it looks binary.
    const text = await file.text();
    if (looksBinary(text)) {
      throw new Error(`"${file.name}" doesn't look like a supported text, spreadsheet, document, or PDF file.`);
    }
    return { name: file.name, text: normalizeNewlines(text), kind: "text" };
  }

  function normalizeNewlines(t) {
    return t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  }

  function looksBinary(t) {
    const sample = t.slice(0, 1000);
    let ctrl = 0;
    for (let i = 0; i < sample.length; i++) {
      const c = sample.charCodeAt(i);
      if (c === 0) return true;
      if (c < 9 || (c > 13 && c < 32)) ctrl++;
    }
    return ctrl / Math.max(1, sample.length) > 0.08;
  }

  /* ---------- Spreadsheets (xlsx/xls) via SheetJS ---------- */
  async function parseSheet(file) {
    if (typeof XLSX === "undefined") throw new Error("Spreadsheet parser failed to load.");
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    return { name: file.name, text: sheetsToText(wb), kind: "sheet", sheets: workbookToSheets(wb) };
  }

  /* Workbook -> [{name, grid:[[cell,...],...]}] (all cells stringified) */
  function workbookToSheets(wb) {
    return wb.SheetNames.map(name => ({
      name,
      grid: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: "" })
        .map(row => row.map(c => (c == null ? "" : String(c))))
    }));
  }

  function sheetsToText(wb) {
    const out = [];
    wb.SheetNames.forEach((name, idx) => {
      if (wb.SheetNames.length > 1) {
        if (idx > 0) out.push("");
        out.push(`──── Sheet: ${name} ────`);
      }
      const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name], { blankrows: false });
      out.push(csv.replace(/\r\n/g, "\n").replace(/\n+$/, ""));
    });
    return out.join("\n");
  }

  /* ---------- Word (.docx) via mammoth ---------- */
  async function parseDocx(file) {
    if (typeof mammoth === "undefined") throw new Error("Word document parser failed to load.");
    const buf = await file.arrayBuffer();
    const res = await mammoth.extractRawText({ arrayBuffer: buf });
    let text = res.value || "";
    // Collapse 3+ blank lines to a single blank line for readable diffs.
    text = text.replace(/\n{3,}/g, "\n\n");

    // Also produce formatted HTML for the visual (rendered-page) diff.
    let docHtml = null;
    try {
      const h = await mammoth.convertToHtml({ arrayBuffer: buf });
      docHtml = (h && h.value && h.value.trim()) ? h.value : null;
    } catch (_) { docHtml = null; }

    return { name: file.name, text: normalizeNewlines(text).trimEnd(), kind: "doc", docHtml };
  }

  /* ---------- PDF via pdf.js ----------
     One document session per file: extract text (reconstructed by
     y-position) AND render each page to an image with word-level
     bounding boxes for the visual diff, then destroy the document.
     Doing everything in a single session — and never keeping two
     documents open at once — keeps pdf.js's worker happy. */
  const PDF_RENDER_SCALE = 1.6;
  const PDF_MAX_VISUAL_PAGES = 60;   // beyond this, skip image rendering

  async function parsePdf(file) {
    if (typeof pdfjsLib === "undefined") throw new Error("PDF parser failed to load.");
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const out = [];
    const pages = [];
    const canRenderVisual = pdf.numPages <= PDF_MAX_VISUAL_PAGES;
    let visualOk = canRenderVisual;

    try {
      for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        const tc = await page.getTextContent();

        // --- reconstructed text (for text-mode diff + export) ---
        if (p > 1) out.push("");
        out.push(`──── Page ${p} ────`);
        const lines = new Map();
        for (const it of tc.items) {
          if (!it.str) continue;
          const y = Math.round(it.transform[5]);
          if (!lines.has(y)) lines.set(y, []);
          lines.get(y).push(it);
        }
        const ys = [...lines.keys()].sort((a, b) => b - a);
        for (const y of ys) {
          const row = lines.get(y)
            .sort((a, b) => a.transform[4] - b.transform[4])
            .map(i => i.str).join("")
            .replace(/\s+$/g, "");
          if (row.trim() !== "") out.push(row);
        }

        // --- rendered image + word boxes (for visual diff) ---
        if (visualOk) {
          try {
            const vp = page.getViewport({ scale: PDF_RENDER_SCALE });
            const canvas = document.createElement("canvas");
            canvas.width = Math.ceil(vp.width);
            canvas.height = Math.ceil(vp.height);
            const ctx = canvas.getContext("2d");
            ctx.fillStyle = "#ffffff";
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            await page.render({ canvasContext: ctx, viewport: vp }).promise;

            const words = [];
            for (const item of tc.items) {
              const str = item.str;
              if (!str || !str.trim()) continue;
              const tx = pdfjsLib.Util.transform(vp.transform, item.transform);
              const fontH = Math.hypot(tx[2], tx[3]);
              const totalW = item.width * vp.scale;
              const x0 = tx[4];
              const yTop = tx[5] - fontH;
              const charW = totalW / Math.max(1, str.length);
              const re = /\S+/g;
              let m;
              while ((m = re.exec(str))) {
                words.push({ str: m[0], x: x0 + m.index * charW, y: yTop, w: m[0].length * charW, h: fontH });
              }
            }

            pages.push({
              url: canvas.toDataURL("image/png"),
              vw: Math.ceil(vp.width),
              vh: Math.ceil(vp.height),
              words
            });
            canvas.width = canvas.height = 0;
          } catch (_) {
            // rendering failed — disable visual mode but keep text
            visualOk = false;
            pages.length = 0;
          }
        }
      }
    } finally {
      try { await pdf.destroy(); } catch (_) {}
    }

    return {
      name: file.name,
      text: out.join("\n"),
      kind: "pdf",
      pageCount: pdf.numPages,
      pdfPages: (visualOk && pages.length) ? pages : null
    };
  }

  /* ============================================================
     Jupyter notebooks (.ipynb JSON  &  nbconvert .html exports)
     Both are normalized to the same cell model so they can be
     compared against each other:
       nbCells: [{ type:'code'|'markdown', source, html?, outputs:[...] }]
       outputs: [{kind:'text', text} | {kind:'image', src} | {kind:'html', html}]
     ============================================================ */
  function stripAnsi(s) {
    return String(s).replace(/\x1b\[[0-9;]*m/g, "").replace(/\x1b\][^\x07]*\x07/g, "");
  }
  function joinMaybe(v) {
    return Array.isArray(v) ? v.join("") : (v == null ? "" : String(v));
  }
  function renderMarkdown(src) {
    if (typeof marked !== "undefined") {
      try { return marked.parse(src, { breaks: true, gfm: true }); } catch (_) {}
    }
    // minimal fallback
    return "<p>" + escHtml(src).replace(/\n{2,}/g, "</p><p>").replace(/\n/g, "<br>") + "</p>";
  }
  function escHtml(s) {
    return String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  async function parseNotebookJson(file) {
    let nb;
    try { nb = JSON.parse(await file.text()); }
    catch (_) { throw new Error(`"${file.name}" is not valid notebook JSON.`); }
    const rawCells = Array.isArray(nb.cells) ? nb.cells
      : (nb.worksheets && nb.worksheets[0] && nb.worksheets[0].cells) || [];
    const cells = [];
    for (const c of rawCells) {
      const type = c.cell_type === "markdown" ? "markdown"
                 : c.cell_type === "code" ? "code" : "raw";
      const source = joinMaybe(c.source || c.input);
      if (type === "markdown") {
        cells.push({ type: "markdown", source, html: renderMarkdown(source), outputs: [] });
      } else if (type === "code") {
        cells.push({ type: "code", source, outputs: extractIpynbOutputs(c.outputs || []) });
      } else {
        // raw cell — show as plain preformatted text
        cells.push({ type: "code", source, outputs: [], raw: true });
      }
    }
    return { name: file.name, text: notebookToText(cells), kind: "notebook", nbCells: cells };
  }

  function extractIpynbOutputs(outs) {
    const result = [];
    for (const o of outs) {
      const t = o.output_type;
      if (t === "stream") {
        result.push({ kind: "text", text: stripAnsi(joinMaybe(o.text)) });
      } else if (t === "error") {
        result.push({ kind: "text", text: stripAnsi(joinMaybe(o.traceback).replace(/\n$/, "")), err: true });
      } else if (t === "execute_result" || t === "display_data") {
        const d = o.data || {};
        if (d["image/png"]) {
          result.push({ kind: "image", src: "data:image/png;base64," + joinMaybe(d["image/png"]).replace(/\s/g, "") });
        } else if (d["image/jpeg"]) {
          result.push({ kind: "image", src: "data:image/jpeg;base64," + joinMaybe(d["image/jpeg"]).replace(/\s/g, "") });
        } else if (d["image/svg+xml"]) {
          result.push({ kind: "html", html: joinMaybe(d["image/svg+xml"]) });
        } else if (d["text/html"]) {
          result.push({ kind: "html", html: joinMaybe(d["text/html"]) });
        } else if (d["text/plain"]) {
          result.push({ kind: "text", text: stripAnsi(joinMaybe(d["text/plain"])) });
        }
      }
    }
    return result;
  }

  /* Detect + parse an nbconvert HTML export. Returns null if it doesn't
     look like a notebook (so the caller treats it as ordinary HTML). */
  function tryParseNotebookHtml(html, name) {
    let doc;
    try { doc = new DOMParser().parseFromString(html, "text/html"); }
    catch (_) { return null; }

    // Lab/notebook 7 template uses .jp-Cell; classic uses .cell.code_cell etc.
    let cellEls = doc.querySelectorAll(".jp-Notebook .jp-Cell, .jp-Cell");
    let template = "lab";
    if (!cellEls.length) {
      cellEls = doc.querySelectorAll("#notebook .cell, .cell.code_cell, .cell.text_cell, div.cell");
      template = "classic";
    }
    if (!cellEls.length) return null;

    const cells = [];
    cellEls.forEach(el => {
      const isMd = template === "lab"
        ? el.classList.contains("jp-MarkdownCell")
        : el.classList.contains("text_cell");
      const isCode = template === "lab"
        ? el.classList.contains("jp-CodeCell")
        : el.classList.contains("code_cell");

      if (isMd) {
        const md = template === "lab"
          ? el.querySelector(".jp-RenderedMarkdown")
          : el.querySelector(".text_cell_render, .rendered_html");
        const node = md || el;
        cells.push({ type: "markdown", source: textOf(node), html: node.innerHTML, outputs: [] });
      } else if (isCode || !isMd) {
        // code cell (or unknown → treat as code)
        const inputEl = template === "lab"
          ? el.querySelector(".jp-InputArea-editor, .jp-CodeMirrorEditor, .highlight, pre")
          : el.querySelector(".input_area pre, .input_area, .highlight, pre");
        const source = inputEl ? textOf(inputEl).replace(/\n$/, "") : "";
        const outEls = template === "lab"
          ? el.querySelectorAll(".jp-OutputArea-output")
          : el.querySelectorAll(".output_area .output_subarea, .output_subarea");
        const outputs = [];
        outEls.forEach(o => {
          const img = o.querySelector("img");
          if (img && img.getAttribute("src")) { outputs.push({ kind: "image", src: img.getAttribute("src") }); return; }
          const table = o.querySelector("table");
          if (table) { outputs.push({ kind: "html", html: table.outerHTML }); return; }
          const pre = o.querySelector("pre");
          const txt = pre ? textOf(pre) : textOf(o);
          if (txt.trim()) outputs.push({ kind: "text", text: stripAnsi(txt.replace(/\n$/, "")) });
        });
        if (source.trim() === "" && outputs.length === 0) return; // skip empties
        cells.push({ type: "code", source, outputs });
      }
    });

    if (!cells.length) return null;
    return { name, text: notebookToText(cells), kind: "notebook", nbCells: cells };
  }

  function textOf(node) {
    return (node.textContent || "").replace(/\u00a0/g, " ");
  }

  /* Plain-text rendering of a notebook for the Text-mode fallback diff. */
  function notebookToText(cells) {
    const out = [];
    cells.forEach((c, i) => {
      if (i > 0) out.push("");
      if (c.type === "markdown") {
        out.push("# ── Markdown cell ──");
        out.push(c.source.trimEnd());
      } else {
        out.push("# ── Code cell ──");
        out.push(c.source.trimEnd());
        c.outputs.forEach(o => {
          if (o.kind === "text") out.push(o.text.split("\n").map(l => "  | " + l).join("\n"));
          else if (o.kind === "image") out.push("  [image output]");
          else if (o.kind === "html") out.push("  [rich output]");
        });
      }
    });
    return out.join("\n");
  }

  window.DiffParse = { parseFile, kindLabel, ext };
})();
