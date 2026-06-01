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
      const dbx = tryParseDatabricksHtml(raw, file.name);
      if (dbx) return dbx;
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

  /* Read the document's own default typography (font, size, line &
     paragraph spacing) out of word/styles.xml so the visual page diff
     renders with the same metrics Microsoft Word would use. mammoth
     strips this information from its HTML, so we recover it directly
     from the .docx package (which is a zip) via JSZip. */
  async function extractDocxStyle(buf) {
    const out = { font: null, sizePt: null, lineHeight: null, afterPt: null, beforePt: null };
    if (typeof JSZip === "undefined") return out;
    try {
      const zip = await JSZip.loadAsync(buf);
      const f = zip.file("word/styles.xml");
      if (!f) return out;
      const xml = await f.async("string");
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      const child = (parent, tag) => parent ? parent.getElementsByTagNameNS(NS, tag)[0] : null;
      const attr = (el, name) => (el ? el.getAttributeNS(NS, name) : null);

      const readRPr = (rPr) => {
        if (!rPr) return;
        const rFonts = child(rPr, "rFonts");
        const fam = rFonts && (attr(rFonts, "ascii") || attr(rFonts, "hAnsi") || attr(rFonts, "cs"));
        if (fam && !out.font) out.font = fam;
        const sz = child(rPr, "sz");
        const v = sz && attr(sz, "val");
        if (v && out.sizePt == null) out.sizePt = parseInt(v, 10) / 2; // val is in half-points
      };
      const readPPr = (pPr) => {
        if (!pPr) return;
        const sp = child(pPr, "spacing");
        if (!sp) return;
        const line = attr(sp, "line");
        const lineRule = attr(sp, "lineRule");
        if (line && out.lineHeight == null && (lineRule === "auto" || !lineRule)) {
          out.lineHeight = parseInt(line, 10) / 240; // 240ths of a line
        }
        const after = attr(sp, "after");
        if (after != null && out.afterPt == null) out.afterPt = parseInt(after, 10) / 20; // twips → pt
        const before = attr(sp, "before");
        if (before != null && out.beforePt == null) out.beforePt = parseInt(before, 10) / 20;
      };

      const root = doc.documentElement;
      const docDefaults = child(root, "docDefaults");
      if (docDefaults) {
        readRPr(child(child(docDefaults, "rPrDefault"), "rPr"));
        readPPr(child(child(docDefaults, "pPrDefault"), "pPr"));
      }
      // The default paragraph style (usually "Normal") overrides docDefaults.
      const styles = root.getElementsByTagNameNS(NS, "style");
      for (let i = 0; i < styles.length; i++) {
        const st = styles[i];
        const id = (attr(st, "styleId") || "").toLowerCase();
        const isDefaultPara = attr(st, "type") === "paragraph" &&
          (attr(st, "default") === "1" || attr(st, "default") === "true");
        if (id === "normal" || id === "standard" || isDefaultPara) {
          readRPr(child(st, "rPr"));
          readPPr(child(st, "pPr"));
          if (id === "normal" || id === "standard") break;
        }
      }
      return out;
    } catch (_) { return out; }
  }

  async function parseDocx(file) {
    if (typeof mammoth === "undefined") throw new Error("Word document parser failed to load.");
    const buf = await file.arrayBuffer();
    // Keep an untouched copy so the high-fidelity renderer (docx-preview) can
    // re-render the document with its real layout, fonts, colours and images.
    const docBuffer = buf.slice(0);
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

    // Recover the document's native typography (fallback render).
    const docStyle = await extractDocxStyle(buf);

    return { name: file.name, text: normalizeNewlines(text).trimEnd(), kind: "doc", docHtml, docStyle, docBuffer };
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
    // Keep an untouched copy so we can re-render pages at higher resolution
    // for OCR later (pdf.js may detach the buffer it's handed).
    const ocrBuffer = buf.slice(0);
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
      pdfPages: (visualOk && pages.length) ? pages : null,
      ocrBuffer
    };
  }

  /* Re-render PDF pages to high-resolution bitmaps for OCR. Small text
     (footnotes, fine print) is illegible to the recognizer at the modest
     scale we use for the on-screen diff, so we rasterise at a higher scale
     here. Returns [{ url, vw, vh }] in the hi-res pixel space. */
  async function renderPdfImages(buffer, scale) {
    if (typeof pdfjsLib === "undefined") throw new Error("PDF renderer unavailable.");
    const pdf = await pdfjsLib.getDocument({ data: buffer.slice(0) }).promise;
    const out = [];
    try {
      for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        // Cap the long edge so very large pages don't blow up memory.
        let s = scale;
        const base = page.getViewport({ scale: 1 });
        const longEdge = Math.max(base.width, base.height) * s;
        if (longEdge > 3400) s = scale * (3400 / longEdge);
        const vp = page.getViewport({ scale: s });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(vp.width);
        canvas.height = Math.ceil(vp.height);
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
        out.push({ url: canvas.toDataURL("image/png"), vw: canvas.width, vh: canvas.height });
        canvas.width = canvas.height = 0;
      }
    } finally {
      try { await pdf.destroy(); } catch (_) {}
    }
    return out;
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

  /* ---- Databricks notebook HTML export ----
     Databricks embeds the whole notebook as a base64 + URL-encoded JSON
     model in `var __DATABRICKS_NOTEBOOK_MODEL = '...'`. Each "command"
     is a cell; a leading %md marks markdown, other %-magics (%sql, %sh,
     %run, …) are code. Outputs live in command.results. */
  function tryParseDatabricksHtml(html, name) {
    if (html.indexOf("__DATABRICKS_NOTEBOOK_MODEL") === -1) return null;
    const m = html.match(/__DATABRICKS_NOTEBOOK_MODEL\s*=\s*'([^']+)'/);
    if (!m) return null;
    let model;
    try { model = JSON.parse(decodeURIComponent(atob(m[1]))); }
    catch (_) { return null; }
    const cmds = (model.commands || []).slice().sort((a, b) => (a.position || 0) - (b.position || 0));
    if (!cmds.length) return null;

    const cells = [];
    for (const c of cmds) {
      const raw = c.command || "";
      const mdm = raw.match(/^%md(?:-sandbox)?[ \t]*\n?/);
      const title = c.showCommandTitle && c.commandTitle ? c.commandTitle : "";
      if (mdm) {
        const src = raw.slice(mdm[0].length);
        cells.push({ type: "markdown", source: src, html: renderMarkdown(src), outputs: [], title });
      } else {
        const outputs = c.hideCommandResult ? [] : extractDbxOutputs(c.results);
        // keep the source verbatim (incl. any %sql/%sh/%run magic line)
        cells.push({ type: "code", source: raw, outputs, title });
      }
    }
    if (!cells.length) return null;
    return { name: name, text: notebookToText(cells), kind: "notebook", nbCells: cells };
  }

  function extractDbxOutputs(results) {
    if (!results) return [];
    const subs = results.type === "listResults" ? (results.data || []) : [results];
    const out = [];
    for (const r of subs) {
      if (!r) continue;
      const t = r.type;
      if (t === "ansi" || t === "text") {
        out.push({ kind: "text", text: stripAnsi(typeof r.data === "string" ? r.data : joinMaybe(r.data)) });
      } else if (t === "mimeBundle" && r.data) {
        const d = r.data;
        if (d["image/png"]) out.push({ kind: "image", src: "data:image/png;base64," + String(d["image/png"]).replace(/\s/g, "") });
        else if (d["image/jpeg"]) out.push({ kind: "image", src: "data:image/jpeg;base64," + String(d["image/jpeg"]).replace(/\s/g, "") });
        else if (d["text/html"]) out.push({ kind: "html", html: sanitizeHtml(joinMaybe(d["text/html"])) });
        else if (d["image/svg+xml"]) out.push({ kind: "html", html: joinMaybe(d["image/svg+xml"]) });
        else if (d["text/plain"]) out.push({ kind: "text", text: stripAnsi(joinMaybe(d["text/plain"])) });
      } else if (t === "htmlSandbox" || t === "html") {
        out.push({ kind: "html", html: sanitizeHtml(String(r.data)) });
      } else if (t === "image") {
        out.push({ kind: "image", src: String(r.data) });
      } else if (t === "table" && r.data) {
        out.push({ kind: "html", html: dbxTableHtml(r.data, r.schema || (r.arguments && r.arguments.schema)) });
      } else if (t === "error") {
        const txt = r.summary || r.cause || r.data || "";
        out.push({ kind: "text", text: stripAnsi(joinMaybe(txt)), err: true });
      } else if (typeof r.data === "string") {
        out.push({ kind: "text", text: stripAnsi(r.data) });
      }
    }
    return out;
  }

  function dbxTableHtml(data, schema) {
    const rows = Array.isArray(data) ? data : [];
    const head = schema && schema.map
      ? "<tr>" + schema.map(s => "<th>" + escHtml(s.name || s) + "</th>").join("") + "</tr>"
      : "";
    const body = rows.slice(0, 100).map(r => {
      const cells = Array.isArray(r) ? r : Object.values(r);
      return "<tr>" + cells.map(c => "<td>" + escHtml(c == null ? "" : String(c)) + "</td>").join("") + "</tr>";
    }).join("");
    return "<table>" + head + body + "</table>";
  }

  /* Strip <script>/<style> and inline event handlers from embedded output
     HTML so rendering a notebook can't execute arbitrary code. */
  function sanitizeHtml(html) {
    let s = String(html)
      .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, "")
      .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, "")
      .replace(/\son\w+\s*=\s*"[^"]*"/gi, "")
      .replace(/\son\w+\s*=\s*'[^']*'/gi, "")
      .replace(/\son\w+\s*=\s*[^\s>]+/gi, "");
    return s;
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

  window.DiffParse = { parseFile, kindLabel, ext, renderPdfImages };
})();
