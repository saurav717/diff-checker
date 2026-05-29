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
      paste: "Pasted text"
    })[kind] || "Text";
  }

  async function parseFile(file) {
    const e = ext(file.name);

    if (e === "xlsx" || e === "xls" || e === "xlsm") return parseSheet(file);
    if (e === "docx") return parseDocx(file);
    if (e === "pdf") return parsePdf(file);

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
    return { name: file.name, text: normalizeNewlines(text).trimEnd(), kind: "doc" };
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

  window.DiffParse = { parseFile, kindLabel, ext };
})();
