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

  /* ---------- PDF via pdf.js (reconstruct lines by y-position) ---------- */
  async function parsePdf(file) {
    if (typeof pdfjsLib === "undefined") throw new Error("PDF parser failed to load.");
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const out = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
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
    }
    return { name: file.name, text: out.join("\n"), kind: "pdf" };
  }

  window.DiffParse = { parseFile, kindLabel, ext };
})();
