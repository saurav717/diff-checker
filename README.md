# Diff Checker

A side-by-side file comparison tool that runs entirely in your browser. Upload, drop, or paste two files and see exactly what changed. Supports code, text, Markdown, CSV/TSV, XLSX, DOCX, and PDF.

## Live site

Once deployed (see below), it will be available at:
`https://<your-username>.github.io/diff-checker/`

## Local use

Just open `index.html` in any modern browser — no build step or server required.

## Tech

- Pure HTML/CSS/JS, no framework
- Parsing & diffing via CDN libraries (`diff`, `xlsx`, `mammoth`, `pdf.js`)
- All processing happens client-side; no files are uploaded anywhere
