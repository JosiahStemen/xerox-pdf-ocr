import {
  extractTextFromPdf,
  tryExtractEmbeddedText,
} from "./ocr-engine.js";
import {
  parseExcelFile,
  getSheetNames,
  getSheetData,
  getColumnHeaders,
  getColumnValues,
  compareWithPdfText,
  buildCompareReport,
} from "./excel-compare.js";

const state = {
  pdfFile: null,
  excelFile: null,
  workbook: null,
  extractedText: "",
  structuredRows: [],
  compareResults: [],
  abortController: null,
};

const $ = (sel) => document.querySelector(sel);

function showToast(msg) {
  const toast = $("#toast");
  toast.textContent = msg;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2800);
}

function setupDropZone(zoneId, inputId, onFile) {
  const zone = $(zoneId);
  const input = $(inputId);

  zone.addEventListener("click", (e) => {
    if (e.target.closest("button") && !e.target.closest(`#${inputId.replace("-input", "-browse")}`)) return;
    if (!e.target.closest(".file-chip__remove")) input.click();
  });

  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("drag-over");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  });

  input.addEventListener("change", () => {
    if (input.files[0]) onFile(input.files[0]);
  });
}

function setPdfFile(file) {
  if (!file || file.type !== "application/pdf") {
    showToast("Please select a PDF file");
    return;
  }
  state.pdfFile = file;
  $("#pdf-name").textContent = file.name;
  $("#pdf-chip").hidden = false;
  updateStartButton();
}

function clearPdf() {
  state.pdfFile = null;
  state.extractedText = "";
  state.structuredRows = [];
  $("#pdf-input").value = "";
  $("#pdf-chip").hidden = true;
  $("#results-section").hidden = true;
  updateStartButton();
}

async function setExcelFile(file) {
  const valid = /\.(xlsx|xls|csv)$/i.test(file.name) ||
    file.type.includes("spreadsheet") ||
    file.type.includes("excel") ||
    file.type === "text/csv";
  if (!valid) {
    showToast("Please select an Excel or CSV file");
    return;
  }
  try {
    state.excelFile = file;
    state.workbook = await parseExcelFile(file);
    $("#excel-name").textContent = file.name;
    $("#excel-chip").hidden = false;
    $("#compare-section").hidden = false;
    populateSheetSelect();
    showToast("Excel loaded — run OCR first, then compare");
  } catch {
    showToast("Could not read Excel file");
  }
}

function clearExcel() {
  state.excelFile = null;
  state.workbook = null;
  $("#excel-input").value = "";
  $("#excel-chip").hidden = true;
  $("#compare-section").hidden = true;
}

function updateStartButton() {
  $("#start-btn").disabled = !state.pdfFile;
}

function populateSheetSelect() {
  const select = $("#sheet-select");
  select.innerHTML = "";
  for (const name of getSheetNames(state.workbook)) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    select.appendChild(opt);
  }
  populateColumnSelect();
  select.onchange = populateColumnSelect;
}

function populateColumnSelect() {
  const sheetName = $("#sheet-select").value;
  const rows = getSheetData(state.workbook, sheetName);
  const headers = getColumnHeaders(rows);
  const select = $("#column-select");
  select.innerHTML = "";
  headers.forEach((h, i) => {
    const opt = document.createElement("option");
    opt.value = i;
    opt.textContent = h;
    select.appendChild(opt);
  });
}

function setProgress(percent, label, detail = "") {
  $("#progress-section").hidden = false;
  $("#progress-fill").style.width = `${percent}%`;
  $("#progress-percent").textContent = `${Math.round(percent)}%`;
  $("#progress-label").textContent = label;
  $("#progress-detail").textContent = detail;
}

function hideProgress() {
  $("#progress-section").hidden = true;
}

function renderTextOutput() {
  $("#text-output").textContent = state.extractedText;
  renderStructuredTable();
  $("#results-section").hidden = false;
}

function renderStructuredTable() {
  const head = $("#structured-head");
  const body = $("#structured-body");
  head.innerHTML = "";
  body.innerHTML = "";

  if (!state.structuredRows.length) return;

  const maxCols = Math.max(...state.structuredRows.map((r) => r.length));
  for (let c = 0; c < maxCols; c++) {
    const th = document.createElement("th");
    th.textContent = `Col ${c + 1}`;
    head.appendChild(th);
  }

  for (const row of state.structuredRows) {
    const tr = document.createElement("tr");
    for (let c = 0; c < maxCols; c++) {
      const td = document.createElement("td");
      td.textContent = row[c] || "";
      tr.appendChild(td);
    }
    body.appendChild(tr);
  }
}

function renderCompareResults(results) {
  state.compareResults = results;
  const body = $("#compare-body");
  body.innerHTML = "";

  let found = 0, missing = 0, fuzzy = 0;
  for (const r of results) {
    if (r.status === "found") found++;
    else if (r.status === "fuzzy") fuzzy++;
    else missing++;

    const tr = document.createElement("tr");
    const badgeClass =
      r.status === "found" ? "status-badge--found" :
      r.status === "fuzzy" ? "status-badge--fuzzy" : "status-badge--missing";

    tr.innerHTML = `
      <td>${r.row}</td>
      <td>${escapeHtml(r.excelValue)}</td>
      <td><span class="status-badge ${badgeClass}">${r.status}</span></td>
      <td>${escapeHtml(r.match || r.snippet || "—")}</td>
    `;
    body.appendChild(tr);
  }

  const summary = $("#compare-summary");
  summary.hidden = false;
  summary.innerHTML = `
    <div class="summary-stat summary-stat--green">
      <div class="summary-stat__value">${found}</div>
      <div class="summary-stat__label">Found</div>
    </div>
    <div class="summary-stat summary-stat--amber">
      <div class="summary-stat__value">${fuzzy}</div>
      <div class="summary-stat__label">Fuzzy match</div>
    </div>
    <div class="summary-stat summary-stat--red">
      <div class="summary-stat__value">${missing}</div>
      <div class="summary-stat__label">Missing</div>
    </div>
  `;
}

function escapeHtml(str) {
  const d = document.createElement("div");
  d.textContent = str;
  return d.innerHTML;
}

async function startExtraction() {
  if (!state.pdfFile) return;

  state.abortController = new AbortController();
  $("#start-btn").disabled = true;
  $("#cancel-btn").hidden = false;
  $("#start-label").textContent = "Processing…";

  try {
    setProgress(2, "Checking for embedded text…");

    const embedded = await tryExtractEmbeddedText(state.pdfFile);
    if (embedded && embedded.length > 50) {
      const useEmbedded = confirm(
        "This PDF contains selectable text (not just images). Use the embedded text instead of OCR? It's faster and more accurate.\n\nClick OK to use embedded text, or Cancel to run OCR anyway."
      );
      if (useEmbedded) {
        state.extractedText = embedded;
        state.structuredRows = embedded.split("\n").map((l) => l.split(/\s{2,}|\t/).map((c) => c.trim()));
        setProgress(100, "Done", "Used embedded PDF text");
        renderTextOutput();
        showToast("Extracted embedded text successfully");
        return;
      }
    }

    setProgress(5, "Loading OCR engine…", "First run downloads language data (~15 MB)");

    const result = await extractTextFromPdf(
      state.pdfFile,
      {
        renderScale: $("#render-scale").value,
        ocrMode: $("#ocr-mode").value,
        preprocess: $("#preprocess").value,
        signal: state.abortController.signal,
      },
      (p) => {
        const base = ((p.page - 1) / p.totalPages) * 90 + 5;
        const pageSlice = (p.pageProgress || 0) * (90 / p.totalPages);
        const pct = base + pageSlice;
        const phase = p.phase === "render" ? "Rendering page" : "OCR on page";
        setProgress(pct, `${phase} ${p.page} of ${p.totalPages}`, "Xerox scan optimization active");
      }
    );

    state.extractedText = result.fullText;
    state.structuredRows = result.structuredRows;
    setProgress(100, "Complete", `${result.pages.length} page(s) processed`);
    renderTextOutput();
    showToast("Text extraction complete");

    if (state.workbook) runComparison();
  } catch (err) {
    if (err.name === "AbortError") {
      showToast("Cancelled");
    } else {
      console.error(err);
      showToast("OCR failed — try Maximum resolution or a smaller PDF");
    }
  } finally {
    hideProgress();
    $("#start-btn").disabled = false;
    $("#cancel-btn").hidden = true;
    $("#start-label").textContent = "Extract Text";
    state.abortController = null;
  }
}

function runComparison() {
  if (!state.workbook || !state.extractedText) {
    showToast("Need both PDF text and Excel file");
    return;
  }

  const sheetName = $("#sheet-select").value;
  const rows = getSheetData(state.workbook, sheetName);
  const headers = getColumnHeaders(rows);
  const hasHeaders = headers[0] !== "Column 1" || rows[0]?.some((c) => isNaN(Number(c)));
  const colIndex = parseInt($("#column-select").value, 10);
  const values = getColumnValues(rows, colIndex, hasHeaders);
  const mode = $("#match-mode").value;

  const results = compareWithPdfText(values, state.extractedText, mode);
  renderCompareResults(results);
  $("#compare-section").hidden = false;
}

function setupTabs() {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("tab--active"));
      tab.classList.add("tab--active");
      const isPlain = tab.dataset.tab === "plain";
      $("#text-output").hidden = !isPlain;
      $("#structured-wrap").hidden = isPlain;
    });
  });
}

function downloadText() {
  if (!state.extractedText) return;
  downloadBlob(state.extractedText, "extracted-text.txt", "text/plain");
}

function downloadCompare() {
  if (!state.compareResults.length) return;
  downloadBlob(buildCompareReport(state.compareResults), "comparison-report.csv", "text/csv");
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function init() {
  setupDropZone("#pdf-drop-zone", "#pdf-input", setPdfFile);
  setupDropZone("#excel-drop-zone", "#excel-input", setExcelFile);

  $("#pdf-browse").addEventListener("click", (e) => {
    e.stopPropagation();
    $("#pdf-input").click();
  });
  $("#excel-browse").addEventListener("click", (e) => {
    e.stopPropagation();
    $("#excel-input").click();
  });

  $("#pdf-clear").addEventListener("click", (e) => { e.stopPropagation(); clearPdf(); });
  $("#excel-clear").addEventListener("click", (e) => { e.stopPropagation(); clearExcel(); });

  $("#start-btn").addEventListener("click", startExtraction);
  $("#cancel-btn").addEventListener("click", () => state.abortController?.abort());
  $("#run-compare").addEventListener("click", runComparison);
  $("#copy-text").addEventListener("click", async () => {
    if (!state.extractedText) return;
    await navigator.clipboard.writeText(state.extractedText);
    showToast("Copied to clipboard");
  });
  $("#download-text").addEventListener("click", downloadText);
  $("#download-compare").addEventListener("click", downloadCompare);

  setupTabs();
}

init();