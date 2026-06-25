export function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

export function getSheetNames(workbook) {
  return workbook.SheetNames;
}

export function getSheetData(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
}

export function getColumnHeaders(rows) {
  if (!rows.length) return [];
  const first = rows[0];
  const hasHeaders = first.some((c) => typeof c === "string" && c.trim() && isNaN(Number(c)));
  if (hasHeaders) return first.map((h, i) => (h ? String(h) : `Column ${i + 1}`));
  const maxCols = Math.max(...rows.map((r) => r.length));
  return Array.from({ length: maxCols }, (_, i) => `Column ${i + 1}`);
}

export function getColumnValues(rows, colIndex, hasHeaders = true) {
  const start = hasHeaders ? 1 : 0;
  const values = [];
  for (let i = start; i < rows.length; i++) {
    const val = rows[i]?.[colIndex];
    if (val !== undefined && val !== null && String(val).trim() !== "") {
      values.push({ row: i + 1, value: String(val).trim() });
    }
  }
  return values;
}

export function compareWithPdfText(excelValues, pdfText, mode) {
  const normalizedPdf = normalizeText(pdfText);

  return excelValues.map(({ row, value }) => {
    const result = findMatch(value, pdfText, normalizedPdf, mode);
    return { row, excelValue: value, ...result };
  });
}

function normalizeText(text) {
  return text
    .toLowerCase()
    .replace(/[\s\-_.,;:'"()]/g, "")
    .replace(/[oO]/g, "0")
    .replace(/[l|¦]/g, "1");
}

function findMatch(value, rawText, normalizedPdf, mode) {
  if (!value) return { status: "missing", match: "", snippet: "" };

  if (mode === "exact") {
    if (rawText.includes(value)) {
      return { status: "found", match: value, snippet: getSnippet(rawText, value) };
    }
    return { status: "missing", match: "", snippet: "" };
  }

  if (mode === "normalized") {
    const normVal = normalizeText(value);
    if (normalizedPdf.includes(normVal)) {
      return { status: "found", match: value, snippet: findSnippetNormalized(rawText, value) };
    }
    return { status: "missing", match: "", snippet: "" };
  }

  const fuzzy = fuzzyFind(value, rawText);
  if (fuzzy.score >= 0.85) {
    return { status: "found", match: fuzzy.match, snippet: fuzzy.snippet };
  }
  if (fuzzy.score >= 0.65) {
    return { status: "fuzzy", match: fuzzy.match, snippet: fuzzy.snippet };
  }
  return { status: "missing", match: "", snippet: "" };
}

function getSnippet(text, value, radius = 40) {
  const idx = text.indexOf(value);
  if (idx === -1) return "";
  const start = Math.max(0, idx - radius);
  const end = Math.min(text.length, idx + value.length + radius);
  let snippet = text.slice(start, end);
  if (start > 0) snippet = "…" + snippet;
  if (end < text.length) snippet += "…";
  return snippet;
}

function findSnippetNormalized(rawText, value) {
  const words = rawText.split(/\s+/);
  const normVal = normalizeText(value);
  for (let i = 0; i < words.length; i++) {
    let combined = "";
    for (let j = i; j < Math.min(i + 8, words.length); j++) {
      combined += words[j];
      if (normalizeText(combined).includes(normVal)) {
        const snippet = words.slice(Math.max(0, i - 2), j + 3).join(" ");
        return snippet.length > 80 ? "…" + snippet.slice(0, 80) + "…" : snippet;
      }
    }
  }
  return "";
}

function fuzzyFind(value, text) {
  const lines = text.split("\n");
  let best = { score: 0, match: "", snippet: "" };

  for (const line of lines) {
    const tokens = line.split(/\s{2,}|\t|,\s*/).map((t) => t.trim()).filter(Boolean);
    for (const token of tokens) {
      const score = similarity(normalizeText(value), normalizeText(token));
      if (score > best.score) {
        best = { score, match: token, snippet: line.trim().slice(0, 100) };
      }
    }
    const lineScore = similarity(normalizeText(value), normalizeText(line));
    if (lineScore > best.score) {
      best = { score: lineScore, match: line.trim().slice(0, 60), snippet: line.trim().slice(0, 100) };
    }
  }

  return best;
}

function similarity(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const longer = a.length >= b.length ? a : b;
  const shorter = a.length >= b.length ? b : a;
  if (longer.includes(shorter) && shorter.length / longer.length > 0.6) {
    return 0.8 + 0.2 * (shorter.length / longer.length);
  }
  const dist = levenshtein(a, b);
  return 1 - dist / Math.max(a.length, b.length);
}

function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i - 1] === b[j - 1]
        ? dp[i - 1][j - 1]
        : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
    }
  }
  return dp[m][n];
}

export function buildCompareReport(results) {
  const header = "Row,Excel Value,Status,Match in PDF,Snippet";
  const rows = results.map((r) =>
    [r.row, `"${r.excelValue.replace(/"/g, '""')}"`, r.status, `"${(r.match || "").replace(/"/g, '""')}"`, `"${(r.snippet || "").replace(/"/g, '""')}"`].join(",")
  );
  return [header, ...rows].join("\n");
}