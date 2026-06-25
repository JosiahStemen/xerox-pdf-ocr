const PSM_MODES = {
  document: "3",
  sparse: "11",
  "single-block": "6",
};

export function preprocessCanvas(source, mode) {
  const canvas = document.createElement("canvas");
  canvas.width = source.width;
  canvas.height = source.height;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(source, 0, 0);

  if (mode === "none") return canvas;

  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = imageData.data;

  for (let i = 0; i < data.length; i += 4) {
    const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
    data[i] = data[i + 1] = data[i + 2] = gray;
  }

  if (mode === "full" || mode === "contrast") {
    const contrast = mode === "full" ? 1.4 : 1.25;
    const factor = (259 * (contrast * 80 + 255)) / (255 * (259 - contrast * 80));
    for (let i = 0; i < data.length; i += 4) {
      data[i] = clamp(factor * (data[i] - 128) + 128);
      data[i + 1] = clamp(factor * (data[i + 1] - 128) + 128);
      data[i + 2] = clamp(factor * (data[i + 2] - 128) + 128);
    }
  }

  if (mode === "full") {
    applySharpen(imageData);
    applyAdaptiveThreshold(imageData);
  }

  ctx.putImageData(imageData, 0, 0);
  return canvas;
}

function clamp(v) {
  return Math.max(0, Math.min(255, Math.round(v)));
}

function applySharpen(imageData) {
  const { data, width, height } = imageData;
  const copy = new Uint8ClampedArray(data);
  const kernel = [0, -1, 0, -1, 5, -1, 0, -1, 0];

  for (let y = 1; y < height - 1; y++) {
    for (let x = 1; x < width - 1; x++) {
      let sum = 0;
      let ki = 0;
      for (let ky = -1; ky <= 1; ky++) {
        for (let kx = -1; kx <= 1; kx++) {
          const idx = ((y + ky) * width + (x + kx)) * 4;
          sum += copy[idx] * kernel[ki++];
        }
      }
      const idx = (y * width + x) * 4;
      data[idx] = data[idx + 1] = data[idx + 2] = clamp(sum);
    }
  }
}

function applyAdaptiveThreshold(imageData) {
  const { data, width, height } = imageData;
  const blockSize = 15;
  const C = 10;
  const gray = new Float32Array(width * height);

  for (let i = 0, p = 0; i < data.length; i += 4, p++) {
    gray[p] = data[i];
  }

  const integral = buildIntegral(gray, width, height);

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const x1 = Math.max(0, x - blockSize);
      const y1 = Math.max(0, y - blockSize);
      const x2 = Math.min(width - 1, x + blockSize);
      const y2 = Math.min(height - 1, y + blockSize);
      const count = (x2 - x1 + 1) * (y2 - y1 + 1);
      const sum = rectSum(integral, width, x1, y1, x2, y2);
      const mean = sum / count;
      const idx = (y * width + x) * 4;
      const val = gray[y * width + x] < mean - C ? 0 : 255;
      data[idx] = data[idx + 1] = data[idx + 2] = val;
    }
  }
}

function buildIntegral(gray, width, height) {
  const integral = new Float64Array((width + 1) * (height + 1));
  for (let y = 0; y < height; y++) {
    let rowSum = 0;
    for (let x = 0; x < width; x++) {
      rowSum += gray[y * width + x];
      const idx = (y + 1) * (width + 1) + (x + 1);
      integral[idx] = rowSum + integral[y * (width + 1) + (x + 1)];
    }
  }
  return integral;
}

function rectSum(integral, width, x1, y1, x2, y2) {
  const w = width + 1;
  const A = integral[y1 * w + x1];
  const B = integral[y1 * w + (x2 + 1)];
  const C = integral[(y2 + 1) * w + x1];
  const D = integral[(y2 + 1) * w + (x2 + 1)];
  return D - B - C + A;
}

export async function extractTextFromPdf(pdfFile, options, onProgress) {
  const pdfjs = await import("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.min.mjs");
  pdfjs.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.worker.min.mjs";

  const arrayBuffer = await pdfFile.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
  const numPages = pdf.numPages;
  const scale = parseFloat(options.renderScale);
  const psm = PSM_MODES[options.ocrMode] || "3";

  let worker = null;
  const pages = [];
  let fullText = "";

  try {
    worker = await Tesseract.createWorker("eng", 1, {
      logger: (m) => {
        if (m.status === "recognizing text" && onProgress) {
          const pageProgress = m.progress || 0;
          onProgress({
            phase: "ocr",
            page: pages.length + 1,
            totalPages: numPages,
            pageProgress,
          });
        }
      },
    });

    await worker.setParameters({
      tessedit_pageseg_mode: psm,
      preserve_interword_spaces: "1",
      tessedit_char_whitelist: "",
    });

    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

      onProgress?.({ phase: "render", page: pageNum, totalPages: numPages, pageProgress: 0 });

      const page = await pdf.getPage(pageNum);
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport }).promise;

      const processed = preprocessCanvas(canvas, options.preprocess);

      onProgress?.({ phase: "ocr", page: pageNum, totalPages: numPages, pageProgress: 0 });

      const result = await worker.recognize(processed);
      const pageText = postProcessOcrText(result.data.text);
      const structured = buildStructuredFromHocr(result.data);

      pages.push({ pageNum, text: pageText, structured, confidence: result.data.confidence });
      fullText += (pageNum > 1 ? "\n\n--- Page " + pageNum + " ---\n\n" : "") + pageText;

      onProgress?.({ phase: "ocr", page: pageNum, totalPages: numPages, pageProgress: 1 });
    }

    const structuredRows = mergeStructuredPages(pages);
    return { fullText: postProcessOcrText(fullText), pages, structuredRows };
  } finally {
    if (worker) await worker.terminate();
  }
}

function postProcessOcrText(text) {
  return text
    .replace(/\r\n/g, "\n")
    .replace(/[|¦]/g, "I")
    .replace(/(\d)[oO](\d)/g, "$10$2")
    .replace(/([a-zA-Z])[|¦]([a-zA-Z])/g, "$1l$2")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/^\s+|\s+$/gm, (m, offset, str) => {
      const line = str.slice(str.lastIndexOf("\n", offset) + 1, str.indexOf("\n", offset));
      return line.trim() === "" ? "" : m.trimEnd();
    })
    .trim();
}

function buildStructuredFromHocr(data) {
  const words = (data.words || []).filter((w) => w.text.trim() && w.confidence > 30);
  if (!words.length) {
    return data.text
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => line.split(/\s{2,}|\t/).map((c) => c.trim()));
  }

  const lines = groupWordsIntoLines(words);
  return lines.map((line) => clusterIntoColumns(line));
}

function groupWordsIntoLines(words) {
  const sorted = [...words].sort((a, b) => a.bbox.y0 - b.bbox.y0 || a.bbox.x0 - b.bbox.x0);
  const lines = [];
  const threshold = 12;

  for (const word of sorted) {
    const cy = (word.bbox.y0 + word.bbox.y1) / 2;
    let line = lines.find((l) => Math.abs(l.y - cy) < threshold);
    if (!line) {
      line = { y: cy, words: [] };
      lines.push(line);
    }
    line.words.push(word);
    line.y = (line.y * (line.words.length - 1) + cy) / line.words.length;
  }

  return lines
    .sort((a, b) => a.y - b.y)
    .map((l) => l.words.sort((a, b) => a.bbox.x0 - b.bbox.x0));
}

function clusterIntoColumns(words) {
  if (!words.length) return [];
  const cols = [];
  const gapThreshold = 25;

  for (const word of words) {
    const last = cols[cols.length - 1];
    if (!last || word.bbox.x0 - last.x1 > gapThreshold) {
      cols.push({ x0: word.bbox.x0, x1: word.bbox.x1, text: word.text });
    } else {
      last.text += " " + word.text;
      last.x1 = word.bbox.x1;
    }
  }

  return cols.map((c) => c.text.trim());
}

function mergeStructuredPages(pages) {
  const all = [];
  for (const p of pages) {
    for (const row of p.structured) {
      if (row.some((c) => c)) all.push(row);
    }
  }
  return all;
}

export async function tryExtractEmbeddedText(pdfFile) {
  try {
    const pdfjs = await import("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.min.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.worker.min.mjs";
    const pdf = await pdfjs.getDocument({ data: await pdfFile.arrayBuffer() }).promise;
    let text = "";
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const pageText = content.items.map((item) => item.str).join(" ");
      if (pageText.trim()) text += (i > 1 ? "\n\n" : "") + pageText;
    }
    return text.trim() || null;
  } catch {
    return null;
  }
}