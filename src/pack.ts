/**
 * Building the case pack: Office → PDF conversion, merging and bookmarks.
 * Everything runs in the browser through the Nutrient Web SDK (WebAssembly); no server is involved.
 */
import NutrientViewer from "@nutrient-sdk/viewer";
import type { DocumentStore } from "./sharepoint";

type Instance = Awaited<ReturnType<typeof NutrientViewer.load>>;

export const sdkBaseUrl = `${window.location.origin}${import.meta.env.BASE_URL}`;
export const licenseKey: string | undefined = import.meta.env.VITE_NUTRIENT_LICENSE_KEY || undefined;

/**
 * Office engine variant. "faster" (SDK default) downloads a ~18 MB (compressed) ahead-of-time compiled
 * module and converts quicker; "smaller" downloads ~3.5 MB but converts slower. Override with ?engine=smaller|faster.
 */
const engineParam = new URLSearchParams(window.location.search).get("engine");
export const processorEngine =
  engineParam === "smaller" ? NutrientViewer.ProcessorEngine.smallerSize : NutrientViewer.ProcessorEngine.fasterProcessing;
/** SDK options shared by every load / convert / preload call. */
export const sdkConfig = { baseUrl: sdkBaseUrl, licenseKey, processorEngine };

export interface PackSource {
  name: string;
  path: string;
  size: number;
  kind: "docx" | "xlsx" | "pdf" | "other";
  /** First page of this document in the combined pack (0-based). */
  startPage: number;
  pageCount: number;
  convertMs: number;
}

export interface BuildResult {
  sources: PackSource[];
  combined: ArrayBuffer;
  totalPages: number;
  timings: { downloadMs: number; convertMs: number; mergeMs: number; totalMs: number; findMs?: number };
}

export function fileKind(name: string): PackSource["kind"] {
  const ext = name.toLowerCase().split(".").pop();
  if (ext === "docx" || ext === "doc") return "docx";
  if (ext === "xlsx" || ext === "xls") return "xlsx";
  if (ext === "pdf") return "pdf";
  return "other";
}

// ------------------------------------------------------------------ engine warm-up
/**
 * The SDK loads two WebAssembly engines lazily: the PDF core (~11 MB) on the first document load and
 * the Office engine (~80 MB, .NET runtime) on the first Word/Excel conversion. Without a warm-up the
 * user pays that download + compile on their first click.
 *
 * warmUpEngines() runs once when the page is idle: it preloads the PDF worker and converts a tiny
 * blank DOCX, which makes the Office engine fetch, compile and initialise exactly the files it needs
 * (the file set varies by browser, so hand-listing URLs to prefetch would be brittle).
 * Real conversions wait for the warm-up so the two never run concurrently.
 */
let engineReady: Promise<void> | null = null;

export function warmUpEngines(onState: (state: "loading" | "ready" | "failed", ms: number) => void): void {
  if (engineReady) return;
  const t0 = performance.now();
  onState("loading", 0);
  // Fire-and-forget: a later NutrientViewer.load() reuses this request rather than starting another.
  void NutrientViewer.preloadWorker({ ...sdkConfig } as Parameters<typeof NutrientViewer.preloadWorker>[0]).catch(() => {});
  engineReady = fetch(`${import.meta.env.BASE_URL}warmup.docx`)
    .then((r) => r.arrayBuffer())
    .then((docx) => convertOffice(docx))
    .then(
      () => onState("ready", performance.now() - t0),
      (e) => { console.warn("Engine warm-up failed; first conversion will load on demand.", e); onState("failed", performance.now() - t0); },
    );
}

function convertOffice(data: ArrayBuffer): Promise<ArrayBuffer> {
  return NutrientViewer.convertToPDF(
    { document: data, ...sdkConfig } as Parameters<typeof NutrientViewer.convertToPDF>[0],
    null,
    { splitExcelSheetsIntoPages: true, spreadsheetRenderOnlyPrintArea: true },
  );
}

/** Converts one file to PDF. PDFs pass straight through; Word and Excel use the SDK's built-in Office engine. */
export async function toPdf(name: string, data: ArrayBuffer): Promise<ArrayBuffer> {
  const kind = fileKind(name);
  if (kind === "pdf") return data;
  if (kind === "other") throw new Error(`Unsupported file type: ${name}`);
  if (engineReady) await engineReady;
  return convertOffice(data);
}

/**
 * Downloads the selected files from the store (into memory), converts each to PDF, then merges them
 * in a headless SDK instance and adds one bookmark per source document. Returns the combined PDF.
 */
/** Converted PDFs held in memory, keyed by store path, so a document previewed in step 1 isn't converted twice. */
export type PdfCache = Map<string, { pdf: ArrayBuffer; convertMs: number }>;

/** Downloads one file (into memory) and returns it as PDF, using the cache when it was already converted. */
export async function getPdf(
  store: DocumentStore,
  file: { name: string; path: string },
  cache: PdfCache,
): Promise<{ pdf: ArrayBuffer; convertMs: number; downloadMs: number }> {
  const hit = cache.get(file.path);
  if (hit) return { pdf: hit.pdf.slice(0), convertMs: hit.convertMs, downloadMs: 0 };
  const d0 = performance.now();
  const data = await store.download(file.path);
  const downloadMs = performance.now() - d0;
  const c0 = performance.now();
  const pdf = await toPdf(file.name, data);
  const convertMs = performance.now() - c0;
  cache.set(file.path, { pdf: pdf.slice(0), convertMs });
  return { pdf, convertMs, downloadMs };
}

export async function buildPack(
  store: DocumentStore,
  files: { name: string; path: string; size: number }[],
  onProgress: (msg: string) => void,
  cache: PdfCache = new Map(),
): Promise<BuildResult> {
  const t0 = performance.now();
  let downloadMs = 0;
  let convertMs = 0;

  const pdfs: { file: (typeof files)[number]; pdf: ArrayBuffer; ms: number }[] = [];
  for (const [i, file] of files.entries()) {
    onProgress(`${cache.has(file.path) ? "Using converted" : "Downloading and converting"} ${i + 1}/${files.length}: ${file.name}`);
    const r = await getPdf(store, file, cache);
    downloadMs += r.downloadMs;
    convertMs += r.convertMs;
    pdfs.push({ file, pdf: r.pdf, ms: r.convertMs });
  }

  onProgress("Combining into one pack…");
  const m0 = performance.now();
  let instance: Instance | undefined;
  try {
    // Headless: no UI is rendered, the SDK is used purely as a PDF processing engine.
    instance = await NutrientViewer.load({
      ...sdkConfig,
      document: pdfs[0].pdf,
      headless: true,
    });

    const sources: PackSource[] = [];
    sources.push(sourceFor(pdfs[0], 0, instance.totalPageCount));
    for (const entry of pdfs.slice(1)) {
      const start = instance.totalPageCount;
      await instance.applyOperations([
        {
          type: "importDocument",
          afterPageIndex: start - 1,
          treatImportedDocumentAsOnePage: false,
          document: new Blob([entry.pdf], { type: "application/pdf" }),
        },
      ]);
      sources.push(sourceFor(entry, start, instance.totalPageCount - start));
    }

    // One bookmark per source document so reviewers can navigate the pack.
    const bookmarks = sources.map(
      (s) =>
        new NutrientViewer.Bookmark({
          name: s.name.replace(/\.[^.]+$/, ""),
          action: new NutrientViewer.Actions.GoToAction({ pageIndex: s.startPage }),
        }),
    );
    await instance.create(bookmarks);

    const combined = await instance.exportPDF();
    const totalPages = instance.totalPageCount;
    const mergeMs = performance.now() - m0;
    return {
      sources,
      combined,
      totalPages,
      timings: { downloadMs, convertMs, mergeMs, totalMs: performance.now() - t0 },
    };
  } finally {
    if (instance) NutrientViewer.unload(instance);
  }
}

function sourceFor(
  entry: { file: { name: string; path: string; size: number }; ms: number },
  startPage: number,
  pageCount: number,
): PackSource {
  return {
    name: entry.file.name,
    path: entry.file.path,
    size: entry.file.size,
    kind: fileKind(entry.file.name),
    startPage,
    pageCount,
    convertMs: Math.round(entry.ms),
  };
}
