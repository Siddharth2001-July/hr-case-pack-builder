import NutrientViewer from "@nutrient-sdk/viewer";
import { MockSharePointStore, type DocumentStore, type StoreItem } from "./sharepoint";
import { buildPack, fileKind, getPdf, licenseKey, processorEngine, sdkConfig, warmUpEngines, type BuildResult, type PdfCache } from "./pack";
import { findPersonalData, listMarks, verifyRemoved, RULE_LABELS, type Mark, type RuleOptions } from "./redaction";
import { buildAudit } from "./audit";

type Instance = Awaited<ReturnType<typeof NutrientViewer.load>>;

// ------------------------------------------------------------------ DOM helpers
const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
function el<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, string> = {}, ...kids: (Node | string)[]) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) k === "class" ? (e.className = v) : e.setAttribute(k, v);
  e.append(...kids);
  return e;
}
const fmtSize = (n: number) => (n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(0)} KB` : `${(n / 1048576).toFixed(1)} MB`);
const fmtMs = (ms: number) => (ms < 1000 ? `${Math.round(ms)} ms` : `${(ms / 1000).toFixed(1)} s`);
const typeChip = (name: string) => {
  const k = name.endsWith(".json") ? "json" : fileKind(name);
  return el("span", { class: `type type-${k}` }, k.toUpperCase());
};

function status(msg: string, busy = false) {
  $("status").textContent = msg;
  document.querySelector(".statusbar")!.classList.toggle("busy", busy);
}

function setStep(active: "build" | "find" | "review" | "apply" | null, done: string[] = []) {
  for (const s of ["build", "find", "review", "apply"]) {
    const li = $(`step-${s}`);
    li.classList.toggle("active", s === active || (active === "review" && s === "apply"));
    li.classList.toggle("done", done.includes(s));
  }
}

// ------------------------------------------------------------------ state
let store: DocumentStore;
let caseId: string | null = null;
let caseFiles: (StoreItem & { selected: boolean })[] = [];
let built: BuildResult | null = null;
let viewer: Instance | null = null;
let startedAt = new Date();
let marksFound: Mark[] = [];
let removedByReviewer: Mark[] = [];
let ruleOpts: RuleOptions | null = null;
let busy = false;
/** Converted PDFs for the open case (memory only), shared by preview and build. */
let pdfCache: PdfCache = new Map();
let previewing: string | null = null;
/** Paths whose PDF conversion has finished in the background (shown as "ready" in the list). */
let converted = new Set<string>();

const viewerEl = $("viewer");

const TOOLBAR = new Set([
  "sidebar-thumbnails", "sidebar-bookmarks", "pager", "pan", "zoom-out", "zoom-in", "zoom-mode",
  "spacer", "search", "redact-text-highlighter", "redact-rectangle",
]);

async function showInViewer(pdf: ArrayBuffer, opts: { review: boolean }) {
  if (viewer) NutrientViewer.unload(viewer);
  viewer = null;
  $("viewer-empty").classList.add("hidden");
  viewer = await NutrientViewer.load({
    container: viewerEl,
    ...sdkConfig,
    document: pdf.slice(0),
    toolbarItems: NutrientViewer.defaultToolbarItems.filter((i) => TOOLBAR.has(i.type) && (opts.review || !i.type.startsWith("redact"))),
    initialViewState: new NutrientViewer.ViewState({
      sidebarMode: NutrientViewer.SidebarMode.BOOKMARKS,
      readOnly: !opts.review,
    }),
  });
  viewer.setAnnotationCreatorName(store.user.displayName);
  if (opts.review) {
    // Keep the review list in sync with what the reviewer does in the viewer itself.
    viewer.addEventListener("annotations.delete", (deleted) => {
      for (const a of deleted) {
        const m = marksFound.find((x) => x.id === a.id);
        if (m && !removedByReviewer.some((r) => r.id === m.id)) removedByReviewer.push(m);
      }
      void renderMarks();
    });
    viewer.addEventListener("annotations.create", () => void renderMarks());
  }
}

// ------------------------------------------------------------------ library pane
async function showCases() {
  caseId = null;
  $("file-panel").classList.add("hidden");
  $("back-to-cases").classList.add("hidden");
  const list = $("case-list");
  list.classList.remove("hidden");
  list.replaceChildren();
  for (const f of await store.listCaseFolders()) {
    const btn = el("button", { class: "case-card" }, el("div", { class: "folder" }, `📁 ${f.name}`), el("div", { class: "muted small" }, f.title));
    btn.onclick = () => openCase(f.name, f.title);
    list.append(btn);
  }
}

async function openCase(id: string, title: string) {
  if (busy) return;
  resetWorkflow();
  caseId = id;
  pdfCache = new Map();
  converted = new Set();
  previewing = null;
  $("case-list").classList.add("hidden");
  $("file-panel").classList.remove("hidden");
  $("back-to-cases").classList.remove("hidden");
  $("case-id").textContent = id;
  $("case-title").textContent = title;
  status(`Opening ${id}…`, true);
  const items = await store.listFolder(id);
  caseFiles = items.filter((i) => !i.isFolder && fileKind(i.name) !== "other").map((i) => ({ ...i, selected: true }));
  renderFiles();
  await renderOutputs();
  setStep("build");
  status(`${id}: ${caseFiles.length} documents`);
  void preconvert(id);
}

/**
 * Converts the case's Word/Excel files to PDF in the background, one at a time, as soon as the case
 * opens, so View and Build usually find the result ready. Stops if the user switches case.
 */
async function preconvert(id: string) {
  for (const f of [...caseFiles]) {
    if (caseId !== id) return;
    try {
      await getPdf(store, f, pdfCache);
      if (caseId !== id) return;
      converted.add(f.path);
      if (!busy) renderFiles();
    } catch (e) {
      console.warn(`Background conversion failed for ${f.name}`, e);
    }
  }
}

function renderFiles() {
  const ul = $("file-list");
  ul.replaceChildren();
  caseFiles.forEach((f, i) => {
    const cb = el("input", { type: "checkbox" }) as HTMLInputElement;
    cb.checked = f.selected;
    cb.disabled = busy || built !== null;
    cb.onchange = () => { f.selected = cb.checked; updateBuildButton(); };
    const up = el("button", { title: "Move up" }, "▲");
    const down = el("button", { title: "Move down" }, "▼");
    up.onclick = () => move(i, -1);
    down.onclick = () => move(i, 1);
    (up as HTMLButtonElement).disabled = (down as HTMLButtonElement).disabled = built !== null;
    const view = el("button", { class: "icon-btn" }, "View") as HTMLButtonElement;
    view.disabled = busy || built !== null;
    view.title = built ? "Start another pack to preview single documents" : "Preview this document";
    view.onclick = () => previewFile(f);
    const name = el("div", { class: "file-name clickable", title: `${f.name}\nClick to preview` }, f.name,
      el("div", { class: "file-meta" }, `${fmtSize(f.size)} · ${f.modifiedBy}`,
        ...(converted.has(f.path) ? [el("span", { class: "ready", title: "Converted to PDF in the background" }, " · ready")] : [])));
    name.onclick = () => { if (!view.disabled) void previewFile(f); };
    ul.append(
      el("li", { class: previewing === f.path ? "previewing" : "" }, cb, typeChip(f.name), name, view,
        el("div", { class: "order-btns" }, up, down)),
    );
  });
  updateBuildButton();
}

/** Shows one source document on its own (Word/Excel converted on the fly), read-only. */
async function previewFile(f: StoreItem) {
  if (busy || built) return;
  busy = true;
  previewing = f.path;
  renderFiles();
  try {
    status(`${converted.has(f.path) ? "Opening" : "Converting"} ${f.name}…`, true);
    const r = await getPdf(store, f, pdfCache);
    converted.add(f.path);
    await showInViewer(r.pdf, { review: false });
    status(`Previewing ${f.name} (read-only)${f.name.toLowerCase().endsWith(".pdf") ? "" : ` · converted in ${fmtMs(r.convertMs)}`}`);
  } catch (e) {
    console.error(e);
    status(`Preview failed: ${(e as Error).message}`);
  } finally {
    busy = false;
    renderFiles();
  }
}

function move(i: number, d: number) {
  const j = i + d;
  if (j < 0 || j >= caseFiles.length) return;
  [caseFiles[i], caseFiles[j]] = [caseFiles[j], caseFiles[i]];
  renderFiles();
}

function updateBuildButton() {
  const n = caseFiles.filter((f) => f.selected).length;
  const b = $<HTMLButtonElement>("btn-build");
  b.textContent = built ? "Pack built" : `Build case pack (${n} document${n === 1 ? "" : "s"})`;
  b.disabled = busy || n === 0 || built !== null;
}

async function renderOutputs(highlight: string[] = []) {
  const ul = $("output-list");
  const items = (await store.listFolder(`${caseId}/Case Pack`)).filter((i) => !i.isFolder);
  ul.replaceChildren();
  if (items.length === 0) { ul.append(el("li", { class: "empty" }, "Nothing saved yet")); return; }
  for (const it of items) {
    const open = el("button", { class: "icon-btn", title: "Open" }, "Open");
    const dl = el("button", { class: "icon-btn", title: "Download a copy" }, "⤓");
    open.onclick = () => openOutput(it);
    dl.onclick = () => downloadOutput(it);
    ul.append(
      el("li", {}, typeChip(it.name),
        el("div", { class: "file-name", title: it.name }, outputLabel(it.name),
          el("div", { class: "file-meta" }, `${fmtSize(it.size)} · ${it.modified.toLocaleTimeString("en-GB")} `,
            ...(highlight.includes(it.path) ? [el("span", { class: "new" }, "NEW")] : []))),
        open, dl),
    );
  }
}

/** Plain-English label for each saved output; the full file name is in the tooltip. */
function outputLabel(name: string): string {
  if (name.endsWith("_REDACTED.pdf")) return "Redacted pack (to share)";
  if (name.endsWith("_COMBINED.pdf")) return "Combined pack (unredacted)";
  if (name.endsWith("_audit.json")) return "Audit log";
  return name;
}

async function openOutput(it: StoreItem) {
  const data = await store.download(it.path);
  if (it.name.endsWith(".json")) {
    window.open(URL.createObjectURL(new Blob([data], { type: "application/json" })), "_blank");
    return;
  }
  previewing = null;
  renderFiles();
  status(`Viewing ${outputLabel(it.name)}: ${it.name} (read-only)`);
  await showInViewer(data, { review: false });
}

async function downloadOutput(it: StoreItem) {
  const data = await store.download(it.path);
  const a = el("a", { href: URL.createObjectURL(new Blob([data])), download: it.name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.getAttribute("href")!), 1000);
}

// ------------------------------------------------------------------ step 1: build
async function onBuild() {
  if (!caseId) return;
  const files = caseFiles.filter((f) => f.selected);
  busy = true;
  renderFiles();
  startedAt = new Date();
  try {
    previewing = null;
    built = await buildPack(store, files, (m) => status(m, true), pdfCache);
    status("Opening the pack…", true);
    const v0 = performance.now();
    await showInViewer(built.combined, { review: true });
    built.timings.totalMs += performance.now() - v0;
    const t = built.timings;
    $("build-result").classList.remove("hidden");
    $("build-result").replaceChildren(
      el("div", { class: "ok" }, `✓ ${built.sources.length} documents → ${built.totalPages} pages`),
      el("table", {},
        ...built.sources.map((s) => el("tr", {}, el("td", {}, s.name), el("td", { class: "muted" }, `${s.pageCount} p · ${fmtMs(s.convertMs)}`))),
        el("tr", {}, el("td", {}, el("b", {}, "Convert / combine / total")), el("td", {}, el("b", {}, `${fmtMs(t.convertMs)} / ${fmtMs(t.mergeMs)} / ${fmtMs(t.totalMs)}`))),
      ),
    );
    $<HTMLButtonElement>("btn-find").disabled = false;
    setStep("find", ["build"]);
    status(`Pack built: ${built.totalPages} pages in ${fmtMs(t.totalMs)}`);
  } catch (e) {
    console.error(e);
    built = null;
    status(`Build failed: ${(e as Error).message}`);
  } finally {
    busy = false;
    renderFiles();
  }
}

// ------------------------------------------------------------------ step 2: find
async function onFind() {
  if (!viewer) return;
  ruleOpts = {
    names: $<HTMLTextAreaElement>("names").value.split(/\n|,/).map((s) => s.trim()).filter(Boolean),
    matchNameParts: $<HTMLInputElement>("opt-parts").checked,
    suNumbers: $<HTMLInputElement>("opt-su").checked,
    emails: $<HTMLInputElement>("opt-email").checked,
    dob: $<HTMLInputElement>("opt-dob").checked,
    otherDates: $<HTMLInputElement>("opt-dates").checked,
  };
  busy = true;
  $<HTMLButtonElement>("btn-find").disabled = true;
  try {
    // Re-running replaces earlier rule-generated marks (hand-drawn marks are kept).
    const old = (await listMarks(viewer)).filter((m) => m.rule !== "manual").map((m) => m.id);
    if (old.length) await viewer.delete(old);
    removedByReviewer = [];
    const f0 = performance.now();
    marksFound = await findPersonalData(viewer, ruleOpts, (m) => status(m, true));
    if (built) built.timings.findMs = performance.now() - f0;
    await renderMarks();
    setStep("review", ["build", "find"]);
    $<HTMLButtonElement>("btn-apply").disabled = false;
    status(`${marksFound.length} items marked for review`);
  } catch (e) {
    console.error(e);
    status(`Search failed: ${(e as Error).message}`);
  } finally {
    busy = false;
    $<HTMLButtonElement>("btn-find").disabled = false;
    $("btn-find").textContent = "Search again";
  }
}

// ------------------------------------------------------------------ step 3: review
async function renderMarks() {
  if (!viewer) return;
  const marks = (await listMarks(viewer)).sort((a, b) => a.pageIndex - b.pageIndex);
  const counts = new Map<string, number>();
  marks.forEach((m) => counts.set(RULE_LABELS[m.rule], (counts.get(RULE_LABELS[m.rule]) ?? 0) + 1));
  $("rule-chips").replaceChildren(
    el("span", { class: "chip" }, el("b", {}, `${marks.length}`), " marked"),
    ...[...counts].map(([k, n]) => el("span", { class: "chip" }, `${k}: ${n}`)),
    ...(removedByReviewer.length ? [el("span", { class: "chip" }, `Kept by reviewer: ${removedByReviewer.length}`)] : []),
  );
  const ul = $("mark-list");
  ul.replaceChildren();
  if (marks.length === 0) { ul.append(el("li", { class: "empty" }, "No marks")); return; }
  for (const m of marks) {
    const keep = el("button", { class: "icon-btn", title: "Don't redact this" }, "Keep");
    keep.onclick = async (ev) => { ev.stopPropagation(); await viewer!.delete(m.id); };
    const li = el("li", {},
      el("div", { class: "mark-text", title: m.text }, m.text, el("div", { class: "mark-rule" }, RULE_LABELS[m.rule])),
      el("span", { class: "mark-page" }, `p${m.pageIndex + 1}`),
      keep);
    li.onclick = async () => {
      const a = (await viewer!.getAnnotations(m.pageIndex)).find((x) => x.id === m.id);
      if (a) { viewer!.jumpToRect(m.pageIndex, a.boundingBox); viewer!.setSelectedAnnotations(NutrientViewer.Immutable.List([m.id])); }
    };
    ul.append(li);
  }
}

// ------------------------------------------------------------------ step 4: apply & save
async function onApply() {
  if (!viewer || !built || !caseId || !ruleOpts) return;
  const toApply = await listMarks(viewer);
  $("confirm-text").textContent = `${toApply.length} marked item${toApply.length === 1 ? "" : "s"} will be redacted across ${built.totalPages} pages.`;
  const dlg = $<HTMLDialogElement>("confirm");
  dlg.showModal();
  const answer = await new Promise<string>((r) => dlg.addEventListener("close", () => r(dlg.returnValue), { once: true }));
  if (answer !== "ok") return;

  busy = true;
  $<HTMLButtonElement>("btn-apply").disabled = true;
  $<HTMLButtonElement>("btn-find").disabled = true;
  try {
    status("Applying redactions…", true);
    const a0 = performance.now();
    await viewer.applyRedactions();
    const applyMs = performance.now() - a0;

    status("Checking that redacted content is gone…", true);
    // Values the reviewer chose to keep elsewhere are excluded, since they are expected to remain.
    const kept = new Set(removedByReviewer.map((m) => m.text.toLowerCase()));
    const verification = await verifyRemoved(viewer, toApply.map((m) => m.text).filter((t) => !kept.has(t.toLowerCase())));

    const redacted = await viewer.exportPDF();
    const stamp = new Date().toISOString().slice(0, 16).replace(/[-:]/g, "").replace("T", "-");
    const base = `${caseId}/Case Pack/CasePack_${caseId}_${stamp}`;
    const outputs: { name: string; path: string; size: number }[] = [];
    const saved: string[] = [];

    status("Saving to SharePoint…", true);
    const s0 = performance.now();
    if ($<HTMLInputElement>("opt-save-combined").checked) {
      const it = await store.upload(`${base}_COMBINED.pdf`, built.combined, "application/pdf");
      outputs.push({ name: it.name, path: it.path, size: it.size });
    }
    const red = await store.upload(`${base}_REDACTED.pdf`, redacted, "application/pdf");
    outputs.push({ name: red.name, path: red.path, size: red.size });
    const saveMs = performance.now() - s0;

    const audit = buildAudit({
      caseId,
      user: store.user,
      startedAt,
      sources: built.sources,
      totalPages: built.totalPages,
      rules: { ...ruleOpts, names: ruleOpts.names.length },
      marksFound,
      marksRemovedByReviewer: removedByReviewer,
      marksApplied: toApply,
      verification: { checked: verification.checked, stillFound: verification.stillFound.length },
      outputs,
      timingsMs: Object.fromEntries(
        Object.entries({ ...built.timings, applyMs, saveMs }).map(([k, v]) => [k, Math.round(v)]),
      ),
    });
    const auditBytes = new TextEncoder().encode(JSON.stringify(audit, null, 2));
    const au = await store.upload(`${base}_audit.json`, auditBytes.buffer as ArrayBuffer, "application/json");
    saved.push(...outputs.map((o) => o.path), au.path);

    await renderOutputs(saved);
    const passed = verification.stillFound.length === 0;
    $("apply-result").classList.remove("hidden");
    $("apply-result").replaceChildren(
      el("div", { class: "ok" }, `✓ ${toApply.length} items redacted, ${outputs.length + 1} files saved to Case Pack/`),
      el("div", { class: passed ? "ok" : "bad" },
        passed
          ? `✓ Verified: none of the ${verification.checked} redacted values can be found in the output`
          : `⚠ ${verification.stillFound.length} redacted value(s) still found: ${verification.stillFound.join(", ")}`),
      el("div", { class: "muted" }, `Apply ${fmtMs(applyMs)} · save ${fmtMs(saveMs)}`),
    );
    $("btn-reset").classList.remove("hidden");
    setStep(null, ["build", "find", "review", "apply"]);
    status(`Saved ${caseId} case pack to SharePoint`);
    // Show the final redacted output, read-only.
    await showInViewer(redacted, { review: false });
  } catch (e) {
    console.error(e);
    status(`Apply failed: ${(e as Error).message}`);
    $<HTMLButtonElement>("btn-apply").disabled = false;
  } finally {
    busy = false;
  }
}

// ------------------------------------------------------------------ reset
function resetWorkflow() {
  if (viewer) NutrientViewer.unload(viewer);
  viewer = null;
  built = null;
  marksFound = [];
  removedByReviewer = [];
  ruleOpts = null;
  $("viewer-empty").classList.remove("hidden");
  for (const id of ["build-result", "apply-result", "btn-reset"]) $(id).classList.add("hidden");
  $<HTMLButtonElement>("btn-find").disabled = true;
  $("btn-find").textContent = "Find and mark";
  $<HTMLButtonElement>("btn-apply").disabled = true;
  $("rule-chips").replaceChildren();
  $("mark-list").replaceChildren(el("li", { class: "empty" }, "No marks yet"));
  setStep(null);
}

// ------------------------------------------------------------------ boot
async function main() {
  store = await MockSharePointStore.create();
  $("user").textContent = store.user.displayName;
  $("site-path").textContent = `${store.siteName} / ${store.libraryName}`;
  $("lib-title").textContent = store.libraryName;
  $("sdk-version").textContent = `Nutrient Web SDK ${NutrientViewer.version} · ${processorEngine === NutrientViewer.ProcessorEngine.smallerSize ? "compact" : "fast"} engine${licenseKey ? "" : " · trial mode"}`;
  $("btn-build").onclick = onBuild;
  $("btn-find").onclick = onFind;
  $("btn-apply").onclick = onApply;
  $("btn-reset").onclick = () => caseId && openCase(caseId, $("case-title").textContent ?? "");
  $("back-to-cases").onclick = () => { if (!busy) { resetWorkflow(); void showCases(); } };
  await showCases();
  setStep(null);
  status("Ready. Open a case to start.");
  // Start loading the PDF and Office engines straight away (the download is the slow part on a first visit).
  warmUpEngines((state, ms) => {
    const e = $("engine-status");
    e.textContent =
      state === "loading" ? "Preparing document engines…" : state === "ready" ? `Document engines ready (${fmtMs(ms)})` : "";
    e.classList.toggle("loading", state === "loading");
  });
}

main().catch((e) => {
  console.error(e);
  status(`Failed to start: ${(e as Error).message}`);
});
