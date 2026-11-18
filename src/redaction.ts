/**
 * Finding personal data and turning it into redaction *marks* (not yet applied), so an HR adviser
 * can review them. Applying is a separate, explicit step (applyRedactions), after which the
 * content under each mark is removed from the PDF itself, not just covered.
 */
import NutrientViewer from "@nutrient-sdk/viewer";

type Instance = Awaited<ReturnType<typeof NutrientViewer.load>>;
type Rect = InstanceType<typeof NutrientViewer.Geometry.Rect>;

export type RuleId = "names" | "emails" | "suNumbers" | "dob" | "otherDates" | "manual";

export const RULE_LABELS: Record<RuleId, string> = {
  names: "Names",
  emails: "Email addresses",
  suNumbers: "Service user numbers",
  dob: "Dates of birth",
  otherDates: "Other dates",
  manual: "Added by reviewer",
};

export interface RuleOptions {
  names: string[];
  matchNameParts: boolean;
  emails: boolean;
  suNumbers: boolean;
  dob: boolean;
  otherDates: boolean;
}

export interface Mark {
  id: string;
  rule: RuleId;
  text: string;
  pageIndex: number;
}

const MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec";
/** Numeric dates: 14/03/2012, 14-03-2012, 14.03.12 */
const NUMERIC_DATE = String.raw`\b\d{1,2}[\/\-.]\d{1,2}[\/\-.](?:\d{4}|\d{2})\b`;
/** Written dates: 14 March 2012, 2nd Nov 2010 */
const WRITTEN_DATE = String.raw`\b\d{1,2}(?:st|nd|rd|th)?\s+(?:${MONTHS})\s+\d{4}\b`;
/** Service user numbers: SU followed by five digits. */
const SU_NUMBER = String.raw`\bSU\d{5}\b`;
/** Text near a date that indicates it is a date of birth. */
const DOB_CONTEXT = /\b(DOB|D\.O\.B|date of birth|born|birth date)\b/i;

const escapeRegex = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Runs every enabled rule and creates redaction marks. Returns the marks created. */
export async function findPersonalData(
  instance: Instance,
  opts: RuleOptions,
  onProgress: (msg: string) => void,
): Promise<Mark[]> {
  const created: Mark[] = [];
  const taken: { pageIndex: number; rect: Rect }[] = [];

  const run = async (rule: RuleId, term: string, searchType: "text" | "regex" | "preset", filter?: (preview: string) => boolean) => {
    const results = await instance.search(term, {
      searchType: searchType as never,
      caseSensitive: false,
      searchInAnnotations: false,
    });
    const annots: InstanceType<typeof NutrientViewer.Annotations.RedactionAnnotation>[] = [];
    const meta: Omit<Mark, "id">[] = [];
    for (const r of results) {
      if (r.pageIndex == null || r.rectsOnPage.size === 0) continue;
      if (filter && !filter(r.previewText)) continue;
      const box = NutrientViewer.Geometry.Rect.union(r.rectsOnPage);
      // Skip anything already covered by an earlier rule (e.g. "Thornton" inside "Ellie-Mae Thornton").
      if (taken.some((t) => t.pageIndex === r.pageIndex && covers(t.rect, box))) continue;
      taken.push({ pageIndex: r.pageIndex, rect: box });
      const text =
        r.locationInPreview != null && r.lengthInPreview != null
          ? r.previewText.substr(r.locationInPreview, r.lengthInPreview)
          : term;
      annots.push(
        new NutrientViewer.Annotations.RedactionAnnotation({
          pageIndex: r.pageIndex,
          rects: r.rectsOnPage,
          boundingBox: box,
          overlayText: "",
          customData: { rule, text },
        }),
      );
      meta.push({ rule, text, pageIndex: r.pageIndex });
    }
    if (annots.length === 0) return;
    const saved = await instance.create(annots);
    saved.forEach((a, i) => created.push({ ...meta[i], id: (a as { id: string }).id }));
  };

  // Order matters: a match mostly covered by an earlier one is skipped. Emails and SU numbers go first
  // (so "daniel" inside daniel.okafor@… isn't marked twice), then full names before name parts.
  if (opts.emails) {
    onProgress("Searching for email addresses…");
    await run("emails", NutrientViewer.SearchPattern.EMAIL_ADDRESS, "preset");
  }
  if (opts.suNumbers) {
    onProgress("Searching for service user numbers…");
    await run("suNumbers", SU_NUMBER, "regex");
  }
  const names = [...new Set(opts.names.map((n) => n.trim()).filter(Boolean))];
  for (const name of names) {
    onProgress(`Searching for “${name}”…`);
    await run("names", String.raw`\b${escapeRegex(name).replace(/\s+/g, "\\s+")}\b`, "regex");
  }
  if (opts.matchNameParts) {
    const parts = new Set<string>();
    for (const name of names) for (const part of name.split(/\s+/)) if (part.length >= 3) parts.add(part);
    for (const part of parts) await run("names", String.raw`\b${escapeRegex(part)}\b`, "regex");
  }
  if (opts.dob || opts.otherDates) {
    onProgress("Searching for dates…");
    for (const pattern of [NUMERIC_DATE, WRITTEN_DATE]) {
      if (opts.dob) await run("dob", pattern, "regex", (preview) => DOB_CONTEXT.test(preview));
      if (opts.otherDates) await run("otherDates", pattern, "regex");
    }
  }
  return created;
}

/** True when rect `a` overlaps most of rect `b`. */
function covers(a: Rect, b: Rect): boolean {
  const ix = Math.max(0, Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left));
  const iy = Math.max(0, Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top));
  return ix * iy >= 0.5 * b.width * b.height;
}

/** Reads every redaction mark currently in the document (rule-created and hand-drawn). */
export async function listMarks(instance: Instance): Promise<Mark[]> {
  const out: Mark[] = [];
  for (let i = 0; i < instance.totalPageCount; i++) {
    const annots = await instance.getAnnotations(i);
    for (const a of annots) {
      if (!(a instanceof NutrientViewer.Annotations.RedactionAnnotation)) continue;
      const cd = (a.customData ?? {}) as { rule?: RuleId; text?: string };
      out.push({ id: a.id, rule: cd.rule ?? "manual", text: cd.text ?? "Area drawn by reviewer", pageIndex: a.pageIndex });
    }
  }
  return out;
}

/**
 * After redactions are applied, searches the document again for every value that was redacted.
 * Any hit means content survived. Expected result: 0.
 */
export async function verifyRemoved(instance: Instance, texts: string[]): Promise<{ checked: number; stillFound: string[] }> {
  const unique = [...new Set(texts.filter((t) => t && t !== "Area drawn by reviewer"))];
  const stillFound: string[] = [];
  for (const t of unique) {
    const hits = await instance.search(t, { searchType: "text" as never, caseSensitive: false, searchInAnnotations: false });
    if (hits.size > 0) stillFound.push(t);
  }
  return { checked: unique.length, stillFound };
}
