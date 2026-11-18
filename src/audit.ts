/**
 * Audit record saved next to each case pack.
 * Deliberately contains NO personal data: it records counts, rules, pages and file names,
 * never the redacted values themselves or the names the adviser searched for.
 */
import NutrientViewer from "@nutrient-sdk/viewer";
import type { PackSource } from "./pack";
import { RULE_LABELS, type Mark, type RuleId } from "./redaction";

export interface AuditInput {
  caseId: string;
  user: { displayName: string; email: string };
  startedAt: Date;
  sources: PackSource[];
  totalPages: number;
  rules: { names: number; matchNameParts: boolean; emails: boolean; suNumbers: boolean; dob: boolean; otherDates: boolean };
  marksFound: Mark[];
  marksRemovedByReviewer: Mark[];
  marksApplied: Mark[];
  verification: { checked: number; stillFound: number };
  outputs: { name: string; path: string; size: number }[];
  timingsMs: Record<string, number>;
}

const countByRule = (marks: Mark[]) =>
  marks.reduce<Record<string, number>>((acc, m) => {
    const k = RULE_LABELS[m.rule as RuleId];
    acc[k] = (acc[k] ?? 0) + 1;
    return acc;
  }, {});

export function buildAudit(a: AuditInput) {
  return {
    schema: "case-pack-audit@1",
    caseId: a.caseId,
    preparedBy: a.user,
    startedAt: a.startedAt.toISOString(),
    completedAt: new Date().toISOString(),
    processing: {
      location: "Client-side (browser). No document data sent to any external service.",
      engine: `Nutrient Web SDK ${NutrientViewer.version} (WebAssembly)`,
    },
    sourceDocuments: a.sources.map((s) => ({
      name: s.name,
      path: s.path,
      sizeBytes: s.size,
      type: s.kind,
      packPages: s.pageCount ? `${s.startPage + 1}-${s.startPage + s.pageCount}` : "",
    })),
    pack: { totalPages: a.totalPages },
    redaction: {
      rulesUsed: {
        namesProvided: a.rules.names,
        matchNameParts: a.rules.matchNameParts,
        emailAddresses: a.rules.emails,
        serviceUserNumbers: a.rules.suNumbers,
        datesOfBirth: a.rules.dob,
        otherDates: a.rules.otherDates,
      },
      found: { total: a.marksFound.length, byRule: countByRule(a.marksFound) },
      removedByReviewer: {
        total: a.marksRemovedByReviewer.length,
        items: a.marksRemovedByReviewer.map((m) => ({ rule: RULE_LABELS[m.rule], page: m.pageIndex + 1 })),
      },
      addedByReviewer: a.marksApplied.filter((m) => m.rule === "manual").length,
      applied: { total: a.marksApplied.length, byRule: countByRule(a.marksApplied) },
      method: "Irreversible. Text, images and vector content under each mark removed from the PDF.",
      verification: {
        valuesRechecked: a.verification.checked,
        valuesStillFound: a.verification.stillFound,
        passed: a.verification.stillFound === 0,
      },
    },
    outputs: a.outputs,
    timingsMs: a.timingsMs,
  };
}
