// Generates dummy HR case files (DOCX, XLSX, PDF) for the mock SharePoint library.
// All people, dates and numbers are fictional.
// Output: public/samples/<case>/<file> plus public/samples/library.json (the mock library index)
// and public/samples/expected-pii.json (planted values, used to check redaction coverage).

import { mkdir, writeFile, rm } from "node:fs/promises";
import { join } from "node:path";
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
} from "docx";
import ExcelJS from "exceljs";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

const OUT = new URL("../public/samples/", import.meta.url).pathname;

// ---------------------------------------------------------------- fake people
const P = {
  child1: { name: "Ellie-Mae Thornton", dob: "14/03/2012", dobLong: "14 March 2012", su: "SU04821" },
  child2: { name: "Callum Rafferty", dob: "02/11/2010", dobLong: "2 November 2010", su: "SU17733" },
  parent1: { name: "Joanne Thornton", email: "jo.thornton82@example.com" },
  staff1: { name: "Daniel Okafor", email: "daniel.okafor@example.org", dob: "21/07/1988" },
  staff2: { name: "Priya Natarajan", email: "priya.natarajan@example.org" },
  staff3: { name: "Mark Ellison", email: "mark.ellison@example.org", dob: "09-01-1979" },
  manager: { name: "Helen Carver", email: "helen.carver@example.org" },
};

// ---------------------------------------------------------------- helpers
const h = (text, level = HeadingLevel.HEADING_1) => new Paragraph({ text, heading: level });
const p = (text) => new Paragraph({ children: [new TextRun(text)], spacing: { after: 160 } });
const kv = (rows) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(([k, v]) =>
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: k, bold: true })] })] }),
          new TableCell({ children: [new Paragraph(v)] }),
        ],
      }),
    ),
  });

async function docx(children) {
  const doc = new Document({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}

async function xlsx(sheetName, header, rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName, { pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  ws.addRow(header).font = { bold: true };
  rows.forEach((r) => ws.addRow(r));
  ws.columns.forEach((c) => (c.width = 28));
  ws.getColumn(header.length).width = 60;
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function pdf(title, lines) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  let page = doc.addPage([595, 842]);
  let y = 790;
  page.drawText(title, { x: 50, y, size: 16, font: bold, color: rgb(0.1, 0.1, 0.3) });
  y -= 36;
  for (const line of lines) {
    if (y < 60) { page = doc.addPage([595, 842]); y = 790; }
    page.drawText(line, { x: 50, y, size: 11, font });
    y -= 18;
  }
  return Buffer.from(await doc.save());
}

// Long, mostly generic policy text to give the pack realistic length.
function policyParagraphs(n) {
  const out = [];
  for (let i = 1; i <= n; i++) {
    out.push(h(`Section ${i}: Safeguarding procedure`, HeadingLevel.HEADING_2));
    out.push(p(
      "All staff must record concerns on the day they arise, using the approved form, and pass them to the " +
        "Designated Safeguarding Lead. Records must be factual, dated and signed. Where a concern relates to a " +
        "member of staff, the Local Authority Designated Officer must be informed within one working day.",
    ));
    out.push(p(
      "Managers should make sure that any investigation is proportionate, that the people involved are kept " +
        "informed, and that personal data is only shared with those who need it, in line with UK GDPR.",
    ));
    if (i % 4 === 0) out.push(p(`Example reference for training purposes only: case contact ${P.manager.name}.`));
  }
  return out;
}

// ---------------------------------------------------------------- cases
const cases = [
  {
    id: "HR-2026-0147",
    title: "Allegation – Oakwood Residential Home",
    files: {
      "01 Referral letter.docx": () =>
        docx([
          h("Referral to HR: allegation against staff member"),
          kv([
            ["Case reference", "HR-2026-0147"],
            ["Staff member", P.staff1.name],
            ["Staff email", P.staff1.email],
            ["Staff DOB", P.staff1.dob],
            ["Young person", `${P.child1.name} (${P.child1.su})`],
            ["Young person DOB", P.child1.dobLong],
            ["Referred by", `${P.manager.name}, Registered Manager`],
          ]),
          p(""),
          p(
            `On 3 October 2026 ${P.child1.name} told her key worker, ${P.staff2.name}, that ${P.staff1.name} had ` +
              "spoken to her in a threatening way during the evening routine. The key worker recorded the disclosure " +
              "the same evening and informed the Registered Manager.",
          ),
          p(
            `${P.child1.name}'s mother, ${P.parent1.name} (${P.parent1.email}), was informed on 4 October 2026 and ` +
              "asked to be kept updated on the outcome.",
          ),
          p(`Please contact ${P.manager.name} at ${P.manager.email} with any questions.`),
        ]),
      "02 Incident log.xlsx": () =>
        xlsx(
          "Incident log",
          ["Date", "Time", "Service user", "SU number", "Staff on shift", "Summary"],
          [
            ["01/10/2026", "19:40", P.child1.name, P.child1.su, P.staff1.name, "Refused evening routine, raised voices reported"],
            ["02/10/2026", "20:15", P.child2.name, P.child2.su, P.staff3.name, "Settled well, no concerns"],
            ["03/10/2026", "18:55", P.child1.name, P.child1.su, `${P.staff1.name}; ${P.staff2.name}`, "Disclosure made to key worker"],
            ["04/10/2026", "09:10", P.child1.name, P.child1.su, P.staff2.name, `Mother informed (${P.parent1.email})`],
            ["06/10/2026", "21:30", P.child2.name, P.child2.su, P.staff3.name, "Minor damage to bedroom door"],
          ],
        ),
      "03 Witness statement – P Natarajan.docx": () =>
        docx([
          h("Witness statement"),
          kv([
            ["Name", P.staff2.name],
            ["Role", "Residential Support Worker (key worker)"],
            ["Email", P.staff2.email],
            ["Date of statement", "07/10/2026"],
          ]),
          p(""),
          p(
            `I have been ${P.child1.name}'s key worker since January 2025. On the evening of 3 October 2026 she ` +
              `came to the office and said that ${P.staff1.name} had shouted at her and told her she would "lose her ` +
              "phone for a month\". She was tearful but calm.",
          ),
          p(
            `I also saw ${P.child2.name} (${P.child2.su}) in the corridor at the time but he did not see the ` +
              "conversation. I recorded the disclosure in the log and called the manager on duty.",
          ),
          p(`Signed: ${P.staff2.name}`),
        ]),
      "04 Safeguarding plan.pdf": () =>
        pdf("Safeguarding plan – HR-2026-0147", [
          `Young person: ${P.child1.name}   Service user no: ${P.child1.su}`,
          `Date of birth: ${P.child1.dob}`,
          `Parent/carer: ${P.parent1.name}  (${P.parent1.email})`,
          "",
          "Interim measures agreed 05/10/2026:",
          `1. ${P.staff1.name} not to work directly with the young person pending outcome.`,
          `2. Key worker ${P.staff2.name} to complete daily check-ins.`,
          "3. LADO informed 04/10/2026, reference withheld.",
          "",
          `Plan owner: ${P.manager.name} (${P.manager.email})`,
          "Review date: 19/10/2026",
        ]),
      "05 Safeguarding policy extract.docx": () =>
        docx([h("Example Care Group safeguarding policy (extract)"), ...policyParagraphs(36)]),
    },
  },
  {
    id: "HR-2026-0152",
    title: "Grievance – Education Services",
    files: {
      "01 Grievance letter.docx": () =>
        docx([
          h("Formal grievance"),
          p(`From: ${P.staff3.name} (${P.staff3.email}), DOB ${P.staff3.dob}`),
          p(`To: ${P.manager.name}, Head of Education`),
          p(
            "I wish to raise a formal grievance about the way overtime has been allocated in the autumn term. " +
              `I believe I have been treated less favourably than colleagues, including ${P.staff1.name}.`,
          ),
          p(`Please reply to ${P.staff3.email}.`),
        ]),
      "02 Overtime timeline.xlsx": () =>
        xlsx(
          "Timeline",
          ["Date", "Employee", "Email", "Hours offered", "Notes"],
          [
            ["05/09/2026", P.staff3.name, P.staff3.email, 0, "Not offered"],
            ["05/09/2026", P.staff1.name, P.staff1.email, 6, "Offered by rota"],
            ["12/09/2026", P.staff3.name, P.staff3.email, 4, ""],
            ["19/09/2026", P.staff2.name, P.staff2.email, 8, "Cover for sickness"],
          ],
        ),
      "03 Meeting notes.docx": () =>
        docx([
          h("Grievance meeting notes"),
          kv([
            ["Date", "22 October 2026"],
            ["Present", `${P.staff3.name}, ${P.manager.name}, HR adviser`],
          ]),
          p(""),
          p(
            `${P.staff3.name} explained the background. ${P.manager.name} agreed to review the overtime rota and ` +
              "respond in writing within 10 working days.",
          ),
        ]),
    },
  },
];

// ---------------------------------------------------------------- write
await rm(OUT, { recursive: true, force: true });
const library = { site: "Example Care Group HR", library: "HR Case Files", folders: [] };
for (const c of cases) {
  const dir = join(OUT, c.id);
  await mkdir(dir, { recursive: true });
  const folder = { name: c.id, title: c.title, files: [] };
  for (const [name, make] of Object.entries(c.files)) {
    const buf = await make();
    await writeFile(join(dir, name), buf);
    folder.files.push({ name, path: `samples/${c.id}/${name}`, size: buf.length });
  }
  library.folders.push(folder);
  console.log(`  ${c.id}: ${folder.files.length} files`);
}
await writeFile(join(OUT, "library.json"), JSON.stringify(library, null, 2));

const expected = {
  names: [...new Set(Object.values(P).map((x) => x.name))],
  emails: Object.values(P).map((x) => x.email).filter(Boolean),
  dobs: Object.values(P).flatMap((x) => [x.dob, x.dobLong]).filter(Boolean),
  suNumbers: Object.values(P).map((x) => x.su).filter(Boolean),
};
await writeFile(join(OUT, "expected-pii.json"), JSON.stringify(expected, null, 2));
// Tiny blank document the app converts at idle to pre-load the Office engine (see warmUpEngines in src/pack.ts).
await writeFile(new URL("../public/warmup.docx", import.meta.url), await docx([p(" ")]));
console.log("Samples written to public/samples/ (plus public/warmup.docx)");
