# HR Case Pack Builder

Proof of concept showing how an in-house app can build redacted HR case packs from SharePoint Online documents using the **Nutrient Web SDK**, entirely in the browser. The organisation, people and documents in it are fictional.

```
SharePoint case folder ──► pick and order documents ──► Word / Excel → PDF ──► combine into one pack
      ▲                                                                              │
      │                                                         find personal data (names, DOBs,
      │                                                         emails, SU numbers) as review marks
      │                                                                              │
      └──── save to <case>/Case Pack/ ◄── apply redactions (irreversible) ◄── adviser reviews marks
            COMBINED.pdf · REDACTED.pdf · audit.json
```

## What it demonstrates

| Requirement | How the POC meets it |
|---|---|
| Convert Word and Excel to PDF without desktop Office | The Web SDK's built-in Office engine (WebAssembly) converts DOCX and XLSX in the browser. No Office install, no server. |
| Combine the documents into one case pack | The adviser ticks and orders documents, and they are merged with one bookmark per source document |
| Redact names, dates of birth, emails and service user numbers | Rules for each: names the adviser types (full names and, optionally, first names or surnames on their own), dates next to "DOB / date of birth / born", emails, and `SU` + 5 digits. "All other dates" is optional. |
| A person reviews before anything is removed | Rules only create *marks*. The adviser can jump to each one, choose **Keep** to un-mark it, or draw extra marks with the viewer's redaction tools. |
| Irreversible redaction | `applyRedactions()` removes the text, images and vector content under each mark from the PDF. The app then searches the output again for every redacted value and reports any it still finds (expected: none). |
| Keep the originals, the combined pack and the redacted pack | Originals are never changed. `COMBINED.pdf` (optional) and `REDACTED.pdf` are saved to `<case>/Case Pack/`. |
| Audit trail | `audit.json` records who, when, the source files and page ranges, the rules used, counts found, removed by the reviewer and applied, the verification result and timings. It contains **no personal data**. |
| Data residency, no third-party processing | All processing happens in the user's browser. The only network traffic is to SharePoint (mocked here) and the app's own static files. |
| No local caching | Files are held in memory (`ArrayBuffer`) for the life of the tab and never written to disk |

## Run it

Requires Node 18+.

```bash
npm install
```

```bash
npm run dev
```

Then open http://localhost:5173.

- `npm install` also copies the SDK's WebAssembly assets into `public/nutrient-viewer-lib/` (self-hosted, no CDN).
- Sample case files are already in `public/samples/`. Run `npm run samples` to regenerate them.
- Without a licence key the SDK runs in **trial mode**, which watermarks output. Put a key in `.env` (copy `.env.example`) as `VITE_NUTRIENT_LICENSE_KEY=...`.

### Demo script (about 3 minutes)

1. Open **HR-2026-0147** (an allegation case with 5 documents: 3 Word, 1 Excel, 1 PDF).
2. Click **Build case pack**. You get 11 pages with a bookmark per document, and the timings are shown.
3. Paste the people in the case into **People in this case**:
   ```
   Ellie-Mae Thornton
   Callum Rafferty
   Joanne Thornton
   Daniel Okafor
   Priya Natarajan
   Mark Ellison
   Helen Carver
   ```
4. Click **Find and mark**, then review the list. Click a row to jump to it, and **Keep** anything that shouldn't be redacted (for example the manager's name in the policy extract). Draw a box with the area redaction tool to show a manual mark.
5. Click **Apply redactions and save**. Three files appear under **Case Pack** with a NEW tag. Open `REDACTED.pdf` and the `audit.json`.

All names, dates, emails and SU numbers in the samples are fictional. `public/samples/expected-pii.json` lists the planted values.

## Code tour

| File | What it does |
|---|---|
| `src/sharepoint.ts` | `DocumentStore` interface plus `MockSharePointStore` (in memory). **Replace the mock with Graph calls for production.** Each method's comment gives the Graph endpoint. |
| `src/pack.ts` | Office → PDF (`NutrientViewer.convertToPDF`), merging (`importDocument` operations in a headless instance) and bookmarks |
| `src/redaction.ts` | Detection rules → `RedactionAnnotation` marks, listing marks for review, and the after-apply verification search |
| `src/audit.ts` | Builds the audit JSON (no personal data) |
| `src/main.ts` | UI and workflow |
| `scripts/make-samples.mjs` | Generates the dummy DOCX, XLSX and PDF case files |

### Swapping in real SharePoint (Microsoft Graph)

Implement `DocumentStore` with an Entra ID app registration (MSAL.js for a web app, or MSAL.NET plus WebView2 if this is hosted inside the existing .NET desktop app):

| Method | Graph call |
|---|---|
| `listCaseFolders()` | `GET /sites/{site-id}/drives/{drive-id}/root/children` |
| `listFolder(path)` | `GET /sites/{site-id}/drives/{drive-id}/root:/{path}:/children` |
| `download(path)` | `GET /sites/{site-id}/drives/{drive-id}/root:/{path}:/content`, read into an `ArrayBuffer` |
| `upload(path, data)` | `PUT …/root:/{path}:/content` for files under 4 MB, otherwise `POST …/root:/{path}:/createUploadSession` and upload in chunks |

Nothing else in the app changes.

## Notes for production

- **Performance.** On the sample pack, conversion takes about 0.7 s per Word document and the 11-page pack builds in about 5 s on a laptop. Test with a real 120–150 page pack on the target devices. If it's too slow, run the same conversion on a UK-hosted .NET service (Nutrient .NET SDK) and keep the review UI in the browser.
- **Names.** Names can't be found reliably by pattern alone, so the adviser enters the people in the case. In production this list could be prefilled from the HR case record.
- **Dates of birth.** These are found by context words near the date. Dates in tables with a "DOB" header in a different row may need "All other dates" plus review.
- **Asset hosting.** The SDK assets (~160 MB, lazily loaded) are served from the app. Use the SDK's recommended self-hosting setup for production builds.
- **Scanned PDFs.** These have no text layer, so text rules won't find anything in them. The adviser can mark them by hand, or OCR can be added.
