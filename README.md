# FISLabelStudio — Excel to A4 Label PDFs

A static, Vercel-ready web app that processes an uploaded Excel spreadsheet entirely in the browser and makes printable **A4 portrait address labels** in a customizable grid (default **2 columns × 6 rows**, 12 records per page).

## How it works

1. Upload an Excel `.xlsx` or `.xls` workbook (up to 25 MB).
2. If multiple worksheets exist, choose one. FISLabelStudio reads the fixed columns automatically; **there is no column mapping screen**.
3. Check the live A4 preview. Adjust font size (9–15 pt), left text padding (2–12 mm) and blue divider colour.
4. Download the currently previewed PDF, any single page directly, or **check any combination of pages** and click **Save selected PDFs separately**.

### Excel data columns (starting with row 3)

| Column | Used for | Output example |
|---|---|---|
| C | Name | Monica Lam Mei Wah (f) |
| G | Address | 8-2, 2/F, Jalan Sungei Besi, ... |
| B | Number | `(156)` |

Blank rows are skipped and no number is shown when column B is empty. Long addresses wrap, with automatic font resizing when necessary. Names, addresses and numbers are left-aligned, and the whole text block is **vertically centred**.

### PDF layout: full A4, no outer border

- Exactly **210 × 297 mm, A4 portrait**.
- Grid defaults to **2 × 6**, with 105 × 49.5 mm cells; the new menu supports 1–5 columns and 1–12 rows, uneven widths/heights, custom gaps and A4 orientation.
- **Zero page margins**, **no outside perimeter border**. Blue dividers appear only between cells; they reach the page edges when margins and gaps are zero.
- Left padding of text defaults to 4 mm and can be adjusted; right padding is 2 mm.
- Print at **Actual size / 100%**, on A4 paper. For true edge-to-edge printing, the physical printer must support **borderless A4**; otherwise some separator-line ends may be clipped by the printer's non-printable area. Make one test print before printing a large batch.

### Download choices

- **Download Page N (PDF):** save one currently previewed page as its own PDF.
- **Per-page PDF buttons:** download any page separately from the page list.
- **Bulk select:** check pages 1, 3, 17 or any other combination; use **Select all** or **Clear selection**.
- **Save selected PDFs separately:** on supported Chromium browsers (Chrome / Edge) choose a folder and the checked PDF pages are written as separate `.pdf` files. On browsers lacking folder-picker support, the app requests separate downloads and you may need to allow multiple file downloads in browser settings. No unselected PDF is saved.
- **Download selected PDFs as ZIP:** optional single ZIP containing *only* the checked one-page PDFs, useful when browsers block multiple downloads.
- **Download combined PDF** and **ZIP of all pages** remain available separately.

For 200 records, the app creates 17 pages. Selecting pages 2, 5 and 17 downloads 3 one-page PDFs, not a merged PDF.

### Privacy

All Excel parsing and PDF generation takes place on the user's computer in the browser. There is no backend, database, account, or spreadsheet upload. The web page loads several version-pinned JavaScript libraries and fonts from public CDNs; as with any production web app, review the third-party dependency risk. For internal school-only access, protect the deployment using an authentication/access layer — browser-side processing alone **does not prevent public access to the app URL**.

## Run locally

Requires Node.js 18+ and Python 3:

```bash
cd fairview-label-maker
npm run dev
```

Open http://localhost:4173 in a browser.

```bash
npm run build
```

The build copies static web assets from `public/` to `dist/` without requiring a database or API keys.

## Deploy on Vercel

1. Extract the ZIP and open `fairview-label-maker`.
2. Create and push a GitHub repository containing its files.
3. In Vercel > **Add New Project**, import the repository.
4. Use:

   | Setting | Value |
   | --- | --- |
   | Framework | Other |
   | Build command | `npm run build` |
   | Output directory | `dist` |
   | Environment variables | None |

5. Deploy; subsequent GitHub pushes can trigger redeployment.

A sample 200-record Excel file is included in `examples/` for local testing. The app's sample-button feature was removed, as requested.


## v1.3 layout updates

- **Upload Excel** is now a compact top-right header control instead of a full-width upload panel. Drag and drop an Excel file onto the button or click it.
- **Print settings** sits immediately left of the upload action as a menu-style popover; no large Section 02 panel.
- **Generate & Download** is on the left side of the two-column workspace and **Live Preview** is on the right (sticky on larger screens).
- Workbook tab selector appears compactly under the toolbar if an Excel file contains multiple sheets.
- Column positions remain fixed: number B3:B; name C3:C; address G3:G. Bulk selection, individual/ZIP/combined exports and full-bleed label layout remain unchanged.

## v1.4 — Flexible row and column layout

Open **Print settings** (compact menu to the left of Upload Excel). New controls:

- **Columns:** 1–5. **Rows:** 1–12. Defaults are 2 columns × 6 rows (12 labels/page).
- **A4 orientation:** portrait (default) or landscape.
- **Page margin:** 0–20 mm, uniform on all sides; default 0 mm.
- **Column gap and row gap:** 0–15 mm. If an extreme combination of gaps, margins and label counts would leave too little space, gap values are automatically reduced to leave at least 16 mm of column width and 10 mm of row height. All final sizes are shown in the menu.
- **Customize individual column widths and row heights:** expand the advanced section to enter proportional weights, e.g. `1, 2, 1` for three columns makes the middle one twice as wide as either side column; row height weights work similarly. The app automatically normalizes the weights so they fill the available page area. Use **Make equal** to restore equal widths or heights.
- **Live label dimensions:** the settings menu shows each label's width and height in millimetres, or the range if unequal.
- **Reset to 2 × 6:** restores the original borderless A4 portrait grid, with zero gaps and margins and equal-sized cells. Font/padding/colour settings remain independent.

Live Preview, combined PDFs, one-page PDFs and individually selected bulk downloads **all share the exact same underlying grid geometry**. Changing row or column count resets your selected page checkboxes because the per-page record groupings change. The fixed spreadsheet mapping is still B=number, C=name and G=address, starting with row 3. The PDF retains only internal dividing lines and never draws a perimeter border.

### Practical examples for 200 records

| Grid | Labels per page | Pages |
| --- | ---: | ---: |
| 2 × 6 (default) | 12 | 17 |
| 3 × 5 | 15 | 14 |
| 4 × 5 | 20 | 10 |
| 3 × 4 | 12 | 17 |

**Print caution:** labels on very dense grids may require tiny font or truncation. Review Live Preview and make a test print. For actual edge-to-edge print, use a printer supporting borderless A4.


## v1.5 — Rename PDFs before download

The **Generate & download** panel now contains a **Rename your PDF** field with a live filename preview. Uploading a new Excel workbook pre-fills the filename using the workbook's name (without `.xlsx`/`.xls`); it can be changed at any point before downloading, with no impact on label contents or the live A4 preview. Switching worksheets preserves the customized name.

- **Combined PDF (all pages)** gets the exact chosen name: entering `FIS October Labels` downloads `FIS October Labels.pdf`.
- **Single-page / selected separate PDFs** append their original page positions: `FIS October Labels-page-01-of-17.pdf`, `FIS October Labels-page-03-of-17.pdf`. This prevents filename collisions within a batch.
- **ZIP containing all separate PDFs** and **ZIP containing selected separate PDFs** use the custom name as their prefix. The PDFs **inside** each ZIP also use the same custom name and page suffixes.
- If a user pastes `.pdf` into the field, it is stripped to avoid `file.pdf.pdf`; unsafe path characters are replaced, the name is limited to 80 characters, and a blank filename falls back to `FISLabels`.
- The field is disabled during export to keep all filenames consistent throughout an ongoing download.

Sample: 200 records at 2 × 6 produce 17 pages. Set the filename to `KL October 2026`; the combined file becomes `KL October 2026.pdf`, or select pages 1 and 17 to save two distinct files as `KL October 2026-page-01-of-17.pdf` and `KL October 2026-page-17-of-17.pdf`.
#   f i n d i v l a b e l  
 