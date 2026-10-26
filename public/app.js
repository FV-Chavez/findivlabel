(() => {
  'use strict';

  const A4_PORTRAIT = { width: 210, height: 297 };
  const mmToPt = (mm) => mm * 72 / 25.4;
  const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, Number.isFinite(Number(v)) ? Number(v) : lo));
  const gridState = { colWeights: [1, 1], rowWeights: [1, 1, 1, 1, 1, 1] };

  const byId = (id) => document.getElementById(id);
  const state = { records: [], sheets: [], activeSheet: '', fileName: '', page: 0, cache: null, working: false, selectedPages: new Set() };
  const els = {
    dropzone: byId('dropzone'), fileInput: byId('fileInput'), fileSummary: byId('fileSummary'),
    filename: byId('filename'), fileDescription: byId('fileDescription'),
    clearBtn: byId('clearBtn'), sheetContainer: byId('sheetContainer'),
    sheetSelect: byId('sheetSelect'), fontSize: byId('fontSize'), fontSizeOutput: byId('fontSizeOutput'),
    labelPadding: byId('labelPadding'), paddingOutput: byId('paddingOutput'),
    borderColor: byId('borderColor'), colorOutput: byId('colorOutput'),
    paperGrid: byId('paperGrid'), recordCount: byId('recordCount'), pageCount: byId('pageCount'),
    pageLabel: byId('pageLabel'), prevBtn: byId('prevBtn'), nextBtn: byId('nextBtn'),
    downloadCombined: byId('downloadCombined'), downloadZip: byId('downloadZip'),
    downloadCurrent: byId('downloadCurrent'), currentPageLabel: byId('currentPageLabel'),
    pageDownloads: byId('pageDownloads'), pageListCount: byId('pageListCount'),
    selectAllPages: byId('selectAllPages'), clearSelection: byId('clearSelection'),
    selectedCount: byId('selectedCount'), downloadSelected: byId('downloadSelected'),
    downloadSelectedZip: byId('downloadSelectedZip'), selectedButtonLabel: byId('selectedButtonLabel'),
    zipLabel: byId('zipLabel'), downloadSummary: byId('downloadSummary'), status: byId('status'),
    exportName: byId('exportName'), combinedNamePreview: byId('combinedNamePreview'),
    individualNamePreview: byId('individualNamePreview'),
    columns: byId('columns'), rows: byId('rows'), orientation: byId('orientation'),
    outerMargin: byId('outerMargin'), columnGap: byId('columnGap'), rowGap: byId('rowGap'),
    columnWeights: byId('columnWeights'), rowWeights: byId('rowWeights'),
    equalColumns: byId('equalColumns'), equalRows: byId('equalRows'), resetLayout: byId('resetLayout'),
    pageDimensions: byId('pageDimensions'), gridSummary: byId('gridSummary'), dimensionReadout: byId('dimensionReadout'),
    labelsPerPage: byId('labelsPerPage'), orientationBadge: byId('orientationBadge')
  };

  // Keep common readable Unicode/spacing while removing unsafe OS path characters.
  // Strip known extensions to prevent duplicated .pdf/.zip suffixes.
  function normalizeExportName(value) {
    let name = String(value || '').normalize('NFKC').trim();
    name = name.replace(/\.(?:pdf|zip|xlsx|xls)$/i, '');
    name = name.replace(/[<>:"/\\|?*\u0000-\u001f\u007f]/g, '-');
    name = name.replace(/\s+/g, ' ').replace(/^\.+/, '').replace(/[. ]+$/, '').trim();
    name = Array.from(name).slice(0, 80).join('').replace(/[. ]+$/, '');
    if (!name) return 'FISLabels';
    if (/^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i.test(name)) return `${name}-labels`;
    return name;
  }

  const baseName = () => normalizeExportName(els.exportName.value);
  const combinedFileName = () => `${baseName()}.pdf`;
  function refreshFilenamePreviews() {
    els.combinedNamePreview.textContent = combinedFileName();
    const digits = Math.max(2, String(Math.max(1, pagesCount())).length);
    els.individualNamePreview.textContent = `${baseName()}-page-${String(1).padStart(digits, '0')}-of-${Math.max(1, pagesCount())}.pdf`;
  }

  function normalizedText(input) {
    if (input == null) return '';
    if (typeof input === 'object') {
      if (Array.isArray(input.richText)) return input.richText.map(part => part.text || '').join('');
      if (input.result != null) return normalizedText(input.result);
      if (input.text != null) return normalizedText(input.text);
      if (input.formula != null) return '';
    }
    return String(input).replace(/\r\n?/g, '\n').replace(/[\t\u00a0]/g, ' ').split('\n').map(line => line.trim().replace(/ {2,}/g, ' ')).join('\n').trim();
  }

  function cleanTotal(input) {
    const str = normalizedText(input);
    return str.replace(/^\((.*)\)$/, '$1').trim();
  }

  function worksheetToRecords(sheet) {
    // Read absolute Excel coordinates, not shifted arrays. This remains correct
    // even when the worksheet's first populated cell is C3 or later.
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    if (!range) return [];
    const records = [];
    const display = (row, column) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      if (!cell) return '';
      // w is the displayed Excel value (important for formatted numbers).
      if (typeof cell.w === 'string') return cell.w;
      if (cell.v != null) return cell.v;
      return '';
    };
    for (let row = 2; row <= range.e.r; row++) {
      const name = normalizedText(display(row, 2)); // Excel column C
      const address = normalizedText(display(row, 6)); // Excel column G
      const total = cleanTotal(display(row, 1)); // Excel column B
      if (!name && !address) continue;
      records.push({ name, address, total, excelRow: row + 1 });
    }
    return records;
  }

  function setStatus(text, error = false) {
    els.status.textContent = text;
    els.status.classList.toggle('error', error);
  }

  function settings() {
    // Inputs are bounded here as well as in the HTML so pasted values cannot
    // create a negative or zero-size label. Ratios are positive and normalized.
    const columns = Math.round(clamp(els.columns.value, 1, 5));
    const rows = Math.round(clamp(els.rows.value, 1, 12));
    const orientation = els.orientation.value === 'landscape' ? 'landscape' : 'portrait';
    const pageW = orientation === 'landscape' ? 297 : 210;
    const pageH = orientation === 'landscape' ? 210 : 297;
    const margin = clamp(els.outerMargin.value, 0, 20);
    // Reserve at least 16 mm horizontally and 10 mm vertically per label,
    // even when the user combines many cells, wide gutters and large margins.
    const colGapLimit = columns === 1 ? 15 : Math.max(0, (pageW - 2 * margin - columns * 16) / (columns - 1));
    const rowGapLimit = rows === 1 ? 15 : Math.max(0, (pageH - 2 * margin - rows * 10) / (rows - 1));
    return {
      columns, rows, orientation, margin,
      colGap: Math.min(clamp(els.columnGap.value, 0, 15), colGapLimit),
      rowGap: Math.min(clamp(els.rowGap.value, 0, 15), rowGapLimit),
      colWeights: Array.from({length: columns}, (_, i) => clamp(gridState.colWeights[i] ?? 1, 0.25, 8)),
      rowWeights: Array.from({length: rows}, (_, i) => clamp(gridState.rowWeights[i] ?? 1, 0.25, 8)),
      fontSize: clamp(els.fontSize.value, 9, 15),
      padding: clamp(els.labelPadding.value, 2, 12),
      color: els.borderColor.value
    };
  }
  function perPage() { const cfg = settings(); return cfg.columns * cfg.rows; }
  function pagesCount() { return Math.ceil(state.records.length / perPage()); }
  function layout(cfg = settings()) {
    const width = cfg.orientation === 'landscape' ? A4_PORTRAIT.height : A4_PORTRAIT.width;
    const height = cfg.orientation === 'landscape' ? A4_PORTRAIT.width : A4_PORTRAIT.height;
    const usableW = width - 2 * cfg.margin - (cfg.columns - 1) * cfg.colGap;
    const usableH = height - 2 * cfg.margin - (cfg.rows - 1) * cfg.rowGap;
    const weightSizes = (weights, space) => {
      const total = weights.reduce((sum, v) => sum + v, 0);
      return weights.map(v => space * v / total);
    };
    const widths = weightSizes(cfg.colWeights, usableW);
    const heights = weightSizes(cfg.rowWeights, usableH);
    const xs = []; const ys = [];
    let x = cfg.margin;
    let y = cfg.margin;
    for (const w of widths) { xs.push(x); x += w + cfg.colGap; }
    for (const h of heights) { ys.push(y); y += h + cfg.rowGap; }
    return { width, height, widths, heights, xs, ys };
  }
  function updateInputs() {
    const cfg = settings();
    els.fontSizeOutput.textContent = `${cfg.fontSize.toFixed(1)} pt`;
    els.paddingOutput.textContent = `${cfg.padding} mm`;
    els.colorOutput.textContent = cfg.color.toUpperCase();
    els.pageDimensions.textContent = `A4 · ${cfg.orientation === 'portrait' ? '210 × 297' : '297 × 210'} mm`;
    els.gridSummary.textContent = `${cfg.columns} × ${cfg.rows} · ${perPage()} labels`;
    els.orientationBadge.textContent = `A4 · ${cfg.orientation.toUpperCase()}`;
    const geo = layout(cfg);
    const fmt = (n) => (Math.round(n * 10) / 10).toLocaleString('en', {maximumFractionDigits: 1});
    const dimension = (values) => {
      const min = Math.min(...values), max = Math.max(...values);
      return max - min < 0.05 ? fmt(min) : `${fmt(min)}–${fmt(max)}`;
    };
    els.dimensionReadout.textContent = `Label size: ${dimension(geo.widths)} mm wide × ${dimension(geo.heights)} mm high`;
    els.columnGap.value = String(cfg.colGap);
    els.rowGap.value = String(cfg.rowGap);
    document.documentElement.style.setProperty('--grid-color', cfg.color);
  }
  function resizeWeights(type, length) {
    const key = type === 'column' ? 'colWeights' : 'rowWeights';
    gridState[key] = Array.from({ length }, (_, i) => gridState[key][i] ?? 1);
  }
  function renderWeights(type) {
    const isCol = type === 'column';
    const holder = isCol ? els.columnWeights : els.rowWeights;
    const weights = isCol ? gridState.colWeights : gridState.rowWeights;
    holder.replaceChildren();
    weights.forEach((weight, i) => {
      const label = document.createElement('label');
      label.className = 'weight-setting';
      label.textContent = `${isCol ? 'Column' : 'Row'} ${i + 1}`;
      const input = document.createElement('input');
      input.type = 'number'; input.inputMode = 'decimal';
      input.min = '0.25'; input.max = '8'; input.step = '0.25';
      input.value = String(weight);
      input.setAttribute('aria-label', `${isCol ? 'Column width' : 'Row height'} ${i + 1} relative size`);
      input.addEventListener('change', () => {
        const number = clamp(input.value, 0.25, 8);
        weights[i] = number;
        input.value = String(number);
        state.cache = null; updateInputs(); updatePreview();
      });
      label.append(input); holder.append(label);
    });
  }
  function resetPaging() {
    // Grid shape changes reorder which records appear on selected pages.
    state.selectedPages.clear(); state.page = 0; state.cache = null;
  }
  function selectedIndices() { return [...state.selectedPages].sort((a, b) => a - b); }

  function refreshSelectionControls() {
    const pageTotal = pagesCount();
    const count = state.selectedPages.size;
    els.selectedCount.textContent = `${count} selected`;
    els.selectAllPages.disabled = state.working || !pageTotal || count === pageTotal;
    els.clearSelection.disabled = state.working || count === 0;
    els.downloadSelected.disabled = state.working || count === 0;
    els.downloadSelectedZip.disabled = state.working || count === 0;
    els.selectedButtonLabel.textContent = count ? `Save ${count} selected PDF${count === 1 ? '' : 's'} separately` : 'Save selected as separate PDFs';
  }

  function updatePreview() {
    const cfg = settings();
    const geometry = layout(cfg);
    const capacity = perPage();
    const pageCount = Math.max(1, pagesCount());
    state.page = Math.min(Math.max(state.page, 0), pageCount - 1);
    els.recordCount.textContent = state.records.length.toLocaleString();
    els.pageCount.textContent = String(pagesCount());
    els.labelsPerPage.textContent = String(capacity);
    els.pageLabel.textContent = `Page ${state.page + 1} of ${pageCount}`;
    els.prevBtn.disabled = state.page === 0;
    els.nextBtn.disabled = state.page >= pageCount - 1;
    const paper = byId('paper');
    paper.style.aspectRatio = `${geometry.width} / ${geometry.height}`;
    els.paperGrid.replaceChildren();
    const records = state.records.slice(state.page * capacity, (state.page + 1) * capacity);
    const pct = (v, total) => `${(100 * v / total).toFixed(5)}%`;
    // The preview shares the PDF's exact millimetre geometry; custom unequal
    // rows and columns, gaps and margins have no CSS-layout rounding drift.
    for (let i = 0; i < capacity; i++) {
      const c = i % cfg.columns, r = Math.floor(i / cfg.columns);
      const cell = document.createElement('div');
      cell.className = 'label-preview';
      Object.assign(cell.style, {
        left: pct(geometry.xs[c], geometry.width),
        top: pct(geometry.ys[r], geometry.height),
        width: pct(geometry.widths[c], geometry.width),
        height: pct(geometry.heights[r], geometry.height),
        paddingLeft: pct(Math.min(cfg.padding, geometry.widths[c] * 0.34), geometry.widths[c]),
        paddingRight: pct(Math.min(2, geometry.widths[c] * 0.10), geometry.widths[c]),
        paddingTop: pct(2, geometry.widths[c]),
        paddingBottom: pct(2, geometry.widths[c])
      });
      const content = document.createElement('div');
      content.className = 'label-lines';
      content.style.fontSize = `${(cfg.fontSize * 25.4 / 72 / geometry.width * 100).toFixed(4)}cqw`;
      const entry = records[i];
      if (entry) {
        const pieces = [entry.name, entry.address, entry.total ? `(${entry.total})` : ''].filter(Boolean);
        for (const piece of pieces) {
          const line = document.createElement('div');
          line.textContent = piece;
          content.append(line);
        }
      }
      cell.append(content);
      els.paperGrid.append(cell);
    }
    for (let c = 0; c < cfg.columns - 1; c++) {
      const line = document.createElement('div');
      line.className = 'preview-divider preview-divider-vertical';
      line.style.left = pct(geometry.xs[c] + geometry.widths[c] + cfg.colGap / 2, geometry.width);
      line.style.top = pct(cfg.margin, geometry.height);
      line.style.height = pct(geometry.height - cfg.margin * 2, geometry.height);
      els.paperGrid.append(line);
    }
    for (let r = 0; r < cfg.rows - 1; r++) {
      const line = document.createElement('div');
      line.className = 'preview-divider preview-divider-horizontal';
      line.style.top = pct(geometry.ys[r] + geometry.heights[r] + cfg.rowGap / 2, geometry.height);
      line.style.left = pct(cfg.margin, geometry.width);
      line.style.width = pct(geometry.width - cfg.margin * 2, geometry.width);
      els.paperGrid.append(line);
    }
    const count = pagesCount();
    els.downloadCombined.disabled = state.working || count === 0;
    els.downloadZip.disabled = state.working || count === 0;
    els.downloadCurrent.disabled = state.working || count === 0;
    els.currentPageLabel.textContent = `Download Page ${state.page + 1} (PDF)`;
    els.pageListCount.textContent = `${count} page${count !== 1 ? 's' : ''}`;
    els.pageDownloads.replaceChildren();
    // Retain only selected page numbers that still exist in this workbook.
    for (const selected of state.selectedPages) {
      if (selected < 0 || selected >= count) state.selectedPages.delete(selected);
    }
    if (count === 0) {
      const empty = document.createElement('p');
      empty.className = 'page-list-empty';
      empty.textContent = 'Upload a spreadsheet to view individual pages.';
      els.pageDownloads.append(empty);
    } else {
      const fragment = document.createDocumentFragment();
      for (let pageIndex = 0; pageIndex < count; pageIndex++) {
        const row = document.createElement('div');
        row.className = `page-download-item${pageIndex === state.page ? ' active' : ''}${state.selectedPages.has(pageIndex) ? ' selected' : ''}`;
        row.setAttribute('role', 'listitem');
        const selection = document.createElement('label');
        selection.className = 'page-select-label';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = state.selectedPages.has(pageIndex);
        checkbox.disabled = state.working;
        checkbox.setAttribute('aria-label', `Select page ${pageIndex + 1} for bulk download`);
        checkbox.addEventListener('change', () => {
          if (checkbox.checked) state.selectedPages.add(pageIndex);
          else state.selectedPages.delete(pageIndex);
          row.classList.toggle('selected', checkbox.checked);
          refreshSelectionControls();
        });
        const details = document.createElement('span');
        details.className = 'page-download-item-label';
        const label = document.createElement('strong');
        label.textContent = `Page ${pageIndex + 1}`;
        const range = document.createElement('small');
        const first = pageIndex * capacity + 1;
        const last = Math.min((pageIndex + 1) * capacity, state.records.length);
        range.textContent = `Records ${first}–${last}`;
        details.append(label, range);
        selection.append(checkbox, details);
        const button = document.createElement('button');
        button.type = 'button';
        button.textContent = 'PDF ↓';
        button.disabled = state.working;
        button.setAttribute('aria-label', `Download page ${pageIndex + 1} as an individual PDF`);
        button.addEventListener('click', () => downloadSingle(pageIndex));
        row.append(selection, button);
        fragment.append(row);
      }
      els.pageDownloads.append(fragment);
    }
    refreshSelectionControls();
    els.exportName.disabled = state.working;
    refreshFilenamePreviews();
    els.zipLabel.textContent = `Download ${count || 'separate'} PDF${count === 1 ? '' : 's'} (.zip)`;
    els.downloadSummary.textContent = count ? `${state.records.length.toLocaleString()} records · ${count} separate PDF page${count !== 1 ? 's' : ''} available` : 'Upload a spreadsheet to enable downloads.';
  }

  function clearData() {
    state.records = [];
    state.fileName = '';
    state.sheets = [];
    state.page = 0;
    state.cache = null;
    state.selectedPages.clear();
    els.exportName.value = 'FISLabels';
    els.fileInput.value = '';
    els.fileSummary.classList.add('hidden');
    els.sheetContainer.classList.add('hidden');
    setStatus('');
    updatePreview();
  }

  function setData(records, fileName, sheetNames = [], selected = '') {
    state.records = records;
    state.fileName = fileName;
    state.page = 0;
    state.cache = null;
    state.selectedPages.clear();
    els.filename.textContent = fileName;
    els.fileDescription.textContent = `${records.length} non-empty records ${selected ? '· ' + selected : ''}`;
    els.fileSummary.classList.remove('hidden');
    els.sheetContainer.classList.toggle('hidden', sheetNames.length < 2);
    els.sheetSelect.replaceChildren();
    for (const name of sheetNames) {
      const item = document.createElement('option');
      item.value = name;
      item.textContent = name;
      item.selected = name === selected;
      els.sheetSelect.append(item);
    }
    updatePreview();
  }

  async function readExcel(file) {
    if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error('Please upload an Excel .xlsx or .xls file.');
    if (file.size > 25 * 1024 * 1024) throw new Error('This spreadsheet is larger than 25 MB. Please use a smaller file.');
    if (typeof XLSX === 'undefined') throw new Error('Excel reader could not load. Check your internet connection and refresh.');
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array', cellDates: false, dense: false });
    if (!workbook.SheetNames.length) throw new Error('No worksheets were found in that Excel file.');
    const sheets = workbook.SheetNames.map(name => ({ name, records: worksheetToRecords(workbook.Sheets[name]) }));
    const initial = sheets.find(sheet => sheet.records.length) || sheets[0];
    // A new workbook starts with a sensible suggestion; users can edit it freely.
    els.exportName.value = normalizeExportName(file.name.replace(/\.(?:xlsx|xls)$/i, ''));
    state.sheets = sheets;
    setData(initial.records, file.name, workbook.SheetNames, initial.name);
    if (!initial.records.length) setStatus('No address labels found. Please check columns C, G and rows starting at row 3.', true);
    else setStatus(`Found ${initial.records.length} labels in “${initial.name}”.`);
  }

  async function upload(file) {
    if (!file) return;
    setStatus('Opening spreadsheet…');
    await new Promise(resolve => setTimeout(resolve, 20));
    try { await readExcel(file); }
    catch (err) { clearData(); setStatus(err.message || 'Could not open this spreadsheet.', true); }
  }

  function hexRgb(hex) {
    const n = parseInt(hex.replace('#', ''), 16);
    return PDFLib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
  }

  // Built-in Helvetica is a fallback only; replace unsupported glyphs predictably.
  // A network-accessible Noto Sans is used normally so names with accents work.
  function safeWinAnsi(value, face) {
    return Array.from(value || '').map(character => {
      if (character === '\n') return '\n';
      try { face.widthOfTextAtSize(character, 12); return character; }
      catch { return '?'; }
    }).join('');
  }

  // Word-wrap against the actual embedded font's metrics to avoid clipping in PDF.
  function wrapLine(line, font, fontSize, maxWidth) {
    if (!line) return [''];
    const words = line.split(/\s+/).filter(Boolean);
    const result = [];
    let current = '';
    const width = text => font.widthOfTextAtSize(text, fontSize);
    for (const word of words) {
      const withWord = current ? `${current} ${word}` : word;
      if (width(withWord) <= maxWidth) { current = withWord; continue; }
      if (current) { result.push(current); current = ''; }
      if (width(word) <= maxWidth) { current = word; continue; }
      // Split long unbroken postal tokens (rather than drawing past the border).
      let chunk = '';
      for (const char of Array.from(word)) {
        if (chunk && width(chunk + char) > maxWidth) { result.push(chunk); chunk = ''; }
        chunk += char;
      }
      current = chunk;
    }
    if (current) result.push(current);
    return result.length ? result : [''];
  }

  function buildLines(record, font, size, textWidth) {
    const out = [];
    const items = [record.name, record.address, record.total ? `(${record.total})` : ''];
    for (const item of items) {
      if (!item) continue;
      for (const rawLine of item.split('\n')) {
        out.push(...wrapLine(rawLine.trim(), font, size, textWidth));
      }
    }
    return out;
  }

  function drawLabel(page, record, column, row, font, cfg, geo) {
    const boxW = mmToPt(geo.widths[column]);
    const boxH = mmToPt(geo.heights[row]);
    const x0 = mmToPt(geo.xs[column]);
    const y0 = mmToPt(geo.height - geo.ys[row] - geo.heights[row]);
    const leftPadding = mmToPt(Math.min(cfg.padding, geo.widths[column] * 0.34));
    const rightPadding = mmToPt(Math.min(2, geo.widths[column] * 0.10));
    const maxWidth = Math.max(1, boxW - leftPadding - rightPadding);
    const maxHeight = Math.max(1, boxH - mmToPt(Math.min(4, geo.heights[row] * 0.22)));
    let size = cfg.fontSize;
    let lines = [];
    // Fit both width and height as much as possible before limiting lines.
    while (size >= 4) {
      lines = buildLines(record, font, size, maxWidth);
      if (lines.length * size * 1.23 <= maxHeight) break;
      size -= 0.25;
    }
    size = Math.max(4, size);
    lines = buildLines(record, font, size, maxWidth);
    const maxLines = Math.max(1, Math.floor(maxHeight / (size * 1.23)));
    if (lines.length > maxLines) {
      lines = lines.slice(0, maxLines);
      let last = lines[lines.length - 1];
      const marker = '...';
      const limit = maxWidth;
      while (last.length && font.widthOfTextAtSize(last + marker, size) > limit) last = last.slice(0, -1);
      // Helvetica does not support the unicode ellipsis; fallback to "...".
      lines[lines.length - 1] = last + (font.widthOfTextAtSize(marker, size) <= limit ? marker : '');
    }
    const lineHeight = size * 1.23;
    const centreY = y0 + boxH / 2;
    const firstBaseline = centreY + (lines.length - 1) * lineHeight / 2 - size * 0.35;
    lines.forEach((line, lineNumber) => {
      page.drawText(line, { x: x0 + leftPadding, y: firstBaseline - lineNumber * lineHeight,
        size, font, color: PDFLib.rgb(0.10, 0.16, 0.20) });
    });
  }

  function drawGrid(page, color, cfg, geo) {
    // Inside dividers only. Even for uneven cells or gaps, never draw a border
    // around the outer page/perimeter.
    for (let c = 0; c < cfg.columns - 1; c++) {
      const x = mmToPt(geo.xs[c] + geo.widths[c] + cfg.colGap / 2);
      page.drawLine({ start: { x, y: mmToPt(cfg.margin) },
        end: { x, y: mmToPt(geo.height - cfg.margin) }, color, thickness: 0.9 });
    }
    for (let r = 0; r < cfg.rows - 1; r++) {
      const y = mmToPt(geo.height - geo.ys[r] - geo.heights[r] - cfg.rowGap / 2);
      page.drawLine({ start: { x: mmToPt(cfg.margin), y },
        end: { x: mmToPt(geo.width - cfg.margin), y }, color, thickness: 0.9 });
    }
  }

  let fontDataPromise = null;
  async function loadFont() {
    if (!fontDataPromise) fontDataPromise = fetch('https://raw.githubusercontent.com/googlefonts/noto-fonts/main/hinted/ttf/NotoSans/NotoSans-Regular.ttf').then(response => {
      if (!response.ok) throw new Error('Could not load the printable font.');
      return response.arrayBuffer();
    }).catch(error => { fontDataPromise = null; throw error; });
    return fontDataPromise;
  }

  async function createCombined() {
    const cfg = settings();
    const key = JSON.stringify([cfg, state.records]);
    if (state.cache?.key === key) return state.cache.bytes;
    if (!window.PDFLib) throw new Error('PDF library could not load. Check your internet connection and refresh.');
    if (!state.records.length) throw new Error('No records to export.');
    const doc = await PDFLib.PDFDocument.create();
    let face;
    let basicFallback = false;
    if (window.fontkit) {
      try {
        doc.registerFontkit(window.fontkit);
        face = await doc.embedFont(await loadFont(), { subset: true });
      } catch (error) { console.warn('Using built-in PDF font fallback:', error); }
    }
    if (!face) {
      face = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
      basicFallback = true;
    }
    doc.setTitle(`Address Labels — A4 (${cfg.columns} × ${cfg.rows})`);
    doc.setSubject('Print-ready name and address labels');
    doc.setCreator('FISLabelStudio');
    const border = hexRgb(cfg.color);
    const geo = layout(cfg);
    for (let pageNum = 0; pageNum < pagesCount(); pageNum++) {
      const page = doc.addPage([mmToPt(geo.width), mmToPt(geo.height)]);
      drawGrid(page, border, cfg, geo);
      const recordsOnPage = state.records.slice(pageNum * perPage(), (pageNum + 1) * perPage());
      recordsOnPage.forEach((record, index) => {
        const printable = basicFallback ? {
          ...record,
          name: safeWinAnsi(record.name, face),
          address: safeWinAnsi(record.address, face),
          total: safeWinAnsi(record.total, face)
        } : record;
        drawLabel(page, printable, index % cfg.columns, Math.floor(index / cfg.columns), face, cfg, geo);
      });
    }
    const bytes = await doc.save();
    state.cache = { key, bytes };
    return bytes;
  }

  // Generate only the selected sheet. This is a real one-page PDF, not a
  // download of the combined document or a ZIP extraction.
  async function createSinglePage(pageIndex) {
    if (!window.PDFLib) throw new Error('PDF library could not load. Check your internet connection and refresh.');
    if (!state.records.length || pageIndex < 0 || pageIndex >= pagesCount()) throw new Error('That page is not available.');
    const cfg = settings();
    const doc = await PDFLib.PDFDocument.create();
    let face;
    let basicFallback = false;
    if (window.fontkit) {
      try {
        doc.registerFontkit(window.fontkit);
        face = await doc.embedFont(await loadFont(), { subset: true });
      } catch (error) { console.warn('Using built-in PDF font fallback:', error); }
    }
    if (!face) {
      face = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
      basicFallback = true;
    }
    doc.setTitle(`Address Labels — Page ${pageIndex + 1} of ${pagesCount()}`);
    doc.setSubject('Individual print-ready A4 address label sheet');
    doc.setCreator('FISLabelStudio');
    const geo = layout(cfg);
    const page = doc.addPage([mmToPt(geo.width), mmToPt(geo.height)]);
    drawGrid(page, hexRgb(cfg.color), cfg, geo);
    const pageRecords = state.records.slice(pageIndex * perPage(), (pageIndex + 1) * perPage());
    pageRecords.forEach((record, index) => {
      const printable = basicFallback ? {
        ...record,
        name: safeWinAnsi(record.name, face),
        address: safeWinAnsi(record.address, face),
        total: safeWinAnsi(record.total, face)
      } : record;
      drawLabel(page, printable, index % cfg.columns, Math.floor(index / cfg.columns), face, cfg, geo);
    });
    return doc.save();
  }

  async function downloadSingle(pageIndex) {
    if (state.working || !state.records.length) return;
    state.working = true;
    updatePreview();
    setStatus(`Generating Page ${pageIndex + 1} of ${pagesCount()} as a separate PDF…`);
    await new Promise(resolve => setTimeout(resolve, 35));
    try {
      const bytes = await createSinglePage(pageIndex);
      saveBlob(new Blob([bytes], { type: 'application/pdf' }), separateName(pageIndex));
      setStatus(`Page ${pageIndex + 1} downloaded as its own A4 PDF (records ${pageIndex * perPage() + 1}–${Math.min((pageIndex + 1) * perPage(), state.records.length)}).`);
    } catch (error) {
      console.error(error);
      setStatus(error?.message || 'Could not create the selected PDF.', true);
    } finally {
      state.working = false;
      updatePreview();
    }
  }

  function separateName(pageIndex) {
    const number = String(pageIndex + 1).padStart(Math.max(2, String(pagesCount()).length), '0');
    return `${baseName()}-page-${number}-of-${pagesCount()}.pdf`;
  }

  // Explicit multi-selection: export *only* the checked pages. Each PDF has
  // exactly one A4 sheet, even when the user selects nonconsecutive pages.
  async function downloadSelectedPages(mode) {
    if (state.working || !state.records.length || !state.selectedPages.size) return;
    const chosen = selectedIndices();
    let directory = null;
    if (mode === 'separate' && typeof window.showDirectoryPicker === 'function') {
      // Directory picker must be called directly from the user click before
      // any lengthy async PDF generation, preserving browser user activation.
      try {
        directory = await window.showDirectoryPicker({ mode: 'readwrite', id: 'fis-labels' });
      } catch (error) {
        if (error?.name === 'AbortError') { setStatus('Folder selection cancelled.'); return; }
        console.warn('Folder picker unavailable; using separate browser downloads.', error);
      }
    }
    state.working = true;
    updatePreview();
    setStatus(`Preparing ${chosen.length} selected, one-page PDFs…`);
    await new Promise(resolve => setTimeout(resolve, 20));
    try {
      const combinedBytes = await createCombined();
      if (!window.PDFLib) throw new Error('PDF library could not load.');
      const source = await PDFLib.PDFDocument.load(combinedBytes);
      if (mode === 'zip' && !window.JSZip) throw new Error('ZIP library could not load. Please refresh.');
      const zip = mode === 'zip' ? new window.JSZip() : null;
      for (const pageIndex of chosen) {
        const onePage = await PDFLib.PDFDocument.create();
        const [copied] = await onePage.copyPages(source, [pageIndex]);
        onePage.addPage(copied);
        const bytes = await onePage.save();
        const filename = separateName(pageIndex);
        if (zip) zip.file(filename, bytes);
        else if (directory) {
          const fileHandle = await directory.getFileHandle(filename, { create: true });
          const writable = await fileHandle.createWritable();
          await writable.write(bytes);
          await writable.close();
        } else {
          // Firefox/Safari and other browsers without directory picker may
          // prompt to allow multiple separate file downloads.
          saveBlob(new Blob([bytes], { type: 'application/pdf' }), filename);
        }
      }
      if (zip) {
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        saveBlob(blob, `${baseName()}-${chosen.length}-selected-separate-pdfs.zip`);
      }
      const pages = chosen.map(i => i + 1).join(', ');
      const result = zip ? 'ZIP of individual PDFs downloaded' : directory ? 'Separate PDFs saved in the chosen folder' : 'Separate PDF downloads requested';
      setStatus(`${result}: ${chosen.length} page${chosen.length === 1 ? '' : 's'} (page${chosen.length === 1 ? '' : 's'} ${pages}).${!directory && !zip && chosen.length > 1 ? ' Allow multiple downloads if your browser asks.' : ''}`);
    } catch (error) {
      console.error(error);
      setStatus(error?.message || 'Could not save the selected PDFs.', true);
    } finally {
      state.working = false;
      updatePreview();
    }
  }

  function saveBlob(blob, name) {
    const href = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = href;
    anchor.download = name;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(href), 20_000);
  }

  async function download(mode) {
    if (state.working || !state.records.length) return;
    state.working = true;
    updatePreview();
    setStatus(mode === 'zip' ? 'Creating separate PDF documents…' : 'Creating combined PDF…');
    await new Promise(resolve => setTimeout(resolve, 35));
    try {
      const merged = await createCombined();
      if (mode === 'combined') {
        saveBlob(new Blob([merged], { type: 'application/pdf' }), combinedFileName());
        setStatus(`Combined PDF downloaded: ${pagesCount()} page${pagesCount() === 1 ? '' : 's'}, ${state.records.length} records.`);
      } else {
        if (!window.JSZip) throw new Error('ZIP library could not load. Check your internet connection and refresh.');
        const source = await PDFLib.PDFDocument.load(merged);
        const zip = new window.JSZip();
        for (let i = 0; i < pagesCount(); i++) {
          const onePage = await PDFLib.PDFDocument.create();
          const [copiedPage] = await onePage.copyPages(source, [i]);
          onePage.addPage(copiedPage);
          const bytes = await onePage.save();
          zip.file(separateName(i), bytes);
        }
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        saveBlob(blob, `${baseName()}-${pagesCount()}-separate-pdfs.zip`);
        setStatus(`ZIP downloaded with ${pagesCount()} individual one-page PDFs.`);
      }
    } catch (error) {
      console.error(error);
      setStatus(error?.message || 'Could not create your PDF. Please try again.', true);
    } finally {
      state.working = false;
      updatePreview();
    }
  }

  const settingsMenu = byId('settingsMenu');
  document.addEventListener('click', event => {
    if (settingsMenu?.open && !settingsMenu.contains(event.target)) settingsMenu.open = false;
  });
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && settingsMenu?.open) {
      settingsMenu.open = false;
      settingsMenu.querySelector('summary')?.focus();
    }
  });
  els.dropzone.addEventListener('click', () => els.fileInput.click());
  els.dropzone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); els.fileInput.click(); } });
  els.dropzone.addEventListener('dragenter', e => { e.preventDefault(); els.dropzone.classList.add('dragover'); });
  els.dropzone.addEventListener('dragover', e => { e.preventDefault(); els.dropzone.classList.add('dragover'); });
  els.dropzone.addEventListener('dragleave', e => { e.preventDefault(); els.dropzone.classList.remove('dragover'); });
  els.dropzone.addEventListener('drop', e => { e.preventDefault(); els.dropzone.classList.remove('dragover'); upload(e.dataTransfer?.files?.[0]); });
  els.fileInput.addEventListener('change', e => upload(e.target.files?.[0]));
  els.exportName.addEventListener('input', refreshFilenamePreviews);
  els.exportName.addEventListener('blur', () => {
    els.exportName.value = normalizeExportName(els.exportName.value);
    refreshFilenamePreviews();
  });
  els.clearBtn.addEventListener('click', clearData);
  els.sheetSelect.addEventListener('change', e => {
    const sheet = state.sheets.find(sheet => sheet.name === e.target.value);
    if (!sheet) return;
    const names = state.sheets.map(item => item.name);
    setData(sheet.records, state.fileName, names, sheet.name);
    setStatus(sheet.records.length ? `Found ${sheet.records.length} records in “${sheet.name}”.` : 'Selected sheet has no usable records.', !sheet.records.length);
  });
  for (const element of [els.fontSize, els.labelPadding, els.borderColor]) {
    element.addEventListener('input', () => { state.cache = null; updateInputs(); updatePreview(); });
  }
  for (const element of [els.columns, els.rows]) {
    element.addEventListener('change', () => {
      const columns = Math.round(clamp(els.columns.value, 1, 5));
      const rows = Math.round(clamp(els.rows.value, 1, 12));
      els.columns.value = String(columns);
      els.rows.value = String(rows);
      resizeWeights('column', columns); resizeWeights('row', rows);
      renderWeights('column'); renderWeights('row');
      resetPaging(); updateInputs(); updatePreview();
      setStatus(`Grid changed to ${columns} columns × ${rows} rows. Page selections were cleared because records were regrouped.`);
    });
  }
  for (const element of [els.outerMargin, els.columnGap, els.rowGap]) {
    element.addEventListener('change', () => {
      const [min, max] = element.id === 'outerMargin' ? [0, 20] : [0, 15];
      element.value = String(clamp(element.value, min, max));
      state.cache = null; updateInputs(); updatePreview();
    });
  }
  els.orientation.addEventListener('change', () => { state.cache = null; updateInputs(); updatePreview(); });
  els.equalColumns.addEventListener('click', () => {
    gridState.colWeights.fill(1); renderWeights('column');
    state.cache = null; updatePreview();
  });
  els.equalRows.addEventListener('click', () => {
    gridState.rowWeights.fill(1); renderWeights('row');
    state.cache = null; updatePreview();
  });
  els.resetLayout.addEventListener('click', () => {
    els.columns.value = '2'; els.rows.value = '6'; els.orientation.value = 'portrait';
    els.outerMargin.value = '0'; els.columnGap.value = '0'; els.rowGap.value = '0';
    gridState.colWeights = [1,1]; gridState.rowWeights = [1,1,1,1,1,1];
    renderWeights('column'); renderWeights('row');
    resetPaging(); updateInputs(); updatePreview();
    setStatus('Layout restored to 2 × 6 labels on A4 portrait. Page selections cleared.');
  });
  els.prevBtn.addEventListener('click', () => { state.page--; updatePreview(); });
  els.nextBtn.addEventListener('click', () => { state.page++; updatePreview(); });
  els.downloadCurrent.addEventListener('click', () => downloadSingle(state.page));
  els.selectAllPages.addEventListener('click', () => { state.selectedPages = new Set(Array.from({ length: pagesCount() }, (_, i) => i)); updatePreview(); });
  els.clearSelection.addEventListener('click', () => { state.selectedPages.clear(); updatePreview(); });
  els.downloadSelected.addEventListener('click', () => downloadSelectedPages('separate'));
  els.downloadSelectedZip.addEventListener('click', () => downloadSelectedPages('zip'));
  els.downloadCombined.addEventListener('click', () => download('combined'));
  els.downloadZip.addEventListener('click', () => download('zip'));
  renderWeights('column'); renderWeights('row');
  updateInputs();
  updatePreview();
})();
