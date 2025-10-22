import ExcelJS from 'exceljs';
import * as fs from 'fs';

const baseFile = 'fulldbbase.xlsx';
const userFile = 'userdb.xlsx';
const templateFile = 'emptydb.xlsx'; // optional template/theme workbook
const outputFile = 'output_merged.xlsx';

function colToLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function normalize(val: unknown): string {
  if (val === null || val === undefined) return '';
  return String(val).trim().toUpperCase();
}

function normalizeHeader(str: unknown): string {
  if (str === null || str === undefined) return '';
  let s = String(str).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  s = s.replace(/[^A-Za-z0-9]+/g, ' ').trim().toUpperCase();
  return s;
}

async function mergeExcelFiles(): Promise<void> {
  if (!fs.existsSync(baseFile)) {
    console.error(`Base file not found: ${baseFile}`);
    process.exit(1);
  }
  if (!fs.existsSync(userFile)) {
    console.error(`User file not found: ${userFile}`);
    process.exit(1);
  }

  // Create output from template (if present) or base to preserve formatting
  const sourceForOutput = fs.existsSync(templateFile) ? templateFile : baseFile;
  fs.copyFileSync(sourceForOutput, outputFile);

  // Load workbooks
  const baseWorkbook = new ExcelJS.Workbook();
  const userWorkbook = new ExcelJS.Workbook();
  await baseWorkbook.xlsx.readFile(baseFile);
  await userWorkbook.xlsx.readFile(userFile);

  // Load output (copied template/base)
  const outWorkbook = new ExcelJS.Workbook();
  await outWorkbook.xlsx.readFile(outputFile);

  const baseSheet = baseWorkbook.worksheets[0];
  const userSheet = userWorkbook.worksheets[0];
  const outSheet = outWorkbook.worksheets[0];

  // Gather base headers across many columns (including formatted empties)
  const baseHeaders: (string | undefined)[] = [];
  const firstRow = baseSheet.getRow(1);
  const maxScan = Math.max(baseSheet.columnCount || 0, 50);
  let lastColNumber = 0;
  for (let i = 1; i <= maxScan; i++) {
    const cell = firstRow.getCell(i);
    const hasFmt = (cell as any)?.style || (cell as any)?.border || (cell as any)?.fill || (cell as any)?.font || (cell as any)?.alignment;
    if ((cell.value !== null && cell.value !== undefined) || hasFmt) {
      lastColNumber = i;
      if (cell.value !== null && cell.value !== undefined) baseHeaders[i] = String(cell.value);
    }
  }
  console.log(`Detected last column with content/formatting: ${lastColNumber}`);
  console.log('Headers found:');
  baseHeaders.forEach((h, i) => { if (h) console.log(`  Column ${i} (${colToLetter(i)}): ${h}`); });

  // Collect user headers robustly
  const userHeaders: (string | undefined)[] = [];
  const userFirstRow = userSheet.getRow(1);
  const userMaxScan = Math.max(userSheet.columnCount || 0, 50);
  for (let i = 1; i <= userMaxScan; i++) {
    const cell = userFirstRow.getCell(i);
    if (cell && cell.value !== null && cell.value !== undefined) userHeaders[i] = String(cell.value);
  }

  // Find key columns
  let cinBaseCol = -1;
  let cinUserCol = -1;
  let salaireUserCol = -1;
  baseHeaders.forEach((header, idx) => { if (header && normalizeHeader(header).includes('CIN')) cinBaseCol = idx; });
  userHeaders.forEach((header, idx) => {
    const h = normalizeHeader(header);
    if (header && h.includes('CIN')) cinUserCol = idx;
    if (header && (h.includes('SALAIRE') || h.includes('NET'))) salaireUserCol = idx;
  });

  console.log(`CIN in base at ${cinBaseCol} (${colToLetter(cinBaseCol)})`);
  console.log(`CIN in user at ${cinUserCol} (${colToLetter(cinUserCol)})`);
  console.log(`Salaire in user at ${salaireUserCol} (${colToLetter(salaireUserCol)})`);

  if (cinBaseCol === -1 || cinUserCol === -1) {
    console.error("❌ Could not find 'CIN' column in one of the files.");
    return;
  }
  if (salaireUserCol === -1) {
    console.error("❌ Could not find 'SALAIRE' or 'NET' column in user file.");
    return;
  }

  // Build salary map from user sheet
  const salaryMap = new Map<string, ExcelJS.CellValue>();
  userSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const cin = row.getCell(cinUserCol).value;
      const salaire = row.getCell(salaireUserCol).value;
      if (cin) salaryMap.set(normalize(cin), salaire);
    }
  });
  console.log(`Built salary map with ${salaryMap.size} entries`);

  console.log(`Base file has ${baseSheet.rowCount - 1} rows`);
  console.log(`User file has ${userSheet.rowCount - 1} rows`);

  // Always copy BASE headers and rows into the OUTPUT first
  const outFirstRow = outSheet.getRow(1);
  for (let i = 1; i < baseHeaders.length; i++) {
    if (baseHeaders[i] !== undefined) {
      const hCell = outFirstRow.getCell(i);
      hCell.value = baseHeaders[i] as string;
    }
  }
  let lastDataCol = 0; baseHeaders.forEach((h, i) => { if (h) lastDataCol = i; });
  const baseMaxRows = baseSheet.rowCount;
  for (let r = 2; r <= baseMaxRows; r++) {
    const baseRow = baseSheet.getRow(r);
    const outRow = outSheet.getRow(r);
    for (let c = 1; c <= lastDataCol; c++) {
      outRow.getCell(c).value = baseRow.getCell(c).value;
    }
  }

  // Rebuild outHeaders after copy
  const outHeaders: (string | undefined)[] = [];
  for (let i = 1; i <= Math.max(outSheet.columnCount || 0, lastDataCol); i++) {
    const cell = outFirstRow.getCell(i);
    if (cell && cell.value !== null && cell.value !== undefined) outHeaders[i] = String(cell.value);
  }

  // Identify 'FONCTION' column in OUTPUT; fallback to P (16) if not found
  let fonctionCol = -1;
  outHeaders.forEach((header, idx) => { if (header && normalizeHeader(header).includes('FONCTION')) fonctionCol = idx; });
  if (fonctionCol === -1) {
    fonctionCol = 16; // P
    const fCell = outFirstRow.getCell(fonctionCol);
    if (!fCell.value) fCell.value = 'FONCTION';
    console.log("'FONCTION' header not found in output; defaulting to column P (16)");
  }

  // Capture current indices of CATEGORIE / DATE EMB / DATE SORTIE before any insertion
  const getHeaderIndex = (name: string): number => {
    let idx = -1;
    outHeaders.forEach((h, i) => { if (h && normalizeHeader(h).includes(name)) idx = i; });
    return idx;
  };
  let categorieIdx = getHeaderIndex('CATEGORIE');
  let dateEmbIdx = getHeaderIndex('DATE EMB');
  let dateSortieIdx = getHeaderIndex('DATE SORTIE');

  const fonctionColBeforeInsert = fonctionCol;
  // Idempotent insert: if the column immediately after FONCTION already is 'SALAIRE NET', don't insert again
  let salaryColIdx = fonctionCol + 1;
  const existingHeaderAfterFonction = outSheet.getRow(1).getCell(salaryColIdx).value;
  if (!existingHeaderAfterFonction || normalizeHeader(existingHeaderAfterFonction).trim() !== 'SALAIRE NET') {
    outSheet.spliceColumns(fonctionCol + 1, 0);
    console.log(`Inserted a new column after FONCTION at ${colToLetter(fonctionCol + 1)} in output`);
    // Adjust captured indices that were after the insertion point
    if (categorieIdx > fonctionColBeforeInsert) categorieIdx += 1;
    if (dateEmbIdx > fonctionColBeforeInsert) dateEmbIdx += 1;
    if (dateSortieIdx > fonctionColBeforeInsert) dateSortieIdx += 1;
  } else {
    console.log(`SALAIRE NET already present at ${colToLetter(salaryColIdx)}; skipping insert`);
  }

  // Move existing columns to target positions R/S/T if needed (preserve data/styles)
  const copyColumn = (src: number, dst: number) => {
    if (src <= 0 || dst <= 0 || src === dst) return;
    const maxRows = Math.max(outSheet.rowCount, baseSheet.rowCount);
    // Copy cell values and styles
    for (let r = 1; r <= maxRows; r++) {
      const fromCell = outSheet.getRow(r).getCell(src);
      const toCell = outSheet.getRow(r).getCell(dst);
      toCell.value = fromCell.value;
      const style = (fromCell as any)?.style;
      if (style) (toCell as any).style = style;
    }
  };

  const Q = 17, R = 18, S = 19, T = 20;
  // Ensure CATEGORIE is at R
  if (categorieIdx > 0 && categorieIdx !== R) {
    copyColumn(categorieIdx, R);
  }
  // Ensure DATE EMB at S
  if (dateEmbIdx > 0 && dateEmbIdx !== S) {
    copyColumn(dateEmbIdx, S);
  }
  // Ensure DATE SORTIE at T
  if (dateSortieIdx > 0 && dateSortieIdx !== T) {
    copyColumn(dateSortieIdx, T);
  }

  // Ensure headers Q..T are exactly as required
  salaryColIdx = Q; // enforce Q explicitly
  const headerRow = outSheet.getRow(1);
  const fnHeaderCell = outSheet.getRow(1).getCell(fonctionColBeforeInsert);
  const headerStyle = (fnHeaderCell as any)?.style ? (fnHeaderCell as any).style : undefined;
  const setHeader = (colIndex: number, text: string) => {
    const cell = headerRow.getCell(colIndex);
    cell.value = text;
    if (headerStyle) (cell as any).style = headerStyle;
  };
  setHeader(Q, 'SALAIRE NET');
  setHeader(R, 'CATEGORIE');
  setHeader(S, 'DATE EMB');
  setHeader(T, 'DATE SORTIE');
  console.log(`Ensured headers: Q=SALAIRE NET, R=CATEGORIE, S=DATE EMB, T=DATE SORTIE`);

  // Debug: Show headers P..T after layout
  const hdrRow = outSheet.getRow(1);
  const headerPT: string[] = [];
  for (let c = 16; c <= 20; c++) { headerPT.push(`${colToLetter(c)}=${String(hdrRow.getCell(c).value ?? '')}`); }
  console.log('Headers P..T:', headerPT.join(', '));

  // Deduplicate SALAIRE NET: keep only at Q, clear any other SALAIRE NET columns (e.g., D)
  const clearColumn = (colIndex: number) => {
    const maxRows = Math.max(outSheet.rowCount, baseSheet.rowCount);
    for (let r = 1; r <= maxRows; r++) {
      const cell = outSheet.getRow(r).getCell(colIndex);
      cell.value = r === 1 ? '' : null;
    }
  };
  // Scan headers to find duplicate SALAIRE NET columns
  const outHeaderRow = outSheet.getRow(1);
  const scanMaxCols = Math.max(outSheet.columnCount || 0, lastDataCol, 50);
  const duplicateSalaryCols: number[] = [];
  for (let c = 1; c <= scanMaxCols; c++) {
    const val = outHeaderRow.getCell(c).value;
    if (val && normalizeHeader(val) === 'SALAIRE NET' && c !== Q) {
      duplicateSalaryCols.push(c);
    }
  }
  if (duplicateSalaryCols.length > 0) {
    duplicateSalaryCols.forEach(c => {
      clearColumn(c);
      console.log(`Cleared duplicate SALAIRE NET at ${colToLetter(c)}`);
    });
  }

  // Ensure column T style matches others (clone from S if available)
  const cloneColumnStyle = (refCol: number, dstCol: number) => {
    if (refCol <= 0 || dstCol <= 0 || refCol === dstCol) return;
    const refColumn = outSheet.getColumn(refCol);
    const dstColumn = outSheet.getColumn(dstCol);
    if (refColumn && typeof refColumn.width === 'number') {
      dstColumn.width = refColumn.width;
    }
    const maxRows = Math.max(outSheet.rowCount, baseSheet.rowCount);
    for (let r = 1; r <= maxRows; r++) {
      const fromCell = outSheet.getRow(r).getCell(refCol) as any;
      const toCell = outSheet.getRow(r).getCell(dstCol) as any;
      if (fromCell && fromCell.style) {
        // Deep clone to avoid shared references
        toCell.style = JSON.parse(JSON.stringify(fromCell.style));
      }
    }
  };
  // Prefer S as the style source; fall back to R, then Q
  if (S > 0) {
    cloneColumnStyle(S, T);
    console.log(`Aligned column ${colToLetter(T)} style/width to match ${colToLetter(S)}`);
  } else if (R > 0) {
    cloneColumnStyle(R, T);
    console.log(`Aligned column ${colToLetter(T)} style/width to match ${colToLetter(R)}`);
  } else if (Q > 0) {
    cloneColumnStyle(Q, T);
    console.log(`Aligned column ${colToLetter(T)} style/width to match ${colToLetter(Q)}`);
  }

  // Fill salary values using CIN from BASE rows
  let matchCount = 0;
  let mismatchCount = 0;

  const userCinSamples = Array.from(salaryMap.keys()).slice(0, 10);
  console.log(`User CIN samples: ${JSON.stringify(userCinSamples)}`);
  const baseCinSamples: string[] = [];
  for (let r = 2; r <= Math.min(baseSheet.rowCount, 50); r++) {
    const v = baseSheet.getRow(r).getCell(cinBaseCol).value;
    if (v) baseCinSamples.push(normalize(v));
    if (baseCinSamples.length >= 10) break;
  }
  console.log(`Base CIN samples: ${JSON.stringify(baseCinSamples)}`);

  for (let r = 2; r <= baseSheet.rowCount; r++) {
    const baseRow = baseSheet.getRow(r);
    const cin = baseRow.getCell(cinBaseCol).value;
    if (cin) {
      const salaire = salaryMap.get(normalize(cin));
      const outRow = outSheet.getRow(r);
      const salaireCell = outRow.getCell(17); // column Q
      if (salaire !== undefined) { matchCount++; salaireCell.value = salaire; } else { mismatchCount++; salaireCell.value = ''; }
      const fonctionCell = outRow.getCell(fonctionColBeforeInsert);
      if ((fonctionCell as any)?.style) (salaireCell as any).style = (fonctionCell as any).style;
    } else {
      mismatchCount++;
    }
  }

  await outWorkbook.xlsx.writeFile(outputFile);
  console.log(`Created output file: ${outputFile}`);
  console.log(`Matched ${matchCount} records`);
  console.log(`${mismatchCount} records without matches`);
}

mergeExcelFiles().catch(err => {
  console.error('Error:', (err as Error).message);
  process.exit(1);
});
