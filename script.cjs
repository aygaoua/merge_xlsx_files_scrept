const ExcelJS = require('exceljs');
const fs = require('fs');

const baseFile = "emptydb3amer1.xlsx";
const userFile = "emptySalaires3amer2.xlsx";
const templateFile = "emptydb.xlsx"; // optional template/theme workbook
const outputFile = "output_merged.xlsx";

if (!fs.existsSync(baseFile)) {
  console.error(`Base file not found: ${baseFile}`);
  process.exit(1);
}
if (!fs.existsSync(userFile)) {
  console.error(`User file not found: ${userFile}`);
  process.exit(1);
}

async function mergeExcelFiles() {
  // First, physically copy the template if present (to preserve theme/formatting),
  // else copy the base file
  const sourceForOutput = fs.existsSync(templateFile) ? templateFile : baseFile;
  fs.copyFileSync(sourceForOutput, outputFile);
  
  // Now load both files
  const baseWorkbook = new ExcelJS.Workbook();
  const userWorkbook = new ExcelJS.Workbook();
  
  await baseWorkbook.xlsx.readFile(outputFile);  // Load the copy
  await userWorkbook.xlsx.readFile(userFile);
  
  const baseSheet = baseWorkbook.worksheets[0];
  const userSheet = userWorkbook.worksheets[0];
  const outputSheet = baseWorkbook.worksheets[0]; // not used; will reassign after read output
  
  // Load the output workbook (copied template/base)
  const outWorkbook = new ExcelJS.Workbook();
  await outWorkbook.xlsx.readFile(outputFile);
  const outSheet = outWorkbook.worksheets[0];
  
  // Prepare headers and some helpers
  const baseHeaders = [];
  const userHeaders = [];
  let lastColNumber = 0;

  const colToLetter = (n) => {
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };

  // Scan first row of BASE to collect headers and detect last formatted column
  const firstRow = baseSheet.getRow(1);
  const maxScan = Math.max(baseSheet.columnCount || 0, 50);
  for (let i = 1; i <= maxScan; i++) {
    const cell = firstRow.getCell(i);
    if (cell && (cell.value !== null && cell.value !== undefined || cell.style || cell.border || cell.fill || cell.font || cell.alignment)) {
      lastColNumber = i;
      if (cell.value !== null && cell.value !== undefined) baseHeaders[i] = cell.value;
    }
  }
  console.log(`Detected last column with content/formatting: ${lastColNumber}`);
  console.log('Headers found:');
  baseHeaders.forEach((header, idx) => { if (header) console.log(`  Column ${idx} (${colToLetter(idx)}): ${header}`); });

  // Collect USER headers robustly (scan a range to handle many empty columns)
  const userFirstRow = userSheet.getRow(1);
  const userMaxScan = Math.max(userSheet.columnCount || 0, 50);
  for (let i = 1; i <= userMaxScan; i++) {
    const cell = userFirstRow.getCell(i);
    if (cell && cell.value !== null && cell.value !== undefined) {
      userHeaders[i] = cell.value;
    }
  }

  const normalize = str => String(str).trim().toUpperCase();
  const normalizeHeader = (str) => {
    if (str === null || str === undefined) return "";
    let s = String(str).normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // strip accents
    s = s.replace(/[^A-Za-z0-9]+/g, ' ').trim().toUpperCase();
    return s;
  };

  // Find key columns
  let cinBaseCol = -1;
  let cinUserCol = -1;
  let salaireUserCol = -1;
  baseHeaders.forEach((header, idx) => { if (header && normalizeHeader(header).includes("CIN")) cinBaseCol = idx; });
  userHeaders.forEach((header, idx) => {
    const h = normalizeHeader(header);
    if (header && h.includes("CIN")) cinUserCol = idx;
    if (header && (h.includes("SALAIRE") || h.includes("NET"))) salaireUserCol = idx;
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
  const salaryMap = new Map();
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

  // Prepare OUTPUT: always copy BASE headers and rows into the OUTPUT first
  const outFirstRow = outSheet.getRow(1);
  // Write headers from BASE
  for (let i = 1; i < baseHeaders.length; i++) {
    if (baseHeaders[i] !== undefined) {
      const hCell = outFirstRow.getCell(i);
      hCell.value = baseHeaders[i];
    }
  }
  // Determine last data column in base by headers
  let lastDataCol = 0; baseHeaders.forEach((h, i) => { if (h) lastDataCol = i; });
  // Copy rows from BASE to OUTPUT
  const baseMaxRows = baseSheet.rowCount;
  for (let r = 2; r <= baseMaxRows; r++) {
    const baseRow = baseSheet.getRow(r);
    const outRow = outSheet.getRow(r);
    for (let c = 1; c <= lastDataCol; c++) {
      outRow.getCell(c).value = baseRow.getCell(c).value;
    }
  }

  // Rebuild outHeaders after copy
  const outHeaders = [];
  for (let i = 1; i <= Math.max(outSheet.columnCount || 0, lastDataCol); i++) {
    const cell = outFirstRow.getCell(i);
    if (cell && cell.value !== null && cell.value !== undefined) outHeaders[i] = cell.value;
  }

  // Identify 'FONCTION' column in OUTPUT; fallback to P (16) if not found
  let fonctionCol = -1;
  outHeaders.forEach((header, idx) => { if (header && normalizeHeader(header).includes("FONCTION")) fonctionCol = idx; });
  if (fonctionCol === -1) {
    fonctionCol = 16; // P
    const fCell = outFirstRow.getCell(fonctionCol);
    if (!fCell.value) fCell.value = "FONCTION";
    console.log("'FONCTION' header not found in output; defaulting to column P (16)");
  }

  const fonctionColBeforeInsert = fonctionCol;
  // Idempotent insert: if the column immediately after FONCTION already is 'SALAIRE NET', don't insert again
  let salaryColIdx = fonctionCol + 1;
  const existingHeaderAfterFonction = outSheet.getRow(1).getCell(salaryColIdx).value;
  if (!existingHeaderAfterFonction || normalizeHeader(existingHeaderAfterFonction) !== 'SALAIRE NET') {
    outSheet.spliceColumns(fonctionCol + 1, 0);
    console.log(`Inserted a new column after FONCTION at ${colToLetter(fonctionCol + 1)} in output`);
  } else {
    console.log(`SALAIRE NET already present at ${colToLetter(salaryColIdx)}; skipping insert`);
  }

  // Ensure headers Q..T are exactly as required
  salaryColIdx = fonctionCol + 1; // Q
  const headerRow = outSheet.getRow(1);
  const fnHeaderCell = outSheet.getRow(1).getCell(fonctionColBeforeInsert);
  const headerStyle = fnHeaderCell && fnHeaderCell.style ? fnHeaderCell.style : undefined;
  const setHeader = (colIndex, text) => {
    const cell = headerRow.getCell(colIndex);
    cell.value = text;
    if (headerStyle) cell.style = headerStyle;
  };
  setHeader(17, 'SALAIRE NET');
  setHeader(18, 'CATEGORIE');
  setHeader(19, 'DATE EMB');
  setHeader(20, 'DATE SORTIE');
  console.log(`Ensured headers: Q=SALAIRE NET, R=CATEGORIE, S=DATE EMB, T=DATE SORTIE`);

  // Debug: Show headers P..T after layout
  const hdrRow = outSheet.getRow(1);
  const headerPT = [];
  for (let c = 16; c <= 20; c++) { headerPT.push(`${colToLetter(c)}=${hdrRow.getCell(c).value || ''}`); }
  console.log('Headers P..T:', headerPT.join(', '));

  // Fill salary values and copy row styles from FONCTION column
  let matchCount = 0;
  let mismatchCount = 0;

  // Debug: sample CINs from user and base for comparison
  const userCinSamples = Array.from(salaryMap.keys()).slice(0, 10);
  console.log(`User CIN samples: ${JSON.stringify(userCinSamples)}`);
  const baseCinSamples = [];
  for (let r = 2; r <= Math.min(baseSheet.rowCount, 50); r++) {
    const v = baseSheet.getRow(r).getCell(cinBaseCol).value;
    if (v) baseCinSamples.push(normalize(v));
    if (baseCinSamples.length >= 10) break;
  }
  console.log(`Base CIN samples: ${JSON.stringify(baseCinSamples)}`);

  // Now fill SALAIRE NET values using CIN from BASE rows
  for (let r = 2; r <= baseSheet.rowCount; r++) {
    const baseRow = baseSheet.getRow(r);
    const cin = baseRow.getCell(cinBaseCol).value;
    if (cin) {
      const salaire = salaryMap.get(normalize(cin));
      const outRow = outSheet.getRow(r);
      const salaireCell = outRow.getCell(salaryColIdx);
      if (salaire !== undefined) { matchCount++; salaireCell.value = salaire; } else { mismatchCount++; salaireCell.value = ""; }
      const fonctionCell = outRow.getCell(fonctionColBeforeInsert);
      if (fonctionCell && fonctionCell.style) salaireCell.style = fonctionCell.style;
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
  console.error("Error:", err.message);
  process.exit(1);
});
