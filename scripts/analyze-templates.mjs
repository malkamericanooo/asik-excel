#!/usr/bin/env node
import * as XLSX from 'xlsx';
import { readFileSync, readdirSync } from 'fs';
import { resolve, join } from 'path';

const ROOT = resolve(import.meta.dirname, '..');

// Find all xlsx files in the project root and fixtures
function findExcelFiles(dir) {
  const files = [];
  try {
    for (const f of readdirSync(dir, { withFileTypes: true })) {
      if (f.isFile() && f.name.endsWith('.xlsx') && !f.name.startsWith('~')) {
        files.push(join(dir, f.name));
      }
    }
  } catch (_) {}
  return files;
}

const templatePath = resolve(ROOT, 'public/templates/master-template.xlsx');
const fixturePath = resolve(ROOT, 'src/fixtures');

// Analyze the master template
console.log('=== MASTER TEMPLATE ===');
console.log('Path:', templatePath);
const tplBuf = readFileSync(templatePath);
const tplWb = XLSX.read(tplBuf, { type: 'buffer', cellStyles: true });
console.log('Sheet Names:', JSON.stringify(tplWb.SheetNames));
console.log();

for (const sheetName of tplWb.SheetNames) {
  const ws = tplWb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1:A1');
  console.log(`--- Sheet: "${sheetName}" ---`);
  console.log(`  Range: ${ws['!ref']}`);
  console.log(`  Rows: ${range.e.r + 1}, Cols: ${range.e.c + 1}`);
  
  // Print first 12 rows to see header structure
  console.log('  First 12 rows:');
  for (let r = 0; r <= Math.min(11, range.e.r); r++) {
    const row = [];
    for (let c = 0; c <= Math.min(range.e.c, 49); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell?.v !== undefined && cell.v !== null && cell.v !== '') {
        row.push(`${XLSX.utils.encode_col(c)}${r+1}=${JSON.stringify(cell.v)}`);
      }
    }
    if (row.length > 0) console.log(`    Row ${r+1}: ${row.join(', ')}`);
  }
  
  // Find summary rows (look for "Jumlah")
  for (let r = range.e.r; r >= 0; r--) {
    for (let c = 0; c <= Math.min(range.e.c, 10); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell?.v && String(cell.v).includes('Jumlah')) {
        console.log(`  Summary found at row ${r+1}, col ${XLSX.utils.encode_col(c)}: "${cell.v}"`);
        // Print summary block (4 rows)
        for (let sr = r; sr <= Math.min(r + 5, range.e.r); sr++) {
          const srow = [];
          for (let sc = 0; sc <= Math.min(range.e.c, 49); sc++) {
            const saddr = XLSX.utils.encode_cell({ r: sr, c: sc });
            const scell = ws[saddr];
            if (scell?.v !== undefined && scell.v !== null && scell.v !== '') {
              srow.push(`${XLSX.utils.encode_col(sc)}${sr+1}=${JSON.stringify(scell.v)}`);
            }
          }
          if (srow.length > 0) console.log(`    Summary Row ${sr+1}: ${srow.join(', ')}`);
        }
        break;
      }
    }
  }
  
  // Print merges
  if (ws['!merges'] && ws['!merges'].length > 0) {
    console.log(`  Merges (${ws['!merges'].length}):`);
    for (const m of ws['!merges'].slice(0, 30)) {
      console.log(`    ${XLSX.utils.encode_range(m)}`);
    }
    if (ws['!merges'].length > 30) console.log(`    ... and ${ws['!merges'].length - 30} more`);
  }
  console.log();
}

// Analyze fixture files
console.log('\n=== FIXTURE FILES ===');
const fixtureFiles = findExcelFiles(fixturePath);
for (const filePath of fixtureFiles) {
  const buf = readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer' });
  console.log(`\n--- ${filePath.split('/').pop()} ---`);
  console.log(`  Sheets: ${JSON.stringify(wb.SheetNames)}`);
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    console.log(`  Sheet "${sn}": ${rows.length} rows`);
    if (rows.length > 0) {
      console.log(`    Header: ${JSON.stringify(rows[0])}`);
      if (rows.length > 1) {
        console.log(`    Row 2: ${JSON.stringify(rows[1])}`);
      }
      if (rows.length > 2) {
        console.log(`    Row 3: ${JSON.stringify(rows[2])}`);
      }
    }
  }
}

// Also analyze the "template-terisi" fixture
const terisiPath = resolve(fixturePath, 'template-terisi.xlsx');
console.log('\n=== TEMPLATE TERISI (FILLED EXAMPLE) ===');
const terisiWb = XLSX.read(readFileSync(terisiPath), { type: 'buffer' });
console.log('Sheet Names:', JSON.stringify(terisiWb.SheetNames));

for (const sheetName of terisiWb.SheetNames) {
  const ws = terisiWb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1:A1');
  console.log(`\n--- Sheet: "${sheetName}" ---`);
  console.log(`  Range: ${ws['!ref']}`);
  console.log(`  Rows: ${range.e.r + 1}, Cols: ${range.e.c + 1}`);
  
  // Print first 8 rows
  console.log('  First 8 rows:');
  for (let r = 0; r <= Math.min(7, range.e.r); r++) {
    const row = [];
    for (let c = 0; c <= Math.min(range.e.c, 49); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell?.v !== undefined && cell.v !== null && cell.v !== '') {
        row.push(`${XLSX.utils.encode_col(c)}${r+1}=${JSON.stringify(cell.v)}`);
      }
    }
    if (row.length > 0) console.log(`    Row ${r+1}: ${row.join(', ')}`);
  }
  
  // Print data rows (count children)
  let dataRows = 0;
  for (let r = 6; r <= range.e.r; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 0 });
    const cell = ws[addr];
    if (cell?.v && typeof cell.v === 'number' && cell.v > 0) {
      dataRows++;
    }
  }
  console.log(`  Data rows (numbered): ${dataRows}`);
  
  // Find summary
  for (let r = range.e.r; r >= 0; r--) {
    for (let c = 0; c <= Math.min(range.e.c, 10); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell?.v && String(cell.v).includes('Jumlah')) {
        console.log(`  Summary at row ${r+1}: "${cell.v}"`);
        for (let sr = r; sr <= Math.min(r + 5, range.e.r); sr++) {
          const srow = [];
          for (let sc = 0; sc <= Math.min(range.e.c, 49); sc++) {
            const saddr = XLSX.utils.encode_cell({ r: sr, c: sc });
            const scell = ws[saddr];
            if (scell?.v !== undefined && scell.v !== null && scell.v !== '') {
              srow.push(`${XLSX.utils.encode_col(sc)}${sr+1}=${JSON.stringify(scell.v)}`);
            }
          }
          if (srow.length > 0) console.log(`    Summary Row ${sr+1}: ${srow.join(', ')}`);
        }
        break;
      }
    }
  }
}
