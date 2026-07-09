import * as XLSX from 'xlsx';
import { readFileSync, writeFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { execSync } from 'child_process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const templatePath = resolve(root, 'public/templates/master-template.xlsx');
const outDir = resolve(root, 'tmp-diagnose');

function loadTemplate() {
  const buf = readFileSync(templatePath);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

function sampleMasterData() {
  return {
    MABUUN: [{
      nama: 'Test Anak', jk: 'L', tanggalLahirSerial: 45748, tanggalLahirStr: '2025-04-01',
      nik: '123', namaOrangTua: 'Ortu', alamat: 'MABUUN', vaccines: { DPT_1: 45825 },
    }],
    KASIAU: [], PEMBATAAN: [], SULINGAN: [], MABURAI: [], 'LUAR WILAYAH': [], Kejar: [],
  };
}

// Minimal copy of build logic variants for diagnosis
function buildVariant(templateBuffer, { cellStylesRead, cellStylesWrite, stripStylesOnWrite }) {
  const wb = XLSX.read(templateBuffer, { type: 'array', cellStyles: cellStylesRead });
  const ws = wb.Sheets['MABUUN'];
  if (ws['C3']) ws['C3'].v = ': Juni 2026';
  ws['A7'] = { v: 1, t: 'n' };
  ws['B7'] = { v: 'Test Anak', t: 's' };

  if (stripStylesOnWrite) {
    for (const sheetName of wb.SheetNames) {
      const sheet = wb.Sheets[sheetName];
      for (const addr of Object.keys(sheet)) {
        if (addr.startsWith('!')) continue;
        const cell = sheet[addr];
        if (cell?.s) delete cell.s;
        if (cell?.z) delete cell.z;
      }
    }
  }

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: cellStylesWrite });
}

function inspectXlsx(path, label) {
  console.log(`\n=== ${label} ===`);
  console.log('size:', readFileSync(path).length);
  try {
    const wb = XLSX.read(readFileSync(path), { type: 'buffer', cellStyles: true });
    const ws = wb.Sheets['MABUUN'];
    console.log('sheets:', wb.SheetNames.length);
    console.log('MABUUN ref:', ws['!ref']);
    console.log('A7:', ws['A7']?.v, 'B7:', ws['B7']?.v);
    console.log('merges:', ws['!merges']?.length ?? 0);
  } catch (e) {
    console.log('xlsx read error:', e.message);
  }

  try {
    const xmlCheck = execSync(
      `unzip -p "${path}" xl/worksheets/sheet1.xml 2>/dev/null | head -c 500`,
      { encoding: 'utf8' },
    );
    console.log('sheet1.xml head:', xmlCheck.slice(0, 200).replace(/\s+/g, ' '));
  } catch {
    console.log('unzip failed');
  }

  try {
    execSync(`unzip -t "${path}"`, { stdio: 'pipe' });
    console.log('zip integrity: OK');
  } catch (e) {
    console.log('zip integrity: FAIL', e.stderr?.toString()?.slice(0, 200));
  }
}

import { mkdirSync } from 'fs';

mkdirSync(outDir, { recursive: true });
const template = loadTemplate();

const variants = [
  ['template-original', () => readFileSync(templatePath)],
  ['read-styles_write-styles', () => buildVariant(template, { cellStylesRead: true, cellStylesWrite: true, stripStylesOnWrite: false })],
  ['read-styles_write-no-styles', () => buildVariant(template, { cellStylesRead: true, cellStylesWrite: false, stripStylesOnWrite: false })],
  ['read-no-styles_write-no-styles', () => buildVariant(template, { cellStylesRead: false, cellStylesWrite: false, stripStylesOnWrite: false })],
  ['read-styles_strip_write-no-styles', () => buildVariant(template, { cellStylesRead: true, cellStylesWrite: false, stripStylesOnWrite: true })],
];

for (const [name, fn] of variants) {
  const outPath = resolve(outDir, `${name}.xlsx`);
  writeFileSync(outPath, fn());
  inspectXlsx(outPath, name);
}

// Full build via dynamic import of compiled test path - use vitest inline instead
console.log('\n=== Running integration export ===');
