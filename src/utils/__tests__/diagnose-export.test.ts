import { describe, it } from 'vitest';
import * as XLSX from 'xlsx';
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { resolve } from 'path';
import { execSync } from 'child_process';
import { buildMasterExcel } from '../masterExcel';
import { createEmptyMasterData, parseAndMergeAsikFile } from '../asikParser';
import { dateStringToExcelSerial } from '../dateUtils';

const TEMPLATE = resolve(__dirname, '../../../public/templates/master-template.xlsx');
const OUT = resolve(__dirname, '../../../tmp-diagnose');

function inspect(path: string, label: string) {
  const info: string[] = [`\n=== ${label} ===`, `size: ${readFileSync(path).length}`];
  try {
    execSync(`unzip -t "${path}"`, { stdio: 'pipe' });
    info.push('zip: OK');
  } catch {
    info.push('zip: FAIL');
  }
  const wb = XLSX.read(readFileSync(path), { type: 'buffer' });
  const ws = wb.Sheets['MABUUN'];
  info.push(`ref: ${ws['!ref']}`, `A7: ${ws['A7']?.v}`, `merges: ${ws['!merges']?.length ?? 0}`);
  // Check styles.xml size
  try {
    const styles = execSync(`unzip -p "${path}" xl/styles.xml 2>/dev/null | wc -c`, { encoding: 'utf8' }).trim();
    info.push(`styles.xml bytes: ${styles}`);
  } catch { /* ignore */ }
  console.log(info.join('\n'));
}

describe('diagnose excel corruption', () => {
  it('compare export variants', async () => {
    mkdirSync(OUT, { recursive: true });
    const templateBuf = readFileSync(TEMPLATE);
    const templateBuffer = templateBuf.buffer.slice(
      templateBuf.byteOffset,
      templateBuf.byteOffset + templateBuf.byteLength,
    );

    const master = createEmptyMasterData();
    master.MABUUN.push({
      nama: 'Test Anak', jk: 'L',
      tanggalLahirSerial: dateStringToExcelSerial('2025-04-01'),
      tanggalLahirStr: '2025-04-01', nik: '123', namaOrangTua: 'Ortu',
      alamat: 'MABUUN', vaccines: { DPT_1: dateStringToExcelSerial('2026-06-17') },
    });

    // Current production path
    const prodBlob = buildMasterExcel(master, 6, 2026, templateBuffer);
    const prodPath = resolve(OUT, 'production-export.xlsx');
    writeFileSync(prodPath, Buffer.from(await prodBlob.arrayBuffer()));

    // Variant: strip styles before write
    const wb = XLSX.read(templateBuffer, { type: 'array', cellStyles: true });
    for (const sn of wb.SheetNames) {
      const ws = wb.Sheets[sn];
      ws['C3'] && (ws['C3'].v = ': Juni 2026');
      for (const addr of Object.keys(ws)) {
        if (addr.startsWith('!')) continue;
        const cell = ws[addr];
        if (cell?.s) delete cell.s;
      }
    }
    const noStylePath = resolve(OUT, 'no-styles-export.xlsx');
    writeFileSync(noStylePath, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: false }));

    // Variant: read without styles, write without styles - full rebuild
    const wb2 = XLSX.read(templateBuffer, { type: 'array', cellStyles: false });
    const noStyleReadPath = resolve(OUT, 'read-no-styles.xlsx');
    writeFileSync(noStyleReadPath, XLSX.write(wb2, { type: 'buffer', bookType: 'xlsx', cellStyles: false }));

    inspect(TEMPLATE, 'template original');
    inspect(prodPath, 'production buildMasterExcel');
    inspect(noStylePath, 'strip styles before write');
    inspect(noStyleReadPath, 'read/write no styles');

    // Check for invalid style indices in production file
    try {
      const stylesXml = execSync(`unzip -p "${prodPath}" xl/styles.xml`, { encoding: 'utf8' });
      const cellXfs = (stylesXml.match(/<cellXfs count="(\d+)"/) ?? [])[1];
      console.log(`\nproduction cellXfs count: ${cellXfs}`);
      const sheet1 = execSync(`unzip -p "${prodPath}" xl/worksheets/sheet1.xml`, { encoding: 'utf8' });
      const styleRefs = [...sheet1.matchAll(/ s="(\d+)"/g)].map((m) => Number(m[1]));
      const maxStyle = Math.max(...styleRefs, -1);
      console.log(`sheet1 max style index: ${maxStyle}, cellXfs: ${cellXfs}`);
      if (cellXfs && maxStyle >= Number(cellXfs)) {
        console.log('*** CORRUPTION: style index out of range! ***');
      }
    } catch (e) {
      console.log('style check error:', e);
    }
  });
});
