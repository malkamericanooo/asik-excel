import { describe, it, expect, beforeAll } from 'vitest';
import * as XLSX from 'xlsx';
import { readFileSync } from 'fs';
import { resolve } from 'path';
import { parseAndMergeAsikFile, createEmptyMasterData } from '../asikParser';
import { buildMasterExcel, findSummaryStartRow } from '../masterExcel';
import type { MasterData, ProcessResult } from '../../types';

const FIXTURES = resolve(__dirname, '../../fixtures');
const TEMPLATE_PATH = resolve(__dirname, '../../../public/templates/master-template.xlsx');

let templateBuffer: ArrayBuffer;
let contohDpt1Buffer: ArrayBuffer;

beforeAll(() => {
  const tpl = readFileSync(TEMPLATE_PATH);
  templateBuffer = tpl.buffer.slice(tpl.byteOffset, tpl.byteOffset + tpl.byteLength);

  const dpt1 = readFileSync(resolve(FIXTURES, 'contoh-dpt1.xlsx'));
  contohDpt1Buffer = dpt1.buffer.slice(dpt1.byteOffset, dpt1.byteOffset + dpt1.byteLength);
});

function emptyResult(): ProcessResult {
  return { added: 0, updated: 0, moved: 0, skipped: 0, logs: [] };
}

describe('integration: contoh-dpt1 → master excel', () => {
  it('classifies kelurahan correctly from ASIK export', () => {
    const masterData = createEmptyMasterData();
    const result = emptyResult();
    parseAndMergeAsikFile(contohDpt1Buffer, masterData, result);

    expect(result.added).toBe(10);
    expect(masterData.MABUUN).toHaveLength(5);
    expect(masterData.PEMBATAAN).toHaveLength(2);
    expect(masterData.MABURAI).toHaveLength(1);
    expect(masterData.Kejar).toHaveLength(1);
    expect(masterData['LUAR WILAYAH']).toHaveLength(1);

    const luar = masterData['LUAR WILAYAH'][0];
    expect(luar.nama).toBe('Arshaka Zayendra Ahmad');
    expect(luar.alamat).toBe('TANTA HULU');

    const kejar = masterData.Kejar[0];
    expect(kejar.nama).toBe('Jiana Arabella');
    expect(kejar.alamat).toBe('Kejar - PEMBATAAN');
  });

  it('routes Belimbing to LUAR WILAYAH not MABUUN', () => {
    const masterData = createEmptyMasterData();
    const result = emptyResult();
    const buf = makeSimpleBuffer([
      { nama: 'Anak Belimbing', kelurahan: 'Belimbing Raya Rt.14', status: 'ideal' },
      { nama: 'Anak Mabuun', kelurahan: 'Mabuun', status: 'ideal' },
    ]);
    parseAndMergeAsikFile(buf, masterData, result, { selectedVaccine: 'DPT_1' });

    expect(masterData['LUAR WILAYAH']).toHaveLength(1);
    expect(masterData['LUAR WILAYAH'][0].nama).toBe('Anak Belimbing');
    expect(masterData.MABUUN).toHaveLength(1);
    expect(masterData.MABUUN[0].nama).toBe('Anak Mabuun');
  });

  it('builds master excel matching template structure', async () => {
    const masterData = createEmptyMasterData();
    parseAndMergeAsikFile(contohDpt1Buffer, masterData, emptyResult());

    const outBuf = await buildMasterExcel(masterData, 6, 2026, templateBuffer).arrayBuffer();
    const wb = XLSX.read(outBuf, { type: 'array' });

    expect(wb.SheetNames).toEqual([
      'MABUUN', 'KASIAU', 'PEMBATAAN', 'SULINGAN', 'MABURAI', 'LUAR WILAYAH', 'Kejar',
    ]);

    const mabuun = wb.Sheets['MABUUN'];
    expect(mabuun['A5']?.v).toBe('NO');
    expect(mabuun['H5']?.v).toBe('HB0 (<24 JAM)');
    expect(mabuun['P5']?.v).toBe('DPT/HB-Hib (1)');
    expect(mabuun['A7']?.v).toBe(1);
    expect(mabuun['B7']?.v).toBeTruthy();

    const luar = wb.Sheets['LUAR WILAYAH'];
    expect(luar['B7']?.v).toBe('Arshaka Zayendra Ahmad');
    expect(luar['G7']?.v).toBe('TANTA HULU');
    expect(luar['P7']?.v).toBeGreaterThan(0);

    const summaryStart = findSummaryStartRow(mabuun);
    expect(mabuun[XLSX.utils.encode_cell({ r: summaryStart, c: 6 })]?.v).toBe('Jumlah');
    const countRow = summaryStart + 2;
    expect(mabuun[XLSX.utils.encode_cell({ r: countRow, c: 15 })]?.v).toBe(5);
  });
});

function makeSimpleBuffer(
  rows: { nama: string; kelurahan: string; status: string }[],
): ArrayBuffer {
  const header = [
    'Nama Anak', 'Tanggal Lahir', 'Nama Orang Tua',
    'Klasifikasi Kelurahan', 'Status Imunisasi', 'Tanggal Imunisasi',
  ];
  const aoa = [
    header,
    ...rows.map((r) => [
      r.nama, '2026-04-01', 'Ortu Test', r.kelurahan, r.status, '2026-06-17',
    ]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
}
